const graphRoot = "https://graph.microsoft.com/v1.0";

const fieldMatchers = [
  { key: "email", patterns: [/e-?mail/i, /email address/i] },
  { key: "contactName", patterns: [/contact.*name/i, /^name$/i, /person/i] },
  { key: "company", patterns: [/company/i, /client/i, /insured/i] },
  { key: "coverholder", patterns: [/coverholder/i] },
  { key: "managingAgent", patterns: [/managing.*agent/i, /\bma\b/i] },
  { key: "broker", patterns: [/broker/i] },
  { key: "phone", patterns: [/phone/i, /telephone/i, /mobile/i] },
  { key: "role", patterns: [/role/i, /title/i, /position/i, /purpose/i] },
];

const instructionPatterns = [
  /instruction/i,
  /special/i,
  /note/i,
  /remark/i,
  /requirement/i,
  /guidance/i,
  /preference/i,
];

export function parseWorkbookLinks(rawValue = "") {
  return rawValue
    .split(/\r?\n|,/)
    .map((value) => value.trim())
    .filter(Boolean)
    .map((url, index) => ({
      id: `workbook-${index + 1}`,
      label: `Workbook ${index + 1}`,
      url,
    }));
}

export function contactWorkbookStatus(config) {
  const sources = parseWorkbookLinks(config.contactWorkbookLinksRaw);
  return {
    configured: sources.length > 0,
    count: sources.length,
    missing: sources.length > 0 ? [] : ["TRACKER_CONTACT_WORKBOOK_LINKS"],
  };
}

export async function readLinkedContactWorkbooks({
  sources,
  getAccessToken,
  fetchImpl = fetch,
}) {
  const token = await getAccessToken();
  const sourceResults = [];
  const contacts = [];
  const warnings = [];

  for (const source of sources) {
    try {
      const result = await readWorkbookSource({ source, token, fetchImpl });
      sourceResults.push(result.source);
      contacts.push(...result.contacts);
      warnings.push(...result.warnings);
    } catch (error) {
      sourceResults.push({
        id: source.id,
        label: source.label,
        status: "error",
        workbookName: "",
        worksheetCount: 0,
        rowCount: 0,
        error: error instanceof Error ? error.message : "Workbook read failed.",
      });
    }
  }

  return {
    generatedAt: new Date().toISOString(),
    sources: sourceResults,
    contacts: dedupeContacts(contacts),
    warnings,
  };
}

async function readWorkbookSource({ source, token, fetchImpl }) {
  const driveItem = await graphGet(
    `/shares/${shareIdFromUrl(source.url)}/driveItem`,
    token,
    fetchImpl,
  );
  const driveId = driveItem?.parentReference?.driveId;
  const itemId = driveItem?.id;
  if (!driveId || !itemId) {
    throw new Error("Graph did not return a drive ID and item ID for this workbook link.");
  }

  const worksheets = await graphGet(
    `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(
      itemId,
    )}/workbook/worksheets`,
    token,
    fetchImpl,
  );
  const worksheetRows = [];
  const warnings = [];
  const worksheetList = Array.isArray(worksheets?.value) ? worksheets.value : [];

  for (const worksheet of worksheetList) {
    const worksheetId = worksheet.id || worksheet.name;
    if (!worksheetId) continue;
    try {
      const range = await graphGet(
        `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(
          itemId,
        )}/workbook/worksheets/${encodeURIComponent(
          worksheetId,
        )}/usedRange(valuesOnly=true)`,
        token,
        fetchImpl,
      );
      const values = Array.isArray(range?.values) ? range.values : [];
      const parsed = parseWorksheetRows({
        source,
        workbookName: driveItem.name || source.label,
        worksheetName: worksheet.name || worksheetId,
        values,
      });
      worksheetRows.push(...parsed.contacts);
      warnings.push(...parsed.warnings);
    } catch (error) {
      warnings.push({
        sourceId: source.id,
        worksheetName: worksheet.name || worksheetId,
        message:
          error instanceof Error
            ? error.message
            : "Worksheet could not be read.",
      });
    }
  }

  return {
    source: {
      id: source.id,
      label: source.label,
      status: "ok",
      workbookName: driveItem.name || source.label,
      worksheetCount: worksheetList.length,
      rowCount: worksheetRows.length,
    },
    contacts: worksheetRows,
    warnings,
  };
}

export function parseWorksheetRows({
  source,
  workbookName,
  worksheetName,
  values,
}) {
  const nonEmptyRows = values.filter((row) =>
    Array.isArray(row) && row.some((cell) => String(cell ?? "").trim()),
  );
  if (nonEmptyRows.length < 2) {
    return {
      contacts: [],
      warnings: [
        {
          sourceId: source.id,
          worksheetName,
          message: "Worksheet has no header row plus data rows.",
        },
      ],
    };
  }

  const headers = nonEmptyRows[0].map((cell) => String(cell ?? "").trim());
  const contacts = [];
  const warnings = [];
  for (const row of nonEmptyRows.slice(1)) {
    const raw = rowToObject(headers, row);
    const email = pickField(raw, "email");
    const company =
      pickField(raw, "company") ||
      pickField(raw, "coverholder") ||
      pickField(raw, "managingAgent") ||
      pickField(raw, "broker");
    const contactName = pickField(raw, "contactName");
    const instructions = instructionValues(raw);
    if (!email && !company && !contactName && instructions.length === 0) continue;
    if (email && !isLikelyEmail(email)) {
      warnings.push({
        sourceId: source.id,
        worksheetName,
        message: `Possible invalid email: ${email}`,
      });
    }
    contacts.push({
      id: contactId(source.id, worksheetName, email, company, contactName),
      sourceId: source.id,
      sourceLabel: source.label,
      workbookName,
      worksheetName,
      company,
      coverholder: pickField(raw, "coverholder"),
      managingAgent: pickField(raw, "managingAgent"),
      broker: pickField(raw, "broker"),
      contactName,
      email,
      phone: pickField(raw, "phone"),
      role: pickField(raw, "role"),
      specialInstructions: instructions,
      raw,
    });
  }

  return { contacts, warnings };
}

function rowToObject(headers, row) {
  const raw = {};
  headers.forEach((header, index) => {
    if (!header) return;
    raw[header] = String(row[index] ?? "").trim();
  });
  return raw;
}

function pickField(raw, key) {
  const matcher = fieldMatchers.find((item) => item.key === key);
  if (!matcher) return "";
  const header = Object.keys(raw).find((candidate) =>
    matcher.patterns.some((pattern) => pattern.test(candidate)),
  );
  return header ? raw[header] : "";
}

function instructionValues(raw) {
  return Object.entries(raw)
    .filter(([header, value]) =>
      instructionPatterns.some((pattern) => pattern.test(header)) &&
      String(value ?? "").trim(),
    )
    .map(([header, value]) => ({
      label: header,
      value: String(value).trim(),
    }));
}

function dedupeContacts(contacts) {
  const seen = new Set();
  return contacts.filter((contact) => {
    const key =
      contact.email?.toLowerCase() ||
      `${contact.company}|${contact.contactName}|${contact.worksheetName}`.toLowerCase();
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function contactId(sourceId, worksheetName, email, company, contactName) {
  return [
    sourceId,
    worksheetName,
    email || company || contactName || "contact",
  ]
    .join(":")
    .toLowerCase()
    .replace(/[^a-z0-9:.-]+/g, "-")
    .slice(0, 140);
}

function isLikelyEmail(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value);
}

function shareIdFromUrl(url) {
  const encoded = Buffer.from(url).toString("base64url");
  return `u!${encoded}`;
}

async function graphGet(path, token, fetchImpl) {
  const response = await fetchImpl(`${graphRoot}${path}`, {
    headers: { authorization: `Bearer ${token}` },
  });
  if (!response.ok) {
    throw new Error(`Graph ${response.status}: ${await response.text()}`);
  }
  return response.json();
}
