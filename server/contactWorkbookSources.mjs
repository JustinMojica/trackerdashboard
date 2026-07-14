const graphRoot = "https://graph.microsoft.com/v1.0";
const instructionSheetSampleRange = "A1:Z40";
const graphRequestTimeoutMs = 12000;
const worksheetReadConcurrency = 8;

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
  const results = await Promise.all(
    sources.map(async (source) => {
      try {
        return await readWorkbookSource({ source, token, fetchImpl });
      } catch (error) {
        return {
          source: {
            id: source.id,
            label: source.label,
            status: "error",
            workbookName: "",
            worksheetCount: 0,
            rowCount: 0,
            error: error instanceof Error ? error.message : "Workbook read failed.",
          },
          contacts: [],
          warnings: [],
        };
      }
    }),
  );
  const sourceResults = results.map((result) => result.source);
  const contacts = results.flatMap((result) => result.contacts);
  const warnings = results.flatMap((result) => result.warnings);

  for (const source of sources) {
    if (!sourceResults.some((result) => result.id === source.id)) {
      sourceResults.push({
        id: source.id,
        label: source.label,
        status: "error",
        workbookName: "",
        worksheetCount: 0,
        rowCount: 0,
        error: "Workbook read did not return a result.",
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

  const worksheetResults = await mapWithConcurrency(
    worksheetList,
    worksheetReadConcurrency,
    async (worksheet) => {
    const worksheetId = worksheet.id || worksheet.name;
      if (!worksheetId) return null;
      if (isNonContactWorksheet(worksheet.name || worksheetId)) return null;
    try {
      const range = await graphGet(
        `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(
          itemId,
        )}/workbook/worksheets/${encodeURIComponent(
          worksheetId,
        )}/range(address='${instructionSheetSampleRange}')`,
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
        return parsed;
    } catch (error) {
        return {
          contacts: [],
          warnings: [
            {
              sourceId: source.id,
              worksheetName: worksheet.name || worksheetId,
              message:
                error instanceof Error
                  ? error.message
                  : "Worksheet could not be read.",
            },
          ],
        };
    }
    },
  );
  for (const result of worksheetResults.filter(Boolean)) {
    worksheetRows.push(...result.contacts);
    warnings.push(...result.warnings);
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
  if (isPlaceholderInstructionSheet(worksheetName, nonEmptyRows)) {
    return { contacts: [], warnings: [] };
  }
  if (isClientInstructionSheet(nonEmptyRows)) {
    return parseClientInstructionSheet({
      source,
      workbookName,
      worksheetName,
      rows: nonEmptyRows,
    });
  }
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
    const email = uniqueEmails(extractEmails(pickField(raw, "email"))).join("; ");
    const company =
      pickField(raw, "company") ||
      pickField(raw, "coverholder") ||
      pickField(raw, "managingAgent") ||
      pickField(raw, "broker");
    const contactName = pickField(raw, "contactName");
    const instructions = instructionValues(raw);
    if (!email && !company && !contactName && instructions.length === 0) continue;
    if (email && uniqueEmails(extractEmails(email)).length === 0) {
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
      emails: {
        dca: [],
        coverholder: [],
        report: [],
        invoice: [],
        all: uniqueEmails(extractEmails(email)),
      },
      phone: pickField(raw, "phone"),
      role: pickField(raw, "role"),
      specialInstructions: instructions,
      raw,
    });
  }

  return { contacts, warnings };
}

function parseClientInstructionSheet({
  source,
  workbookName,
  worksheetName,
  rows,
}) {
  const primaryHeaders = rows[0] ?? [];
  const primaryValues = rows[1] ?? [];
  const secondaryHeaderIndex = rows.findIndex((row) =>
    row.some((cell) => String(cell ?? "").toLowerCase().includes("report submission email")),
  );
  const secondaryHeaders = secondaryHeaderIndex >= 0 ? rows[secondaryHeaderIndex] : [];
  const secondaryValues =
    secondaryHeaderIndex >= 0 ? rows[secondaryHeaderIndex + 1] ?? [] : [];
  const clientBlock = cleanCell(primaryValues[0]);
  const dcaContactBlock = cleanCell(primaryValues[1]);
  const coverholderContactBlock = cleanCell(primaryValues[2]);
  const company = firstLine(clientBlock) || worksheetName;
  const reportSubmission = fieldValueFromRows(rows, [/report submission email/i]) || cleanCell(secondaryValues[0]);
  const invoiceSubmission = fieldValueFromRows(rows, [/invoice submission email/i]) || cleanCell(secondaryValues[1]);
  const onsiteRemote = fieldValueFromRows(rows, [/onsite\/remote preference/i, /onsite.*remote/i]) || cleanCell(secondaryValues[2]);
  const paymentTerms = fieldValueFromRows(rows, [/fees and payment terms/i, /payment terms/i]) || cleanCell(secondaryValues[3]);
  const other = fieldValueFromRows(rows, [/^other$/i]) || cleanCell(secondaryValues[4]);
  const notes = fieldValueFromRows(rows, [/notes\/comments/i, /notes/i, /comments/i]) || cleanCell(secondaryValues[5]);
  const dcaEmails = uniqueEmails(extractEmails(dcaContactBlock));
  const coverholderEmails = uniqueEmails(extractEmails(coverholderContactBlock));
  const reportEmails = uniqueEmails(extractEmails(reportSubmission));
  const invoiceEmails = uniqueEmails(extractEmails(invoiceSubmission));
  const sheetEmails = uniqueEmails(extractEmails(rows.flat().map(cleanCell).join("\n")));
  const allEmails = uniqueEmails([
    ...dcaEmails,
    ...coverholderEmails,
    ...reportEmails,
    ...invoiceEmails,
    ...sheetEmails,
  ]);
  const email = allEmails.join("; ");
  const instructions = [];
  addInstruction(instructions, primaryHeaders[3], primaryValues[3]);
  addInstruction(instructions, primaryHeaders[4], primaryValues[4]);
  addInstruction(instructions, primaryHeaders[5], primaryValues[5]);
  addInstruction(instructions, secondaryHeaders[0] || "Report Submission Email", reportSubmission);
  addInstruction(instructions, secondaryHeaders[1] || "Invoice Submission Email", invoiceSubmission);
  addInstruction(instructions, secondaryHeaders[2] || "Onsite/Remote Preference", onsiteRemote);
  addInstruction(instructions, secondaryHeaders[3] || "Fees and Payment Terms", paymentTerms);
  addInstruction(instructions, secondaryHeaders[4] || "Other", other);
  addInstruction(instructions, secondaryHeaders[5] || "Notes/Comments", notes);

  return {
    contacts: [
      {
        id: contactId(source.id, worksheetName, email, company, worksheetName),
        sourceId: source.id,
        sourceLabel: source.label,
        workbookName,
        worksheetName,
        company,
        coverholder: contactNameFromBlock(coverholderContactBlock),
        managingAgent: company,
        broker: "",
        contactName: contactNameFromBlock(dcaContactBlock) || company,
        email,
        emails: {
          dca: dcaEmails,
          coverholder: coverholderEmails,
          report: reportEmails,
          invoice: invoiceEmails,
          all: allEmails,
        },
        phone: firstPhone([clientBlock, dcaContactBlock, coverholderContactBlock]),
        role: "Client instruction sheet",
        specialInstructions: instructions,
        raw: {
          client: clientBlock,
          dcaContact: dcaContactBlock,
          coverholderContact: coverholderContactBlock,
          reportSubmission,
          invoiceSubmission,
          onsiteRemote,
          paymentTerms,
          other,
          notes,
        },
      },
    ],
    warnings: email
      ? []
      : [
          {
            sourceId: source.id,
            worksheetName,
            message: "Instruction sheet loaded but no email address was detected.",
          },
        ],
  };
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

function isClientInstructionSheet(rows) {
  return String(rows[0]?.[0] ?? "")
    .toLowerCase()
    .includes("client name/address");
}

function isPlaceholderInstructionSheet(worksheetName, rows) {
  const normalizedName = String(worksheetName || "").trim().toLowerCase();
  if (/^blank\s*\d*$/i.test(normalizedName)) return true;
  const text = rows.flat().map(cleanCell).join(" ").toLowerCase();
  if (!text) return true;
  const hasInstructionHeaders =
    text.includes("client name/address") ||
    text.includes("dca primary contact") ||
    text.includes("cover holder primary contact");
  const hasUsableValue = extractEmails(text).length > 0 || rows.length > 2;
  return hasInstructionHeaders && !hasUsableValue;
}

function isNonContactWorksheet(name) {
  const normalized = String(name).toLowerCase();
  return (
    normalized.includes("london calls") ||
    normalized.includes("instructions") ||
    normalized.includes("template")
  );
}

function cleanCell(value) {
  return String(value ?? "")
    .replace(/\r\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function fieldValueFromRows(rows, labelPatterns) {
  for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex] ?? [];
    for (let columnIndex = 0; columnIndex < row.length; columnIndex += 1) {
      const label = cleanCell(row[columnIndex]);
      if (!labelPatterns.some((pattern) => pattern.test(label))) continue;
      const rightValue = cleanCell(row[columnIndex + 1]);
      const rightLooksLikeHeader =
        rightValue &&
        (labelPatterns.some((pattern) => pattern.test(rightValue)) ||
          (/email|preference|terms|other|notes|comments/i.test(rightValue) &&
            extractEmails(rightValue).length === 0));
      if (rightValue && !rightLooksLikeHeader) return rightValue;
      const belowValue = cleanCell(rows[rowIndex + 1]?.[columnIndex]);
      if (belowValue && !labelPatterns.some((pattern) => pattern.test(belowValue))) {
        return belowValue;
      }
    }
  }
  return "";
}

function firstLine(value) {
  return cleanCell(value)
    .split(/\n/)
    .map((line) => line.trim())
    .find(Boolean) || "";
}

function addInstruction(instructions, label, value) {
  const cleanedValue = cleanCell(value);
  const cleanedLabel = cleanCell(label);
  if (!cleanedLabel || !cleanedValue) return;
  if (/^none( specified| received| provided)?$/i.test(cleanedValue)) return;
  instructions.push({
    label: cleanedLabel,
    value: cleanedValue,
  });
}

function extractEmails(value) {
  return cleanCell(value).match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi) ?? [];
}

function uniqueEmails(values) {
  const seen = new Set();
  return values
    .map((value) => value.trim().replace(/[.;,]+$/, ""))
    .filter((value) => {
      const key = value.toLowerCase();
      if (!key || seen.has(key)) return false;
      seen.add(key);
      return true;
    });
}

function contactNameFromBlock(value) {
  const lines = cleanCell(value)
    .split(/\n/)
    .map((line) =>
      line
        .replace(/^email:\s*/i, "")
        .replace(/^e:\s*/i, "")
        .replace(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi, "")
        .replace(/\b(email|e-mail|e):?\s*/gi, "")
        .replace(/[;|,]+$/g, "")
        .trim(),
    )
    .filter(Boolean)
    .filter((line) => !isLikelyEmail(line) && !/^tel|^d:|^m:/i.test(line));
  return lines[0] ?? "";
}

function firstPhone(values) {
  const match = values
    .map(cleanCell)
    .join("\n")
    .match(/(?:tel(?:ephone)?|direct|d|m)?[:\s]*(\+?\d[\d\s().-]{6,}\d)/i);
  return match?.[1]?.trim() ?? "";
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

async function mapWithConcurrency(items, concurrency, mapper) {
  const results = new Array(items.length);
  let nextIndex = 0;
  const workers = Array.from(
    { length: Math.min(concurrency, items.length) },
    async () => {
      while (nextIndex < items.length) {
        const currentIndex = nextIndex;
        nextIndex += 1;
        results[currentIndex] = await mapper(items[currentIndex], currentIndex);
      }
    },
  );
  await Promise.all(workers);
  return results;
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
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), graphRequestTimeoutMs);
  let response;
  try {
    response = await fetchImpl(`${graphRoot}${path}`, {
      headers: { authorization: `Bearer ${token}` },
      signal: controller.signal,
    });
  } finally {
    clearTimeout(timeout);
  }
  if (!response.ok) {
    throw new Error(`Graph ${response.status}: ${await response.text()}`);
  }
  return response.json();
}
