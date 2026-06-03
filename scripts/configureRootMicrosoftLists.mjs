import { existsSync, readFileSync, writeFileSync } from "node:fs";
import { resolve } from "node:path";
import { userToTrackerFields } from "../server/accessUserStore.mjs";
import { projectToTrackerFields } from "../server/projectStore.mjs";

const graphBaseUrl = "https://graph.microsoft.com/v1.0";
const outputPath = resolve("deploy-artifacts", "microsoft-lists-root-site.json");
const siteUrlArg = process.argv.find((arg) => arg.startsWith("--site-url="));
const targetSiteUrl = siteUrlArg?.slice("--site-url=".length) || "";

const stageChoices = [
  "Intake",
  "Registration",
  "Quote",
  "Scheduling",
  "Pre-Audit",
  "Audit Fieldwork",
  "Reporting",
  "Invoicing",
  "Closed",
];
const assignmentStatusChoices = ["New", "In Progress", "Blocked", "On Hold", "Completed"];
const quoteStatusChoices = ["Not Started", "Drafting", "Sent", "Accepted", "Rejected"];
const progressChoices = ["Not Started", "In Progress", "Complete", "Not Required"];
const reportStatusChoices = ["Not Started", "Drafting", "Review", "Issued"];
const invoiceStatusChoices = ["Not Started", "Prepared", "Sent", "Paid"];

const schemas = [
  {
    key: "auditAssignments",
    displayName: "Audit Assignments",
    columns: [
      text("TrackerAssignmentId", "Tracker assignment ID", true, true),
      text("AssignmentNumber", "Assignment number", true, true),
      choice("AssignmentSource", "Assignment source", ["Email", "DAM"], true),
      choice("AssignmentType", "Assignment type", ["DCA", "CH", "MGA", "Company Contract"], true),
      text("AuditEntity", "Audit entity", true, true),
      text("ClientCoverholderCode", "Client / coverholder code", false, true),
      text("Broker", "Broker"),
      text("LeadAuditor", "Lead auditor"),
      choice("CurrentStage", "Current stage", stageChoices, true),
      choice("AssignmentStatus", "Assignment status", assignmentStatusChoices, true),
      choice("QuoteStatus", "Quote status", quoteStatusChoices),
      currency("QuoteAmount", "Quote amount"),
      text("TentativeAuditWeek", "Tentative audit week"),
      dateColumn("ConfirmedAuditDate", "Confirmed audit date"),
      choice("AuditType", "Audit type", ["Remote", "Onsite"]),
      booleanColumn("BaaReceived", "BAA received"),
      booleanColumn("EndorsementsReceived", "Endorsements received"),
      booleanColumn("PremiumBdxReceived", "Premium BDX received"),
      choice("PreAuditQuestionnaireStatus", "Pre-audit questionnaire status", progressChoices),
      choice("DocumentRequestStatus", "Document request status", progressChoices),
      dateColumn("DocumentRequestDate", "Document request date"),
      dateColumn("BrokerLastChasedDate", "Broker last chased date"),
      dateColumn("BrokerExpectedResponseDate", "Broker expected response date"),
      booleanColumn("FileSelectionCompleted", "File selection completed"),
      booleanColumn("TestingSheetCompleted", "Testing sheet completed"),
      dateColumn("FindingsSentDate", "Findings sent date"),
      dateColumn("CoverholderResponseReceivedDate", "Coverholder response received date"),
      choice("ReportStatus", "Report status", reportStatusChoices),
      choice("InvoiceStatus", "Invoice status", invoiceStatusChoices),
      booleanColumn("PaymentReceived", "Payment received"),
      choice("DamSubmissionStatus", "DAM submission status", ["Not Required", "Not Started", "Submitted", "Accepted"]),
      dateColumn("DueDate", "Due date"),
      dateColumn("LastUpdatedDate", "Last updated date"),
      note("Labels", "Labels"),
      note("NextAction", "Next action"),
      note("Blockers", "Blockers"),
      note("TrackerProjectJson", "Tracker project JSON"),
    ],
  },
  {
    key: "auditTeamMembers",
    displayName: "Audit Team Members",
    columns: [
      text("TeamMemberKey", "Team member key", true, true),
      text("TrackerAssignmentId", "Tracker assignment ID", true, true),
      text("AssignmentNumber", "Assignment number", true, true),
      text("PersonName", "Person name", true, true),
      choice("TeamRole", "Team role", ["Lead Auditor", "Supporting Auditor"], true),
      booleanColumn("ActiveOnAssignment", "Active on assignment"),
    ],
  },
  {
    key: "auditComments",
    displayName: "Audit Comments",
    columns: [
      text("TrackerCommentId", "Tracker comment ID", true, true),
      text("TrackerAssignmentId", "Tracker assignment ID", true, true),
      text("AssignmentNumber", "Assignment number", true, true),
      dateColumn("CommentCreatedAt", "Comment created at"),
      text("CommentAuthor", "Comment author"),
      note("CommentBody", "Comment body"),
    ],
  },
  {
    key: "auditChecklistItems",
    displayName: "Audit Checklist Items",
    columns: [
      text("TrackerChecklistItemId", "Tracker checklist item ID", true, true),
      text("TrackerAssignmentId", "Tracker assignment ID", true, true),
      text("AssignmentNumber", "Assignment number", true, true),
      text("ChecklistKey", "Checklist key", true, true),
      text("ChecklistStage", "Checklist stage", false, true),
      text("ChecklistItem", "Checklist item", true),
      booleanColumn("Completed", "Completed"),
    ],
  },
  {
    key: "auditStatusHistory",
    displayName: "Audit Status History",
    columns: [
      text("TrackerHistoryId", "Tracker history ID", true, true),
      text("TrackerAssignmentId", "Tracker assignment ID", true, true),
      text("AssignmentNumber", "Assignment number", true, true),
      dateColumn("ChangedAt", "Changed at"),
      text("ChangedBy", "Changed by"),
      choice("FromStage", "From stage", stageChoices),
      choice("ToStage", "To stage", stageChoices),
      note("StageNote", "Stage note"),
    ],
  },
  {
    key: "auditActivityLog",
    displayName: "Audit Activity Log",
    columns: [
      text("TrackerEventId", "Tracker event ID", true, true),
      text("TrackerAssignmentId", "Tracker assignment ID", true, true),
      text("AssignmentNumber", "Assignment number", true, true),
      dateColumn("OccurredAt", "Occurred at"),
      choice("EventType", "Event type", ["field", "stage", "comment", "document", "checklist", "team", "finance", "system"], true),
      text("ActorName", "Actor name", false, true),
      text("ActorUsername", "Actor username"),
      text("Summary", "Summary", true),
      note("Detail", "Detail"),
      text("FieldName", "Field name"),
      note("PreviousValue", "Previous value"),
      note("NewValue", "New value"),
      text("SourceList", "Source list"),
      text("SourceRecordId", "Source record ID"),
    ],
  },
  {
    key: "trackerUsers",
    displayName: "Tracker Users",
    columns: [
      text("TrackerUsername", "Tracker username", true, true),
      text("FullName", "Full name", true, true),
      text("Email", "Email", false, true),
      choice("Role", "Role", ["Admin", "Audit Manager", "Auditor", "Finance", "Read Only"], true),
      choice("PermissionGroup", "Permission group", ["Admin", "Audit Manager", "Auditor", "Finance", "Read Only"], true),
      booleanColumn("Active", "Active"),
      booleanColumn("EmailVerified", "Email verified"),
      choice("AccessRequestStatus", "Access request status", ["Approved", "Pending Verification", "Pending Approval", "Rejected"], true),
      choice("DefaultVisibility", "Default visibility", ["Role Default", "All Projects", "Assigned Projects", "Finance Records"]),
      dateColumn("RequestedAt", "Requested at"),
      dateColumn("ApprovedAt", "Approved at"),
      text("ApprovedBy", "Approved by"),
      note("RejectionReason", "Rejection reason"),
      note("VerificationCodeHash", "Verification code hash"),
      dateColumn("VerificationSentAt", "Verification sent at"),
    ],
  },
];

const env = loadEnvFile(resolve("server.env"));
const token = await getGraphToken(env);
const site = targetSiteUrl
  ? await graphRequest(sitePathFromUrl(targetSiteUrl))
  : await graphRequest("/sites/root");
const existingLists = await listSiteLists(site.id);
const listIds = {};
const created = [];
const reused = [];

for (const schema of schemas) {
  const existing = existingLists.find((list) => list.displayName === schema.displayName);
  if (existing) {
    listIds[schema.key] = existing.id;
    reused.push(schema.displayName);
    continue;
  }
  const list = await graphRequest(`/sites/${encodeSegment(site.id)}/lists`, {
    method: "POST",
    body: JSON.stringify(toGraphListCreateRequest(schema)),
  });
  listIds[schema.key] = list.id;
  created.push(schema.displayName);
}

const users = loadUsers(env);
const projects = loadProjects();
await upsertRows(site.id, listIds.trackerUsers, "Email", users.map(userToTrackerFields));
await upsertRows(site.id, listIds.auditAssignments, "TrackerAssignmentId", projects.map(projectToTrackerFields));

const result = {
  generatedAt: new Date().toISOString(),
  site: {
    id: site.id,
    name: site.displayName || site.name || "Root site",
    webUrl: site.webUrl,
  },
  listIds,
  created,
  reused,
  seeded: {
    trackerUsers: users.length,
    auditAssignments: projects.length,
  },
};

writeFileSync(outputPath, `${JSON.stringify(result, null, 2)}\n`);
console.log(`SITE=${result.site.name}`);
console.log(`CREATED=${created.length}`);
console.log(`REUSED=${reused.length}`);
console.log(`SEEDED_USERS=${users.length}`);
console.log(`SEEDED_PROJECTS=${projects.length}`);
console.log(`OUTPUT=${outputPath}`);

async function getGraphToken(config) {
  const body = new URLSearchParams({
    client_id: required(config.MICROSOFT_CLIENT_ID, "MICROSOFT_CLIENT_ID"),
    client_secret: required(config.MICROSOFT_CLIENT_SECRET, "MICROSOFT_CLIENT_SECRET"),
    scope: "https://graph.microsoft.com/.default",
    grant_type: "client_credentials",
  });
  const response = await fetch(
    `https://login.microsoftonline.com/${required(config.MICROSOFT_TENANT_ID, "MICROSOFT_TENANT_ID")}/oauth2/v2.0/token`,
    { method: "POST", body },
  );
  const payload = await response.json();
  if (!response.ok || !payload.access_token) {
    throw new Error(`Graph token failed: ${payload.error || response.status}`);
  }
  return payload.access_token;
}

async function listSiteLists(siteId) {
  const lists = [];
  let path = `/sites/${encodeSegment(siteId)}/lists?$select=id,displayName`;
  while (path) {
    const payload = await graphRequest(path);
    lists.push(...(payload.value ?? []));
    path = payload["@odata.nextLink"] ?? "";
  }
  return lists;
}

async function upsertRows(siteId, listId, uniqueField, rows) {
  if (!rows.length) return;
  const existingItems = await listItems(siteId, listId);
  const existingByKey = new Map(
    existingItems
      .map((item) => [String(item.fields?.[uniqueField] || "").toLowerCase(), item])
      .filter(([value]) => value),
  );
  for (const fields of rows) {
    const key = String(fields[uniqueField] || "").toLowerCase();
    if (!key) continue;
    const existing = existingByKey.get(key);
    if (existing) {
      await graphRequest(
        `/sites/${encodeSegment(siteId)}/lists/${encodeSegment(listId)}/items/${encodeSegment(existing.id)}/fields`,
        { method: "PATCH", body: JSON.stringify(fields) },
      );
    } else {
      await graphRequest(
        `/sites/${encodeSegment(siteId)}/lists/${encodeSegment(listId)}/items`,
        { method: "POST", body: JSON.stringify({ fields }) },
      );
    }
  }
}

async function listItems(siteId, listId) {
  const items = [];
  let path = `/sites/${encodeSegment(siteId)}/lists/${encodeSegment(listId)}/items?expand=fields`;
  while (path) {
    const payload = await graphRequest(path);
    items.push(...(payload.value ?? []));
    path = payload["@odata.nextLink"] ?? "";
  }
  return items;
}

async function graphRequest(pathOrUrl, init = {}) {
  const url = pathOrUrl.startsWith("https://")
    ? pathOrUrl
    : `${graphBaseUrl}${pathOrUrl}`;
  const response = await fetch(url, {
    ...init,
    headers: {
      authorization: `Bearer ${token}`,
      "content-type": "application/json",
      ...(init.headers ?? {}),
    },
  });
  if (!response.ok) {
    throw new Error(`Graph ${response.status}: ${await response.text()}`);
  }
  if (response.status === 204) return {};
  return response.json();
}

function loadUsers(config) {
  const storePath = resolve("server", "data", "access-users.json");
  if (existsSync(storePath)) {
    const store = JSON.parse(readFileSync(storePath, "utf8"));
    if (Array.isArray(store.users) && store.users.length) return store.users;
  }
  const now = new Date().toISOString();
  return splitCsv(config.TRACKER_ADMIN_EMAILS).map((email) => ({
    email,
    username: usernameFromEmail(email),
    fullName: nameFromEmail(email),
    role: "Admin",
    permissionGroup: "Admin",
    active: true,
    defaultVisibility: "All Projects",
    emailVerified: true,
    accessRequestStatus: "Approved",
    requestedAt: now,
    approvedAt: now,
    approvedBy: "server-config",
    rejectionReason: "",
    verificationCodeHash: "",
    verificationSentAt: "",
  }));
}

function loadProjects() {
  const storePath = resolve("server", "data", "audit-projects.json");
  if (!existsSync(storePath)) return [];
  const store = JSON.parse(readFileSync(storePath, "utf8"));
  return Array.isArray(store.projects) ? store.projects : [];
}

function loadEnvFile(path) {
  const result = {};
  for (const rawLine of readFileSync(path, "utf8").split(/\r?\n/)) {
    const line = rawLine.trim();
    if (!line || line.startsWith("#") || !line.includes("=")) continue;
    const index = line.indexOf("=");
    let value = line.slice(index + 1).trim();
    if (
      (value.startsWith('"') && value.endsWith('"')) ||
      (value.startsWith("'") && value.endsWith("'"))
    ) {
      value = value.slice(1, -1);
    }
    result[line.slice(0, index).trim()] = value;
  }
  return result;
}

function toGraphListCreateRequest(schema) {
  return {
    displayName: schema.displayName,
    columns: schema.columns.map(toGraphColumnDefinition),
    list: { template: "genericList" },
  };
}

function toGraphColumnDefinition(column) {
  const base = {
    name: column.name,
    displayName: column.displayName,
    required: Boolean(column.required),
    indexed: Boolean(column.indexed),
  };
  if (column.type === "note") {
    return {
      ...base,
      text: {
        allowMultipleLines: true,
        appendChangesToExistingText: false,
        linesForEditing: 6,
      },
    };
  }
  if (column.type === "choice") {
    return {
      ...base,
      choice: {
        allowTextEntry: false,
        choices: column.choices ?? [],
        displayAs: "dropDownMenu",
      },
    };
  }
  if (column.type === "currency") return { ...base, currency: {} };
  if (column.type === "dateTime") return { ...base, dateTime: {} };
  if (column.type === "boolean") return { ...base, boolean: {} };
  return {
    ...base,
    text: {
      allowMultipleLines: false,
      appendChangesToExistingText: false,
      linesForEditing: 0,
      maxLength: 255,
    },
  };
}

function text(name, displayName, required = false, indexed = false) {
  return { name, displayName, type: "text", required, indexed };
}

function note(name, displayName) {
  return { name, displayName, type: "note" };
}

function choice(name, displayName, choices, required = false) {
  return { name, displayName, type: "choice", choices, required };
}

function currency(name, displayName) {
  return { name, displayName, type: "currency" };
}

function dateColumn(name, displayName) {
  return { name, displayName, type: "dateTime" };
}

function booleanColumn(name, displayName) {
  return { name, displayName, type: "boolean" };
}

function splitCsv(value) {
  return String(value || "")
    .split(",")
    .map((item) => item.trim().toLowerCase())
    .filter(Boolean);
}

function usernameFromEmail(email) {
  return email
    .split("@")[0]
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ".")
    .replace(/^\.+|\.+$/g, "");
}

function nameFromEmail(email) {
  return usernameFromEmail(email)
    .split(".")
    .filter(Boolean)
    .map((part) => part.slice(0, 1).toUpperCase() + part.slice(1))
    .join(" ");
}

function required(value, name) {
  const textValue = String(value || "").trim();
  if (!textValue) throw new Error(`${name} is required`);
  return textValue;
}

function encodeSegment(value) {
  return encodeURIComponent(value);
}

function sitePathFromUrl(value) {
  const parsed = new URL(value);
  const path = parsed.pathname.replace(/\/$/, "");
  return path
    ? `/sites/${parsed.hostname}:${path}:`
    : `/sites/${parsed.hostname}:/`;
}
