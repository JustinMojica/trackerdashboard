import { existsSync } from "node:fs";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import { dirname } from "node:path";

const graphBaseUrl = "https://graph.microsoft.com/v1.0";

export function projectStoreStatus(config) {
  const mode =
    config.projectStoreMode === "microsoft-lists" ? "microsoft-lists" : "local";
  const missing = [];
  if (mode === "microsoft-lists") {
    if (!hasConfiguredValue(config.projectStoreSiteId)) {
      missing.push("TRACKER_PROJECTS_SITE_ID");
    }
    if (!hasConfiguredValue(config.projectStoreListId)) {
      missing.push("TRACKER_PROJECTS_LIST_ID");
    }
  }
  return {
    mode,
    configured: missing.length === 0,
    missing,
  };
}

export function createProjectStore({ mode, projectsFile, graph }) {
  if (mode === "microsoft-lists") {
    return createMicrosoftListsProjectStore(graph);
  }
  return createLocalProjectStore(projectsFile);
}

function createLocalProjectStore(projectsFile) {
  return {
    async load() {
      if (!existsSync(projectsFile)) return { projects: [] };
      const store = JSON.parse(await readFile(projectsFile, "utf8"));
      return { projects: Array.isArray(store.projects) ? store.projects : [] };
    },
    async save(store) {
      await mkdir(dirname(projectsFile), { recursive: true });
      await writeFile(
        projectsFile,
        `${JSON.stringify({ projects: store.projects ?? [] }, null, 2)}\n`,
      );
    },
  };
}

function createMicrosoftListsProjectStore({ siteId, listId, getAccessToken }) {
  return {
    async load() {
      const items = await listProjectItems({ siteId, listId, getAccessToken });
      return {
        projects: items
          .map((item) => trackerFieldsToProject(item.fields ?? {}))
          .filter((project) => project?.id),
      };
    },
    async save(store) {
      const projects = Array.isArray(store.projects) ? store.projects : [];
      const items = await listProjectItems({ siteId, listId, getAccessToken });
      const existingByProjectId = new Map(
        items
          .map((item) => [String(item.fields?.TrackerAssignmentId || ""), item])
          .filter(([id]) => id),
      );
      const wantedProjectIds = new Set(projects.map((project) => String(project.id)));

      for (const project of projects) {
        const fields = projectToTrackerFields(project);
        const existing = existingByProjectId.get(String(project.id));
        if (existing) {
          await graphRequest(
            getAccessToken,
            `/sites/${encodeGraphSegment(siteId)}/lists/${encodeGraphSegment(
              listId,
            )}/items/${encodeGraphSegment(existing.id)}/fields`,
            {
              method: "PATCH",
              body: JSON.stringify(fields),
            },
          );
        } else {
          await graphRequest(
            getAccessToken,
            `/sites/${encodeGraphSegment(siteId)}/lists/${encodeGraphSegment(
              listId,
            )}/items`,
            {
              method: "POST",
              body: JSON.stringify({ fields }),
            },
          );
        }
      }

      for (const item of items) {
        const projectId = String(item.fields?.TrackerAssignmentId || "");
        if (projectId && !wantedProjectIds.has(projectId)) {
          await graphRequest(
            getAccessToken,
            `/sites/${encodeGraphSegment(siteId)}/lists/${encodeGraphSegment(
              listId,
            )}/items/${encodeGraphSegment(item.id)}`,
            { method: "DELETE" },
          );
        }
      }
    },
  };
}

export function projectToTrackerFields(project) {
  return {
    Title: project.assignmentNumber || project.id,
    TrackerAssignmentId: project.id,
    AssignmentNumber: project.assignmentNumber,
    AssignmentSource: project.assignmentSource,
    AssignmentType: project.assignmentType,
    AuditEntity: project.auditEntity,
    ClientCoverholderCode: project.clientCoverholderCode,
    Broker: project.broker,
    LeadAuditor: project.assignedAuditor,
    CurrentStage: project.currentStage,
    AssignmentStatus: project.assignmentStatus,
    QuoteStatus: project.quoteStatus,
    QuoteAmount: Number(project.quoteAmount) || 0,
    TentativeAuditWeek: project.tentativeAuditWeek,
    ConfirmedAuditDate: nullableDate(project.confirmedAuditDate),
    AuditType: project.auditType,
    BaaReceived: Boolean(project.baaReceived),
    EndorsementsReceived: Boolean(project.endorsementsReceived),
    PremiumBdxReceived: Boolean(project.premiumBdxReceived),
    PreAuditQuestionnaireStatus: project.preAuditQuestionnaireStatus,
    DocumentRequestStatus: project.documentRequestStatus,
    DocumentRequestDate: nullableDate(project.documentRequestDate),
    BrokerLastChasedDate: nullableDate(project.brokerLastChasedDate),
    BrokerExpectedResponseDate: nullableDate(project.brokerExpectedResponseDate),
    FileSelectionCompleted: Boolean(project.fileSelectionCompleted),
    TestingSheetCompleted: Boolean(project.testingSheetCompleted),
    FindingsSentDate: nullableDate(project.findingsSentDate),
    CoverholderResponseReceivedDate: nullableDate(
      project.coverholderResponseReceivedDate,
    ),
    ReportStatus: project.reportStatus,
    InvoiceStatus: project.invoiceStatus,
    PaymentReceived: Boolean(project.paymentReceived),
    DamSubmissionStatus: project.damSubmissionStatus,
    DueDate: nullableDate(project.dueDate),
    LastUpdatedDate: nullableDate(project.lastUpdatedDate),
    Labels: JSON.stringify(project.labels ?? []),
    NextAction: project.nextAction,
    Blockers: project.blockers,
    TrackerProjectJson: JSON.stringify(project),
  };
}

function trackerFieldsToProject(fields) {
  const rawProject = String(fields.TrackerProjectJson || fields.ProjectJson || "");
  if (rawProject) {
    try {
      const project = JSON.parse(rawProject);
      return project && typeof project === "object" ? project : null;
    } catch {
      return null;
    }
  }
  return {
    id: String(fields.TrackerAssignmentId || ""),
    assignmentNumber: String(fields.AssignmentNumber || fields.Title || ""),
    assignmentSource: String(fields.AssignmentSource || "Email"),
    assignmentType: String(fields.AssignmentType || "DCA"),
    auditEntity: String(fields.AuditEntity || ""),
    clientCoverholderCode: String(fields.ClientCoverholderCode || ""),
    broker: String(fields.Broker || ""),
    assignedAuditor: String(fields.LeadAuditor || ""),
    auditTeam: [],
    currentStage: String(fields.CurrentStage || "Intake"),
    assignmentStatus: String(fields.AssignmentStatus || "New"),
    quoteStatus: String(fields.QuoteStatus || "Not Started"),
    quoteAmount: Number(fields.QuoteAmount) || 0,
    tentativeAuditWeek: String(fields.TentativeAuditWeek || ""),
    confirmedAuditDate: stringValue(fields.ConfirmedAuditDate),
    auditType: String(fields.AuditType || "Remote"),
    baaReceived: Boolean(fields.BaaReceived),
    endorsementsReceived: Boolean(fields.EndorsementsReceived),
    premiumBdxReceived: Boolean(fields.PremiumBdxReceived),
    preAuditQuestionnaireStatus: String(
      fields.PreAuditQuestionnaireStatus || "Not Started",
    ),
    documentRequestStatus: String(fields.DocumentRequestStatus || "Not Started"),
    documentRequestDate: stringValue(fields.DocumentRequestDate),
    brokerLastChasedDate: stringValue(fields.BrokerLastChasedDate),
    brokerExpectedResponseDate: stringValue(fields.BrokerExpectedResponseDate),
    fileSelectionCompleted: Boolean(fields.FileSelectionCompleted),
    testingSheetCompleted: Boolean(fields.TestingSheetCompleted),
    findingsSentDate: stringValue(fields.FindingsSentDate),
    coverholderResponseReceivedDate: stringValue(
      fields.CoverholderResponseReceivedDate,
    ),
    reportStatus: String(fields.ReportStatus || "Not Started"),
    invoiceStatus: String(fields.InvoiceStatus || "Not Started"),
    paymentReceived: Boolean(fields.PaymentReceived),
    damSubmissionStatus: String(fields.DamSubmissionStatus || "Not Required"),
    nextAction: String(fields.NextAction || ""),
    blockers: String(fields.Blockers || ""),
    dueDate: stringValue(fields.DueDate),
    lastUpdatedDate: stringValue(fields.LastUpdatedDate),
    labels: safeJsonArray(fields.Labels),
    checklistCompletions: {},
    statusHistory: [],
    comments: [],
    activityEvents: [],
  };
}

async function listProjectItems({ siteId, listId, getAccessToken }) {
  const items = [];
  let path =
    `/sites/${encodeGraphSegment(siteId)}/lists/${encodeGraphSegment(
      listId,
    )}/items?expand=fields`;
  while (path) {
    const response = await graphRequest(getAccessToken, path);
    items.push(...(response.value ?? []));
    path = response["@odata.nextLink"] ?? "";
  }
  return items;
}

async function graphRequest(getAccessToken, pathOrUrl, init = {}) {
  const token = await getAccessToken();
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

function nullableDate(value) {
  const text = String(value || "").trim();
  return text ? text : null;
}

function stringValue(value) {
  return value === undefined || value === null ? "" : String(value).slice(0, 10);
}

function safeJsonArray(value) {
  try {
    const parsed = JSON.parse(String(value || "[]"));
    return Array.isArray(parsed) ? parsed.map(String) : [];
  } catch {
    return [];
  }
}

function encodeGraphSegment(value) {
  return encodeURIComponent(value);
}

function hasConfiguredValue(value) {
  const normalized = String(value ?? "").trim().toLowerCase();
  if (!normalized) return false;
  if (normalized.startsWith("replace-with-")) return false;
  return true;
}
