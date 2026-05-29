export type CentralAuditTeamMember = {
  person: string;
  role: string;
};

export type CentralStatusHistoryItem = {
  id: string;
  changedAt: string;
  changedBy: string;
  fromStage: string;
  toStage: string;
  note: string;
};

export type CentralProjectComment = {
  id: string;
  createdAt: string;
  author: string;
  body: string;
};

export type CentralActivityEvent = {
  id: string;
  createdAt: string;
  actor: string;
  type: string;
  title: string;
  detail: string;
};

export type CentralAuditProject = {
  id: string;
  assignmentNumber: string;
  assignmentSource: string;
  assignmentType: string;
  auditEntity: string;
  clientCoverholderCode: string;
  broker: string;
  assignedAuditor: string;
  auditTeam: CentralAuditTeamMember[];
  reviewer: string;
  currentStage: string;
  assignmentStatus: string;
  quoteStatus: string;
  quoteAmount: number;
  tentativeAuditWeek: string;
  confirmedAuditDate: string;
  auditType: string;
  baaReceived: boolean;
  endorsementsReceived: boolean;
  premiumBdxReceived: boolean;
  preAuditQuestionnaireStatus: string;
  documentRequestStatus: string;
  documentRequestDate: string;
  brokerLastChasedDate: string;
  brokerExpectedResponseDate: string;
  fileSelectionCompleted: boolean;
  testingSheetCompleted: boolean;
  findingsSentDate: string;
  coverholderResponseReceivedDate: string;
  reportStatus: string;
  invoiceStatus: string;
  paymentReceived: boolean;
  damSubmissionStatus: string;
  nextAction: string;
  blockers: string;
  dueDate: string;
  lastUpdatedDate: string;
  labels: string[];
  checklistCompletions: Record<string, boolean>;
  statusHistory: CentralStatusHistoryItem[];
  comments: CentralProjectComment[];
  activityEvents: CentralActivityEvent[];
};

export type CentralPrototypeUser = {
  fullName: string;
  username: string;
  password: string;
  role: string;
  permissionGroup: string;
  email: string;
  active: boolean;
  defaultVisibility: string;
  emailVerified?: boolean;
  accessRequestStatus?: string;
  requestedAt?: string;
  approvedAt?: string;
  approvedBy?: string;
  rejectionReason?: string;
};

type ColumnType = "text" | "note" | "choice" | "number" | "currency" | "dateTime" | "boolean";

export type MicrosoftListColumn = {
  name: string;
  displayName: string;
  type: ColumnType;
  required?: boolean;
  indexed?: boolean;
  choices?: string[];
};

export type MicrosoftListSchema = {
  key:
    | "auditAssignments"
    | "auditTeamMembers"
    | "auditComments"
    | "auditChecklistItems"
    | "auditStatusHistory"
    | "auditActivityLog"
    | "trackerUsers";
  displayName: string;
  purpose: string;
  columns: MicrosoftListColumn[];
};

export type MicrosoftListSeedRow = {
  listKey: MicrosoftListSchema["key"];
  fields: Record<string, string | number | boolean>;
};

export type MicrosoftListsMigrationPackage = {
  app: "audit-assignment-tracker";
  schemaVersion: 3;
  exportedAt: string;
  exportedBy: string;
  storageTarget: "Microsoft Lists / SharePoint";
  schemas: MicrosoftListSchema[];
  graphListCreateRequests: Record<string, unknown>[];
  rows: {
    auditAssignments: MicrosoftListSeedRow[];
    auditTeamMembers: MicrosoftListSeedRow[];
    auditComments: MicrosoftListSeedRow[];
    auditChecklistItems: MicrosoftListSeedRow[];
    auditStatusHistory: MicrosoftListSeedRow[];
    auditActivityLog: MicrosoftListSeedRow[];
    trackerUsers: MicrosoftListSeedRow[];
  };
  totals: {
    lists: number;
    rows: number;
    assignments: number;
    activityLogEvents: number;
  };
};

const stageChoices = [
  "Intake",
  "Registration",
  "Quote",
  "Scheduling",
  "Pre-Audit",
  "File Selection",
  "Audit Fieldwork",
  "Findings",
  "Report Drafting",
  "Final Submission",
  "Invoice",
  "Closed",
];

const assignmentStatusChoices = ["New", "In Progress", "Blocked", "On Hold", "Completed"];
const quoteStatusChoices = ["Not Started", "Drafting", "Sent", "Accepted", "Rejected"];
const progressChoices = ["Not Started", "In Progress", "Complete", "Not Required"];
const reportStatusChoices = ["Not Started", "Drafting", "Review", "Issued"];
const invoiceStatusChoices = ["Not Started", "Prepared", "Sent", "Paid"];
const yesNoChoices = ["true", "false"];

export const microsoftListSchemas: MicrosoftListSchema[] = [
  {
    key: "auditAssignments",
    displayName: "Audit Assignments",
    purpose: "One row per audit assignment; this is the main tracker record.",
    columns: [
      text("TrackerAssignmentId", "Tracker assignment ID", true, true),
      text("AssignmentNumber", "Assignment number", true, true),
      choice("AssignmentSource", "Assignment source", ["Email", "DAM"], true),
      choice("AssignmentType", "Assignment type", ["DCA", "CH", "MGA", "Company Contract"], true),
      text("AuditEntity", "Audit entity", true, true),
      text("ClientCoverholderCode", "Client / coverholder code", false, true),
      text("Broker", "Broker"),
      text("LeadAuditor", "Lead auditor"),
      text("Reviewer", "Reviewer"),
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
    ],
  },
  {
    key: "auditTeamMembers",
    displayName: "Audit Team Members",
    purpose: "One row per assignment team member for lead/support workload reporting.",
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
    purpose: "One row per card comment, linked back to the assignment.",
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
    purpose: "One row per completed or reopened checklist item.",
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
    purpose: "One row per stage movement for cycle-time and stage history reporting.",
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
    purpose: "Append-only accountability log across edits, stages, comments, documents, checklist, team, and finance.",
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
    purpose: "Prototype role map to replace with Microsoft 365 identity groups later.",
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
    ],
  },
];

export function buildMicrosoftListsMigrationPackage(
  projects: CentralAuditProject[],
  users: CentralPrototypeUser[],
  options: { exportedBy: string; exportedAt?: string },
): MicrosoftListsMigrationPackage {
  const rows = {
    auditAssignments: projects.map(projectToAssignmentRow),
    auditTeamMembers: projects.flatMap(projectToTeamRows),
    auditComments: projects.flatMap(projectToCommentRows),
    auditChecklistItems: projects.flatMap(projectToChecklistRows),
    auditStatusHistory: projects.flatMap(projectToStatusHistoryRows),
    auditActivityLog: projects.flatMap(projectToActivityLogRows),
    trackerUsers: users.map(userToRow),
  };
  const rowCount = Object.values(rows).reduce((sum, listRows) => sum + listRows.length, 0);
  return {
    app: "audit-assignment-tracker",
    schemaVersion: 3,
    exportedAt: options.exportedAt ?? new Date().toISOString(),
    exportedBy: options.exportedBy,
    storageTarget: "Microsoft Lists / SharePoint",
    schemas: microsoftListSchemas,
    graphListCreateRequests: microsoftListSchemas.map(toGraphListCreateRequest),
    rows,
    totals: {
      lists: microsoftListSchemas.length,
      rows: rowCount,
      assignments: projects.length,
      activityLogEvents: rows.auditActivityLog.length,
    },
  };
}

export function toGraphListCreateRequest(schema: MicrosoftListSchema) {
  return {
    displayName: schema.displayName,
    columns: schema.columns.map(toGraphColumnDefinition),
    list: {
      template: "genericList",
    },
  };
}

function text(
  name: string,
  displayName: string,
  required = false,
  indexed = false,
): MicrosoftListColumn {
  return { name, displayName, type: "text", required, indexed };
}

function note(name: string, displayName: string): MicrosoftListColumn {
  return { name, displayName, type: "note" };
}

function choice(
  name: string,
  displayName: string,
  choices: string[],
  required = false,
): MicrosoftListColumn {
  return { name, displayName, type: "choice", choices, required };
}

function currency(name: string, displayName: string): MicrosoftListColumn {
  return { name, displayName, type: "currency" };
}

function dateColumn(name: string, displayName: string): MicrosoftListColumn {
  return { name, displayName, type: "dateTime" };
}

function booleanColumn(name: string, displayName: string): MicrosoftListColumn {
  return { name, displayName, type: "boolean" };
}

function toGraphColumnDefinition(column: MicrosoftListColumn) {
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
  if (column.type === "number") return { ...base, number: {} };
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

function projectToAssignmentRow(project: CentralAuditProject): MicrosoftListSeedRow {
  return row("auditAssignments", {
    Title: project.assignmentNumber,
    TrackerAssignmentId: project.id,
    AssignmentNumber: project.assignmentNumber,
    AssignmentSource: project.assignmentSource,
    AssignmentType: project.assignmentType,
    AuditEntity: project.auditEntity,
    ClientCoverholderCode: project.clientCoverholderCode,
    Broker: project.broker,
    LeadAuditor: project.assignedAuditor,
    Reviewer: project.reviewer,
    CurrentStage: project.currentStage,
    AssignmentStatus: project.assignmentStatus,
    QuoteStatus: project.quoteStatus,
    QuoteAmount: project.quoteAmount,
    TentativeAuditWeek: project.tentativeAuditWeek,
    ConfirmedAuditDate: dateOnly(project.confirmedAuditDate),
    AuditType: project.auditType,
    BaaReceived: project.baaReceived,
    EndorsementsReceived: project.endorsementsReceived,
    PremiumBdxReceived: project.premiumBdxReceived,
    PreAuditQuestionnaireStatus: project.preAuditQuestionnaireStatus,
    DocumentRequestStatus: project.documentRequestStatus,
    DocumentRequestDate: dateOnly(project.documentRequestDate),
    BrokerLastChasedDate: dateOnly(project.brokerLastChasedDate),
    BrokerExpectedResponseDate: dateOnly(project.brokerExpectedResponseDate),
    FileSelectionCompleted: project.fileSelectionCompleted,
    TestingSheetCompleted: project.testingSheetCompleted,
    FindingsSentDate: dateOnly(project.findingsSentDate),
    CoverholderResponseReceivedDate: dateOnly(project.coverholderResponseReceivedDate),
    ReportStatus: project.reportStatus,
    InvoiceStatus: project.invoiceStatus,
    PaymentReceived: project.paymentReceived,
    DamSubmissionStatus: project.damSubmissionStatus,
    DueDate: dateOnly(project.dueDate),
    LastUpdatedDate: dateOnly(project.lastUpdatedDate),
    Labels: project.labels.join("; "),
    NextAction: project.nextAction,
    Blockers: project.blockers,
  });
}

function projectToTeamRows(project: CentralAuditProject) {
  return project.auditTeam.map((member) =>
    row("auditTeamMembers", {
      Title: `${project.assignmentNumber} - ${member.person}`,
      TeamMemberKey: `${project.id}|${member.role}|${member.person}`,
      TrackerAssignmentId: project.id,
      AssignmentNumber: project.assignmentNumber,
      PersonName: member.person,
      TeamRole: member.role,
      ActiveOnAssignment: true,
    }),
  );
}

function projectToCommentRows(project: CentralAuditProject) {
  return project.comments.map((comment) =>
    row("auditComments", {
      Title: `${project.assignmentNumber} comment`,
      TrackerCommentId: comment.id,
      TrackerAssignmentId: project.id,
      AssignmentNumber: project.assignmentNumber,
      CommentCreatedAt: normalizeDateTime(comment.createdAt),
      CommentAuthor: comment.author,
      CommentBody: comment.body,
    }),
  );
}

function projectToChecklistRows(project: CentralAuditProject) {
  return Object.entries(project.checklistCompletions).map(([key, completed]) => {
    const [stage, ...itemParts] = key.split(":");
    const item = itemParts.join(":") || key;
    return row("auditChecklistItems", {
      Title: `${project.assignmentNumber} - ${item}`,
      TrackerChecklistItemId: `${project.id}|${key}`,
      TrackerAssignmentId: project.id,
      AssignmentNumber: project.assignmentNumber,
      ChecklistKey: key,
      ChecklistStage: stage,
      ChecklistItem: item,
      Completed: completed,
    });
  });
}

function projectToStatusHistoryRows(project: CentralAuditProject) {
  return project.statusHistory.map((history) =>
    row("auditStatusHistory", {
      Title: `${project.assignmentNumber} ${history.fromStage} to ${history.toStage}`,
      TrackerHistoryId: history.id,
      TrackerAssignmentId: project.id,
      AssignmentNumber: project.assignmentNumber,
      ChangedAt: normalizeDateTime(history.changedAt),
      ChangedBy: history.changedBy,
      FromStage: history.fromStage,
      ToStage: history.toStage,
      StageNote: history.note,
    }),
  );
}

function projectToActivityLogRows(project: CentralAuditProject) {
  const fromEvents = project.activityEvents.map((event) =>
    activityRow(project, {
      id: event.id,
      occurredAt: event.createdAt,
      eventType: normalizeEventType(event.type),
      actorName: event.actor,
      summary: event.title,
      detail: event.detail,
      sourceList: "Audit Assignments",
      sourceRecordId: project.id,
    }),
  );
  const fromStatusHistory = project.statusHistory.map((history) =>
    activityRow(project, {
      id: `activity-${history.id}`,
      occurredAt: history.changedAt,
      eventType: "stage",
      actorName: history.changedBy,
      summary: `${history.fromStage} to ${history.toStage}`,
      detail: history.note,
      fieldName: "CurrentStage",
      previousValue: history.fromStage,
      newValue: history.toStage,
      sourceList: "Audit Status History",
      sourceRecordId: history.id,
    }),
  );
  const fromComments = project.comments.map((comment) =>
    activityRow(project, {
      id: `activity-${comment.id}`,
      occurredAt: comment.createdAt,
      eventType: "comment",
      actorName: comment.author,
      summary: "Comment added",
      detail: comment.body,
      sourceList: "Audit Comments",
      sourceRecordId: comment.id,
    }),
  );
  return [...fromEvents, ...fromStatusHistory, ...fromComments].sort((a, b) =>
    String(a.fields.OccurredAt).localeCompare(String(b.fields.OccurredAt)),
  );
}

function activityRow(
  project: CentralAuditProject,
  event: {
    id: string;
    occurredAt: string;
    eventType: string;
    actorName: string;
    summary: string;
    detail: string;
    fieldName?: string;
    previousValue?: string;
    newValue?: string;
    sourceList: string;
    sourceRecordId: string;
  },
) {
  return row("auditActivityLog", {
    Title: `${project.assignmentNumber} - ${event.summary}`,
    TrackerEventId: event.id,
    TrackerAssignmentId: project.id,
    AssignmentNumber: project.assignmentNumber,
    OccurredAt: normalizeDateTime(event.occurredAt),
    EventType: event.eventType,
    ActorName: event.actorName,
    ActorUsername: usernameFromName(event.actorName),
    Summary: event.summary,
    Detail: event.detail,
    FieldName: event.fieldName ?? "",
    PreviousValue: event.previousValue ?? "",
    NewValue: event.newValue ?? "",
    SourceList: event.sourceList,
    SourceRecordId: event.sourceRecordId,
  });
}

function userToRow(user: CentralPrototypeUser): MicrosoftListSeedRow {
  return row("trackerUsers", {
    Title: user.fullName,
    TrackerUsername: user.username,
    FullName: user.fullName,
    Email: user.email,
    Role: user.role,
    PermissionGroup: user.permissionGroup,
    Active: user.active,
    EmailVerified: user.emailVerified ?? user.active,
    AccessRequestStatus: user.accessRequestStatus ?? "Approved",
    DefaultVisibility: user.defaultVisibility,
    RequestedAt: normalizeDateTime(user.requestedAt ?? ""),
    ApprovedAt: normalizeDateTime(user.approvedAt ?? ""),
    ApprovedBy: user.approvedBy ?? "",
    RejectionReason: user.rejectionReason ?? "",
  });
}

function row(
  listKey: MicrosoftListSchema["key"],
  fields: Record<string, string | number | boolean>,
): MicrosoftListSeedRow {
  return { listKey, fields };
}

function normalizeEventType(type: string) {
  if (type === "field" || type === "stage" || type === "document" || type === "checklist" || type === "team") {
    return type;
  }
  return "system";
}

function usernameFromName(name: string) {
  const normalized = name.trim().toLowerCase().replace(/[^a-z0-9]+/g, ".");
  return normalized.replace(/^\.+|\.+$/g, "");
}

function dateOnly(value: string) {
  if (!value) return "";
  if (/^\d{4}-\d{2}-\d{2}$/.test(value)) return value;
  return normalizeDateTime(value).slice(0, 10);
}

function normalizeDateTime(value: string) {
  if (!value) return "";
  const parsed = Date.parse(value);
  if (!Number.isNaN(parsed)) return new Date(parsed).toISOString();
  if (/^\d{4}-\d{2}-\d{2}$/.test(value)) return `${value}T12:00:00.000Z`;
  return value;
}
