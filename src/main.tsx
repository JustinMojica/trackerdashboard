import React, { FormEvent, useMemo, useState } from "react";
import { createRoot } from "react-dom/client";
import "./styles.css";

export type AssignmentSource = "Email" | "DAM";
export type AssignmentType = "DCA" | "CH" | "MGA" | "Company Contract";
export type AuditType = "Remote" | "Onsite";
export type Stage =
  | "Intake"
  | "Registration"
  | "Quote"
  | "Scheduling"
  | "Pre-Audit"
  | "File Selection"
  | "Audit Fieldwork"
  | "Findings"
  | "Report Drafting"
  | "Final Submission"
  | "Invoice"
  | "Closed";
export type AssignmentStatus =
  | "New"
  | "In Progress"
  | "Blocked"
  | "On Hold"
  | "Completed";
export type QuoteStatus =
  | "Not Started"
  | "Drafting"
  | "Sent"
  | "Accepted"
  | "Rejected";
export type ProgressStatus =
  | "Not Started"
  | "In Progress"
  | "Complete"
  | "Not Required";
export type ReportStatus = "Not Started" | "Drafting" | "Review" | "Issued";
export type InvoiceStatus = "Not Started" | "Prepared" | "Sent" | "Paid";
export type ProjectLabel =
  | "High Priority"
  | "Medium Priority"
  | "Low Priority"
  | "Waiting on Broker";
export type DamSubmissionStatus =
  | "Not Required"
  | "Not Started"
  | "Submitted"
  | "Accepted";

export type StatusHistoryItem = {
  id: string;
  changedAt: string;
  changedBy: string;
  fromStage: Stage;
  toStage: Stage;
  note: string;
};

export type ProjectComment = {
  id: string;
  createdAt: string;
  author: string;
  body: string;
};

export type AuditProject = {
  id: string;
  assignmentNumber: string;
  assignmentSource: AssignmentSource;
  assignmentType: AssignmentType;
  auditEntity: string;
  clientCoverholderCode: string;
  broker: string;
  assignedAuditor: string;
  reviewer: string;
  currentStage: Stage;
  assignmentStatus: AssignmentStatus;
  quoteStatus: QuoteStatus;
  quoteAmount: number;
  tentativeAuditWeek: string;
  confirmedAuditDate: string;
  auditType: AuditType;
  baaReceived: boolean;
  endorsementsReceived: boolean;
  premiumBdxReceived: boolean;
  preAuditQuestionnaireStatus: ProgressStatus;
  documentRequestStatus: ProgressStatus;
  documentRequestDate: string;
  brokerLastChasedDate: string;
  brokerExpectedResponseDate: string;
  fileSelectionCompleted: boolean;
  testingSheetCompleted: boolean;
  findingsSentDate: string;
  coverholderResponseReceivedDate: string;
  reportStatus: ReportStatus;
  invoiceStatus: InvoiceStatus;
  paymentReceived: boolean;
  damSubmissionStatus: DamSubmissionStatus;
  nextAction: string;
  blockers: string;
  dueDate: string;
  lastUpdatedDate: string;
  labels: ProjectLabel[];
  checklistCompletions: Record<string, boolean>;
  statusHistory: StatusHistoryItem[];
  comments: ProjectComment[];
};

type Filters = {
  auditor: string;
  stage: string;
  source: string;
  quoteStatus: string;
  dueDate: string;
  missingDocuments: boolean;
};

type SavedView =
  | "all"
  | "todaysWork"
  | "myAudits"
  | "blocked"
  | "dueThisWeek"
  | "awaitingDocuments";
type ViewMode = "kanban" | "table";

export const stages: Stage[] = [
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

const today = new Date("2026-05-05T12:00:00Z");
const storageKey = "audit-assignment-tracker-projects-v1";
const auditorStorageKey = "audit-assignment-tracker-auditors-v1";
const myAuditorStorageKey = "audit-assignment-tracker-my-auditor-v1";

const assignmentTypeOptions: AssignmentType[] = [
  "DCA",
  "CH",
  "MGA",
  "Company Contract",
];

const labelOptions: ProjectLabel[] = [
  "High Priority",
  "Medium Priority",
  "Low Priority",
  "Waiting on Broker",
];

const defaultAuditorOptions = [
  "Lorraine Mojica",
  "Walter Aviles",
  "Leslie Domenech",
  "Mark James",
  "Justin Mojica",
  "Sheilah Couture",
  "Annabelle J. Crawford Mojica",
  "Molly Aviles",
  "Lindsie Guillermo",
];

export const sampleProjects: AuditProject[] = [
  {
    id: "audit-001",
    assignmentNumber: "AA-2026-0142",
    assignmentSource: "DAM",
    assignmentType: "CH",
    auditEntity: "Northbridge Coverholder Operations",
    clientCoverholderCode: "CH-1048",
    broker: "Northbridge Market Services",
    assignedAuditor: "Lorraine Mojica",
    reviewer: "Owen Price",
    currentStage: "Quote",
    assignmentStatus: "Blocked",
    quoteStatus: "Sent",
    quoteAmount: 12800,
    tentativeAuditWeek: "2026-W20",
    confirmedAuditDate: "",
    auditType: "Remote",
    baaReceived: true,
    endorsementsReceived: false,
    premiumBdxReceived: false,
    preAuditQuestionnaireStatus: "Not Started",
    documentRequestStatus: "In Progress",
    documentRequestDate: "2026-04-28",
    brokerLastChasedDate: "2026-05-01",
    brokerExpectedResponseDate: "2026-05-08",
    fileSelectionCompleted: false,
    testingSheetCompleted: false,
    findingsSentDate: "",
    coverholderResponseReceivedDate: "",
    reportStatus: "Not Started",
    invoiceStatus: "Not Started",
    paymentReceived: false,
    damSubmissionStatus: "Not Started",
    nextAction: "Follow up on DAM quote approval and endorsements.",
    blockers: "",
    dueDate: "2026-05-03",
    lastUpdatedDate: "2026-05-01",
    labels: ["Waiting on Broker", "High Priority"],
    checklistCompletions: {},
    statusHistory: [
      {
        id: "h-001",
        changedAt: "2026-04-25",
        changedBy: "System seed",
        fromStage: "Intake",
        toStage: "Registration",
        note: "Assignment registered from DAM intake.",
      },
      {
        id: "h-002",
        changedAt: "2026-04-28",
        changedBy: "Maya Chen",
        fromStage: "Registration",
        toStage: "Quote",
        note: "Quote prepared and sent in DAM.",
      },
    ],
    comments: [
      {
        id: "c-001",
        createdAt: "2026-05-01 09:15 AM",
        author: "Prototype user",
        body: "Waiting on remaining intake support before moving forward.",
      },
    ],
  },
  {
    id: "audit-002",
    assignmentNumber: "AA-2026-0148",
    assignmentSource: "Email",
    assignmentType: "MGA",
    auditEntity: "Harbor Specialty Program",
    clientCoverholderCode: "CH-2217",
    broker: "Harbor Underwriting Group",
    assignedAuditor: "Walter Aviles",
    reviewer: "Priya Shah",
    currentStage: "Pre-Audit",
    assignmentStatus: "In Progress",
    quoteStatus: "Accepted",
    quoteAmount: 9300,
    tentativeAuditWeek: "2026-W19",
    confirmedAuditDate: "2026-05-08",
    auditType: "Onsite",
    baaReceived: true,
    endorsementsReceived: true,
    premiumBdxReceived: false,
    preAuditQuestionnaireStatus: "Complete",
    documentRequestStatus: "In Progress",
    documentRequestDate: "2026-04-25",
    brokerLastChasedDate: "2026-05-04",
    brokerExpectedResponseDate: "2026-05-07",
    fileSelectionCompleted: false,
    testingSheetCompleted: false,
    findingsSentDate: "",
    coverholderResponseReceivedDate: "",
    reportStatus: "Not Started",
    invoiceStatus: "Not Started",
    paymentReceived: false,
    damSubmissionStatus: "Not Required",
    nextAction: "Receive premium bordereau before file selection.",
    blockers: "",
    dueDate: "2026-05-07",
    lastUpdatedDate: "2026-05-04",
    labels: ["Waiting on Broker"],
    checklistCompletions: {},
    statusHistory: [
      {
        id: "h-003",
        changedAt: "2026-04-20",
        changedBy: "System seed",
        fromStage: "Intake",
        toStage: "Registration",
        note: "Email intake logged.",
      },
      {
        id: "h-004",
        changedAt: "2026-04-22",
        changedBy: "Lena Ortiz",
        fromStage: "Registration",
        toStage: "Quote",
        note: "Email quote template sent.",
      },
      {
        id: "h-005",
        changedAt: "2026-04-29",
        changedBy: "Lena Ortiz",
        fromStage: "Quote",
        toStage: "Pre-Audit",
        note: "Quote accepted and audit date confirmed.",
      },
    ],
    comments: [],
  },
  {
    id: "audit-003",
    assignmentNumber: "AA-2026-0155",
    assignmentSource: "DAM",
    assignmentType: "DCA",
    auditEntity: "Summit Claims Administration",
    clientCoverholderCode: "CH-3094",
    broker: "Summit Specialty Brokers",
    assignedAuditor: "Lorraine Mojica",
    reviewer: "Noah Reed",
    currentStage: "Findings",
    assignmentStatus: "Blocked",
    quoteStatus: "Accepted",
    quoteAmount: 15750,
    tentativeAuditWeek: "2026-W17",
    confirmedAuditDate: "2026-04-23",
    auditType: "Remote",
    baaReceived: true,
    endorsementsReceived: true,
    premiumBdxReceived: true,
    preAuditQuestionnaireStatus: "Complete",
    documentRequestStatus: "Complete",
    documentRequestDate: "2026-04-08",
    brokerLastChasedDate: "2026-04-18",
    brokerExpectedResponseDate: "2026-04-22",
    fileSelectionCompleted: true,
    testingSheetCompleted: true,
    findingsSentDate: "2026-04-30",
    coverholderResponseReceivedDate: "",
    reportStatus: "Not Started",
    invoiceStatus: "Not Started",
    paymentReceived: false,
    damSubmissionStatus: "Not Started",
    nextAction: "Chase coverholder response to findings.",
    blockers: "",
    dueDate: "2026-05-10",
    lastUpdatedDate: "2026-05-04",
    labels: ["High Priority"],
    checklistCompletions: {},
    statusHistory: [
      {
        id: "h-006",
        changedAt: "2026-04-05",
        changedBy: "System seed",
        fromStage: "Intake",
        toStage: "Quote",
        note: "Fast-tracked DAM assignment.",
      },
      {
        id: "h-007",
        changedAt: "2026-04-24",
        changedBy: "Maya Chen",
        fromStage: "Audit Fieldwork",
        toStage: "Findings",
        note: "Testing completed; findings sent.",
      },
    ],
    comments: [
      {
        id: "c-002",
        createdAt: "2026-05-04 02:30 PM",
        author: "Prototype user",
        body: "Findings sent; response is the next gating item.",
      },
    ],
  },
  {
    id: "audit-004",
    assignmentNumber: "AA-2026-0161",
    assignmentSource: "Email",
    assignmentType: "Company Contract",
    auditEntity: "Cedar Binding Authority",
    clientCoverholderCode: "CH-4175",
    broker: "Cedar Risk Partners",
    assignedAuditor: "Justin Mojica",
    reviewer: "Priya Shah",
    currentStage: "Final Submission",
    assignmentStatus: "In Progress",
    quoteStatus: "Accepted",
    quoteAmount: 11200,
    tentativeAuditWeek: "2026-W15",
    confirmedAuditDate: "2026-04-10",
    auditType: "Onsite",
    baaReceived: true,
    endorsementsReceived: true,
    premiumBdxReceived: true,
    preAuditQuestionnaireStatus: "Complete",
    documentRequestStatus: "Complete",
    documentRequestDate: "2026-04-01",
    brokerLastChasedDate: "2026-04-08",
    brokerExpectedResponseDate: "2026-04-12",
    fileSelectionCompleted: true,
    testingSheetCompleted: true,
    findingsSentDate: "2026-04-16",
    coverholderResponseReceivedDate: "2026-04-22",
    reportStatus: "Issued",
    invoiceStatus: "Prepared",
    paymentReceived: false,
    damSubmissionStatus: "Not Required",
    nextAction: "Send final report package and invoice by email.",
    blockers: "",
    dueDate: "2026-05-06",
    lastUpdatedDate: "2026-05-04",
    labels: ["Medium Priority"],
    checklistCompletions: {},
    statusHistory: [
      {
        id: "h-008",
        changedAt: "2026-04-01",
        changedBy: "Jon Bell",
        fromStage: "Quote",
        toStage: "Scheduling",
        note: "Quote accepted.",
      },
      {
        id: "h-009",
        changedAt: "2026-05-01",
        changedBy: "Reviewer",
        fromStage: "Report Drafting",
        toStage: "Final Submission",
        note: "Report approved for final issue.",
      },
    ],
    comments: [],
  },
];

const blankProject = (): AuditProject => ({
  id: `audit-${Date.now()}`,
  assignmentNumber: `AA-2026-${Math.floor(1000 + Math.random() * 9000)}`,
  assignmentSource: "Email",
  assignmentType: "CH",
  auditEntity: "",
  clientCoverholderCode: "",
  broker: "",
  assignedAuditor: "",
  reviewer: "",
  currentStage: "Intake",
  assignmentStatus: "New",
  quoteStatus: "Not Started",
  quoteAmount: 0,
  tentativeAuditWeek: "",
  confirmedAuditDate: "",
  auditType: "Remote",
  baaReceived: false,
  endorsementsReceived: false,
  premiumBdxReceived: false,
  preAuditQuestionnaireStatus: "Not Started",
  documentRequestStatus: "Not Started",
  documentRequestDate: "",
  brokerLastChasedDate: "",
  brokerExpectedResponseDate: "",
  fileSelectionCompleted: false,
  testingSheetCompleted: false,
  findingsSentDate: "",
  coverholderResponseReceivedDate: "",
  reportStatus: "Not Started",
  invoiceStatus: "Not Started",
  paymentReceived: false,
  damSubmissionStatus: "Not Required",
  nextAction: "",
  blockers: "",
  dueDate: "",
  lastUpdatedDate: new Date().toISOString().slice(0, 10),
  labels: [],
  checklistCompletions: {},
  statusHistory: [],
  comments: [],
});

export const requiredDocuments = [
  { key: "baaReceived", label: "BAA received" },
  { key: "endorsementsReceived", label: "Endorsements received" },
  { key: "premiumBdxReceived", label: "Premium BDX received" },
] as const;

const checklistByStage: Record<Stage, string[]> = {
  Intake: [
    "Capture assignment number and source",
    "Record broker and client / coverholder code",
    "Assign audit owner",
  ],
  Registration: [
    "Validate assignment details",
    "Create SharePoint-ready item record",
    "Confirm reviewer",
  ],
  Quote: [
    "Prepare base quote",
    "Confirm quote amount",
    "Record quote decision",
  ],
  Scheduling: [
    "Confirm tentative audit week",
    "Confirm audit date",
    "Confirm remote or onsite audit type",
  ],
  "Pre-Audit": [
    "Send document request",
    "Track BAA, endorsements, and premium BDX",
    "Receive pre-audit questionnaire",
  ],
  "File Selection": [
    "Confirm premium BDX received",
    "Complete sample selection",
    "Notify auditor of selected files",
  ],
  "Audit Fieldwork": [
    "Complete testing sheet",
    "Log exceptions",
    "Prepare findings list",
  ],
  Findings: [
    "Send findings",
    "Receive coverholder response",
    "Complete recommendations / wrap-up decision",
  ],
  "Report Drafting": [
    "Draft report",
    "Reviewer quality check",
    "Resolve review comments",
  ],
  "Final Submission": [
    "Prepare final report package",
    "Submit to the correct channel",
    "Record submission confirmation",
  ],
  Invoice: [
    "Prepare invoice",
    "Issue invoice",
    "Track payment",
    "Confirm payment received",
  ],
  Closed: [
    "Confirm final report, submission, and invoice complete",
    "Archive audit record",
    "Capture lessons learned",
  ],
};

export function withProjectDefaults(project: AuditProject): AuditProject {
  return {
    ...project,
    assignmentType: project.assignmentType ?? "CH",
    auditEntity: project.auditEntity ?? "",
    paymentReceived:
      project.paymentReceived ?? project.invoiceStatus === "Paid",
    labels: project.labels ?? [],
    documentRequestDate: project.documentRequestDate ?? "",
    brokerLastChasedDate: project.brokerLastChasedDate ?? "",
    brokerExpectedResponseDate: project.brokerExpectedResponseDate ?? "",
    checklistCompletions: project.checklistCompletions ?? {},
    comments: project.comments ?? [],
  };
}

function loadProjects(): AuditProject[] {
  const raw = localStorage.getItem(storageKey);
  if (!raw) {
    localStorage.setItem(storageKey, JSON.stringify(sampleProjects));
    return sampleProjects;
  }
  try {
    return (JSON.parse(raw) as AuditProject[]).map(withProjectDefaults);
  } catch {
    return sampleProjects;
  }
}

function saveProjects(projects: AuditProject[]) {
  localStorage.setItem(storageKey, JSON.stringify(projects));
}

function loadAuditors(): string[] {
  const raw = localStorage.getItem(auditorStorageKey);
  if (!raw) {
    localStorage.setItem(
      auditorStorageKey,
      JSON.stringify(defaultAuditorOptions),
    );
    return defaultAuditorOptions;
  }
  try {
    const parsed = JSON.parse(raw) as string[];
    return parsed.length ? parsed : defaultAuditorOptions;
  } catch {
    return defaultAuditorOptions;
  }
}

function saveAuditors(auditors: string[]) {
  localStorage.setItem(auditorStorageKey, JSON.stringify(auditors));
}

function timestampNow() {
  return new Date().toLocaleString("en-US", {
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
  });
}

function checklistKey(stage: Stage, item: string) {
  return `${stage}:${item}`;
}

export function getMissingDocuments(project: AuditProject) {
  return requiredDocuments
    .filter((doc) => !project[doc.key])
    .map((doc) => doc.label);
}

export function computedBlockers(project: AuditProject) {
  const blockers: string[] = [...getMissingDocuments(project)];
  if (
    stages.indexOf(project.currentStage) >= stages.indexOf("Quote") &&
    project.quoteStatus !== "Accepted"
  ) {
    blockers.push("Quote not accepted");
  }
  if (
    stages.indexOf(project.currentStage) >= stages.indexOf("File Selection") &&
    !project.premiumBdxReceived
  ) {
    blockers.push("Premium BDX required before file selection");
  }
  if (
    stages.indexOf(project.currentStage) >= stages.indexOf("Findings") &&
    !project.coverholderResponseReceivedDate
  ) {
    blockers.push(
      "Coverholder response required before recommendations / wrap-up",
    );
  }
  if (project.blockers.trim()) blockers.push(project.blockers.trim());
  return blockers;
}

export function canMoveToStage(project: AuditProject, targetStage: Stage) {
  const targetIndex = stages.indexOf(targetStage);
  if (
    targetIndex >= stages.indexOf("Scheduling") &&
    project.quoteStatus !== "Accepted"
  ) {
    return "Quote must be accepted before moving to Scheduling.";
  }
  if (
    targetIndex >= stages.indexOf("File Selection") &&
    !project.premiumBdxReceived
  ) {
    return "Premium BDX must be received before moving to File Selection.";
  }
  if (
    targetIndex >= stages.indexOf("Report Drafting") &&
    !project.coverholderResponseReceivedDate
  ) {
    return "Coverholder response must be received before recommendations / wrap-up and report drafting.";
  }
  return "";
}

export function daysUntil(dateValue: string) {
  if (!dateValue) return Number.POSITIVE_INFINITY;
  const due = new Date(`${dateValue}T12:00:00Z`);
  return Math.ceil((due.getTime() - today.getTime()) / (1000 * 60 * 60 * 24));
}

export function dueLabel(project: AuditProject) {
  const days = daysUntil(project.dueDate);
  if (!Number.isFinite(days))
    return { text: "No due date", className: "muted" };
  if (days < 0)
    return { text: `${Math.abs(days)}d overdue`, className: "danger" };
  if (days <= 3) return { text: `Due in ${days}d`, className: "warning" };
  return { text: `Due ${project.dueDate}`, className: "ok" };
}

export type DocumentWorkflowAction =
  | "markWaitingOnBroker"
  | "recordBrokerChase"
  | "markDocumentsComplete";

export type ActivityItem = {
  id: string;
  timestamp: string;
  type: "stage" | "comment" | "checklist" | "document";
  title: string;
  detail: string;
  tone?: "ok" | "warning" | "danger" | "muted";
};

function todayIso() {
  return new Date().toISOString().slice(0, 10);
}

function addUniqueLabel(labels: ProjectLabel[], label: ProjectLabel) {
  return labels.includes(label) ? labels : [...labels, label];
}

export function documentReadiness(project: AuditProject) {
  const requiredComplete = requiredDocuments.filter((doc) => project[doc.key]).length;
  const workflowComplete = [
    project.preAuditQuestionnaireStatus === "Complete",
    project.documentRequestStatus === "Complete",
  ].filter(Boolean).length;
  const completeCount = requiredComplete + workflowComplete;
  const totalCount = requiredDocuments.length + 2;
  return {
    completeCount,
    totalCount,
    percent: Math.round((completeCount / totalCount) * 100),
    missingDocuments: getMissingDocuments(project),
    waitingOnBroker: project.labels.includes("Waiting on Broker"),
  };
}

export function applyDocumentWorkflowAction(
  project: AuditProject,
  action: DocumentWorkflowAction,
  date = todayIso(),
): AuditProject {
  if (action === "markWaitingOnBroker") {
    return withProjectDefaults({
      ...project,
      labels: addUniqueLabel(project.labels, "Waiting on Broker"),
      assignmentStatus: project.currentStage === "Closed" ? project.assignmentStatus : "On Hold",
      documentRequestStatus:
        project.documentRequestStatus === "Not Required" ? "In Progress" : project.documentRequestStatus,
      documentRequestDate: project.documentRequestDate || date,
      brokerLastChasedDate: project.brokerLastChasedDate || date,
      nextAction:
        project.nextAction || "Follow up with broker on outstanding documents.",
      lastUpdatedDate: date,
    });
  }
  if (action === "recordBrokerChase") {
    return withProjectDefaults({
      ...project,
      labels: addUniqueLabel(project.labels, "Waiting on Broker"),
      assignmentStatus: project.currentStage === "Closed" ? project.assignmentStatus : "On Hold",
      documentRequestStatus:
        project.documentRequestStatus === "Not Required" ? "In Progress" : project.documentRequestStatus,
      documentRequestDate: project.documentRequestDate || date,
      brokerLastChasedDate: date,
      nextAction: `Broker chased on ${date}; await outstanding documents.`,
      lastUpdatedDate: date,
    });
  }
  return withProjectDefaults({
    ...project,
    baaReceived: true,
    endorsementsReceived: true,
    premiumBdxReceived: true,
    preAuditQuestionnaireStatus: "Complete",
    documentRequestStatus: "Complete",
    labels: project.labels.filter((label) => label !== "Waiting on Broker"),
    assignmentStatus: project.currentStage === "Closed" ? project.assignmentStatus : "In Progress",
    nextAction: "Documents complete; proceed with file selection/readiness review.",
    lastUpdatedDate: date,
  });
}

export function activityTimeline(project: AuditProject): ActivityItem[] {
  const stageItems: ActivityItem[] = project.statusHistory.map((item) => ({
    id: `stage-${item.id}`,
    timestamp: item.changedAt,
    type: "stage",
    title: `${item.fromStage} → ${item.toStage}`,
    detail: `${item.note} · ${item.changedBy}`,
    tone: "ok",
  }));
  const commentItems: ActivityItem[] = project.comments.map((comment) => ({
    id: `comment-${comment.id}`,
    timestamp: comment.createdAt,
    type: "comment",
    title: `Comment from ${comment.author}`,
    detail: comment.body,
    tone: "muted",
  }));
  const checklistItems: ActivityItem[] = Object.entries(project.checklistCompletions)
    .filter(([, complete]) => complete)
    .map(([key]) => ({
      id: `checklist-${key}`,
      timestamp: project.lastUpdatedDate,
      type: "checklist",
      title: "Checklist completed",
      detail: key.split(":").slice(1).join(":") || key,
      tone: "ok",
    }));
  const documentItems: ActivityItem[] = [];
  if (project.documentRequestDate) {
    documentItems.push({
      id: "document-requested",
      timestamp: project.documentRequestDate,
      type: "document",
      title: "Document request sent",
      detail: `Request status: ${project.documentRequestStatus}`,
      tone: "warning",
    });
  }
  if (project.brokerLastChasedDate) {
    documentItems.push({
      id: "broker-chased",
      timestamp: project.brokerLastChasedDate,
      type: "document",
      title: "Broker follow-up recorded",
      detail: project.brokerExpectedResponseDate
        ? `Expected response: ${project.brokerExpectedResponseDate}`
        : "No expected response date set",
      tone: "warning",
    });
  }
  return [...stageItems, ...commentItems, ...checklistItems, ...documentItems].sort(
    (a, b) => Date.parse(b.timestamp) - Date.parse(a.timestamp),
  );
}


function sourceTasks(project: AuditProject) {
  return project.assignmentSource === "DAM"
    ? [
        "Complete DAM quote fields",
        "Upload quote support in DAM",
        "Submit final report through DAM",
      ]
    : [
        "Generate email quote template",
        "Send client-ready quote email",
        "Send final report by email",
      ];
}

export function escapeCsv(value: string | number | boolean) {
  const text = String(value ?? "");
  return `"${text.replace(/"/g, '""')}"`;
}

function exportProjectsToCsv(projects: AuditProject[]) {
  const columns: [
    string,
    (project: AuditProject) => string | number | boolean,
  ][] = [
    ["Assignment Number", (project) => project.assignmentNumber],
    ["Source", (project) => project.assignmentSource],
    ["Assignment Type", (project) => project.assignmentType],
    ["Audit Entity", (project) => project.auditEntity],
    ["Client / Coverholder Code", (project) => project.clientCoverholderCode],
    ["Broker", (project) => project.broker],
    ["Assigned Auditor", (project) => project.assignedAuditor],
    ["Reviewer", (project) => project.reviewer],
    ["Current Stage", (project) => project.currentStage],
    ["Assignment Status", (project) => project.assignmentStatus],
    ["Quote Status", (project) => project.quoteStatus],
    ["Quote Amount", (project) => project.quoteAmount],
    ["Due Date", (project) => project.dueDate],
    ["Labels", (project) => project.labels.join("; ")],
    ["Payment Received", (project) => project.paymentReceived],
    ["Next Action", (project) => project.nextAction],
    ["Blockers", (project) => computedBlockers(project).join("; ")],
  ];
  const csv = [
    columns.map(([label]) => escapeCsv(label)).join(","),
    ...projects.map((project) =>
      columns.map(([, getter]) => escapeCsv(getter(project))).join(","),
    ),
  ].join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `audit-assignments-${new Date().toISOString().slice(0, 10)}.csv`;
  link.click();
  URL.revokeObjectURL(url);
}

function App() {
  const [projects, setProjects] = useState<AuditProject[]>(() =>
    loadProjects(),
  );
  const [selectedId, setSelectedId] = useState(projects[0]?.id ?? "");
  const [editing, setEditing] = useState<AuditProject | null>(null);
  const [filters, setFilters] = useState<Filters>({
    auditor: "",
    stage: "",
    source: "",
    quoteStatus: "",
    dueDate: "",
    missingDocuments: false,
  });
  const [message, setMessage] = useState("");
  const [auditorOptions, setAuditorOptions] = useState<string[]>(() =>
    loadAuditors(),
  );
  const [myAuditor, setMyAuditorState] = useState(
    () => localStorage.getItem(myAuditorStorageKey) || auditorOptions[0] || "",
  );
  const [savedView, setSavedView] = useState<SavedView>("all");
  const [viewMode, setViewMode] = useState<ViewMode>("kanban");
  const [hiddenWorkloadAuditors, setHiddenWorkloadAuditors] = useState<
    string[]
  >([]);
  const [shownZeroLoadAuditors, setShownZeroLoadAuditors] = useState<string[]>([]);

  const auditors = [
    ...auditorOptions,
    ...projects
      .map((project) => project.assignedAuditor)
      .filter((auditor) => auditor && !auditorOptions.includes(auditor)),
  ];
  const zeroLoadAuditors = auditors.filter(
    (auditor) =>
      !projects.some(
        (project) =>
          project.assignedAuditor === auditor && project.currentStage !== "Closed",
      ),
  );
  const effectiveHiddenWorkloadAuditors = Array.from(
    new Set([
      ...hiddenWorkloadAuditors,
      ...zeroLoadAuditors.filter(
        (auditor) => !shownZeroLoadAuditors.includes(auditor),
      ),
    ]),
  );

  const filteredProjects = useMemo(
    () =>
      projects.filter((project) => {
        if (filters.auditor && project.assignedAuditor !== filters.auditor)
          return false;
        if (filters.stage && project.currentStage !== filters.stage)
          return false;
        if (filters.source && project.assignmentSource !== filters.source)
          return false;
        if (filters.quoteStatus && project.quoteStatus !== filters.quoteStatus)
          return false;
        if (filters.dueDate === "overdue" && daysUntil(project.dueDate) >= 0)
          return false;
        if (
          filters.dueDate === "dueSoon" &&
          !(daysUntil(project.dueDate) >= 0 && daysUntil(project.dueDate) <= 3)
        )
          return false;
        if (
          filters.missingDocuments &&
          getMissingDocuments(project).length === 0
        )
          return false;
        if (
          savedView === "todaysWork" &&
          !(
            project.currentStage !== "Closed" &&
            (daysUntil(project.dueDate) <= 1 ||
              computedBlockers(project).length > 0)
          )
        )
          return false;
        if (savedView === "myAudits" && project.assignedAuditor !== myAuditor)
          return false;
        if (
          savedView === "blocked" &&
          computedBlockers(project).length === 0 &&
          project.assignmentStatus !== "Blocked"
        )
          return false;
        if (
          savedView === "dueThisWeek" &&
          !(daysUntil(project.dueDate) >= 0 && daysUntil(project.dueDate) <= 7)
        )
          return false;
        if (
          savedView === "awaitingDocuments" &&
          getMissingDocuments(project).length === 0
        )
          return false;
        return true;
      }),
    [projects, filters, savedView, myAuditor],
  );
  const selectedProject =
    projects.find((project) => project.id === selectedId) ?? projects[0];

  const persist = (nextProjects: AuditProject[]) => {
    setProjects(nextProjects);
    saveProjects(nextProjects);
  };

  const updateAuditors = (nextAuditors: string[]) => {
    setAuditorOptions(nextAuditors);
    saveAuditors(nextAuditors);
    if (nextAuditors.length && !nextAuditors.includes(myAuditor)) {
      setMyAuditorState(nextAuditors[0]);
      localStorage.setItem(myAuditorStorageKey, nextAuditors[0]);
    }
  };

  const updateMyAuditor = (auditor: string) => {
    setMyAuditorState(auditor);
    localStorage.setItem(myAuditorStorageKey, auditor);
  };

  const upsertProject = (project: AuditProject) => {
    const cleanProject = withProjectDefaults({
      ...project,
      quoteAmount: Number(project.quoteAmount) || 0,
      lastUpdatedDate: new Date().toISOString().slice(0, 10),
      damSubmissionStatus:
        project.assignmentSource === "DAM"
          ? project.damSubmissionStatus
          : ("Not Required" as DamSubmissionStatus),
    });
    const exists = projects.some((item) => item.id === cleanProject.id);
    const nextProjects = exists
      ? projects.map((item) =>
          item.id === cleanProject.id ? cleanProject : item,
        )
      : [cleanProject, ...projects];
    persist(nextProjects);
    setSelectedId(cleanProject.id);
    setEditing(null);
    setMessage(exists ? "Project updated." : "Project created.");
  };

  const moveProject = (project: AuditProject, targetStage: Stage) => {
    const blocker = canMoveToStage(project, targetStage);
    if (blocker) {
      setMessage(blocker);
      return;
    }
    const updated: AuditProject = {
      ...project,
      currentStage: targetStage,
      assignmentStatus:
        targetStage === "Closed"
          ? "Completed"
          : computedBlockers({ ...project, currentStage: targetStage }).length
            ? "Blocked"
            : "In Progress",
      lastUpdatedDate: new Date().toISOString().slice(0, 10),
      statusHistory: [
        ...project.statusHistory,
        {
          id: `h-${Date.now()}`,
          changedAt: timestampNow(),
          changedBy: "Prototype user",
          fromStage: project.currentStage,
          toStage: targetStage,
          note: "Stage changed in tracker.",
        },
      ],
    };
    persist(projects.map((item) => (item.id === project.id ? updated : item)));
    setSelectedId(project.id);
    setMessage(`${project.assignmentNumber} moved to ${targetStage}.`);
  };

  const addProjectComment = (
    project: AuditProject,
    comment: ProjectComment,
  ) => {
    const updatedProject = {
      ...project,
      comments: [...project.comments, comment],
      lastUpdatedDate: new Date().toISOString().slice(0, 10),
    };
    persist(
      projects.map((item) => (item.id === project.id ? updatedProject : item)),
    );
    setSelectedId(project.id);
    setMessage(`Comment added to ${project.assignmentNumber}.`);
  };

  const toggleChecklistItem = (project: AuditProject, key: string) => {
    const updatedProject = {
      ...project,
      checklistCompletions: {
        ...project.checklistCompletions,
        [key]: !project.checklistCompletions[key],
      },
      lastUpdatedDate: new Date().toISOString().slice(0, 10),
    };
    persist(
      projects.map((item) => (item.id === project.id ? updatedProject : item)),
    );
    setSelectedId(project.id);
  };

  const updateProjectDocumentWorkflow = (
    project: AuditProject,
    action: DocumentWorkflowAction,
  ) => {
    const updatedProject = applyDocumentWorkflowAction(project, action);
    persist(
      projects.map((item) => (item.id === project.id ? updatedProject : item)),
    );
    setSelectedId(project.id);
    const actionMessage =
      action === "markDocumentsComplete"
        ? "Document readiness marked complete."
        : action === "recordBrokerChase"
          ? "Broker follow-up recorded."
          : "Waiting on Broker applied.";
    setMessage(`${actionMessage} ${project.assignmentNumber} updated.`);
  };

  return (
    <main>
      <header className="hero">
        <div>
          <p className="eyebrow">Local storage prototype</p>
          <h1>Audit Assignment Tracker</h1>
          <p>
            Manage audit assignments from intake through final report, invoice,
            and close-out.
          </p>
        </div>
        <div className="hero-actions">
          <button onClick={() => setEditing(blankProject())}>
            Add project
          </button>
          <button
            className="secondary"
            onClick={() => {
              persist(sampleProjects);
              setSelectedId(sampleProjects[0].id);
            }}
          >
            Reset sample data
          </button>
        </div>
      </header>

      {message && (
        <div className="toast" role="status">
          {message}
        </div>
      )}

      <Dashboard projects={projects} />
      <SavedViews
        savedView={savedView}
        setSavedView={setSavedView}
        myAuditor={myAuditor}
      />
      <WorkloadCounts
        projects={projects}
        auditors={auditors}
        hiddenAuditors={effectiveHiddenWorkloadAuditors}
        toggleAuditorHidden={(auditor) => {
          const isZeroLoad = zeroLoadAuditors.includes(auditor);
          if (effectiveHiddenWorkloadAuditors.includes(auditor)) {
            setHiddenWorkloadAuditors((current) =>
              current.filter((item) => item !== auditor),
            );
            if (isZeroLoad) {
              setShownZeroLoadAuditors((current) =>
                current.includes(auditor) ? current : [...current, auditor],
              );
            }
            return;
          }
          setHiddenWorkloadAuditors((current) =>
            current.includes(auditor) ? current : [...current, auditor],
          );
          if (isZeroLoad) {
            setShownZeroLoadAuditors((current) =>
              current.filter((item) => item !== auditor),
            );
          }
        }}
        showAllAuditors={() => {
          setHiddenWorkloadAuditors([]);
          setShownZeroLoadAuditors(zeroLoadAuditors);
        }}
      />
      <PeopleAdmin
        auditorOptions={auditorOptions}
        setAuditorOptions={updateAuditors}
        myAuditor={myAuditor}
        setMyAuditor={updateMyAuditor}
      />
      <FiltersPanel
        filters={filters}
        setFilters={setFilters}
        auditors={auditors}
      />
      <div className="view-toolbar panel">
        <div className="segmented">
          <button
            className={viewMode === "kanban" ? "active" : "secondary"}
            onClick={() => setViewMode("kanban")}
          >
            Kanban view
          </button>
          <button
            className={viewMode === "table" ? "active" : "secondary"}
            onClick={() => setViewMode("table")}
          >
            Table view
          </button>
        </div>
        <button onClick={() => exportProjectsToCsv(filteredProjects)}>
          Export filtered CSV
        </button>
      </div>
      {viewMode === "kanban" ? (
        <Kanban
          projects={filteredProjects}
          selectedId={selectedProject?.id}
          onSelect={setSelectedId}
          onMove={moveProject}
        />
      ) : (
        <ProjectTable projects={filteredProjects} onSelect={setSelectedId} />
      )}
      {selectedProject && (
        <ProjectDetail
          project={selectedProject}
          onEdit={() => setEditing(selectedProject)}
          onMove={moveProject}
          onAddComment={addProjectComment}
          onToggleChecklist={toggleChecklistItem}
          onDocumentWorkflowAction={updateProjectDocumentWorkflow}
        />
      )}
      {editing && (
        <ProjectForm
          project={editing}
          onCancel={() => setEditing(null)}
          onSave={upsertProject}
          auditorOptions={auditorOptions}
        />
      )}
    </main>
  );
}

function Dashboard({ projects }: { projects: AuditProject[] }) {
  const overdue = projects.filter(
    (project) => daysUntil(project.dueDate) < 0,
  ).length;
  const dueSoon = projects.filter(
    (project) =>
      daysUntil(project.dueDate) >= 0 && daysUntil(project.dueDate) <= 3,
  ).length;
  const blocked = projects.filter(
    (project) =>
      computedBlockers(project).length > 0 ||
      project.assignmentStatus === "Blocked",
  ).length;
  const quoteValue = projects.reduce(
    (sum, project) => sum + project.quoteAmount,
    0,
  );
  return (
    <section className="summary-grid" aria-label="Dashboard summary">
      <SummaryCard
        label="Open projects"
        value={
          projects.filter((project) => project.currentStage !== "Closed").length
        }
      />
      <SummaryCard label="Blocked" value={blocked} tone="danger" />
      <SummaryCard label="Overdue" value={overdue} tone="danger" />
      <SummaryCard label="Due soon" value={dueSoon} tone="warning" />
      <SummaryCard
        label="Quoted value"
        value={quoteValue.toLocaleString("en-US", {
          style: "currency",
          currency: "USD",
          maximumFractionDigits: 0,
        })}
      />
    </section>
  );
}

function SummaryCard({
  label,
  value,
  tone = "",
}: {
  label: string;
  value: string | number;
  tone?: string;
}) {
  return (
    <article className={`summary-card ${tone}`}>
      <span>{label}</span>
      <strong>{value}</strong>
    </article>
  );
}

function SavedViews({
  savedView,
  setSavedView,
  myAuditor,
}: {
  savedView: SavedView;
  setSavedView: (view: SavedView) => void;
  myAuditor: string;
}) {
  const views: { id: SavedView; label: string; helper: string }[] = [
    { id: "all", label: "All audits", helper: "Everything visible" },
    {
      id: "todaysWork",
      label: "Today’s Work",
      helper: "Due, blocked, or next",
    },
    {
      id: "myAudits",
      label: "My audits",
      helper: myAuditor || "Choose me in Admin",
    },
    { id: "blocked", label: "Blocked audits", helper: "Needs attention" },
    { id: "dueThisWeek", label: "Due this week", helper: "Next 7 days" },
    {
      id: "awaitingDocuments",
      label: "Awaiting documents",
      helper: "Missing required docs",
    },
  ];
  return (
    <section className="panel saved-views">
      <div className="section-title">
        <h2>Saved views</h2>
        <span>One-click work queues</span>
      </div>
      <div className="saved-view-grid">
        {views.map((view) => (
          <button
            key={view.id}
            className={
              savedView === view.id ? "saved-view active" : "saved-view"
            }
            onClick={() => setSavedView(view.id)}
          >
            <strong>{view.label}</strong>
            <span>{view.helper}</span>
          </button>
        ))}
      </div>
    </section>
  );
}

function WorkloadCounts({
  projects,
  auditors,
  hiddenAuditors,
  toggleAuditorHidden,
  showAllAuditors,
}: {
  projects: AuditProject[];
  auditors: string[];
  hiddenAuditors: string[];
  toggleAuditorHidden: (auditor: string) => void;
  showAllAuditors: () => void;
}) {
  const buildRow = (auditor: string) => {
    const openProjects = projects.filter(
      (project) =>
        project.assignedAuditor === auditor &&
        project.currentStage !== "Closed",
    );
    const blockedCount = openProjects.filter(
      (project) => computedBlockers(project).length > 0,
    ).length;
    const dueSoonCount = openProjects.filter(
      (project) =>
        daysUntil(project.dueDate) >= 0 && daysUntil(project.dueDate) <= 7,
    ).length;
    const tone =
      openProjects.length >= 4
        ? "high"
        : openProjects.length >= 2
          ? "medium"
          : "low";
    const status =
      openProjects.length === 0
        ? "Available"
        : openProjects.length >= 4
          ? "At capacity"
          : openProjects.length >= 2
            ? "Moderate load"
            : "Light load";
    return {
      auditor,
      openCount: openProjects.length,
      blockedCount,
      dueSoonCount,
      status,
      tone,
    };
  };
  const visibleAuditors = auditors.filter(
    (auditor) => !hiddenAuditors.includes(auditor),
  );
  const workloadRows = visibleAuditors.map(buildRow);
  const totalOpen = auditors
    .map(buildRow)
    .reduce((sum, row) => sum + row.openCount, 0);

  return (
    <section className="panel workload-dashboard">
      <div className="section-title workload-header">
        <div>
          <p className="eyebrow dark">Workload dashboard</p>
          <h2>Auditor workload</h2>
        </div>
        <div className="workload-summary">
          <strong>{totalOpen}</strong>
          <span>open assignments</span>
          <small>
            {visibleAuditors.length} active · {hiddenAuditors.length} minimized
          </small>
        </div>
      </div>
      {hiddenAuditors.length > 0 && (
        <div className="minimized-auditors">
          <span>Minimized</span>
          {hiddenAuditors.map((auditor) => (
            <button
              type="button"
              className="secondary"
              key={auditor}
              onClick={() => toggleAuditorHidden(auditor)}
            >
              Show {auditor}
            </button>
          ))}
          <button type="button" onClick={showAllAuditors}>
            Show all
          </button>
        </div>
      )}
      <p className="workload-meter-note">
        Auditors with no open assignments are minimized by default. Active rows
        show plain workload counts instead of relative bars.
      </p>
      <div className="workload-list">
        {workloadRows.map((row) => (
          <article className={`workload-row ${row.tone}`} key={row.auditor}>
            <div className="workload-person">
              <span className="avatar">
                {row.auditor
                  .split(" ")
                  .map((part) => part[0])
                  .slice(0, 2)
                  .join("")}
              </span>
              <div>
                <strong>{row.auditor}</strong>
                <small>
                  {row.openCount} open · {row.blockedCount} blocked ·{" "}
                  {row.dueSoonCount} due soon
                </small>
              </div>
            </div>
            <div className="workload-stats" aria-label={`${row.auditor} workload summary`}>
              <span className="workload-stat primary">
                <strong>{row.openCount}</strong>
                Open
              </span>
              <span className={row.blockedCount ? "workload-stat danger" : "workload-stat"}>
                <strong>{row.blockedCount}</strong>
                Blocked
              </span>
              <span className={row.dueSoonCount ? "workload-stat warning" : "workload-stat"}>
                <strong>{row.dueSoonCount}</strong>
                Due soon
              </span>
              <span className={`workload-status ${row.tone}`}>{row.status}</span>
            </div>
            <button
              type="button"
              className="link"
              onClick={() => toggleAuditorHidden(row.auditor)}
            >
              Minimize
            </button>
          </article>
        ))}
      </div>
    </section>
  );
}

function PeopleAdmin({
  auditorOptions,
  setAuditorOptions,
  myAuditor,
  setMyAuditor,
}: {
  auditorOptions: string[];
  setAuditorOptions: (auditors: string[]) => void;
  myAuditor: string;
  setMyAuditor: (auditor: string) => void;
}) {
  const availableDefaultAuditors = defaultAuditorOptions.filter(
    (auditor) => !auditorOptions.includes(auditor),
  );
  const [newAuditor, setNewAuditor] = useState("");
  const [quickAuditor, setQuickAuditor] = useState(
    availableDefaultAuditors[0] ?? "",
  );
  const selectedQuickAuditor =
    quickAuditor || availableDefaultAuditors[0] || "";
  const addAuditorName = (name: string) => {
    const cleanName = name.trim();
    if (
      !cleanName ||
      auditorOptions.some(
        (auditor) => auditor.toLowerCase() === cleanName.toLowerCase(),
      )
    )
      return false;
    setAuditorOptions([...auditorOptions, cleanName]);
    return true;
  };
  const addAuditor = (event: FormEvent) => {
    event.preventDefault();
    if (addAuditorName(newAuditor)) setNewAuditor("");
  };
  const quickAddAuditor = () => {
    if (addAuditorName(selectedQuickAuditor)) {
      const remainingAuditors = availableDefaultAuditors.filter(
        (auditor) => auditor !== selectedQuickAuditor,
      );
      setQuickAuditor(remainingAuditors[0] ?? "");
    }
  };
  const removeAuditor = (auditor: string) => {
    const nextAuditors = auditorOptions.filter((item) => item !== auditor);
    setAuditorOptions(nextAuditors);
  };
  return (
    <section className="panel people-admin">
      <div className="section-title">
        <h2>People / admin settings</h2>
        <button
          type="button"
          className="secondary"
          onClick={() =>
            setAuditorOptions([
              ...defaultAuditorOptions,
              ...auditorOptions.filter(
                (auditor) => !defaultAuditorOptions.includes(auditor),
              ),
            ])
          }
        >
          Restore default auditors
        </button>
      </div>
      <div className="admin-grid">
        <form onSubmit={addAuditor} className="add-person-form">
          <label>
            Add custom auditor
            <input
              value={newAuditor}
              placeholder="Type auditor name"
              onChange={(event) => setNewAuditor(event.target.value)}
            />
          </label>
          <button type="submit">Add</button>
        </form>
        <div className="quick-add-auditor">
          <label>
            Quick add default auditor
            <select
              value={selectedQuickAuditor}
              disabled={availableDefaultAuditors.length === 0}
              onChange={(event) => setQuickAuditor(event.target.value)}
            >
              {availableDefaultAuditors.length === 0 ? (
                <option value="">All default auditors are visible</option>
              ) : (
                availableDefaultAuditors.map((auditor) => (
                  <option key={auditor} value={auditor}>
                    {auditor}
                  </option>
                ))
              )}
            </select>
          </label>
          <button
            type="button"
            className="secondary"
            disabled={!selectedQuickAuditor}
            onClick={quickAddAuditor}
          >
            Add selected
          </button>
        </div>
        <label>
          My auditor for saved view
          <select
            value={myAuditor}
            onChange={(event) => setMyAuditor(event.target.value)}
          >
            {auditorOptions.map((auditor) => (
              <option key={auditor} value={auditor}>
                {auditor}
              </option>
            ))}
          </select>
        </label>
      </div>
      <div className="person-chips">
        {auditorOptions.map((auditor) => (
          <span className="person-chip" key={auditor}>
            {auditor}
            <button type="button" onClick={() => removeAuditor(auditor)}>
              ×
            </button>
          </span>
        ))}
      </div>
    </section>
  );
}

function ProjectTable({
  projects,
  onSelect,
}: {
  projects: AuditProject[];
  onSelect: (id: string) => void;
}) {
  return (
    <section className="panel">
      <div className="section-title">
        <h2>Table view</h2>
        <span>{projects.length} rows</span>
      </div>
      <div className="table-wrap">
        <table>
          <thead>
            <tr>
              <th>Assignment</th>
              <th>Type</th>
              <th>Audit entity</th>
              <th>Auditor</th>
              <th>Stage</th>
              <th>Source</th>
              <th>Quote</th>
              <th>Due</th>
              <th>Payment received</th>
              <th>Next action</th>
              <th>Blockers</th>
            </tr>
          </thead>
          <tbody>
            {projects.map((project) => {
              const due = dueLabel(project);
              const blockers = computedBlockers(project);
              return (
                <tr key={project.id} onClick={() => onSelect(project.id)}>
                  <td>
                    <strong>{project.assignmentNumber}</strong>
                    <span>{project.clientCoverholderCode}</span>
                  </td>
                  <td>{project.assignmentType}</td>
                  <td>{project.auditEntity || "—"}</td>
                  <td>{project.assignedAuditor}</td>
                  <td>{project.currentStage}</td>
                  <td>{project.assignmentSource}</td>
                  <td>{project.quoteStatus}</td>
                  <td>
                    <span className={`pill ${due.className}`}>{due.text}</span>
                  </td>
                  <td>{project.paymentReceived ? "Yes" : "No"}</td>
                  <td>{project.nextAction}</td>
                  <td>{blockers.length ? blockers.join("; ") : "—"}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </section>
  );
}

function FiltersPanel({
  filters,
  setFilters,
  auditors,
}: {
  filters: Filters;
  setFilters: (filters: Filters) => void;
  auditors: string[];
}) {
  return (
    <section className="panel">
      <div className="section-title">
        <h2>Filters</h2>
        <button
          className="link"
          onClick={() =>
            setFilters({
              auditor: "",
              stage: "",
              source: "",
              quoteStatus: "",
              dueDate: "",
              missingDocuments: false,
            })
          }
        >
          Clear
        </button>
      </div>
      <div className="filters">
        <Select
          label="Auditor"
          value={filters.auditor}
          options={auditors}
          onChange={(value) => setFilters({ ...filters, auditor: value })}
        />
        <Select
          label="Stage"
          value={filters.stage}
          options={stages}
          onChange={(value) => setFilters({ ...filters, stage: value })}
        />
        <Select
          label="Source"
          value={filters.source}
          options={["Email", "DAM"]}
          onChange={(value) => setFilters({ ...filters, source: value })}
        />
        <Select
          label="Quote status"
          value={filters.quoteStatus}
          options={["Not Started", "Drafting", "Sent", "Accepted", "Rejected"]}
          onChange={(value) => setFilters({ ...filters, quoteStatus: value })}
        />
        <Select
          label="Due date"
          value={filters.dueDate}
          options={[
            ["overdue", "Overdue"],
            ["dueSoon", "Due soon"],
          ]}
          onChange={(value) => setFilters({ ...filters, dueDate: value })}
        />
        <label className="checkbox">
          <input
            type="checkbox"
            checked={filters.missingDocuments}
            onChange={(event) =>
              setFilters({ ...filters, missingDocuments: event.target.checked })
            }
          />{" "}
          Missing documents
        </label>
      </div>
    </section>
  );
}

function Select({
  label,
  value,
  options,
  onChange,
  placeholder = "All",
}: {
  label: string;
  value: string;
  options: (string | [string, string])[];
  onChange: (value: string) => void;
  placeholder?: string;
}) {
  return (
    <label>
      {label}
      <select value={value} onChange={(event) => onChange(event.target.value)}>
        <option value="">{placeholder}</option>
        {options.map((option) => {
          const value = Array.isArray(option) ? option[0] : option;
          const text = Array.isArray(option) ? option[1] : option;
          return (
            <option key={value} value={value}>
              {text}
            </option>
          );
        })}
      </select>
    </label>
  );
}

function Kanban({
  projects,
  selectedId,
  onSelect,
  onMove,
}: {
  projects: AuditProject[];
  selectedId?: string;
  onSelect: (id: string) => void;
  onMove: (project: AuditProject, stage: Stage) => void;
}) {
  const handleDrop = (projectId: string, targetStage: Stage) => {
    const project = projects.find((item) => item.id === projectId);
    if (project) onMove(project, targetStage);
  };

  return (
    <section className="panel">
      <div className="section-title">
        <h2>Kanban by lifecycle stage</h2>
        <span>{projects.length} visible · drag cards between stages</span>
      </div>
      <div className="kanban">
        {stages.map((stage) => (
          <div
            className="column"
            key={stage}
            onDragOver={(event) => event.preventDefault()}
            onDrop={(event) => {
              event.preventDefault();
              handleDrop(event.dataTransfer.getData("text/plain"), stage);
            }}
          >
            <h3>{stage}</h3>
            {projects
              .filter((project) => project.currentStage === stage)
              .map((project) => {
                const due = dueLabel(project);
                return (
                  <article
                    draggable
                    className={`card ${selectedId === project.id ? "selected" : ""}`}
                    key={project.id}
                    onClick={() => onSelect(project.id)}
                    onDragStart={(event) => {
                      event.dataTransfer.setData("text/plain", project.id);
                      event.dataTransfer.effectAllowed = "move";
                    }}
                  >
                    <strong>{project.assignmentNumber}</strong>
                    <span>
                      {project.clientCoverholderCode} ·{" "}
                      {project.assignedAuditor}
                    </span>
                    <span className="pill muted">{project.assignmentType}</span>
                    {project.labels.map((label) => (
                      <span
                        className={`project-label mini ${label.toLowerCase().replace(/ /g, "-")}`}
                        key={label}
                      >
                        {label}
                      </span>
                    ))}
                    {project.comments.length > 0 && (
                      <span className="pill muted">
                        {project.comments.length} comment
                        {project.comments.length === 1 ? "" : "s"}
                      </span>
                    )}
                    <span className={`pill ${due.className}`}>{due.text}</span>
                    {computedBlockers(project).length > 0 && (
                      <span className="pill danger">Blocked</span>
                    )}
                    <select
                      value={project.currentStage}
                      onClick={(event) => event.stopPropagation()}
                      onChange={(event) =>
                        onMove(project, event.target.value as Stage)
                      }
                    >
                      {stages.map((target) => (
                        <option key={target} value={target}>
                          {target}
                        </option>
                      ))}
                    </select>
                  </article>
                );
              })}
          </div>
        ))}
      </div>
    </section>
  );
}

function ProjectDetail({
  project,
  onEdit,
  onMove,
  onAddComment,
  onToggleChecklist,
  onDocumentWorkflowAction,
}: {
  project: AuditProject;
  onEdit: () => void;
  onMove: (project: AuditProject, stage: Stage) => void;
  onAddComment: (project: AuditProject, comment: ProjectComment) => void;
  onToggleChecklist: (project: AuditProject, key: string) => void;
  onDocumentWorkflowAction: (project: AuditProject, action: DocumentWorkflowAction) => void;
}) {
  const blockers = computedBlockers(project);
  return (
    <section className="detail-grid">
      <article className="panel detail">
        <div className="section-title">
          <h2>{project.assignmentNumber}</h2>
          <button onClick={onEdit}>Edit project</button>
        </div>
        {project.labels.length > 0 && (
          <div className="label-strip">
            {project.labels.map((label) => (
              <span
                className={`project-label ${label.toLowerCase().replace(/ /g, "-")}`}
                key={label}
              >
                {label}
              </span>
            ))}
          </div>
        )}
        <div className="meta-grid">
          <Meta label="Source" value={project.assignmentSource} />
          <Meta label="Assignment type" value={project.assignmentType} />
          <Meta label="Audit entity" value={project.auditEntity || "Not set"} />
          <Meta
            label="Client / coverholder code"
            value={project.clientCoverholderCode}
          />
          <Meta label="Broker" value={project.broker} />
          <Meta label="Auditor" value={project.assignedAuditor} />
          <Meta label="Reviewer" value={project.reviewer} />
          <Meta label="Status" value={project.assignmentStatus} />
          <Meta
            label="Quote"
            value={`${project.quoteStatus} · ${project.quoteAmount.toLocaleString("en-US", { style: "currency", currency: "USD", maximumFractionDigits: 0 })}`}
          />
          <Meta
            label="Audit timing"
            value={`${project.tentativeAuditWeek || "No week"} · ${project.confirmedAuditDate || "No date"}`}
          />
          <Meta label="Audit type" value={project.auditType} />
          <Meta
            label="Payment received"
            value={project.paymentReceived ? "Yes" : "No"}
          />
          <Meta label="Last updated" value={project.lastUpdatedDate} />
        </div>
        <h3>Next action</h3>
        <p>{project.nextAction || "No next action recorded."}</p>
        <h3>Blockers</h3>
        {blockers.length ? (
          <ul className="blockers">
            {blockers.map((blocker) => (
              <li key={blocker}>{blocker}</li>
            ))}
          </ul>
        ) : (
          <p>No blockers recorded.</p>
        )}
        <div className="move-row">
          <label>
            Move stage
            <select
              value={project.currentStage}
              onChange={(event) => onMove(project, event.target.value as Stage)}
            >
              {stages.map((stage) => (
                <option key={stage}>{stage}</option>
              ))}
            </select>
          </label>
        </div>
      </article>
      <DocumentReadiness
        project={project}
        onDocumentWorkflowAction={onDocumentWorkflowAction}
      />
      <Checklist project={project} onToggleChecklist={onToggleChecklist} />
      <Comments project={project} onAddComment={onAddComment} />
      <ActivityTimeline project={project} />
    </section>
  );
}

function Meta({ label, value }: { label: string; value: string }) {
  return (
    <div>
      <span>{label}</span>
      <strong>{value}</strong>
    </div>
  );
}

function DocumentReadiness({
  project,
  onDocumentWorkflowAction,
}: {
  project: AuditProject;
  onDocumentWorkflowAction: (project: AuditProject, action: DocumentWorkflowAction) => void;
}) {
  const readiness = documentReadiness(project);
  return (
    <article className="panel document-workflow">
      <div className="section-title">
        <div>
          <h2>Document readiness</h2>
          <span>
            {readiness.completeCount}/{readiness.totalCount} readiness items complete
          </span>
        </div>
        <span className={`readiness-score ${readiness.percent === 100 ? "ok" : "warning"}`}>
          {readiness.percent}%
        </span>
      </div>
      <div className="readiness-track" aria-hidden="true">
        <span style={{ width: `${readiness.percent}%` }} />
      </div>
      <div className="document-status-grid">
        {requiredDocuments.map((doc) => (
          <span
            key={doc.key}
            className={`document-chip ${project[doc.key] ? "complete" : "missing"}`}
          >
            {project[doc.key] ? "✓" : "•"} {doc.label}
          </span>
        ))}
        <span
          className={`document-chip ${
            project.preAuditQuestionnaireStatus === "Complete" ? "complete" : "missing"
          }`}
        >
          Questionnaire: {project.preAuditQuestionnaireStatus}
        </span>
        <span
          className={`document-chip ${
            project.documentRequestStatus === "Complete" ? "complete" : "missing"
          }`}
        >
          Request: {project.documentRequestStatus}
        </span>
      </div>
      <div className="document-dates">
        <Meta label="Requested" value={project.documentRequestDate || "Not sent"} />
        <Meta label="Last chased" value={project.brokerLastChasedDate || "Not chased"} />
        <Meta
          label="Expected response"
          value={project.brokerExpectedResponseDate || "Not set"}
        />
      </div>
      {readiness.missingDocuments.length > 0 ? (
        <p className="readiness-note">
          Missing: {readiness.missingDocuments.join(", ")}
        </p>
      ) : (
        <p className="readiness-note ok-text">All core documents are marked received.</p>
      )}
      <div className="document-actions">
        <button
          type="button"
          className="secondary"
          onClick={() =>
            onDocumentWorkflowAction(project, "markWaitingOnBroker")
          }
        >
          Mark waiting on broker
        </button>
        <button
          type="button"
          className="secondary"
          onClick={() => onDocumentWorkflowAction(project, "recordBrokerChase")}
        >
          Record broker chase
        </button>
        <button
          type="button"
          onClick={() =>
            onDocumentWorkflowAction(project, "markDocumentsComplete")
          }
        >
          Mark documents complete
        </button>
      </div>
    </article>
  );
}

function ActivityTimeline({ project }: { project: AuditProject }) {
  const items = activityTimeline(project);
  return (
    <article className="panel activity-panel">
      <div className="section-title">
        <div>
          <h2>Activity timeline</h2>
          <span>Comments, stage moves, checklist, and document follow-ups</span>
        </div>
      </div>
      {items.length === 0 ? (
        <p>No activity yet.</p>
      ) : (
        <div className="activity-list">
          {items.map((item) => (
            <div className={`activity-item ${item.type}`} key={item.id}>
              <span className={`activity-type ${item.tone || "muted"}`}>
                {item.type}
              </span>
              <strong>{item.title}</strong>
              <small>{item.timestamp}</small>
              <p>{item.detail}</p>
            </div>
          ))}
        </div>
      )}
    </article>
  );
}

function Checklist({
  project,
  onToggleChecklist,
}: {
  project: AuditProject;
  onToggleChecklist: (project: AuditProject, key: string) => void;
}) {
  const sourceSpecific = sourceTasks(project);
  const checklistItems = [
    ...checklistByStage[project.currentStage],
    ...sourceSpecific,
  ];
  const completedCount = checklistItems.filter(
    (item) =>
      project.checklistCompletions[checklistKey(project.currentStage, item)],
  ).length;
  return (
    <article className="panel">
      <div className="section-title">
        <div>
          <h2>Stage checklist</h2>
          <span>
            {completedCount}/{checklistItems.length} complete for{" "}
            {project.currentStage}
          </span>
        </div>
      </div>
      <ul className="checklist interactive-checklist">
        {checklistItems.map((item) => {
          const key = checklistKey(project.currentStage, item);
          return (
            <li
              key={key}
              className={project.checklistCompletions[key] ? "done" : ""}
            >
              <label className="checklist-toggle">
                <input
                  type="checkbox"
                  checked={Boolean(project.checklistCompletions[key])}
                  onChange={() => onToggleChecklist(project, key)}
                />
                <span
                  className={sourceSpecific.includes(item) ? "conditional" : ""}
                >
                  {item}
                </span>
              </label>
            </li>
          );
        })}
      </ul>
      <h3>Document readiness</h3>
      <ul className="document-list">
        {requiredDocuments.map((doc) => (
          <li
            key={doc.key}
            className={project[doc.key] ? "complete" : "missing"}
          >
            {doc.label}
          </li>
        ))}
        <li
          className={
            project.preAuditQuestionnaireStatus === "Complete"
              ? "complete"
              : "missing"
          }
        >
          Pre-audit questionnaire: {project.preAuditQuestionnaireStatus}
        </li>
        <li
          className={
            project.documentRequestStatus === "Complete"
              ? "complete"
              : "missing"
          }
        >
          Document request: {project.documentRequestStatus}
        </li>
      </ul>
    </article>
  );
}

function Comments({
  project,
  onAddComment,
}: {
  project: AuditProject;
  onAddComment: (project: AuditProject, comment: ProjectComment) => void;
}) {
  const [commentBody, setCommentBody] = useState("");
  const [commentAuthor, setCommentAuthor] = useState("Prototype user");
  const addComment = (event: FormEvent) => {
    event.preventDefault();
    const body = commentBody.trim();
    if (!body) return;
    onAddComment(project, {
      id: `comment-${Date.now()}`,
      createdAt: timestampNow(),
      author: commentAuthor.trim() || "Prototype user",
      body,
    });
    setCommentBody("");
  };
  return (
    <article className="panel comments-panel">
      <h2>Card comments</h2>
      <form className="comment-form" onSubmit={addComment}>
        <label>
          Name
          <input
            value={commentAuthor}
            onChange={(event) => setCommentAuthor(event.target.value)}
          />
        </label>
        <label>
          Comment
          <textarea
            value={commentBody}
            placeholder="Add an update, question, or note for this audit card"
            onChange={(event) => setCommentBody(event.target.value)}
          />
        </label>
        <button type="submit">Add comment</button>
      </form>
      <div className="comment-list">
        {project.comments.length === 0 ? (
          <p>No comments yet.</p>
        ) : (
          project.comments
            .slice()
            .reverse()
            .map((comment) => (
              <div className="comment" key={comment.id}>
                <strong>{comment.author}</strong>
                <span>{comment.createdAt}</span>
                <p>{comment.body}</p>
              </div>
            ))
        )}
      </div>
    </article>
  );
}

function History({ project }: { project: AuditProject }) {
  return (
    <article className="panel history">
      <h2>Status history</h2>
      {project.statusHistory.length === 0 ? (
        <p>No stage changes yet.</p>
      ) : (
        project.statusHistory
          .slice()
          .reverse()
          .map((item) => (
            <div key={item.id}>
              <strong>
                {item.fromStage} → {item.toStage}
              </strong>
              <span>
                {item.changedAt} · {item.changedBy}
              </span>
              <p>{item.note}</p>
            </div>
          ))
      )}
    </article>
  );
}

function ProjectForm({
  project,
  onSave,
  onCancel,
  auditorOptions,
}: {
  project: AuditProject;
  onSave: (project: AuditProject) => void;
  onCancel: () => void;
  auditorOptions: string[];
}) {
  const [draft, setDraft] = useState(project);
  const [step, setStep] = useState(0);
  const isNewProject = project.statusHistory.length === 0;
  const steps = [
    "Assignment basics",
    "People",
    "Planning",
    "Documents & quote",
    "Review",
  ];
  const update = <K extends keyof AuditProject>(
    key: K,
    value: AuditProject[K],
  ) => setDraft({ ...draft, [key]: value });
  const toggleLabel = (label: ProjectLabel) => {
    update(
      "labels",
      (draft.labels.includes(label)
        ? draft.labels.filter((item) => item !== label)
        : [...draft.labels, label]) as AuditProject["labels"],
    );
  };
  const submit = (event: FormEvent) => {
    event.preventDefault();
    if (isNewProject && step < steps.length - 1) {
      setStep(step + 1);
      return;
    }
    onSave(draft);
  };

  const basics = (
    <section className="form-section-card">
      <div>
        <h3>Assignment basics</h3>
        <p>
          Capture the intake facts first. This is the minimum record users need
          to start tracking the audit.
        </p>
      </div>
      <div className="form-grid wizard-grid">
        <Input
          label="Assignment number"
          value={draft.assignmentNumber}
          onChange={(value) => update("assignmentNumber", value)}
        />
        <Select
          label="Source"
          value={draft.assignmentSource}
          options={["Email", "DAM"]}
          placeholder="Select source"
          onChange={(value) =>
            update("assignmentSource", value as AssignmentSource)
          }
        />
        <Select
          label="Assignment type"
          value={draft.assignmentType}
          options={assignmentTypeOptions}
          placeholder="Select type"
          onChange={(value) =>
            update("assignmentType", value as AssignmentType)
          }
        />
        <Input
          label="Audit Entity"
          value={draft.auditEntity}
          onChange={(value) => update("auditEntity", value)}
        />
        <Input
          label="Client / coverholder code"
          value={draft.clientCoverholderCode}
          onChange={(value) => update("clientCoverholderCode", value)}
        />
        <Input
          label="Broker"
          value={draft.broker}
          onChange={(value) => update("broker", value)}
        />
      </div>
      <div className="label-picker">
        <span>Labels / tags</span>
        <div>
          {labelOptions.map((label) => (
            <button
              type="button"
              key={label}
              className={`label-option ${label.toLowerCase().replace(/ /g, "-")}${
                draft.labels.includes(label) ? " selected" : ""
              }`}
              onClick={() => toggleLabel(label)}
            >
              {label}
            </button>
          ))}
        </div>
      </div>
    </section>
  );

  const people = (
    <section className="form-section-card">
      <div>
        <h3>People</h3>
        <p>
          Assign ownership so the card immediately appears in workload and saved
          views.
        </p>
      </div>
      <div className="form-grid wizard-grid">
        <Select
          label="Assigned auditor"
          value={draft.assignedAuditor}
          options={auditorOptions}
          placeholder="Select auditor"
          onChange={(value) => update("assignedAuditor", value)}
        />
        <Input
          label="Reviewer"
          value={draft.reviewer}
          onChange={(value) => update("reviewer", value)}
        />
        <Select
          label="Assignment status"
          value={draft.assignmentStatus}
          options={["New", "In Progress", "Blocked", "On Hold", "Completed"]}
          placeholder="Select status"
          onChange={(value) =>
            update("assignmentStatus", value as AssignmentStatus)
          }
        />
        <Input
          label="Due date"
          type="date"
          value={draft.dueDate}
          onChange={(value) => update("dueDate", value)}
        />
      </div>
    </section>
  );

  const planning = (
    <section className="form-section-card">
      <div>
        <h3>Planning</h3>
        <p>
          Schedule the audit and choose how fieldwork will happen. Exact dates
          can be filled in later.
        </p>
      </div>
      <div className="form-grid wizard-grid">
        <Input
          label="Tentative audit week"
          value={draft.tentativeAuditWeek}
          placeholder="2026-W21"
          onChange={(value) => update("tentativeAuditWeek", value)}
        />
        <Input
          label="Confirmed audit date"
          type="date"
          value={draft.confirmedAuditDate}
          onChange={(value) => update("confirmedAuditDate", value)}
        />
        <Select
          label="Audit type"
          value={draft.auditType}
          options={["Remote", "Onsite"]}
          placeholder="Select audit type"
          onChange={(value) => update("auditType", value as AuditType)}
        />
        <Select
          label="Current stage"
          value={draft.currentStage}
          options={stages}
          placeholder="Select stage"
          onChange={(value) => update("currentStage", value as Stage)}
        />
      </div>
    </section>
  );

  const documentsQuote = (
    <section className="form-section-card">
      <div>
        <h3>Documents & quote</h3>
        <p>
          Only capture early-stage readiness here. Later report and invoice
          fields stay hidden until edit mode.
        </p>
      </div>
      <div className="form-grid wizard-grid">
        <Select
          label="Quote status"
          value={draft.quoteStatus}
          options={["Not Started", "Drafting", "Sent", "Accepted", "Rejected"]}
          placeholder="Select quote status"
          onChange={(value) => update("quoteStatus", value as QuoteStatus)}
        />
        <Input
          label="Quote amount"
          value={String(draft.quoteAmount)}
          type="number"
          onChange={(value) => update("quoteAmount", Number(value))}
        />
        <Check
          label="BAA received"
          checked={draft.baaReceived}
          onChange={(value) => update("baaReceived", value)}
        />
        <Check
          label="Endorsements received"
          checked={draft.endorsementsReceived}
          onChange={(value) => update("endorsementsReceived", value)}
        />
        <Check
          label="Premium BDX received"
          checked={draft.premiumBdxReceived}
          onChange={(value) => update("premiumBdxReceived", value)}
        />
        <Select
          label="Pre-audit questionnaire"
          value={draft.preAuditQuestionnaireStatus}
          options={["Not Started", "In Progress", "Complete", "Not Required"]}
          placeholder="Select questionnaire status"
          onChange={(value) =>
            update("preAuditQuestionnaireStatus", value as ProgressStatus)
          }
        />
        <Select
          label="Document request"
          value={draft.documentRequestStatus}
          options={["Not Started", "In Progress", "Complete", "Not Required"]}
          placeholder="Select document status"
          onChange={(value) =>
            update("documentRequestStatus", value as ProgressStatus)
          }
        />
        <Input
          label="Document request date"
          type="date"
          value={draft.documentRequestDate}
          onChange={(value) => update("documentRequestDate", value)}
        />
        <Input
          label="Broker last chased"
          type="date"
          value={draft.brokerLastChasedDate}
          onChange={(value) => update("brokerLastChasedDate", value)}
        />
        <Input
          label="Expected broker response"
          type="date"
          value={draft.brokerExpectedResponseDate}
          onChange={(value) => update("brokerExpectedResponseDate", value)}
        />
      </div>
      <label>
        Next action
        <textarea
          value={draft.nextAction}
          onChange={(event) => update("nextAction", event.target.value)}
        />
      </label>
      <label>
        Manual blockers
        <textarea
          value={draft.blockers}
          onChange={(event) => update("blockers", event.target.value)}
        />
      </label>
    </section>
  );

  const review = (
    <section className="form-section-card review-card">
      <div>
        <h3>Review project before creating</h3>
        <p>
          This card will start in {draft.currentStage}. You can add advanced
          report, findings, and invoice details later from Edit Project.
        </p>
      </div>
      <div className="review-grid">
        <Meta label="Assignment" value={draft.assignmentNumber} />
        <Meta
          label="Source / type"
          value={`${draft.assignmentSource} · ${draft.assignmentType}`}
        />
        <Meta label="Audit Entity" value={draft.auditEntity || "Not set"} />
        <Meta label="Auditor" value={draft.assignedAuditor || "Not assigned"} />
        <Meta label="Reviewer" value={draft.reviewer || "Not assigned"} />
        <Meta label="Due date" value={draft.dueDate || "Not set"} />
        <Meta
          label="Quote"
          value={`${draft.quoteStatus} · ${draft.quoteAmount || 0}`}
        />
        <Meta
          label="Labels"
          value={draft.labels.length ? draft.labels.join(", ") : "None"}
        />
      </div>
    </section>
  );

  const advanced = (
    <section className="form-section-card">
      <div>
        <h3>Advanced lifecycle fields</h3>
        <p>
          These fields are shown in edit mode because they become relevant later
          in the audit lifecycle.
        </p>
      </div>
      <div className="form-grid wizard-grid">
        <Check
          label="File selection completed"
          checked={draft.fileSelectionCompleted}
          onChange={(value) => update("fileSelectionCompleted", value)}
        />
        <Check
          label="Testing sheet completed"
          checked={draft.testingSheetCompleted}
          onChange={(value) => update("testingSheetCompleted", value)}
        />
        <Input
          label="Findings sent date"
          type="date"
          value={draft.findingsSentDate}
          onChange={(value) => update("findingsSentDate", value)}
        />
        <Input
          label="Coverholder response date"
          type="date"
          value={draft.coverholderResponseReceivedDate}
          onChange={(value) => update("coverholderResponseReceivedDate", value)}
        />
        <Select
          label="Report status"
          value={draft.reportStatus}
          options={["Not Started", "Drafting", "Review", "Issued"]}
          placeholder="Select report status"
          onChange={(value) => update("reportStatus", value as ReportStatus)}
        />
        <Select
          label="Invoice status"
          value={draft.invoiceStatus}
          options={["Not Started", "Prepared", "Sent", "Paid"]}
          placeholder="Select invoice status"
          onChange={(value) => {
            update("invoiceStatus", value as InvoiceStatus);
            if (value === "Paid") update("paymentReceived", true);
          }}
        />
        <Check
          label="Payment received"
          checked={draft.paymentReceived}
          onChange={(value) => update("paymentReceived", value)}
        />
        {draft.assignmentSource === "DAM" && (
          <Select
            label="DAM submission"
            value={draft.damSubmissionStatus}
            options={["Not Started", "Submitted", "Accepted"]}
            placeholder="Select DAM status"
            onChange={(value) =>
              update("damSubmissionStatus", value as DamSubmissionStatus)
            }
          />
        )}
      </div>
    </section>
  );

  const createStepContent = [basics, people, planning, documentsQuote, review];

  return (
    <div className="modal-backdrop">
      <form className="modal intake-modal" onSubmit={submit}>
        <div className="section-title modal-title-row">
          <div>
            <p className="eyebrow dark">
              {isNewProject ? "Guided intake" : "Advanced edit"}
            </p>
            <h2>{isNewProject ? "Add project" : "Edit project"}</h2>
          </div>
          <button type="button" className="link" onClick={onCancel}>
            Close
          </button>
        </div>
        {isNewProject ? (
          <>
            <div className="wizard-steps" aria-label="Add project steps">
              {steps.map((stepLabel, index) => (
                <button
                  type="button"
                  key={stepLabel}
                  className={
                    index === step ? "active" : index < step ? "complete" : ""
                  }
                  onClick={() => setStep(index)}
                >
                  <span>{index + 1}</span>
                  {stepLabel}
                </button>
              ))}
            </div>
            {createStepContent[step]}
          </>
        ) : (
          <>
            {basics}
            {people}
            {planning}
            {documentsQuote}
            {advanced}
          </>
        )}
        <div className="modal-actions sticky-actions">
          <button type="button" className="secondary" onClick={onCancel}>
            Cancel
          </button>
          {isNewProject && step > 0 && (
            <button
              type="button"
              className="secondary"
              onClick={() => setStep(step - 1)}
            >
              Back
            </button>
          )}
          <button type="submit">
            {isNewProject && step < steps.length - 1 ? "Next" : "Save project"}
          </button>
        </div>
      </form>
    </div>
  );
}

function Input({
  label,
  value,
  onChange,
  type = "text",
  placeholder = "",
}: {
  label: string;
  value: string;
  onChange: (value: string) => void;
  type?: string;
  placeholder?: string;
}) {
  return (
    <label>
      {label}
      <input
        type={type}
        value={value}
        placeholder={placeholder}
        onChange={(event) => onChange(event.target.value)}
      />
    </label>
  );
}

function Check({
  label,
  checked,
  onChange,
}: {
  label: string;
  checked: boolean;
  onChange: (value: boolean) => void;
}) {
  return (
    <label className="checkbox form-check">
      <input
        type="checkbox"
        checked={checked}
        onChange={(event) => onChange(event.target.checked)}
      />{" "}
      {label}
    </label>
  );
}

const rootElement = typeof document !== "undefined" ? document.getElementById("root") : null;

if (rootElement) {
  createRoot(rootElement).render(<App />);
}
