import React, { FormEvent, useEffect, useMemo, useState } from "react";
import { createRoot } from "react-dom/client";
import {
  approveAccessRequest,
  buildAccessRequestUser,
  canApproveAccessRequest,
  isValidCompanyEmail,
  rejectAccessRequest,
  verifyAccessRequestEmail,
  type AccountRequestStatus,
} from "./accessRequests";
import {
  approveSecureAccessRequest,
  getSecureAccessState,
  logoutSecureAccess,
  rejectSecureAccessRequest,
  secureAccessUrl,
  verifySecureAccessCode,
  type SecureAccessState,
  type SecureAccessUser,
} from "./secureAccessClient";
import {
  hasMicrosoftAuthConfig,
  microsoftAuthScopeLabel,
  refreshMicrosoftGraphToken,
  restoreMicrosoftGraphSession,
  sanitizeMicrosoftAuthConfig,
  signInWithMicrosoft,
  signOutOfMicrosoft,
  type MicrosoftAuthAccount,
  type MicrosoftAuthConfig,
} from "./microsoftAuth";
import {
  buildSyncPackage,
  hasFullMicrosoftListsConfig,
  hasMinimumMicrosoftListsConfig,
  microsoftListLabels,
  missingMicrosoftListLabels,
  pullAssignmentRowsFromMicrosoftLists,
  pushMigrationPackageToMicrosoftLists,
  requiredMicrosoftListKeys,
  sanitizeMicrosoftListsConfig,
  testMicrosoftListsConnection,
  type MicrosoftListKey,
  type MicrosoftListsConnectionConfig,
} from "./microsoftListsClient";
import {
  buildMicrosoftListsMigrationPackage,
  microsoftListSchemas,
} from "./microsoftListsSchema";
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

export type AuditActivityEvent = {
  id: string;
  createdAt: string;
  actor: string;
  type: "team" | "field" | "document" | "checklist" | "stage";
  title: string;
  detail: string;
};

export type AuditTeamRole = "Lead Auditor" | "Supporting Auditor";

export type AuditTeamMember = {
  person: string;
  role: AuditTeamRole;
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
  auditTeam: AuditTeamMember[];
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
  activityEvents: AuditActivityEvent[];
};

type Filters = {
  auditor: string;
  stage: string;
  source: string;
  quoteStatus: string;
  dueDate: string;
  missingDocuments: boolean;
  workState: string;
};

type FilterPreset = {
  id: string;
  label: string;
  filters: Partial<Filters>;
};

type ViewMode = "kanban" | "table";
type StorageMode = "local" | "microsoft-lists";
type UserRole = "Admin" | "Audit Manager" | "Auditor" | "Finance" | "Read Only";
type ProjectVisibility =
  | "Role Default"
  | "All Projects"
  | "Assigned Projects"
  | "Finance Records";

type PrototypeUser = {
  fullName: string;
  username: string;
  password: string;
  role: UserRole;
  permissionGroup: UserRole;
  email: string;
  active: boolean;
  defaultVisibility: ProjectVisibility;
  emailVerified: boolean;
  accessRequestStatus: AccountRequestStatus;
  verificationCode: string;
  requestedAt: string;
  approvedAt: string;
  approvedBy: string;
  rejectionReason: string;
};

type ConfirmationRequest = {
  title: string;
  message: string;
  confirmLabel: string;
  tone?: "default" | "danger";
  onConfirm: () => void | Promise<void>;
};

type CommunicationTemplate = {
  id: string;
  label: string;
  kind: "Email" | "Document";
  purpose: string;
  subject: (project: AuditProject) => string;
  body: (project: AuditProject) => string;
};

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
const currentUserStorageKey = "audit-assignment-tracker-current-user-v1";
const usersStorageKey = "audit-assignment-tracker-users-v1";
const lastExportStorageKey = "audit-assignment-tracker-last-export-v1";
const microsoftListsConfigStorageKey =
  "audit-assignment-tracker-microsoft-lists-config-v1";
const microsoftAuthConfigStorageKey =
  "audit-assignment-tracker-microsoft-auth-config-v1";
const storageModeStorageKey = "audit-assignment-tracker-storage-mode-v1";

const defaultFilters: Filters = {
  auditor: "",
  stage: "",
  source: "",
  quoteStatus: "",
  dueDate: "",
  missingDocuments: false,
  workState: "",
};

const userRoleOptions: UserRole[] = [
  "Admin",
  "Audit Manager",
  "Auditor",
  "Finance",
  "Read Only",
];

const projectVisibilityOptions: ProjectVisibility[] = [
  "Role Default",
  "All Projects",
  "Assigned Projects",
  "Finance Records",
];

const assignmentTypeOptions: AssignmentType[] = [
  "DCA",
  "CH",
  "MGA",
  "Company Contract",
];

const assignmentStatusOptions: AssignmentStatus[] = [
  "New",
  "In Progress",
  "Blocked",
  "On Hold",
  "Completed",
];
const quoteStatusOptions: QuoteStatus[] = [
  "Not Started",
  "Drafting",
  "Sent",
  "Accepted",
  "Rejected",
];
const progressStatusOptions: ProgressStatus[] = [
  "Not Started",
  "In Progress",
  "Complete",
  "Not Required",
];
const reportStatusOptions: ReportStatus[] = [
  "Not Started",
  "Drafting",
  "Review",
  "Issued",
];
const invoiceStatusOptions: InvoiceStatus[] = [
  "Not Started",
  "Prepared",
  "Sent",
  "Paid",
];
const damSubmissionStatusOptions: DamSubmissionStatus[] = [
  "Not Required",
  "Not Started",
  "Submitted",
  "Accepted",
];

const labelOptions: ProjectLabel[] = [
  "High Priority",
  "Medium Priority",
  "Low Priority",
  "Waiting on Broker",
];

function approvedPrototypeUser(
  user: Omit<
    PrototypeUser,
    | "emailVerified"
    | "accessRequestStatus"
    | "verificationCode"
    | "requestedAt"
    | "approvedAt"
    | "approvedBy"
    | "rejectionReason"
  >,
): PrototypeUser {
  return {
    ...user,
    emailVerified: true,
    accessRequestStatus: "Approved",
    verificationCode: "",
    requestedAt: "2026-05-01T12:00:00.000Z",
    approvedAt: "2026-05-01T12:00:00.000Z",
    approvedBy: "System seed",
    rejectionReason: "",
  };
}

const defaultPrototypeUsers: PrototypeUser[] = [
  approvedPrototypeUser({
    fullName: "Lorraine Mojica",
    username: "lorraine.mojica",
    password: "password",
    role: "Auditor",
    permissionGroup: "Auditor",
    email: "lorraine.mojica@[company-domain]",
    active: true,
    defaultVisibility: "Role Default",
  }),
  approvedPrototypeUser({
    fullName: "Walter Aviles",
    username: "walter.aviles",
    password: "password",
    role: "Auditor",
    permissionGroup: "Auditor",
    email: "walter.aviles@[company-domain]",
    active: true,
    defaultVisibility: "Role Default",
  }),
  approvedPrototypeUser({
    fullName: "Leslie Domenech",
    username: "leslie.domenech",
    password: "password",
    role: "Auditor",
    permissionGroup: "Auditor",
    email: "leslie.domenech@[company-domain]",
    active: true,
    defaultVisibility: "Role Default",
  }),
  approvedPrototypeUser({
    fullName: "Mark James",
    username: "mark.james",
    password: "password",
    role: "Audit Manager",
    permissionGroup: "Audit Manager",
    email: "mark.james@[company-domain]",
    active: true,
    defaultVisibility: "Role Default",
  }),
  approvedPrototypeUser({
    fullName: "Justin Mojica",
    username: "justin.mojica",
    password: "password",
    role: "Admin",
    permissionGroup: "Admin",
    email: "justin.mojica@[company-domain]",
    active: true,
    defaultVisibility: "Role Default",
  }),
  approvedPrototypeUser({
    fullName: "Sheilah Couture",
    username: "sheilah.couture",
    password: "password",
    role: "Finance",
    permissionGroup: "Finance",
    email: "sheilah.couture@[company-domain]",
    active: true,
    defaultVisibility: "Role Default",
  }),
  approvedPrototypeUser({
    fullName: "Annabelle J. Crawford Mojica",
    username: "annabelle.crawford.mojica",
    password: "password",
    role: "Read Only",
    permissionGroup: "Read Only",
    email: "annabelle.crawford.mojica@[company-domain]",
    active: true,
    defaultVisibility: "Role Default",
  }),
  approvedPrototypeUser({
    fullName: "Molly Aviles",
    username: "molly.aviles",
    password: "password",
    role: "Auditor",
    permissionGroup: "Auditor",
    email: "molly.aviles@[company-domain]",
    active: true,
    defaultVisibility: "Role Default",
  }),
  approvedPrototypeUser({
    fullName: "Lindsie Guillermo",
    username: "lindsie.guillermo",
    password: "password",
    role: "Auditor",
    permissionGroup: "Auditor",
    email: "lindsie.guillermo@[company-domain]",
    active: true,
    defaultVisibility: "Role Default",
  }),
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
    auditTeam: [
      { person: "Lorraine Mojica", role: "Lead Auditor" },
      { person: "Walter Aviles", role: "Supporting Auditor" },
    ],
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
    activityEvents: [],
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
    auditTeam: [
      { person: "Walter Aviles", role: "Lead Auditor" },
      { person: "Leslie Domenech", role: "Supporting Auditor" },
    ],
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
    activityEvents: [],
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
    auditTeam: [{ person: "Lorraine Mojica", role: "Lead Auditor" }],
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
    activityEvents: [],
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
    auditTeam: [{ person: "Justin Mojica", role: "Lead Auditor" }],
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
    activityEvents: [],
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
  auditTeam: [],
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
  activityEvents: [],
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

function normalizeAuditTeam(project: AuditProject): AuditTeamMember[] {
  const seededTeam: AuditTeamMember[] = (project.auditTeam ?? [])
    .filter((member) => member.person.trim())
    .map((member) => ({
      person: member.person.trim(),
      role:
        member.role === "Lead Auditor"
          ? "Lead Auditor"
          : "Supporting Auditor",
    }));
  const fallbackTeam =
    seededTeam.length > 0
      ? seededTeam
      : project.assignedAuditor
        ? [
            {
              person: project.assignedAuditor.trim(),
              role: "Lead Auditor" as AuditTeamRole,
            },
          ]
        : [];
  const seen = new Set<string>();
  const uniqueTeam = fallbackTeam.filter((member) => {
    if (seen.has(member.person)) return false;
    seen.add(member.person);
    return true;
  });
  if (
    uniqueTeam.length > 0 &&
    !uniqueTeam.some((member) => member.role === "Lead Auditor")
  ) {
    return [
      { ...uniqueTeam[0], role: "Lead Auditor" as AuditTeamRole },
      ...uniqueTeam.slice(1),
    ];
  }
  return uniqueTeam;
}

function primaryAuditor(project: AuditProject) {
  const team = normalizeAuditTeam(project);
  return (
    team.find((member) => member.role === "Lead Auditor")?.person ??
    team[0]?.person ??
    project.assignedAuditor ??
    ""
  );
}

function assignedAuditorNames(project: AuditProject) {
  return normalizeAuditTeam(project).map((member) => member.person);
}

function projectHasAuditor(project: AuditProject, auditor: string) {
  return assignedAuditorNames(project).includes(auditor);
}

function formatAuditTeam(project: AuditProject) {
  const team = normalizeAuditTeam(project);
  if (team.length === 0) return "Not assigned";
  return team
    .map((member) =>
      member.role === "Lead Auditor"
        ? `${member.person} (Lead)`
        : `${member.person} (Support)`,
    )
    .join(", ");
}

function auditTeamRole(project: AuditProject, auditor: string) {
  return normalizeAuditTeam(project).find((member) => member.person === auditor)?.role;
}

function hasFullProjectAccess(user: PrototypeUser) {
  return user.role === "Admin" || user.role === "Audit Manager";
}

function isFinanceProject(project: AuditProject) {
  return (
    project.currentStage === "Final Submission" ||
    project.currentStage === "Invoice" ||
    project.currentStage === "Closed" ||
    project.invoiceStatus !== "Not Started" ||
    project.reportStatus === "Issued"
  );
}

function roleDefaultVisibility(role: UserRole): ProjectVisibility {
  if (role === "Finance") return "Finance Records";
  if (role === "Auditor") return "Assigned Projects";
  return "All Projects";
}

function effectiveVisibility(user: PrototypeUser) {
  return user.defaultVisibility === "Role Default"
    ? roleDefaultVisibility(user.role)
    : user.defaultVisibility;
}

function canViewProject(user: PrototypeUser, project: AuditProject) {
  if (!user.active) return false;
  const visibility = effectiveVisibility(user);
  if (visibility === "All Projects") return true;
  if (visibility === "Finance Records") return isFinanceProject(project);
  return (
    projectHasAuditor(project, user.fullName) ||
    project.reviewer === user.fullName
  );
}

function canEditProject(user: PrototypeUser, project: AuditProject) {
  if (hasFullProjectAccess(user)) return true;
  return user.role === "Auditor" && projectHasAuditor(project, user.fullName);
}

function canCreateProject(user: PrototypeUser) {
  return user.role === "Admin" || user.role === "Audit Manager";
}

function canUpdateFinance(user: PrototypeUser, project: AuditProject) {
  return user.role === "Finance" && isFinanceProject(project);
}

function canComment(user: PrototypeUser, project: AuditProject) {
  return user.role !== "Read Only" && canViewProject(user, project);
}

function visibleProjectMessage(user: PrototypeUser) {
  const visibility = effectiveVisibility(user);
  if (visibility === "All Projects") {
    return "Showing all projects.";
  }
  if (visibility === "Finance Records") {
    return "Showing invoice, final submission, and close-out records.";
  }
  return "Showing projects where you are the lead or supporting auditor.";
}

function projectMatchesWorkState(project: AuditProject, workState: string) {
  if (workState === "blocked") {
    return (
      computedBlockers(project).length > 0 ||
      project.assignmentStatus === "Blocked"
    );
  }
  if (workState === "waitingOnBroker") {
    return project.labels.includes("Waiting on Broker");
  }
  if (workState === "pendingPayment") {
    return isFinanceProject(project) && !project.paymentReceived;
  }
  return true;
}

function buildFilters(overrides: Partial<Filters> = {}): Filters {
  return { ...defaultFilters, ...overrides };
}

function filterPresetsForUser(user: PrototypeUser): FilterPreset[] {
  const shared: FilterPreset[] = [
    { id: "all", label: "All visible", filters: {} },
    { id: "overdue", label: "Overdue", filters: { dueDate: "overdue" } },
    { id: "dueSoon", label: "Due soon", filters: { dueDate: "dueSoon" } },
    {
      id: "blocked",
      label: "Blocked",
      filters: { workState: "blocked" },
    },
    {
      id: "waitingOnBroker",
      label: "Waiting on broker",
      filters: { workState: "waitingOnBroker" },
    },
  ];
  if (user.role === "Auditor") {
    return [
      {
        id: "myAssignments",
        label: "My assignments",
        filters: { auditor: user.fullName },
      },
      {
        id: "myOverdue",
        label: "My overdue",
        filters: { auditor: user.fullName, dueDate: "overdue" },
      },
      {
        id: "myMissingDocs",
        label: "My missing docs",
        filters: { auditor: user.fullName, missingDocuments: true },
      },
      ...shared,
    ];
  }
  if (user.role === "Finance") {
    return [
      {
        id: "pendingPayment",
        label: "Pending payment",
        filters: { workState: "pendingPayment" },
      },
      {
        id: "invoiceStage",
        label: "Invoice stage",
        filters: { stage: "Invoice" },
      },
      ...shared,
    ];
  }
  return shared;
}

function emptyProjectState(user: PrototypeUser) {
  const visibility = effectiveVisibility(user);
  if (visibility === "Finance Records") {
    return {
      title: "No finance records are ready",
      message:
        "Finance records appear after a project reaches final submission, invoice, closed status, or has invoice activity.",
    };
  }
  if (visibility === "Assigned Projects") {
    return {
      title: "No assignments are assigned to you",
      message:
        "Assigned work appears here when your name is listed as lead, support, or reviewer on an active project.",
    };
  }
  if (user.role === "Read Only") {
    return {
      title: "No read-only projects are available",
      message:
        "Read-only users can view projects once project records have been created and made visible.",
    };
  }
  return {
    title: "No projects in the tracker",
    message:
      "Create a project or import a JSON backup to start working from this browser.",
  };
}

function formatCurrency(value: number) {
  return value.toLocaleString("en-US", {
    style: "currency",
    currency: "USD",
    maximumFractionDigits: 0,
  });
}

function formatDateTime(value: string) {
  if (!value) return "Never";
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return value;
  return parsed.toLocaleString("en-US", {
    month: "short",
    day: "numeric",
    year: "numeric",
    hour: "numeric",
    minute: "2-digit",
  });
}

function createActivityEvent(
  type: AuditActivityEvent["type"],
  title: string,
  detail: string,
  actor = "Prototype user",
): AuditActivityEvent {
  return {
    id: `event-${Date.now()}-${Math.random().toString(36).slice(2)}`,
    createdAt: timestampNow(),
    actor,
    type,
    title,
    detail,
  };
}

type DurationRange = "ytd" | "90d" | "7d";

function rangeStart(range: DurationRange) {
  const start = new Date(today);
  if (range === "ytd") return new Date(`${today.getUTCFullYear()}-01-01T00:00:00Z`);
  start.setUTCDate(start.getUTCDate() - (range === "90d" ? 90 : 7));
  return start;
}

function stageDurationMetrics(projects: AuditProject[], range: DurationRange) {
  const start = rangeStart(range);
  const durations = new Map<Stage, number[]>();
  stages.forEach((stage) => durations.set(stage, []));
  projects.forEach((project) => {
    const sorted = project.statusHistory
      .slice()
      .sort((a, b) => Date.parse(a.changedAt) - Date.parse(b.changedAt));
    sorted.forEach((item, index) => {
      const changedAt = new Date(item.changedAt);
      if (Number.isNaN(changedAt.getTime()) || changedAt < start) return;
      const prior = sorted[index - 1];
      const enteredAt = prior ? new Date(prior.changedAt) : new Date(`${project.lastUpdatedDate}T12:00:00Z`);
      const durationDays = Math.max(0, Math.round((changedAt.getTime() - enteredAt.getTime()) / 86400000));
      durations.get(item.fromStage)?.push(durationDays);
    });
  });
  return stages
    .map((stage) => {
      const values = durations.get(stage) ?? [];
      const average = values.length
        ? values.reduce((sum, value) => sum + value, 0) / values.length
        : 0;
      return { stage, average, count: values.length };
    })
    .filter((metric) => metric.count > 0);
}

export function withProjectDefaults(project: AuditProject): AuditProject {
  const auditTeam = normalizeAuditTeam(project);
  return {
    ...project,
    assignmentType: project.assignmentType ?? "CH",
    auditEntity: project.auditEntity ?? "",
    assignedAuditor: primaryAuditor({ ...project, auditTeam }),
    auditTeam,
    paymentReceived:
      project.paymentReceived ?? project.invoiceStatus === "Paid",
    labels: project.labels ?? [],
    documentRequestDate: project.documentRequestDate ?? "",
    brokerLastChasedDate: project.brokerLastChasedDate ?? "",
    brokerExpectedResponseDate: project.brokerExpectedResponseDate ?? "",
    checklistCompletions: project.checklistCompletions ?? {},
    comments: project.comments ?? [],
    activityEvents: project.activityEvents ?? [],
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

function loadStorageMode(): StorageMode {
  return localStorage.getItem(storageModeStorageKey) === "microsoft-lists"
    ? "microsoft-lists"
    : "local";
}

function saveStorageMode(mode: StorageMode) {
  localStorage.setItem(storageModeStorageKey, mode);
}

function emptyMicrosoftListsConfig(): MicrosoftListsConnectionConfig {
  return { siteId: "", listIds: {} };
}

function loadMicrosoftListsConfig(): MicrosoftListsConnectionConfig {
  const raw = localStorage.getItem(microsoftListsConfigStorageKey);
  if (!raw) return emptyMicrosoftListsConfig();
  try {
    return sanitizeMicrosoftListsConfig(
      JSON.parse(raw) as MicrosoftListsConnectionConfig,
    );
  } catch {
    return emptyMicrosoftListsConfig();
  }
}

function saveMicrosoftListsConfig(config: MicrosoftListsConnectionConfig) {
  localStorage.setItem(
    microsoftListsConfigStorageKey,
    JSON.stringify(sanitizeMicrosoftListsConfig(config)),
  );
}

function emptyMicrosoftAuthConfig(): MicrosoftAuthConfig {
  return { clientId: "", tenantId: "" };
}

function loadMicrosoftAuthConfig(): MicrosoftAuthConfig {
  const raw = localStorage.getItem(microsoftAuthConfigStorageKey);
  if (!raw) return emptyMicrosoftAuthConfig();
  try {
    return sanitizeMicrosoftAuthConfig(JSON.parse(raw) as MicrosoftAuthConfig);
  } catch {
    return emptyMicrosoftAuthConfig();
  }
}

function saveMicrosoftAuthConfig(config: MicrosoftAuthConfig) {
  localStorage.setItem(
    microsoftAuthConfigStorageKey,
    JSON.stringify(sanitizeMicrosoftAuthConfig(config)),
  );
}

function microsoftFieldsToProject(
  fields: Record<string, string | number | boolean | null>,
): AuditProject {
  const project = blankProject();
  const textValue = (key: string) => String(fields[key] ?? "");
  const numberValue = (key: string) => Number(fields[key] ?? 0) || 0;
  const booleanValue = (key: string) => fields[key] === true;
  const labels = textValue("Labels")
    .split(";")
    .map((label) => label.trim())
    .filter((label): label is ProjectLabel =>
      labelOptions.includes(label as ProjectLabel),
    );

  return withProjectDefaults({
    ...project,
    id: textValue("TrackerAssignmentId") || project.id,
    assignmentNumber: textValue("AssignmentNumber") || project.assignmentNumber,
    assignmentSource:
      textValue("AssignmentSource") === "DAM" ? "DAM" : "Email",
    assignmentType: assignmentTypeOptions.includes(
      textValue("AssignmentType") as AssignmentType,
    )
      ? (textValue("AssignmentType") as AssignmentType)
      : "CH",
    auditEntity: textValue("AuditEntity"),
    clientCoverholderCode: textValue("ClientCoverholderCode"),
    broker: textValue("Broker"),
    assignedAuditor: textValue("LeadAuditor"),
    auditTeam: textValue("LeadAuditor")
      ? [{ person: textValue("LeadAuditor"), role: "Lead Auditor" }]
      : [],
    reviewer: textValue("Reviewer"),
    currentStage: stages.includes(textValue("CurrentStage") as Stage)
      ? (textValue("CurrentStage") as Stage)
      : "Intake",
    assignmentStatus: assignmentStatusOptions.includes(
      textValue("AssignmentStatus") as AssignmentStatus,
    )
      ? (textValue("AssignmentStatus") as AssignmentStatus)
      : "New",
    quoteStatus: quoteStatusOptions.includes(
      textValue("QuoteStatus") as QuoteStatus,
    )
      ? (textValue("QuoteStatus") as QuoteStatus)
      : "Not Started",
    quoteAmount: numberValue("QuoteAmount"),
    tentativeAuditWeek: textValue("TentativeAuditWeek"),
    confirmedAuditDate: textValue("ConfirmedAuditDate").slice(0, 10),
    auditType: textValue("AuditType") === "Onsite" ? "Onsite" : "Remote",
    baaReceived: booleanValue("BaaReceived"),
    endorsementsReceived: booleanValue("EndorsementsReceived"),
    premiumBdxReceived: booleanValue("PremiumBdxReceived"),
    preAuditQuestionnaireStatus: progressStatusOptions.includes(
      textValue("PreAuditQuestionnaireStatus") as ProgressStatus,
    )
      ? (textValue("PreAuditQuestionnaireStatus") as ProgressStatus)
      : "Not Started",
    documentRequestStatus: progressStatusOptions.includes(
      textValue("DocumentRequestStatus") as ProgressStatus,
    )
      ? (textValue("DocumentRequestStatus") as ProgressStatus)
      : "Not Started",
    documentRequestDate: textValue("DocumentRequestDate").slice(0, 10),
    brokerLastChasedDate: textValue("BrokerLastChasedDate").slice(0, 10),
    brokerExpectedResponseDate: textValue("BrokerExpectedResponseDate").slice(0, 10),
    fileSelectionCompleted: booleanValue("FileSelectionCompleted"),
    testingSheetCompleted: booleanValue("TestingSheetCompleted"),
    findingsSentDate: textValue("FindingsSentDate").slice(0, 10),
    coverholderResponseReceivedDate: textValue("CoverholderResponseReceivedDate").slice(0, 10),
    reportStatus: reportStatusOptions.includes(
      textValue("ReportStatus") as ReportStatus,
    )
      ? (textValue("ReportStatus") as ReportStatus)
      : "Not Started",
    invoiceStatus: invoiceStatusOptions.includes(
      textValue("InvoiceStatus") as InvoiceStatus,
    )
      ? (textValue("InvoiceStatus") as InvoiceStatus)
      : "Not Started",
    paymentReceived: booleanValue("PaymentReceived"),
    damSubmissionStatus: damSubmissionStatusOptions.includes(
      textValue("DamSubmissionStatus") as DamSubmissionStatus,
    )
      ? (textValue("DamSubmissionStatus") as DamSubmissionStatus)
      : "Not Required",
    nextAction: textValue("NextAction"),
    blockers: textValue("Blockers"),
    dueDate: textValue("DueDate").slice(0, 10),
    lastUpdatedDate: textValue("LastUpdatedDate").slice(0, 10) || todayIso(),
    labels,
  });
}

function withUserDefaults(user: Partial<PrototypeUser>): PrototypeUser {
  const role = user.role ?? "Auditor";
  const emailVerified = user.emailVerified ?? Boolean(user.active ?? true);
  const accessRequestStatus =
    user.accessRequestStatus ??
    (user.active === false ? "Pending Approval" : "Approved");
  return {
    fullName: user.fullName?.trim() || "New User",
    username: user.username?.trim().toLowerCase() || "new.user",
    password: user.password || "password",
    role,
    permissionGroup: user.permissionGroup ?? role,
    email: user.email?.trim() || "new.user@[company-domain]",
    active: user.active ?? true,
    defaultVisibility: user.defaultVisibility ?? "Role Default",
    emailVerified,
    accessRequestStatus,
    verificationCode: user.verificationCode ?? "",
    requestedAt: user.requestedAt ?? "",
    approvedAt: user.approvedAt ?? "",
    approvedBy: user.approvedBy ?? "",
    rejectionReason: user.rejectionReason ?? "",
  };
}

function loadPrototypeUsers(): PrototypeUser[] {
  const raw = localStorage.getItem(usersStorageKey);
  if (!raw) {
    localStorage.setItem(usersStorageKey, JSON.stringify(defaultPrototypeUsers));
    return defaultPrototypeUsers;
  }
  try {
    const parsed = JSON.parse(raw) as Partial<PrototypeUser>[];
    const users = parsed.map(withUserDefaults);
    return users.length ? users : defaultPrototypeUsers;
  } catch {
    return defaultPrototypeUsers;
  }
}

function savePrototypeUsers(users: PrototypeUser[]) {
  localStorage.setItem(usersStorageKey, JSON.stringify(users));
}

function secureUserToPrototypeUser(user: SecureAccessUser): PrototypeUser {
  return withUserDefaults({
    fullName: user.fullName,
    username: user.username,
    password: "",
    role: user.role,
    permissionGroup: user.permissionGroup,
    email: user.email,
    active: user.active,
    defaultVisibility: user.defaultVisibility,
    emailVerified: user.emailVerified,
    accessRequestStatus: user.accessRequestStatus,
    verificationCode: "",
    requestedAt: user.requestedAt,
    approvedAt: user.approvedAt,
    approvedBy: user.approvedBy,
    rejectionReason: user.rejectionReason,
  });
}

function authenticateUser(
  username: string,
  password: string,
  users: PrototypeUser[],
) {
  const normalizedUsername = username.trim().toLowerCase();
  return (
    users.find(
      (user) =>
        user.active &&
        user.username === normalizedUsername &&
        user.password === password,
    ) ?? null
  );
}

function loginAccessMessage(username: string, users: PrototypeUser[]) {
  const normalizedUsername = username.trim().toLowerCase();
  const user = users.find((candidate) => candidate.username === normalizedUsername);
  if (!user) return "Invalid prototype username or password.";
  if (!user.emailVerified) {
    return "Confirm your company email before an admin can approve access.";
  }
  if (user.accessRequestStatus === "Pending Approval") {
    return "Your account request is waiting for admin approval.";
  }
  if (user.accessRequestStatus === "Rejected") {
    return user.rejectionReason || "Your account request was rejected.";
  }
  if (!user.active) return "Your account is inactive.";
  return "Invalid prototype username or password.";
}

function saveCurrentUsername(username: string) {
  if (!username) {
    localStorage.removeItem(currentUserStorageKey);
    return;
  }
  localStorage.setItem(currentUserStorageKey, username);
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

export function daysUntil(dateValue: string | undefined, todayDate = today) {
  if (!dateValue) return Number.POSITIVE_INFINITY;
  const due = new Date(`${dateValue}T12:00:00Z`);
  return Math.ceil((due.getTime() - todayDate.getTime()) / 86400000);
}

export function dueLabel(project: AuditProject) {
  const days = daysUntil(project.dueDate);
  if (!Number.isFinite(days))
    return { text: "No due date", className: "muted" };
  if (days < 0)
    return { text: `${Math.abs(days)}d overdue`, className: "danger" };
  if (days === 0) return { text: "Due today", className: "warning" };
  if (days <= 3) return { text: `Due in ${days}d`, className: "warning" };
  return { text: `Due ${project.dueDate}`, className: "ok" };
}

function pushUnique(items: string[], item: string) {
  if (!items.includes(item)) items.push(item);
}

function recommendedNextSteps(
  project: AuditProject,
  today = new Date("2026-05-05T12:00:00Z"),
) {
  const steps: string[] = [];
  const missingDocuments = getMissingDocuments(project);
  const dueInDays = daysUntil(project.dueDate, today);

  if (missingDocuments.length > 0) {
    pushUnique(
      steps,
      `Chase missing documents: ${missingDocuments.join(", ")}.`,
    );
  }
  if (project.quoteStatus !== "Accepted") {
    pushUnique(steps, "Confirm quote status and capture the client decision.");
  }
  if (!project.assignedAuditor && normalizeAuditTeam(project).length === 0) {
    pushUnique(steps, "Assign a lead auditor before advancing the audit.");
  }
  if (dueInDays < 0) {
    pushUnique(steps, "Escalate the overdue assignment and reset the due date.");
  } else if (dueInDays <= 3) {
    pushUnique(
      steps,
      "Prioritize this audit because it is due within three days.",
    );
  }
  if (
    stages.indexOf(project.currentStage) >= stages.indexOf("Findings") &&
    !project.coverholderResponseReceivedDate
  ) {
    pushUnique(
      steps,
      "Request the coverholder response so wrap-up can continue.",
    );
  }
  if (project.nextAction.trim()) {
    pushUnique(
      steps,
      `Complete the recorded next action: ${project.nextAction.trim()}`,
    );
  }

  const stageAction: Record<Stage, string> = {
    Intake: "Validate intake fields and move the audit into registration.",
    Registration: "Confirm reviewer assignment and prepare the quote record.",
    Quote: "Move to scheduling once the quote is accepted.",
    Scheduling:
      "Confirm the audit date, audit week, and remote or onsite format.",
    "Pre-Audit": "Complete document readiness before file selection.",
    "File Selection": "Finish sample selection and notify the audit team.",
    "Audit Fieldwork": "Complete testing, log exceptions, and prepare findings.",
    Findings:
      "Send findings follow-up and record the coverholder response date.",
    "Report Drafting": "Route the draft report to reviewer quality check.",
    "Final Submission":
      "Send the final report package through the correct channel.",
    Invoice: "Issue the invoice and track payment through receipt.",
    Closed: "Archive the record and capture lessons learned.",
  };
  pushUnique(steps, stageAction[project.currentStage]);
  pushUnique(
    steps,
    "Update comments or the audit trail with the latest owner-facing note.",
  );
  pushUnique(
    steps,
    "Review blockers and clear any stale labels before the next stage move.",
  );

  return steps.slice(0, 5);
}

export type DocumentWorkflowAction =
  | "markWaitingOnBroker"
  | "recordBrokerChase"
  | "clearWaitingOnBroker"
  | "markDocumentsComplete";

export type ActivityItem = {
  id: string;
  timestamp: string;
  type: "stage" | "comment" | "checklist" | "document" | "team";
  title: string;
  detail: string;
  tone?: "ok" | "warning" | "danger" | "muted";
};

const activityTypeOptions: ActivityItem["type"][] = [
  "stage",
  "comment",
  "checklist",
  "document",
  "team",
];

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
  if (action === "clearWaitingOnBroker") {
    return withProjectDefaults({
      ...project,
      labels: project.labels.filter((label) => label !== "Waiting on Broker"),
      assignmentStatus:
        project.currentStage === "Closed" ? project.assignmentStatus : "In Progress",
      nextAction:
        project.nextAction ||
        "Broker completed their action; review received support and continue readiness.",
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
  const seenWorkflowEvents = new Set<string>();
  const dedupeWorkflowTitles = new Set([
    "Waiting on Broker applied",
    "Broker chase recorded",
    "Waiting on Broker cleared",
    "Documents marked complete",
  ]);
  const eventItems: ActivityItem[] = (project.activityEvents ?? [])
    .filter((event) => {
      if (event.title === "Comment added") return false;
      if (!dedupeWorkflowTitles.has(event.title)) return true;
      const eventDay = event.createdAt.split(",")[0] || event.createdAt;
      const dedupeKey = `${event.title}:${eventDay}`;
      if (seenWorkflowEvents.has(dedupeKey)) return false;
      seenWorkflowEvents.add(dedupeKey);
      return true;
    })
    .map((event) => ({
      id: `event-${event.id}`,
      timestamp: event.createdAt,
      type: event.type === "field" ? "stage" : event.type,
      title: event.title,
      detail: `${event.detail} · ${event.actor}`,
      tone:
        event.type === "document"
          ? "warning"
          : event.type === "team" || event.type === "checklist"
            ? "ok"
            : "muted",
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
  return [...eventItems, ...stageItems, ...commentItems, ...checklistItems, ...documentItems].sort(
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

const communicationTemplates: CommunicationTemplate[] = [
  {
    id: "document-request",
    label: "Document request",
    kind: "Email",
    purpose: "Send the first document package request to the broker/contact owner.",
    subject: (project) =>
      `Document request - ${project.assignmentNumber} - ${project.auditEntity}`,
    body: (project) =>
      `Hello,\n\nPlease provide the required audit support for ${project.auditEntity} (${project.clientCoverholderCode}).\n\nRequired items:\n- Binding authority agreement\n- Endorsements\n- Premium bordereaux\n- Completed pre-audit questionnaire, if applicable\n\nRequested by: ${project.documentRequestDate || todayIso()}\nExpected response: ${project.brokerExpectedResponseDate || "TBD"}\n\nAudit team: ${formatAuditTeam(project)}\n\nThank you.`,
  },
  {
    id: "pre-audit-questionnaire",
    label: "Pre-audit questionnaire",
    kind: "Document",
    purpose: "Create the standard questionnaire request text before automation exists.",
    subject: (project) =>
      `Pre-audit questionnaire - ${project.assignmentNumber} - ${project.auditEntity}`,
    body: (project) =>
      `Pre-audit questionnaire request\n\nAssignment: ${project.assignmentNumber}\nAudit entity: ${project.auditEntity}\nBroker: ${project.broker}\nAudit type: ${project.auditType}\nTentative audit week: ${project.tentativeAuditWeek || "TBD"}\nConfirmed audit date: ${project.confirmedAuditDate || "TBD"}\n\nPlease complete the questionnaire and return it with supporting documents so the audit team can confirm readiness before fieldwork.`,
  },
  {
    id: "quote-email",
    label: "Quote email",
    kind: "Email",
    purpose: "Prepare the quote wording for review before sending.",
    subject: (project) =>
      `Quote for ${project.assignmentNumber} - ${project.auditEntity}`,
    body: (project) =>
      `Hello,\n\nPlease find the audit quote details for ${project.auditEntity}.\n\nAssignment: ${project.assignmentNumber}\nAssignment type: ${project.assignmentType}\nAudit type: ${project.auditType}\nQuote status: ${project.quoteStatus}\nQuote amount: ${formatCurrency(project.quoteAmount)}\nTentative audit week: ${project.tentativeAuditWeek || "TBD"}\n\nPlease confirm acceptance or advise if any changes are required.\n\nThank you.`,
  },
  {
    id: "findings-follow-up",
    label: "Findings follow-up",
    kind: "Email",
    purpose: "Send findings and start the coverholder response cycle.",
    subject: (project) =>
      `Findings response requested - ${project.assignmentNumber} - ${project.auditEntity}`,
    body: (project) =>
      `Hello,\n\nFindings have been issued for ${project.auditEntity}.\n\nAssignment: ${project.assignmentNumber}\nFindings sent: ${project.findingsSentDate || todayIso()}\nCoverholder response received: ${project.coverholderResponseReceivedDate || "Not yet received"}\n\nPlease provide responses and supporting evidence for each finding so the audit team can continue report finalization.\n\nThank you.`,
  },
  {
    id: "invoice-note",
    label: "Invoice note",
    kind: "Document",
    purpose: "Prepare invoice handoff details for finance.",
    subject: (project) =>
      `Invoice handoff - ${project.assignmentNumber} - ${project.auditEntity}`,
    body: (project) =>
      `Invoice handoff\n\nAssignment: ${project.assignmentNumber}\nAudit entity: ${project.auditEntity}\nClient / coverholder code: ${project.clientCoverholderCode}\nQuote amount: ${formatCurrency(project.quoteAmount)}\nInvoice status: ${project.invoiceStatus}\nPayment received: ${project.paymentReceived ? "Yes" : "No"}\nReport status: ${project.reportStatus}\nDAM submission: ${project.damSubmissionStatus}\n\nNext action: ${project.nextAction || "No next action recorded."}`,
  },
];

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
    ["Lead Auditor", (project) => primaryAuditor(project)],
    ["Audit Team", (project) => formatAuditTeam(project)],
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

function exportProjectsToJson(projects: AuditProject[]) {
  const payload = {
    exportedAt: new Date().toISOString(),
    app: "audit-assignment-tracker",
    version: 1,
    projects,
  };
  downloadJsonFile(
    payload,
    `audit-assignments-${new Date().toISOString().slice(0, 10)}.json`,
  );
}

function exportMicrosoftListsPackage(
  projects: AuditProject[],
  users: PrototypeUser[],
  exportedBy: string,
) {
  const payload = buildMicrosoftListsMigrationPackage(projects, users, {
    exportedBy,
  });
  downloadJsonFile(
    payload,
    `audit-microsoft-lists-package-${new Date().toISOString().slice(0, 10)}.json`,
  );
}

function downloadJsonFile(payload: unknown, filename: string) {
  const blob = new Blob([JSON.stringify(payload, null, 2)], {
    type: "application/json;charset=utf-8",
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
}

function normalizeImportedProjects(payload: unknown) {
  const candidate =
    payload &&
    typeof payload === "object" &&
    "projects" in payload &&
    Array.isArray((payload as { projects?: unknown }).projects)
      ? (payload as { projects: unknown[] }).projects
      : payload;
  if (!Array.isArray(candidate)) return null;
  return (candidate as AuditProject[]).map(withProjectDefaults);
}

function App() {
  const [projects, setProjects] = useState<AuditProject[]>(() =>
    loadProjects(),
  );
  const [users, setUsers] = useState<PrototypeUser[]>(() =>
    loadPrototypeUsers(),
  );
  const [secureAccess, setSecureAccess] = useState<SecureAccessState | null>(
    null,
  );
  const [secureAccessLoading, setSecureAccessLoading] = useState(true);
  const [selectedId, setSelectedId] = useState(projects[0]?.id ?? "");
  const [editing, setEditing] = useState<AuditProject | null>(null);
  const [filters, setFilters] = useState<Filters>(defaultFilters);
  const [message, setMessage] = useState("");
  const [viewMode, setViewMode] = useState<ViewMode>("kanban");
  const [durationRange, setDurationRange] = useState<DurationRange>("ytd");
  const [storageMode, setStorageMode] = useState<StorageMode>(() =>
    loadStorageMode(),
  );
  const [microsoftListsConfig, setMicrosoftListsConfig] =
    useState<MicrosoftListsConnectionConfig>(() => loadMicrosoftListsConfig());
  const [microsoftAuthConfig, setMicrosoftAuthConfig] =
    useState<MicrosoftAuthConfig>(() => loadMicrosoftAuthConfig());
  const [microsoftAccount, setMicrosoftAccount] =
    useState<MicrosoftAuthAccount | null>(null);
  const [graphAccessToken, setGraphAccessToken] = useState("");
  const [microsoftAuthStatus, setMicrosoftAuthStatus] = useState(
    "Add a Microsoft Entra app client ID, then sign in to get a Graph token.",
  );
  const [storageStatus, setStorageStatus] = useState(
    "Using browser storage until Microsoft Lists is configured.",
  );
  const [storageSyncing, setStorageSyncing] = useState(false);
  const [lastExportedAt, setLastExportedAt] = useState(
    () => localStorage.getItem(lastExportStorageKey) ?? "",
  );
  const [confirmation, setConfirmation] = useState<ConfirmationRequest | null>(
    null,
  );
  const [hiddenWorkloadAuditors, setHiddenWorkloadAuditors] = useState<
    string[]
  >([]);
  const [shownZeroLoadAuditors, setShownZeroLoadAuditors] = useState<string[]>([]);

  const signedInUser =
    secureAccess?.authenticated && secureAccess.user
      ? secureUserToPrototypeUser(secureAccess.user)
      : null;

  const refreshSecureAccess = async () => {
    setSecureAccessLoading(true);
    try {
      setSecureAccess(await getSecureAccessState());
      setMessage("");
    } catch (error) {
      setSecureAccess({
        configured: false,
        authenticated: false,
        status: "setup-required",
        setup: {
          configured: false,
          missing: ["secure access server"],
          redirectUri: "http://127.0.0.1:8787/api/auth/callback",
          frontendOrigin: window.location.origin,
        },
      });
      setMessage(error instanceof Error ? error.message : "Secure access failed.");
    } finally {
      setSecureAccessLoading(false);
    }
  };

  useEffect(() => {
    void refreshSecureAccess();
  }, []);

  useEffect(() => {
    let cancelled = false;
    if (!hasMicrosoftAuthConfig(microsoftAuthConfig)) {
      setMicrosoftAccount(null);
      setGraphAccessToken("");
      setMicrosoftAuthStatus(
        "Add a Microsoft Entra app client ID, then sign in to get a Graph token.",
      );
      return;
    }
    setMicrosoftAuthStatus("Checking saved Microsoft sign-in session...");
    void restoreMicrosoftGraphSession(microsoftAuthConfig)
      .then((session) => {
        if (cancelled) return;
        if (!session) {
          setMicrosoftAccount(null);
          setGraphAccessToken("");
          setMicrosoftAuthStatus("Microsoft sign-in is configured but not signed in.");
          return;
        }
        setMicrosoftAccount(session.account);
        setGraphAccessToken(session.accessToken);
        setMicrosoftAuthStatus(
          session.accessToken
            ? `Signed in as ${session.account.username}.`
            : `Signed in as ${session.account.username}; refresh token before syncing.`,
        );
      })
      .catch((error) => {
        if (cancelled) return;
        setMicrosoftAccount(null);
        setGraphAccessToken("");
        setMicrosoftAuthStatus(
          error instanceof Error
            ? error.message
            : "Could not restore Microsoft sign-in.",
        );
      });
    return () => {
      cancelled = true;
    };
  }, [microsoftAuthConfig.clientId, microsoftAuthConfig.tenantId]);
  const auditorOptions = users
    .filter((user) => user.active && user.role !== "Finance" && user.role !== "Read Only")
    .map((user) => user.fullName);
  const auditors = Array.from(
    new Set([
      ...auditorOptions,
      ...projects.flatMap((project) => assignedAuditorNames(project)),
    ]),
  );
  const visibleProjects = useMemo(
    () =>
      signedInUser
        ? projects.filter((project) => canViewProject(signedInUser, project))
        : [],
    [projects, signedInUser],
  );
  const zeroLoadAuditors = auditors.filter(
    (auditor) =>
      !visibleProjects.some(
        (project) =>
          projectHasAuditor(project, auditor) && project.currentStage !== "Closed",
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
      visibleProjects.filter((project) => {
        if (filters.auditor && !projectHasAuditor(project, filters.auditor))
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
          filters.workState &&
          !projectMatchesWorkState(project, filters.workState)
        )
          return false;
        return true;
      }),
    [visibleProjects, filters],
  );
  const selectedProject =
    visibleProjects.find((project) => project.id === selectedId) ??
    visibleProjects[0];
  const pendingAccessRequests =
    secureAccess?.pendingRequests?.map(secureUserToPrototypeUser) ?? [];

  const confirmAccountEmail = async (verificationCode: string) => {
    try {
      await verifySecureAccessCode(verificationCode);
      await refreshSecureAccess();
      setMessage("Email confirmed. Your profile is now waiting for admin approval.");
    } catch (error) {
      setMessage(
        error instanceof Error ? error.message : "Verification code was not accepted.",
      );
    }
  };

  const signOut = async () => {
    await logoutSecureAccess();
    await refreshSecureAccess();
    setEditing(null);
    setSelectedId("");
    setMessage("");
  };

  if (secureAccessLoading || !signedInUser) {
    return (
      <LoginScreen
        access={secureAccess}
        loading={secureAccessLoading}
        message={message}
        onVerifyEmail={confirmAccountEmail}
        onRefresh={() => void refreshSecureAccess()}
      />
    );
  }

  const persist = (nextProjects: AuditProject[]) => {
    setProjects(nextProjects);
    saveProjects(nextProjects);
  };

  const requestConfirmation = (request: ConfirmationRequest) => {
    setConfirmation(request);
  };

  const recordExport = (label: string) => {
    const exportedAt = new Date().toISOString();
    setLastExportedAt(exportedAt);
    localStorage.setItem(lastExportStorageKey, exportedAt);
    setMessage(`${label} exported.`);
  };

  const handleExportCsv = () => {
    exportProjectsToCsv(filteredProjects);
    recordExport("Filtered CSV");
  };

  const handleExportJson = () => {
    exportProjectsToJson(visibleProjects);
    recordExport("JSON backup");
  };

  const handleExportMicrosoftListsPackage = () => {
    if (!hasFullProjectAccess(signedInUser)) {
      setMessage("Only admins and audit managers can export the Microsoft Lists package.");
      return;
    }
    exportMicrosoftListsPackage(
      visibleProjects,
      signedInUser.role === "Admin" ? users : [],
      signedInUser.fullName,
    );
    recordExport("Microsoft Lists package");
  };

  const saveConnectionConfig = (config: MicrosoftListsConnectionConfig) => {
    const cleanConfig = sanitizeMicrosoftListsConfig(config);
    setMicrosoftListsConfig(cleanConfig);
    saveMicrosoftListsConfig(cleanConfig);
    setStorageStatus("Microsoft Lists connection settings saved locally.");
  };

  const saveMicrosoftSignInConfig = (config: MicrosoftAuthConfig) => {
    const cleanConfig = sanitizeMicrosoftAuthConfig(config);
    setMicrosoftAuthConfig(cleanConfig);
    saveMicrosoftAuthConfig(cleanConfig);
    setGraphAccessToken("");
    setMicrosoftAccount(null);
    setMicrosoftAuthStatus(
      hasMicrosoftAuthConfig(cleanConfig)
        ? "Microsoft sign-in settings saved. Sign in to authorize Graph access."
        : "Add a Microsoft Entra app client ID, then sign in to get a Graph token.",
    );
  };

  const switchStorageMode = (mode: StorageMode) => {
    setStorageMode(mode);
    saveStorageMode(mode);
    setStorageStatus(
      mode === "microsoft-lists"
        ? "Microsoft Lists mode selected. Test the connection before syncing records."
        : "Browser storage mode selected. Records save in this browser only.",
    );
  };

  const microsoftListsSession = (accessToken = graphAccessToken.trim()) => ({
    ...microsoftListsConfig,
    accessToken,
  });

  const applyMicrosoftGraphSession = (session: {
    accessToken: string;
    account: MicrosoftAuthAccount;
  }) => {
    setMicrosoftAccount(session.account);
    setGraphAccessToken(session.accessToken);
    setMicrosoftAuthStatus(`Signed in as ${session.account.username}.`);
    return microsoftListsSession(session.accessToken);
  };

  const requireGraphSession = async () => {
    if (!hasMinimumMicrosoftListsConfig(microsoftListsConfig)) {
      setStorageStatus("Site ID and Audit Assignments list ID are required first.");
      return null;
    }
    if (graphAccessToken.trim()) {
      return microsoftListsSession();
    }
    if (!hasMicrosoftAuthConfig(microsoftAuthConfig)) {
      setStorageStatus("Configure Microsoft sign-in before using Microsoft Lists.");
      return null;
    }
    try {
      setMicrosoftAuthStatus("Requesting Microsoft Graph access...");
      return applyMicrosoftGraphSession(
        await refreshMicrosoftGraphToken(microsoftAuthConfig),
      );
    } catch (error) {
      setMicrosoftAuthStatus(
        error instanceof Error
          ? error.message
          : "Microsoft sign-in failed.",
      );
      setStorageStatus("Microsoft Graph access is required before syncing.");
      return null;
    }
  };

  const handleMicrosoftSignIn = async () => {
    if (!hasMicrosoftAuthConfig(microsoftAuthConfig)) {
      setMicrosoftAuthStatus("Save a Microsoft Entra app client ID first.");
      return;
    }
    setStorageSyncing(true);
    try {
      applyMicrosoftGraphSession(await signInWithMicrosoft(microsoftAuthConfig));
      setStorageStatus("Microsoft Graph sign-in ready for Lists sync.");
    } catch (error) {
      setMicrosoftAuthStatus(
        error instanceof Error
          ? error.message
          : "Microsoft sign-in failed.",
      );
    } finally {
      setStorageSyncing(false);
    }
  };

  const handleMicrosoftRefreshToken = async () => {
    if (!hasMicrosoftAuthConfig(microsoftAuthConfig)) {
      setMicrosoftAuthStatus("Save a Microsoft Entra app client ID first.");
      return;
    }
    setStorageSyncing(true);
    try {
      applyMicrosoftGraphSession(
        await refreshMicrosoftGraphToken(microsoftAuthConfig),
      );
      setStorageStatus("Microsoft Graph token refreshed.");
    } catch (error) {
      setMicrosoftAuthStatus(
        error instanceof Error
          ? error.message
          : "Could not refresh Microsoft Graph token.",
      );
    } finally {
      setStorageSyncing(false);
    }
  };

  const handleMicrosoftSignOut = async () => {
    setStorageSyncing(true);
    try {
      if (hasMicrosoftAuthConfig(microsoftAuthConfig)) {
        await signOutOfMicrosoft(microsoftAuthConfig);
      }
      setMicrosoftAccount(null);
      setGraphAccessToken("");
      setMicrosoftAuthStatus("Signed out of Microsoft.");
      setStorageStatus("Microsoft Graph sign-in is disconnected.");
    } catch (error) {
      setMicrosoftAuthStatus(
        error instanceof Error
          ? error.message
          : "Could not sign out of Microsoft.",
      );
    } finally {
      setStorageSyncing(false);
    }
  };

  const testConnection = async () => {
    const session = await requireGraphSession();
    if (!session) return;
    setStorageSyncing(true);
    try {
      await testMicrosoftListsConnection(session);
      setStorageStatus("Microsoft Lists connection verified.");
      switchStorageMode("microsoft-lists");
    } catch (error) {
      setStorageStatus(
        error instanceof Error
          ? error.message
          : "Microsoft Lists connection failed.",
      );
    } finally {
      setStorageSyncing(false);
    }
  };

  const syncToMicrosoftLists = async () => {
    const session = await requireGraphSession();
    if (!session) return;
    if (!hasFullMicrosoftListsConfig(microsoftListsConfig)) {
      setStorageStatus(
        `Missing list IDs: ${missingMicrosoftListLabels(microsoftListsConfig).join(", ")}.`,
      );
      return;
    }
    setStorageSyncing(true);
    try {
      const syncPackage = buildSyncPackage(
        visibleProjects,
        signedInUser.role === "Admin" ? users : [],
        signedInUser.fullName,
      );
      const summary = await pushMigrationPackageToMicrosoftLists(
        session,
        syncPackage,
      );
      const problemText = summary.errors.length
        ? ` ${summary.errors.length} issues need review.`
        : "";
      setStorageStatus(
        `Synced to Microsoft Lists: ${summary.created} created, ${summary.updated} updated, ${summary.skipped} skipped.${problemText}`,
      );
      setStorageMode("microsoft-lists");
      saveStorageMode("microsoft-lists");
    } catch (error) {
      setStorageStatus(
        error instanceof Error
          ? error.message
          : "Microsoft Lists sync failed.",
      );
    } finally {
      setStorageSyncing(false);
    }
  };

  const pullFromMicrosoftLists = async () => {
    const session = await requireGraphSession();
    if (!session) return;
    requestConfirmation({
      title: "Load assignments from Microsoft Lists?",
      message:
        "This replaces the project records in this browser with rows from the configured Audit Assignments list.",
      confirmLabel: "Load from Lists",
      tone: "danger",
      onConfirm: async () => {
        setStorageSyncing(true);
        try {
          const rows = await pullAssignmentRowsFromMicrosoftLists(session);
          const nextProjects = rows.map(microsoftFieldsToProject);
          if (nextProjects.length === 0) {
            setStorageStatus("Microsoft Lists returned no assignment rows.");
            return;
          }
          persist(nextProjects);
          setSelectedId(nextProjects[0].id);
          setStorageMode("microsoft-lists");
          saveStorageMode("microsoft-lists");
          setStorageStatus(
            `${nextProjects.length} assignments loaded from Microsoft Lists.`,
          );
        } catch (error) {
          setStorageStatus(
            error instanceof Error
              ? error.message
              : "Could not load assignments from Microsoft Lists.",
          );
        } finally {
          setStorageSyncing(false);
        }
      },
    });
  };

  const resetSampleProjects = () => {
    requestConfirmation({
      title: "Reset sample data?",
      message:
        "This replaces the current project records in this browser with the starter sample data.",
      confirmLabel: "Reset sample data",
      tone: "danger",
      onConfirm: () => {
        persist(sampleProjects);
        setSelectedId(sampleProjects[0].id);
        setMessage("Sample data reset.");
      },
    });
  };

  const queueProjectImport = (file: File) => {
    if (!hasFullProjectAccess(signedInUser)) {
      setMessage("Only admins and audit managers can import project data.");
      return;
    }
    requestConfirmation({
      title: "Import JSON backup?",
      message: `Importing ${file.name} replaces the current project records in this browser.`,
      confirmLabel: "Import JSON",
      tone: "danger",
      onConfirm: () => importProjectsFromFile(file),
    });
  };

  const upsertProject = (project: AuditProject) => {
    const exists = projects.some((item) => item.id === project.id);
    const originalProject = projects.find((item) => item.id === project.id);
    if (
      (exists && originalProject && !canEditProject(signedInUser, originalProject)) ||
      (!exists && !canCreateProject(signedInUser))
    ) {
      setMessage("Your role cannot save that project.");
      return;
    }
    const cleanProject = withProjectDefaults({
      ...project,
      quoteAmount: Number(project.quoteAmount) || 0,
      lastUpdatedDate: new Date().toISOString().slice(0, 10),
      damSubmissionStatus:
        project.assignmentSource === "DAM"
          ? project.damSubmissionStatus
          : ("Not Required" as DamSubmissionStatus),
      activityEvents: [
        ...(project.activityEvents ?? []),
        createActivityEvent(
          "field",
          exists ? "Project edited" : "Project created",
          exists
            ? "Project fields were updated from Edit Project."
            : "Project was created from guided intake.",
          signedInUser.fullName,
        ),
      ],
    });
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
    if (!canEditProject(signedInUser, project)) {
      setMessage("Your role cannot move this project.");
      return;
    }
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
      activityEvents: [
        ...(project.activityEvents ?? []),
        createActivityEvent(
          "stage",
          `${project.currentStage} → ${targetStage}`,
          "Stage changed in tracker.",
          signedInUser.fullName,
        ),
      ],
      statusHistory: [
        ...project.statusHistory,
        {
          id: `h-${Date.now()}`,
          changedAt: timestampNow(),
          changedBy: signedInUser.fullName,
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

  const removeProjectLabel = (project: AuditProject, label: ProjectLabel) => {
    if (!canEditProject(signedInUser, project)) {
      setMessage("Your role cannot update labels on this project.");
      return;
    }
    if (!project.labels.includes(label)) return;
    const updatedProject = withProjectDefaults({
      ...project,
      labels: project.labels.filter((item) => item !== label),
      assignmentStatus:
        label === "Waiting on Broker" && project.assignmentStatus === "On Hold"
          ? "In Progress"
          : project.assignmentStatus,
      lastUpdatedDate: new Date().toISOString().slice(0, 10),
    });
    persist(
      projects.map((item) => (item.id === project.id ? updatedProject : item)),
    );
    setSelectedId(project.id);
    setMessage(`${label} removed from ${project.assignmentNumber}.`);
  };

  const addProjectComment = (
    project: AuditProject,
    comment: ProjectComment,
  ) => {
    if (!canComment(signedInUser, project)) {
      setMessage("Your role cannot add comments to this project.");
      return;
    }
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
    if (!canEditProject(signedInUser, project)) {
      setMessage("Your role cannot update checklist items on this project.");
      return;
    }
    const updatedProject = {
      ...project,
      checklistCompletions: {
        ...project.checklistCompletions,
        [key]: !project.checklistCompletions[key],
      },
      activityEvents: [
        ...(project.activityEvents ?? []),
        createActivityEvent(
          "checklist",
          project.checklistCompletions[key]
            ? "Checklist item reopened"
            : "Checklist item completed",
          key.split(":").slice(1).join(":") || key,
          signedInUser.fullName,
        ),
      ],
      lastUpdatedDate: new Date().toISOString().slice(0, 10),
    };
    persist(
      projects.map((item) => (item.id === project.id ? updatedProject : item)),
    );
    setSelectedId(project.id);
  };

  const addSupportingAuditor = (project: AuditProject, auditor: string) => {
    if (!hasFullProjectAccess(signedInUser)) {
      setMessage("Only admins and audit managers can change audit teams.");
      return;
    }
    if (!auditor || projectHasAuditor(project, auditor)) return;
    const updatedProject = withProjectDefaults({
      ...project,
      auditTeam: [
        ...normalizeAuditTeam(project),
        { person: auditor, role: "Supporting Auditor" },
      ],
      activityEvents: [
        ...(project.activityEvents ?? []),
        createActivityEvent(
          "team",
          "Supporting auditor added",
          `${auditor} joined as a supporting auditor.`,
          signedInUser.fullName,
        ),
      ],
      lastUpdatedDate: new Date().toISOString().slice(0, 10),
    });
    persist(
      projects.map((item) => (item.id === project.id ? updatedProject : item)),
    );
    setSelectedId(project.id);
    setMessage(`${auditor} added to ${project.assignmentNumber}.`);
  };

  const updateProjectDocumentWorkflow = (
    project: AuditProject,
    action: DocumentWorkflowAction,
  ) => {
    if (!canEditProject(signedInUser, project)) {
      setMessage("Your role cannot update document workflow on this project.");
      return;
    }
    const readiness = documentReadiness(project);
    if (action === "markDocumentsComplete" && readiness.percent === 100) {
      setMessage(`${project.assignmentNumber} documents are already complete.`);
      return;
    }
    if (action === "markWaitingOnBroker" && readiness.waitingOnBroker) {
      setMessage(`${project.assignmentNumber} is already waiting on broker.`);
      return;
    }
    if (action === "recordBrokerChase") {
      const alreadyChasedToday = project.brokerLastChasedDate === todayIso();
      if (alreadyChasedToday) {
        setMessage(`${project.assignmentNumber} already has a broker chase today.`);
        return;
      }
    }
    if (action === "clearWaitingOnBroker" && !readiness.waitingOnBroker) {
      setMessage(`${project.assignmentNumber} is not marked waiting on broker.`);
      return;
    }
    const eventTitle =
      action === "markDocumentsComplete"
        ? "Documents marked complete"
        : action === "recordBrokerChase"
          ? "Broker chase recorded"
          : action === "clearWaitingOnBroker"
            ? "Waiting on Broker cleared"
            : "Waiting on Broker applied";
    const updatedProject = {
      ...applyDocumentWorkflowAction(project, action),
      activityEvents: [
        ...(project.activityEvents ?? []),
        createActivityEvent("document", eventTitle, eventTitle, signedInUser.fullName),
      ],
    };
    persist(
      projects.map((item) => (item.id === project.id ? updatedProject : item)),
    );
    setSelectedId(project.id);
    setMessage(`${eventTitle}. ${project.assignmentNumber} updated.`);
  };

  const updateProjectFinance = (
    project: AuditProject,
    invoiceStatus: InvoiceStatus,
    paymentReceived: boolean,
  ) => {
    if (!canUpdateFinance(signedInUser, project) && !hasFullProjectAccess(signedInUser)) {
      setMessage("Your role cannot update invoice or payment fields.");
      return;
    }
    const updatedProject = withProjectDefaults({
      ...project,
      invoiceStatus,
      paymentReceived: invoiceStatus === "Paid" ? true : paymentReceived,
      lastUpdatedDate: new Date().toISOString().slice(0, 10),
      activityEvents: [
        ...(project.activityEvents ?? []),
        createActivityEvent(
          "field",
          "Finance fields updated",
          `Invoice status set to ${invoiceStatus}; payment received is ${
            invoiceStatus === "Paid" || paymentReceived ? "yes" : "no"
          }.`,
          signedInUser.fullName,
        ),
      ],
    });
    persist(
      projects.map((item) => (item.id === project.id ? updatedProject : item)),
    );
    setSelectedId(project.id);
    setMessage(`${project.assignmentNumber} finance fields updated.`);
  };

  const importProjectsFromFile = async (file: File) => {
    if (!hasFullProjectAccess(signedInUser)) {
      setMessage("Only admins and audit managers can import project data.");
      return;
    }
    try {
      const text = await file.text();
      const importedProjects = normalizeImportedProjects(JSON.parse(text));
      if (!importedProjects || importedProjects.length === 0) {
        setMessage("Import file did not contain any projects.");
        return;
      }
      persist(importedProjects);
      setSelectedId(importedProjects[0].id);
      setMessage(`${importedProjects.length} projects imported from JSON.`);
    } catch {
      setMessage("Could not import that JSON file.");
    }
  };

  const upsertUser = (
    originalUsername: string | null,
    draftUser: PrototypeUser,
  ) => {
    if (signedInUser.role !== "Admin") {
      setMessage("Only admins can manage users.");
      return;
    }
    const cleanUser = withUserDefaults({
      ...draftUser,
      username: draftUser.username.trim().toLowerCase(),
      permissionGroup: draftUser.role,
    });
    if (!cleanUser.fullName.trim() || !cleanUser.username.trim()) {
      setMessage("User name and username are required.");
      return;
    }
    if (
      users.some(
        (user) =>
          user.username === cleanUser.username &&
          user.username !== originalUsername,
      )
    ) {
      setMessage("That username already exists.");
      return;
    }
    if (originalUsername === signedInUser.username && !cleanUser.active) {
      setMessage("You cannot deactivate your own signed-in account.");
      return;
    }
    const nextUsers = originalUsername
      ? users.map((user) =>
          user.username === originalUsername ? cleanUser : user,
        )
      : [...users, cleanUser];
    if (!nextUsers.some((user) => user.active && user.role === "Admin")) {
      setMessage("At least one active Admin user is required.");
      return;
    }
    setUsers(nextUsers);
    savePrototypeUsers(nextUsers);
    setMessage(`${cleanUser.fullName} saved.`);
  };

  const approveUserRequest = async (username: string) => {
    if (signedInUser.role !== "Admin") {
      setMessage("Only admins can approve account requests.");
      return;
    }
    const user = pendingAccessRequests.find(
      (candidate) => candidate.username === username,
    );
    if (!user || !canApproveAccessRequest(user)) {
      setMessage("That account request must have a verified email before approval.");
      return;
    }
    try {
      await approveSecureAccessRequest(user.email);
      await refreshSecureAccess();
      setMessage(`${user.fullName} approved for tracker access.`);
    } catch (error) {
      setMessage(error instanceof Error ? error.message : "Approval failed.");
    }
  };

  const rejectUserRequest = async (username: string) => {
    if (signedInUser.role !== "Admin") {
      setMessage("Only admins can reject account requests.");
      return;
    }
    const user = pendingAccessRequests.find(
      (candidate) => candidate.username === username,
    );
    if (!user) {
      setMessage("Account request was not found.");
      return;
    }
    try {
      await rejectSecureAccessRequest(user.email);
      await refreshSecureAccess();
      setMessage(`${user.fullName} rejected.`);
    } catch (error) {
      setMessage(error instanceof Error ? error.message : "Rejection failed.");
    }
  };

  const resetUsers = () => {
    if (signedInUser.role !== "Admin") {
      setMessage("Only admins can reset users.");
      return;
    }
    setUsers(defaultPrototypeUsers);
    savePrototypeUsers(defaultPrototypeUsers);
    setMessage("Prototype users reset.");
  };

  const confirmResetUsers = () => {
    requestConfirmation({
      title: "Reset prototype users?",
      message:
        "This restores the starter users, roles, passwords, and visibility settings in this browser.",
      confirmLabel: "Reset users",
      tone: "danger",
      onConfirm: resetUsers,
    });
  };

  return (
    <main>
      <header className="hero">
        <div>
          <p className="eyebrow">Prototype login: {signedInUser.role}</p>
          <h1>Audit Assignment Tracker</h1>
          <p>
            Manage audit assignments from intake through final report, invoice,
            and close-out.
          </p>
        </div>
        <div className="hero-actions">
          <button
            disabled={!canCreateProject(signedInUser)}
            onClick={() => setEditing(blankProject())}
          >
            Add project
          </button>
          <button
            className="secondary"
            disabled={!hasFullProjectAccess(signedInUser)}
            onClick={resetSampleProjects}
          >
            Reset sample data
          </button>
          <button className="secondary" onClick={signOut}>
            Sign out
          </button>
        </div>
      </header>

      {message && (
        <div className="toast" role="status">
          {message}
        </div>
      )}

      <AccessBanner user={signedInUser} visibleCount={visibleProjects.length} />
      {signedInUser.role === "Admin" && (
        <UserManagementPanel
          users={users}
          pendingRequests={pendingAccessRequests}
          onSaveUser={upsertUser}
          onApproveRequest={approveUserRequest}
          onRejectRequest={rejectUserRequest}
          onResetUsers={confirmResetUsers}
        />
      )}
      {hasFullProjectAccess(signedInUser) && (
        <CentralStoragePanel
          projects={visibleProjects}
          users={signedInUser.role === "Admin" ? users : []}
          exportedBy={signedInUser.fullName}
          mode={storageMode}
          config={microsoftListsConfig}
          authConfig={microsoftAuthConfig}
          account={microsoftAccount}
          accessToken={graphAccessToken}
          authStatus={microsoftAuthStatus}
          status={storageStatus}
          syncing={storageSyncing}
          onExport={handleExportMicrosoftListsPackage}
          onModeChange={switchStorageMode}
          onSaveConfig={saveConnectionConfig}
          onSaveAuthConfig={saveMicrosoftSignInConfig}
          onMicrosoftSignIn={() => void handleMicrosoftSignIn()}
          onMicrosoftSignOut={() => void handleMicrosoftSignOut()}
          onMicrosoftRefreshToken={() => void handleMicrosoftRefreshToken()}
          onTestConnection={() => void testConnection()}
          onSyncToLists={() => void syncToMicrosoftLists()}
          onPullFromLists={() => void pullFromMicrosoftLists()}
        />
      )}
      <Dashboard projects={visibleProjects} />
      <TodaysWork projects={visibleProjects} onSelect={setSelectedId} />
      <CycleTimeDashboard
        projects={visibleProjects}
        range={durationRange}
        setRange={setDurationRange}
      />
      <WorkloadCounts
        projects={visibleProjects}
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
      <FiltersPanel
        filters={filters}
        setFilters={setFilters}
        presets={filterPresetsForUser(signedInUser)}
        auditors={
          hasFullProjectAccess(signedInUser) || signedInUser.role === "Read Only"
            ? auditors
            : auditors.filter((auditor) =>
                visibleProjects.some((project) => projectHasAuditor(project, auditor)),
              )
        }
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
            Audit table
          </button>
        </div>
        <button onClick={handleExportCsv}>
          Export filtered CSV
        </button>
        <button
          className="secondary"
          onClick={handleExportJson}
        >
          Export JSON backup
        </button>
        <button
          className="secondary"
          disabled={!hasFullProjectAccess(signedInUser)}
          onClick={handleExportMicrosoftListsPackage}
        >
          Export Lists package
        </button>
        <span className="last-export">
          Last export: {formatDateTime(lastExportedAt)}
        </span>
        <label className="import-control">
          Import JSON
          <input
            type="file"
            accept="application/json,.json"
            disabled={!hasFullProjectAccess(signedInUser)}
            onChange={(event) => {
              const file = event.target.files?.[0];
              if (file) queueProjectImport(file);
              event.currentTarget.value = "";
            }}
          />
        </label>
      </div>
      {viewMode === "kanban" ? (
        <Kanban
          projects={filteredProjects}
          selectedId={selectedProject?.id}
          onSelect={setSelectedId}
          onMove={moveProject}
          onRemoveLabel={removeProjectLabel}
          currentUser={signedInUser}
        />
      ) : (
        <ProjectTable projects={filteredProjects} onSelect={setSelectedId} />
      )}
      {selectedProject && (
        <ProjectDetail
          project={selectedProject}
          onEdit={() => setEditing(selectedProject)}
          onMove={moveProject}
          onRemoveLabel={removeProjectLabel}
          onAddComment={addProjectComment}
          onToggleChecklist={toggleChecklistItem}
          onDocumentWorkflowAction={updateProjectDocumentWorkflow}
          auditors={auditors}
          onAddSupportingAuditor={addSupportingAuditor}
          currentUser={signedInUser}
          onUpdateFinance={updateProjectFinance}
        />
      )}
      {!selectedProject && (
        <section className="panel empty-state">
          <h2>{emptyProjectState(signedInUser).title}</h2>
          <p>{emptyProjectState(signedInUser).message}</p>
        </section>
      )}
      {editing && (
        <ProjectForm
          project={editing}
          onCancel={() => setEditing(null)}
          onSave={upsertProject}
          auditorOptions={auditorOptions}
        />
      )}
      {confirmation && (
        <ConfirmDialog
          request={confirmation}
          onCancel={() => setConfirmation(null)}
          onConfirm={() => {
            void Promise.resolve(confirmation.onConfirm()).finally(() =>
              setConfirmation(null),
            );
          }}
        />
      )}
    </main>
  );
}

function LoginScreen({
  access,
  loading,
  message,
  onVerifyEmail,
  onRefresh,
}: {
  access: SecureAccessState | null;
  loading: boolean;
  message: string;
  onVerifyEmail: (verificationCode: string) => void | Promise<void>;
  onRefresh: () => void;
}) {
  const [verificationCode, setVerificationCode] = useState("");

  const submitEmailVerification = (event: FormEvent) => {
    event.preventDefault();
    void onVerifyEmail(verificationCode);
  };
  const signInUrl = access?.signInUrl ?? secureAccessUrl("/api/auth/start?mode=signin");
  const requestAccessUrl =
    access?.requestAccessUrl ?? secureAccessUrl("/api/auth/start?mode=request");
  const pendingVerification =
    access?.status === "pending-verification" ||
    access?.user?.accessRequestStatus === "Pending Verification";
  const pendingApproval =
    access?.status === "pending-approval" ||
    access?.user?.accessRequestStatus === "Pending Approval";
  const rejected =
    access?.status === "rejected" || access?.user?.accessRequestStatus === "Rejected";

  return (
    <main className="login-shell">
      <section className="microsoft-login-stack">
        <div className="microsoft-login-card">
          <div className="microsoft-brand">
            <span className="microsoft-mark" aria-hidden="true">
              <span />
              <span />
              <span />
              <span />
            </span>
            <strong>Microsoft</strong>
          </div>
          <h1>Sign in</h1>
          <p className="microsoft-login-copy">
            Use your company Microsoft account to access the Audit Assignment Tracker.
          </p>
          {loading && <div className="toast">Checking secure access...</div>}
          {access?.status === "setup-required" && (
            <div className="secure-setup-box">
              <strong>Secure access server is not configured.</strong>
              <span>Missing: {access.setup?.missing.join(", ")}</span>
              <span>Register redirect URI: {access.setup?.redirectUri}</span>
            </div>
          )}
          {pendingVerification && (
            <form className="verification-form" onSubmit={submitEmailVerification}>
              <p>
                A verification code was sent to {access?.user?.email}. Enter it before
                admin approval.
              </p>
              <Input
                label="Verification code"
                value={verificationCode}
                onChange={setVerificationCode}
                placeholder="6-digit code"
              />
              <button type="submit">Confirm code</button>
            </form>
          )}
          {pendingApproval && (
            <div className="secure-setup-box">
              <strong>Email confirmed.</strong>
              <span>Your profile is waiting for admin approval.</span>
            </div>
          )}
          {rejected && (
            <div className="secure-setup-box">
              <strong>Access request rejected.</strong>
              <span>{access?.user?.rejectionReason || "Contact an admin."}</span>
            </div>
          )}
        {message && (
          <div className="toast" role="status">
            {message}
          </div>
        )}
          <div className="microsoft-login-actions">
            <a className="microsoft-primary-link" href={signInUrl}>
              Next
            </a>
            <button type="button" className="secondary" onClick={onRefresh}>
              Back
            </button>
          </div>
          <p className="microsoft-create-line">
            No account? <a href={requestAccessUrl}>Create one!</a>
          </p>
          <a className="microsoft-help-link" href={requestAccessUrl}>
            Request tracker access
          </a>
        </div>
        <a className="signin-options-card" href={signInUrl}>
          <span aria-hidden="true">⌕</span>
          Sign-in options
        </a>
      </section>
    </main>
  );
}

function ConfirmDialog({
  request,
  onCancel,
  onConfirm,
}: {
  request: ConfirmationRequest;
  onCancel: () => void;
  onConfirm: () => void;
}) {
  return (
    <div className="modal-backdrop confirmation-backdrop">
      <section
        className="panel confirmation-modal"
        role="dialog"
        aria-modal="true"
        aria-labelledby="confirmation-title"
      >
        <div>
          <p className="eyebrow dark">Confirm</p>
          <h2 id="confirmation-title">{request.title}</h2>
          <p>{request.message}</p>
        </div>
        <div className="modal-actions">
          <button type="button" className="secondary" onClick={onCancel}>
            Cancel
          </button>
          <button
            type="button"
            className={request.tone === "danger" ? "danger-button" : undefined}
            onClick={onConfirm}
          >
            {request.confirmLabel}
          </button>
        </div>
      </section>
    </div>
  );
}

function AccessBanner({
  user,
  visibleCount,
}: {
  user: PrototypeUser;
  visibleCount: number;
}) {
  return (
    <section className="panel access-banner">
      <div>
        <span className="avatar">
          {user.fullName
            .split(" ")
            .map((part) => part[0])
            .slice(0, 2)
            .join("")}
        </span>
        <div>
          <p className="eyebrow dark">Signed in</p>
          <h2>{user.fullName}</h2>
          <span>
            {user.username} · {user.permissionGroup} · {visibleCount} visible projects
          </span>
        </div>
      </div>
      <strong>{visibleProjectMessage(user)}</strong>
    </section>
  );
}

function CentralStoragePanel({
  projects,
  users,
  exportedBy,
  mode,
  config,
  authConfig,
  account,
  accessToken,
  authStatus,
  status,
  syncing,
  onExport,
  onModeChange,
  onSaveConfig,
  onSaveAuthConfig,
  onMicrosoftSignIn,
  onMicrosoftSignOut,
  onMicrosoftRefreshToken,
  onTestConnection,
  onSyncToLists,
  onPullFromLists,
}: {
  projects: AuditProject[];
  users: PrototypeUser[];
  exportedBy: string;
  mode: StorageMode;
  config: MicrosoftListsConnectionConfig;
  authConfig: MicrosoftAuthConfig;
  account: MicrosoftAuthAccount | null;
  accessToken: string;
  authStatus: string;
  status: string;
  syncing: boolean;
  onExport: () => void;
  onModeChange: (mode: StorageMode) => void;
  onSaveConfig: (config: MicrosoftListsConnectionConfig) => void;
  onSaveAuthConfig: (config: MicrosoftAuthConfig) => void;
  onMicrosoftSignIn: () => void;
  onMicrosoftSignOut: () => void;
  onMicrosoftRefreshToken: () => void;
  onTestConnection: () => void;
  onSyncToLists: () => void;
  onPullFromLists: () => void;
}) {
  const [draftConfig, setDraftConfig] = useState(config);
  const [draftAuthConfig, setDraftAuthConfig] = useState(authConfig);
  useEffect(() => {
    setDraftConfig(config);
  }, [config]);
  useEffect(() => {
    setDraftAuthConfig(authConfig);
  }, [authConfig]);
  const packagePreview = useMemo(
    () =>
      buildMicrosoftListsMigrationPackage(projects, users, {
        exportedBy,
        exportedAt: "preview",
      }),
    [projects, users, exportedBy],
  );
  const updateListId = (key: MicrosoftListKey, value: string) => {
    setDraftConfig((current) => ({
      ...current,
      listIds: {
        ...current.listIds,
        [key]: value,
      },
    }));
  };
  const hasMinimumConfig = hasMinimumMicrosoftListsConfig(config);
  const hasFullConfig = hasFullMicrosoftListsConfig(config);
  const hasAuthConfig = hasMicrosoftAuthConfig(authConfig);
  const canRequestGraph = hasAuthConfig || Boolean(accessToken.trim());

  return (
    <section className="panel central-storage">
      <div className="section-title">
        <div>
          <p className="eyebrow dark">Central storage</p>
          <h2>Microsoft Lists foundation</h2>
          <span>
            {mode === "microsoft-lists"
              ? "Microsoft Lists mode selected; sync actions use the configured Graph connection."
              : "Browser storage active; Microsoft Lists connection is ready to configure."}
          </span>
        </div>
        <div className="storage-mode-toggle">
          <button
            type="button"
            className={mode === "local" ? "active" : "secondary"}
            onClick={() => onModeChange("local")}
          >
            Browser
          </button>
          <button
            type="button"
            className={mode === "microsoft-lists" ? "active" : "secondary"}
            onClick={() => onModeChange("microsoft-lists")}
          >
            Microsoft Lists
          </button>
        </div>
      </div>
      <div className={`storage-status ${mode === "microsoft-lists" ? "connected" : ""}`}>
        <strong>{mode === "microsoft-lists" ? "Microsoft Lists mode" : "Browser storage mode"}</strong>
        <span>{status}</span>
      </div>
      <div className="storage-stats">
        <span>
          <strong>{packagePreview.totals.lists}</strong>
          Lists
        </span>
        <span>
          <strong>{packagePreview.totals.assignments}</strong>
          Assignments
        </span>
        <span>
          <strong>{packagePreview.totals.activityLogEvents}</strong>
          Activity events
        </span>
        <span>
          <strong>{packagePreview.totals.rows}</strong>
          Seed rows
        </span>
      </div>
      <div className="storage-config">
        <Input
          label="SharePoint site ID"
          value={draftConfig.siteId}
          placeholder="contoso.sharepoint.com,site-id,web-id"
          onChange={(value) =>
            setDraftConfig((current) => ({ ...current, siteId: value }))
          }
        />
        <Input
          label="Microsoft Entra app client ID"
          value={draftAuthConfig.clientId}
          placeholder="Application (client) ID"
          onChange={(value) =>
            setDraftAuthConfig((current) => ({ ...current, clientId: value }))
          }
        />
        <Input
          label="Tenant ID or domain"
          value={draftAuthConfig.tenantId}
          placeholder="organizations, tenant ID, or domain"
          onChange={(value) =>
            setDraftAuthConfig((current) => ({ ...current, tenantId: value }))
          }
        />
        <div className="storage-list-id-grid">
          {requiredMicrosoftListKeys.map((key) => (
            <Input
              key={key}
              label={`${microsoftListLabels[key]} list ID`}
              value={draftConfig.listIds[key] ?? ""}
              placeholder="SharePoint list ID"
              onChange={(value) => updateListId(key, value)}
            />
          ))}
        </div>
      </div>
      <div className="auth-card">
        <div>
          <strong>
            {account ? `Signed in as ${account.name}` : "Microsoft sign-in"}
          </strong>
          <span>{authStatus}</span>
          <small>Requested Graph scopes: {microsoftAuthScopeLabel()}</small>
        </div>
        <div className="storage-actions">
          <button type="button" onClick={() => onSaveAuthConfig(draftAuthConfig)}>
            Save sign-in
          </button>
          <button
            type="button"
            className="secondary"
            disabled={syncing || !hasAuthConfig}
            onClick={onMicrosoftSignIn}
          >
            Sign in with Microsoft
          </button>
          <button
            type="button"
            className="secondary"
            disabled={syncing || !hasAuthConfig}
            onClick={onMicrosoftRefreshToken}
          >
            Refresh token
          </button>
          <button
            type="button"
            className="secondary"
            disabled={syncing || !account}
            onClick={onMicrosoftSignOut}
          >
            Sign out
          </button>
        </div>
      </div>
      <div className="storage-actions">
        <button type="button" onClick={() => onSaveConfig(draftConfig)}>
          Save connection
        </button>
        <button
          type="button"
          className="secondary"
          disabled={syncing || !hasMinimumConfig || !canRequestGraph}
          onClick={onTestConnection}
        >
          Test connection
        </button>
        <button
          type="button"
          className="secondary"
          disabled={syncing || !hasFullConfig || !canRequestGraph}
          onClick={onSyncToLists}
        >
          Sync to Lists
        </button>
        <button
          type="button"
          className="secondary"
          disabled={syncing || !hasMinimumConfig || !canRequestGraph}
          onClick={onPullFromLists}
        >
          Load from Lists
        </button>
        <button type="button" className="secondary" onClick={onExport}>
          Export Lists package
        </button>
      </div>
      <div className="storage-list-chips" aria-label="Microsoft Lists schema">
        {microsoftListSchemas.map((schema) => (
          <span key={schema.key}>{schema.displayName}</span>
        ))}
      </div>
    </section>
  );
}

function UserManagementPanel({
  users,
  pendingRequests,
  onSaveUser,
  onApproveRequest,
  onRejectRequest,
  onResetUsers,
}: {
  users: PrototypeUser[];
  pendingRequests: PrototypeUser[];
  onSaveUser: (originalUsername: string | null, user: PrototypeUser) => void;
  onApproveRequest: (username: string) => void;
  onRejectRequest: (username: string) => void;
  onResetUsers: () => void;
}) {
  const firstUsername = users[0]?.username ?? "";
  const [selectedUsername, setSelectedUsername] = useState(firstUsername);
  const selectedUser =
    users.find((user) => user.username === selectedUsername) ?? users[0];
  const [draft, setDraft] = useState<PrototypeUser>(
    selectedUser ??
      withUserDefaults({
        fullName: "",
        username: "",
      }),
  );
  const [isNewUser, setIsNewUser] = useState(false);

  const loadUser = (user: PrototypeUser) => {
    setSelectedUsername(user.username);
    setDraft(user);
    setIsNewUser(false);
  };

  const startNewUser = () => {
    setSelectedUsername("");
    setDraft(
      withUserDefaults({
        fullName: "",
        username: "",
        email: "",
        role: "Auditor",
        permissionGroup: "Auditor",
      }),
    );
    setIsNewUser(true);
  };

  const updateDraft = <K extends keyof PrototypeUser>(
    key: K,
    value: PrototypeUser[K],
  ) => {
    const nextDraft = { ...draft, [key]: value };
    if (key === "role") {
      nextDraft.permissionGroup = value as UserRole;
    }
    setDraft(nextDraft);
  };

  return (
    <section className="panel user-management">
      <div className="section-title">
        <div>
          <p className="eyebrow dark">Admin</p>
          <h2>User management</h2>
          <span>Edit prototype users, roles, active status, and visibility.</span>
        </div>
        <div className="segmented">
          <button type="button" onClick={startNewUser}>
            Add user
          </button>
          <button type="button" className="secondary" onClick={onResetUsers}>
            Reset users
          </button>
        </div>
      </div>
      <div className="user-management-grid">
        {pendingRequests.length > 0 && (
          <div className="access-request-queue">
            <div>
              <h3>Account requests</h3>
              <span>{pendingRequests.length} waiting for review</span>
            </div>
            {pendingRequests.map((user) => (
              <article className="access-request-card" key={user.username}>
                <div>
                  <strong>{user.fullName}</strong>
                  <span>{user.email}</span>
                  <small>
                    {user.emailVerified
                      ? "Email confirmed"
                      : "Waiting on email confirmation"}
                  </small>
                </div>
                <div className="storage-actions">
                  <button
                    type="button"
                    disabled={!canApproveAccessRequest(user)}
                    onClick={() => onApproveRequest(user.username)}
                  >
                    Approve
                  </button>
                  <button
                    type="button"
                    className="secondary"
                    onClick={() => onRejectRequest(user.username)}
                  >
                    Reject
                  </button>
                </div>
              </article>
            ))}
          </div>
        )}
        <div className="user-list" aria-label="Prototype users">
          {users.map((user) => (
            <button
              type="button"
              className={
                user.username === selectedUsername ? "selected secondary" : "secondary"
              }
              key={user.username}
              onClick={() => loadUser(user)}
            >
              <strong>{user.fullName}</strong>
              <span className="user-meta">
                <span>{user.role}</span>
                <span
                  className={`user-status-badge ${
                    user.active && user.accessRequestStatus === "Approved"
                      ? "active"
                      : "inactive"
                  }`}
                >
                  {user.accessRequestStatus === "Approved"
                    ? user.active
                      ? "Active"
                      : "Inactive"
                    : user.accessRequestStatus}
                </span>
              </span>
            </button>
          ))}
        </div>
        <form
          className="user-editor"
          onSubmit={(event) => {
            event.preventDefault();
            onSaveUser(isNewUser ? null : selectedUsername, draft);
            setSelectedUsername(draft.username.trim().toLowerCase());
            setIsNewUser(false);
          }}
        >
          <div className="form-grid user-editor-grid">
            <Input
              label="Full name"
              value={draft.fullName}
              onChange={(value) => updateDraft("fullName", value)}
            />
            <Input
              label="Username"
              value={draft.username}
              onChange={(value) => updateDraft("username", value)}
            />
            <Input
              label="Email"
              value={draft.email}
              onChange={(value) => updateDraft("email", value)}
            />
            <Input
              label="Test password"
              value={draft.password}
              onChange={(value) => updateDraft("password", value)}
            />
            <Select
              label="Role"
              value={draft.role}
              options={userRoleOptions}
              placeholder="Select role"
              onChange={(value) => updateDraft("role", value as UserRole)}
            />
            <Select
              label="Default visibility"
              value={draft.defaultVisibility}
              options={projectVisibilityOptions}
              placeholder="Select visibility"
              onChange={(value) =>
                updateDraft("defaultVisibility", value as ProjectVisibility)
              }
            />
            <Check
              label="Active user"
              checked={draft.active}
              onChange={(value) => updateDraft("active", value)}
            />
            <Check
              label="Email verified"
              checked={draft.emailVerified}
              onChange={(value) => updateDraft("emailVerified", value)}
            />
          </div>
          <div className="modal-actions">
            <button type="submit">Save user</button>
          </div>
        </form>
      </div>
    </section>
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
    <section className="summary-grid" aria-label="Audit dashboard summary">
      <SummaryCard
        label="Open projects"
        value={
          projects.filter((project) => project.currentStage !== "Closed").length
        }
      />
      <SummaryCard label="Blocked" value={blocked} tone="danger" />
      <SummaryCard label="Overdue" value={overdue} tone="danger" />
      <SummaryCard label="Due in 3 days" value={dueSoon} tone="warning" />
      <SummaryCard
        label="Pipeline value"
        value={quoteValue.toLocaleString("en-US", {
          style: "currency",
          currency: "USD",
          maximumFractionDigits: 0,
        })}
      />
    </section>
  );
}

function sortByUrgency(projects: AuditProject[]) {
  return projects
    .slice()
    .sort((a, b) => daysUntil(a.dueDate) - daysUntil(b.dueDate));
}

function TodaysWork({
  projects,
  onSelect,
}: {
  projects: AuditProject[];
  onSelect: (id: string) => void;
}) {
  const openProjects = projects.filter((project) => project.currentStage !== "Closed");
  const queues = [
    {
      label: "Overdue",
      tone: "queue-danger",
      items: sortByUrgency(
        openProjects.filter((project) => daysUntil(project.dueDate) < 0),
      ),
    },
    {
      label: "Due soon",
      tone: "queue-warning",
      items: sortByUrgency(
        openProjects.filter(
          (project) =>
            daysUntil(project.dueDate) >= 0 && daysUntil(project.dueDate) <= 3,
        ),
      ),
    },
    {
      label: "Blocked",
      tone: "queue-danger",
      items: sortByUrgency(
        openProjects.filter((project) => projectMatchesWorkState(project, "blocked")),
      ),
    },
    {
      label: "Waiting on broker",
      tone: "queue-warning",
      items: sortByUrgency(
        openProjects.filter((project) =>
          projectMatchesWorkState(project, "waitingOnBroker"),
        ),
      ),
    },
  ];
  return (
    <section className="panel todays-work">
      <div className="section-title">
        <div>
          <p className="eyebrow dark">Today</p>
          <h2>Today's work</h2>
          <span>Priority queues from your visible projects</span>
        </div>
      </div>
      <div className="todays-work-grid">
        {queues.map((queue) => (
          <article className={`todays-work-queue ${queue.tone}`} key={queue.label}>
            <div className="queue-heading">
              <strong>{queue.label}</strong>
              <span>{queue.items.length}</span>
            </div>
            {queue.items.length === 0 ? (
              <p>No items</p>
            ) : (
              <div className="queue-list">
                {queue.items.slice(0, 4).map((project) => {
                  const due = dueLabel(project);
                  return (
                    <button
                      type="button"
                      className="queue-item"
                      key={project.id}
                      onClick={() => onSelect(project.id)}
                    >
                      <strong>{project.assignmentNumber}</strong>
                      <span>{project.auditEntity || project.clientCoverholderCode}</span>
                      <small className={`pill ${due.className}`}>{due.text}</small>
                    </button>
                  );
                })}
              </div>
            )}
          </article>
        ))}
      </div>
    </section>
  );
}

function CycleTimeDashboard({
  projects,
  range,
  setRange,
}: {
  projects: AuditProject[];
  range: DurationRange;
  setRange: (range: DurationRange) => void;
}) {
  const metrics = stageDurationMetrics(projects, range);
  const rangeLabel =
    range === "ytd" ? "YTD" : range === "90d" ? "last 3 months" : "last week";
  return (
    <section className="panel cycle-dashboard">
      <div className="section-title">
        <div>
          <p className="eyebrow dark">Cycle time</p>
          <h2>Average time in stage</h2>
          <span>Lifecycle movement timing for {rangeLabel}</span>
        </div>
        <div className="segmented">
          <button
            type="button"
            className={range === "ytd" ? "active" : "secondary"}
            onClick={() => setRange("ytd")}
          >
            YTD
          </button>
          <button
            type="button"
            className={range === "90d" ? "active" : "secondary"}
            onClick={() => setRange("90d")}
          >
            3 months
          </button>
          <button
            type="button"
            className={range === "7d" ? "active" : "secondary"}
            onClick={() => setRange("7d")}
          >
            1 week
          </button>
        </div>
      </div>
      {metrics.length === 0 ? (
        <p>No stage moves in this period yet.</p>
      ) : (
        <div className="cycle-grid">
          {metrics.map((metric) => (
            <article className="cycle-card" key={metric.stage}>
              <span>{metric.stage}</span>
              <strong>{metric.average.toFixed(1)}d</strong>
              <small>{metric.count} move{metric.count === 1 ? "" : "s"}</small>
            </article>
          ))}
        </div>
      )}
    </section>
  );
}

function HoverLink({ label, helper }: { label: string; helper: string }) {
  return (
    <span className="hover-link" tabIndex={0}>
      {label}
      <span role="tooltip">{helper}</span>
    </span>
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
        projectHasAuditor(project, auditor) &&
        project.currentStage !== "Closed",
    );
    const blockedCount = openProjects.filter(
      (project) => computedBlockers(project).length > 0,
    ).length;
    const dueSoonCount = openProjects.filter(
      (project) =>
        daysUntil(project.dueDate) >= 0 && daysUntil(project.dueDate) <= 7,
    ).length;
    const leadCount = openProjects.filter(
      (project) => auditTeamRole(project, auditor) === "Lead Auditor",
    ).length;
    const supportCount = openProjects.filter(
      (project) => auditTeamRole(project, auditor) === "Supporting Auditor",
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
          ? "High load"
          : openProjects.length >= 2
            ? "Moderate load"
            : "Light load";
    return {
      auditor,
      openCount: openProjects.length,
      blockedCount,
      dueSoonCount,
      leadCount,
      supportCount,
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
          <p className="eyebrow dark">Team workload</p>
          <h2>Open assignments by auditor</h2>
        </div>
        <div className="workload-summary">
          <strong>{totalOpen}</strong>
          <span>open auditor assignments</span>
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
                  {row.openCount} open · {row.leadCount} lead · {row.supportCount} support ·{" "}
                  {row.blockedCount} blocked
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
                Due in 3 days
              </span>
              <span className="workload-stat">
                <strong>{row.leadCount}</strong>
                Lead
              </span>
              <span className="workload-stat">
                <strong>{row.supportCount}</strong>
                Support
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
        <h2>Audit table</h2>
        <span>{projects.length} rows</span>
      </div>
      <div className="table-wrap">
        <table>
          <thead>
            <tr>
              <th>Assignment</th>
              <th>Type</th>
              <th>Audit entity</th>
              <th>Audit team</th>
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
                  <td>{formatAuditTeam(project)}</td>
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
  presets,
  auditors,
}: {
  filters: Filters;
  setFilters: (filters: Filters) => void;
  presets: FilterPreset[];
  auditors: string[];
}) {
  const activeFilterCount = Object.values(filters).filter(Boolean).length;
  return (
    <section className="panel">
      <div className="section-title">
        <div>
          <h2>Filters</h2>
          <span>{activeFilterCount} active filter{activeFilterCount === 1 ? "" : "s"}</span>
        </div>
        <button
          className="link"
          onClick={() => setFilters(buildFilters())}
        >
          Clear filters
        </button>
      </div>
      <div className="saved-views" aria-label="Saved filter views">
        <span>Saved views</span>
        {presets.map((preset) => (
          <button
            type="button"
            className="secondary"
            key={preset.id}
            onClick={() => setFilters(buildFilters(preset.filters))}
          >
            {preset.label}
          </button>
        ))}
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
            ["dueSoon", "Due in 3 days"],
          ]}
          onChange={(value) => setFilters({ ...filters, dueDate: value })}
        />
        <Select
          label="Work state"
          value={filters.workState}
          options={[
            ["blocked", "Blocked"],
            ["waitingOnBroker", "Waiting on broker"],
            ["pendingPayment", "Pending payment"],
          ]}
          onChange={(value) => setFilters({ ...filters, workState: value })}
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
  onRemoveLabel,
  currentUser,
}: {
  projects: AuditProject[];
  selectedId?: string;
  onSelect: (id: string) => void;
  onMove: (project: AuditProject, stage: Stage) => void;
  onRemoveLabel: (project: AuditProject, label: ProjectLabel) => void;
  currentUser: PrototypeUser;
}) {
  const handleDrop = (projectId: string, targetStage: Stage) => {
    const project = projects.find((item) => item.id === projectId);
    if (project) onMove(project, targetStage);
  };

  return (
    <section className="panel">
      <div className="section-title">
        <h2>Lifecycle board</h2>
        <span>{projects.length} visible · drag cards between stages</span>
      </div>
      <div className="kanban">
        {stages.map((stage) => {
          const stageProjects = projects.filter(
            (project) => project.currentStage === stage,
          );
          return (
            <div
              className="column"
              key={stage}
              onDragOver={(event) => event.preventDefault()}
              onDrop={(event) => {
                event.preventDefault();
                handleDrop(event.dataTransfer.getData("text/plain"), stage);
              }}
            >
              <div className="column-heading">
                <h3>{stage}</h3>
                <span>{stageProjects.length}</span>
              </div>
              {stageProjects.map((project) => {
                const due = dueLabel(project);
                const editable = canEditProject(currentUser, project);
                return (
                  <article
                    draggable={editable}
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
                      {formatAuditTeam(project)}
                    </span>
                    <span className="pill muted">{project.assignmentType}</span>
                    {project.labels.map((label) => (
                      <LabelChip
                        label={label}
                        key={label}
                        compact
                        onRemove={
                          editable
                            ? (event) => {
                                event.stopPropagation();
                                onRemoveLabel(project, label);
                              }
                            : undefined
                        }
                      />
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
                      disabled={!editable}
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
          );
        })}
      </div>
    </section>
  );
}

function labelClassName(label: ProjectLabel) {
  return label.toLowerCase().replace(/ /g, "-");
}

function LabelChip({
  label,
  compact = false,
  onRemove,
}: {
  label: ProjectLabel;
  compact?: boolean;
  onRemove?: (event: React.MouseEvent<HTMLButtonElement>) => void;
}) {
  return (
    <span
      className={`project-label removable ${compact ? "mini" : ""} ${labelClassName(
        label,
      )}`}
    >
      {label}
      {onRemove && (
        <button
          type="button"
          aria-label={`Remove ${label}`}
          onClick={onRemove}
          onMouseDown={(event) => event.stopPropagation()}
        >
          ×
        </button>
      )}
    </span>
  );
}

function ProjectDetail({
  project,
  onEdit,
  onMove,
  onRemoveLabel,
  onAddComment,
  onToggleChecklist,
  onDocumentWorkflowAction,
  auditors,
  onAddSupportingAuditor,
  currentUser,
  onUpdateFinance,
}: {
  project: AuditProject;
  onEdit: () => void;
  onMove: (project: AuditProject, stage: Stage) => void;
  onRemoveLabel: (project: AuditProject, label: ProjectLabel) => void;
  onAddComment: (project: AuditProject, comment: ProjectComment) => void;
  onToggleChecklist: (project: AuditProject, key: string) => void;
  onDocumentWorkflowAction: (
    project: AuditProject,
    action: DocumentWorkflowAction,
  ) => void;
  auditors: string[];
  onAddSupportingAuditor: (project: AuditProject, auditor: string) => void;
  currentUser: PrototypeUser;
  onUpdateFinance: (
    project: AuditProject,
    invoiceStatus: InvoiceStatus,
    paymentReceived: boolean,
  ) => void;
}) {
  const blockers = computedBlockers(project);
  const nextSteps = recommendedNextSteps(project);
  const canEdit = canEditProject(currentUser, project);
  const canManageFinance =
    canUpdateFinance(currentUser, project) || hasFullProjectAccess(currentUser);
  return (
    <section className="detail-grid">
      <article className="panel detail">
        <div className="section-title">
          <h2>{project.assignmentNumber}</h2>
          <button disabled={!canEdit} onClick={onEdit}>Edit project</button>
        </div>
        {project.labels.length > 0 && (
          <div className="label-strip">
            {project.labels.map((label) => (
              <LabelChip
                label={label}
                key={label}
                onRemove={canEdit ? () => onRemoveLabel(project, label) : undefined}
              />
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
          <Meta label="Audit team" value={formatAuditTeam(project)} />
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
        <div className="next-action-heading">
          <h3>Next action</h3>
          <HoverLink
            label="Why these steps?"
            helper="Recommended next steps combine blockers, due date, document readiness, stage, and the recorded next action."
          />
        </div>
        <p>{project.nextAction || "No next action recorded."}</p>
        <div className="recommended-next">
          <strong>Recommended next steps</strong>
          <ul>
            {nextSteps.map((step) => (
              <li key={step}>{step}</li>
            ))}
          </ul>
        </div>
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
              disabled={!canEdit}
              onChange={(event) => onMove(project, event.target.value as Stage)}
            >
              {stages.map((stage) => (
                <option key={stage}>{stage}</option>
              ))}
            </select>
          </label>
        </div>
      </article>
      <AuditTeamPanel
        project={project}
        auditors={auditors}
        onAddSupportingAuditor={onAddSupportingAuditor}
        canManageTeam={hasFullProjectAccess(currentUser)}
      />
      <DocumentReadiness
        project={project}
        onDocumentWorkflowAction={onDocumentWorkflowAction}
        canUpdateDocuments={canEdit}
      />
      <FinancePanel
        key={project.id}
        project={project}
        canUpdateFinance={canManageFinance}
        onUpdateFinance={onUpdateFinance}
      />
      <Checklist
        project={project}
        onToggleChecklist={onToggleChecklist}
        canUpdateChecklist={canEdit}
      />
      <Comments
        project={project}
        currentUser={currentUser}
        canAddComment={canComment(currentUser, project)}
        onAddComment={onAddComment}
      />
      <TemplateLibrary project={project} />
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

function AuditTeamPanel({
  project,
  auditors,
  onAddSupportingAuditor,
  canManageTeam,
}: {
  project: AuditProject;
  auditors: string[];
  onAddSupportingAuditor: (project: AuditProject, auditor: string) => void;
  canManageTeam: boolean;
}) {
  const [selectedAuditor, setSelectedAuditor] = useState(
    auditors.find((auditor) => !projectHasAuditor(project, auditor)) ?? "",
  );
  const availableAuditors = auditors.filter(
    (auditor) => !projectHasAuditor(project, auditor),
  );
  const selected = selectedAuditor || availableAuditors[0] || "";
  return (
    <article className="panel team-panel">
      <div className="section-title">
        <div>
          <h2>Audit team</h2>
          <span>Add help after intake when someone jumps in.</span>
        </div>
      </div>
      <div className="team-list">
        {normalizeAuditTeam(project).map((member) => (
          <span className="team-member" key={member.person}>
            <strong>{member.person}</strong>
            <small>{member.role}</small>
          </span>
        ))}
      </div>
      <div className="add-support-row">
        <label>
        Add supporting auditor
          <select
            value={selected}
            disabled={!canManageTeam || availableAuditors.length === 0}
            onChange={(event) => setSelectedAuditor(event.target.value)}
          >
            {availableAuditors.length === 0 ? (
              <option value="">All auditors are already assigned</option>
            ) : (
              availableAuditors.map((auditor) => (
                <option key={auditor} value={auditor}>
                  {auditor}
                </option>
              ))
            )}
          </select>
        </label>
        <button
          type="button"
          disabled={!canManageTeam || !selected}
          onClick={() => {
            onAddSupportingAuditor(project, selected);
            setSelectedAuditor(
              availableAuditors.find((auditor) => auditor !== selected) ?? "",
            );
          }}
        >
          Add support
        </button>
      </div>
    </article>
  );
}

function DocumentReadiness({
  project,
  onDocumentWorkflowAction,
  canUpdateDocuments,
}: {
  project: AuditProject;
  onDocumentWorkflowAction: (
    project: AuditProject,
    action: DocumentWorkflowAction,
  ) => void;
  canUpdateDocuments: boolean;
}) {
  const readiness = documentReadiness(project);
  const documentsComplete = readiness.percent === 100;
  const waitingOnBroker = readiness.waitingOnBroker;
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
        {requiredDocuments.map((doc) => {
          const shortLabel = doc.label.replace(" received", "");
          const complete = Boolean(project[doc.key]);
          return (
            <span
              key={doc.key}
              className={`document-chip ${complete ? "complete" : "pending"}`}
            >
              {complete ? "Complete" : "Missing"}: {shortLabel}
            </span>
          );
        })}
        <span
          className={`document-chip ${
            project.preAuditQuestionnaireStatus === "Complete" ? "complete" : "pending"
          }`}
        >
          {project.preAuditQuestionnaireStatus === "Complete" ? "Complete" : project.preAuditQuestionnaireStatus}: Questionnaire
        </span>
        <span
          className={`document-chip ${
            project.documentRequestStatus === "Complete" ? "complete" : "pending"
          }`}
        >
          {project.documentRequestStatus === "Complete" ? "Complete" : project.documentRequestStatus}: Document request
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
          disabled={!canUpdateDocuments || waitingOnBroker}
          onClick={() =>
            onDocumentWorkflowAction(project, "markWaitingOnBroker")
          }
        >
          {waitingOnBroker ? "Waiting on broker" : "Mark waiting on broker"}
        </button>
        <button
          type="button"
          className="secondary"
          disabled={!canUpdateDocuments || project.brokerLastChasedDate === todayIso()}
          onClick={() => onDocumentWorkflowAction(project, "recordBrokerChase")}
        >
          {project.brokerLastChasedDate === todayIso()
            ? "Broker chased today"
            : "Record broker chase"}
        </button>
        <button
          type="button"
          className="secondary"
          disabled={!canUpdateDocuments || !waitingOnBroker}
          onClick={() => onDocumentWorkflowAction(project, "clearWaitingOnBroker")}
        >
          {waitingOnBroker ? "Clear waiting label" : "Waiting label cleared"}
        </button>
        <button
          type="button"
          disabled={!canUpdateDocuments || documentsComplete}
          onClick={() =>
            onDocumentWorkflowAction(project, "markDocumentsComplete")
          }
        >
          {documentsComplete ? "Documents complete" : "Mark documents complete"}
        </button>
      </div>
    </article>
  );
}

function FinancePanel({
  project,
  canUpdateFinance,
  onUpdateFinance,
}: {
  project: AuditProject;
  canUpdateFinance: boolean;
  onUpdateFinance: (
    project: AuditProject,
    invoiceStatus: InvoiceStatus,
    paymentReceived: boolean,
  ) => void;
}) {
  const [invoiceStatus, setInvoiceStatus] = useState(project.invoiceStatus);
  const [paymentReceived, setPaymentReceived] = useState(project.paymentReceived);

  return (
    <article className="panel finance-panel">
      <div className="section-title">
        <div>
          <h2>Finance</h2>
          <span>Invoice and payment status</span>
        </div>
      </div>
      <div className="finance-grid">
        <Select
          label="Invoice status"
          value={invoiceStatus}
          options={["Not Started", "Prepared", "Sent", "Paid"]}
          placeholder="Select invoice status"
          onChange={(value) => {
            const nextStatus = value as InvoiceStatus;
            setInvoiceStatus(nextStatus);
            if (nextStatus === "Paid") setPaymentReceived(true);
          }}
        />
        <Check
          label="Payment received"
          checked={paymentReceived}
          onChange={setPaymentReceived}
        />
      </div>
      <button
        type="button"
        disabled={!canUpdateFinance}
        onClick={() =>
          onUpdateFinance(project, invoiceStatus, paymentReceived)
        }
      >
        Save finance
      </button>
    </article>
  );
}

function TemplateLibrary({ project }: { project: AuditProject }) {
  const [selectedTemplateId, setSelectedTemplateId] = useState(
    communicationTemplates[0].id,
  );
  const [copyMessage, setCopyMessage] = useState("");
  const template =
    communicationTemplates.find((item) => item.id === selectedTemplateId) ??
    communicationTemplates[0];
  const subject = template.subject(project);
  const body = template.body(project);

  const copyText = async (label: string, value: string) => {
    try {
      await navigator.clipboard.writeText(value);
      setCopyMessage(`${label} copied.`);
    } catch {
      setCopyMessage("Copy failed. Select the preview text manually.");
    }
  };

  return (
    <article className="panel template-panel">
      <div className="section-title">
        <div>
          <p className="eyebrow dark">Templates</p>
          <h2>Email and document templates</h2>
          <span>{template.purpose}</span>
        </div>
      </div>
      <div className="template-controls">
        <Select
          label="Template"
          value={selectedTemplateId}
          options={communicationTemplates.map((item) => [
            item.id,
            `${item.label} (${item.kind})`,
          ])}
          placeholder="Select template"
          onChange={(value) =>
            setSelectedTemplateId(value || communicationTemplates[0].id)
          }
        />
      </div>
      <div className="template-preview">
        <label>
          Subject
          <input readOnly value={subject} />
        </label>
        <label>
          Body
          <textarea readOnly rows={10} value={body} />
        </label>
      </div>
      <div className="template-actions">
        <button
          type="button"
          className="secondary"
          onClick={() => void copyText("Subject", subject)}
        >
          Copy subject
        </button>
        <button type="button" onClick={() => void copyText("Body", body)}>
          Copy body
        </button>
        {copyMessage && <span>{copyMessage}</span>}
      </div>
    </article>
  );
}

function ActivityTimeline({ project }: { project: AuditProject }) {
  const items = activityTimeline(project);
  const [query, setQuery] = useState("");
  const [typeFilter, setTypeFilter] = useState("");
  const normalizedQuery = query.trim().toLowerCase();
  const filteredItems = items.filter((item) => {
    if (typeFilter && item.type !== typeFilter) return false;
    if (!normalizedQuery) return true;
    return [item.type, item.title, item.detail, item.timestamp]
      .join(" ")
      .toLowerCase()
      .includes(normalizedQuery);
  });
  return (
    <article className="panel activity-panel">
      <div className="section-title">
        <div>
          <h2>Audit trail</h2>
          <span>
            Filterable history of comments, stage moves, checklist changes, and
            document follow-ups
          </span>
        </div>
      </div>
      {items.length === 0 ? (
        <p>No activity yet.</p>
      ) : (
        <>
          <div className="activity-filters">
            <Input
              label="Search audit trail"
              value={query}
              onChange={setQuery}
              placeholder="Search dates, users, actions, or notes"
            />
            <Select
              label="Event type"
              value={typeFilter}
              options={activityTypeOptions}
              placeholder="All event types"
              onChange={setTypeFilter}
            />
          </div>
          <div className="activity-summary">
            <span>{filteredItems.length} shown</span>
            <span>{items.length} total</span>
            <span>{items.filter((item) => item.type === "document").length} document</span>
            <span>{items.filter((item) => item.type === "team").length} team</span>
            <span>{items.filter((item) => item.type === "stage").length} stage</span>
          </div>
          {filteredItems.length === 0 ? (
            <p>No activity matches the current search.</p>
          ) : (
            <div className="activity-list">
              {filteredItems.map((item) => (
                <div className={`activity-item ${item.type}`} key={item.id}>
                  <div className="activity-marker" aria-hidden="true" />
                  <div>
                    <div className="activity-heading">
                      <span className={`activity-type ${item.tone || "muted"}`}>
                        {item.type}
                      </span>
                      <small>{item.timestamp}</small>
                    </div>
                    <strong>{item.title}</strong>
                    <p>{item.detail}</p>
                  </div>
                </div>
              ))}
            </div>
          )}
        </>
      )}
    </article>
  );
}

function Checklist({
  project,
  onToggleChecklist,
  canUpdateChecklist,
}: {
  project: AuditProject;
  onToggleChecklist: (project: AuditProject, key: string) => void;
  canUpdateChecklist: boolean;
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
    <article className="panel checklist-panel">
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
                  disabled={!canUpdateChecklist}
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
  currentUser,
  canAddComment,
  onAddComment,
}: {
  project: AuditProject;
  currentUser: PrototypeUser;
  canAddComment: boolean;
  onAddComment: (project: AuditProject, comment: ProjectComment) => void;
}) {
  const [commentBody, setCommentBody] = useState("");
  const addComment = (event: FormEvent) => {
    event.preventDefault();
    const body = commentBody.trim();
    if (!body || !canAddComment) return;
    onAddComment(project, {
      id: `comment-${Date.now()}`,
      createdAt: timestampNow(),
      author: currentUser.fullName,
      body,
    });
    setCommentBody("");
  };
  return (
    <article className="panel comments-panel">
      <h2>Card comments</h2>
      <form className="comment-form" onSubmit={addComment}>
        <span className="comment-author">Commenting as {currentUser.fullName}</span>
        <label>
          Comment
          <textarea
            value={commentBody}
            disabled={!canAddComment}
            placeholder="Add an update, question, or note for this audit card"
            onChange={(event) => setCommentBody(event.target.value)}
          />
        </label>
        <button type="submit" disabled={!canAddComment}>Add comment</button>
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
  const draftAuditTeam = normalizeAuditTeam(draft);
  const leadAuditor = primaryAuditor(draft);
  const updateLeadAuditor = (auditor: string) => {
    const supportingTeam = draftAuditTeam
      .filter((member) => member.person !== auditor)
      .map((member) => ({
        ...member,
        role:
          member.role === "Lead Auditor"
            ? ("Supporting Auditor" as AuditTeamRole)
            : member.role,
      }));
    setDraft({
      ...draft,
      assignedAuditor: auditor,
      auditTeam: auditor
        ? [{ person: auditor, role: "Lead Auditor" }, ...supportingTeam]
        : supportingTeam,
    });
  };
  const toggleSupportingAuditor = (auditor: string) => {
    if (auditor === leadAuditor) return;
    const existing = draftAuditTeam.some((member) => member.person === auditor);
    setDraft({
      ...draft,
      auditTeam: existing
        ? draftAuditTeam.filter((member) => member.person !== auditor)
        : [
            ...draftAuditTeam,
            { person: auditor, role: "Supporting Auditor" },
          ],
    });
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
          Assign ownership so the card immediately appears in workload reporting
          and auditor filters.
        </p>
      </div>
      <div className="form-grid wizard-grid">
        <Select
          label="Lead auditor"
          value={leadAuditor}
          options={auditorOptions}
          placeholder="Select lead auditor"
          onChange={updateLeadAuditor}
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
      <div className="team-picker">
        <span>Supporting auditors</span>
        <p>Choose any additional auditors working this assignment with the lead.</p>
        <div>
          {auditorOptions
            .filter((auditor) => auditor !== leadAuditor)
            .map((auditor) => {
              const selected = draftAuditTeam.some(
                (member) => member.person === auditor,
              );
              return (
                <button
                  type="button"
                  key={auditor}
                  className={selected ? "team-option selected" : "team-option"}
                  onClick={() => toggleSupportingAuditor(auditor)}
                >
                  {selected ? "✓ " : "+ "}
                  {auditor}
                </button>
              );
            })}
        </div>
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
        <Meta label="Audit team" value={formatAuditTeam(draft)} />
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
