import React, { FormEvent, useEffect, useMemo, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import {
  canApproveAccessRequest,
  type AccountRequestStatus,
} from "./accessRequests";
import {
  approveSecureAccessRequest,
  createOutlookCalendarEvent,
  getLinkedContactSources,
  getSecureAccessState,
  getSecureSystemHealth,
  logoutSecureAccess,
  rejectSecureAccessRequest,
  secureAccessUrl,
  updateSecureAccessUser,
  verifySecureAccessCode,
  type AccessApprovalUpdate,
  type LinkedContact,
  type LinkedContactSourcesResponse,
  type SecureAccessState,
  type SecureAccessUser,
  type SecureSystemHealth,
} from "./secureAccessClient";
import {
  buildMicrosoftListsMigrationPackage,
} from "./microsoftListsSchema";
import {
  getServerProjects,
  saveServerProjects,
} from "./projectDataClient";
import mosaicLogoUrl from "./assets/mosaic-logo-transparent.png";
import "./styles.css";

export type AssignmentSource = "Email" | "DAM";
export type AssignmentType = "DCA" | "CH" | "MGA" | "Company Contract";
export type AuditType = "Remote" | "Onsite";
export type AuditStructure = "Solo" | "Coordinated";
export type CalendarSyncStatus =
  | "Not Synced"
  | "Ready to Sync"
  | "Synced"
  | "Conflict Review";
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

export type ProjectDocumentKey =
  | "baaReceived"
  | "endorsementsReceived"
  | "premiumBdxReceived"
  | "dcaAgreementReceived"
  | "claimsBdxReceived";

export type RequiredDocument = {
  key: ProjectDocumentKey;
  label: string;
};

export type ManagingAgentWorkstream = {
  id: string;
  managingAgentName: string;
  managingAgentCode: string;
  leadAuditor: string;
  supportAuditors: string[];
  currentStage: Stage;
  assignmentStatus: AssignmentStatus;
  dueDate: string;
  documentRequestStatus: ProgressStatus;
  baaReceived: boolean;
  endorsementsReceived: boolean;
  premiumBdxReceived: boolean;
  dcaAgreementReceived: boolean;
  claimsBdxReceived: boolean;
  blockers: string;
  nextAction: string;
  completed: boolean;
  waived: boolean;
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
  auditStructure: AuditStructure;
  managingAgentWorkstreams: ManagingAgentWorkstream[];
  currentStage: Stage;
  assignmentStatus: AssignmentStatus;
  quoteStatus: QuoteStatus;
  quoteAmount: number;
  tentativeAuditWeek: string;
  confirmedAuditDate: string;
  schedulingNotes: string;
  calendarSyncStatus: CalendarSyncStatus;
  calendarEventId: string;
  calendarEventWebLink: string;
  calendarEventLastSyncedAt: string;
  auditLocation: string;
  auditRemoteLink: string;
  auditDurationHours: number;
  auditStartTime: string;
  linkedContactId: string;
  linkedContactSource: string;
  auditType: AuditType;
  baaReceived: boolean;
  endorsementsReceived: boolean;
  premiumBdxReceived: boolean;
  dcaAgreementReceived: boolean;
  claimsBdxReceived: boolean;
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
  archived?: boolean;
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
type AppSection =
  | "dashboard"
  | "assignments"
  | "scheduling"
  | "command"
  | "reports"
  | "admin";
type AdminTab = "users" | "contacts" | "activity" | "storage" | "health";
type UserRole = "Admin" | "Audit Manager" | "Auditor" | "Finance" | "Read Only";
type ProjectVisibility =
  | "Role Default"
  | "All Projects"
  | "Assigned Projects"
  | "Finance Records";

type PrototypeUser = {
  fullName: string;
  username: string;
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

type TemplateReceiverKind =
  | "DCA contact"
  | "Coverholder contact"
  | "Invoice contact"
  | "Report contact"
  | "Project contact";

type TemplateReceiver = {
  kind: TemplateReceiverKind;
  name: string;
  email: string;
  company: string;
  source: string;
  confidence: "Matched workbook" | "Needs review";
  guidance: string;
};

type WorkflowGate = {
  label: string;
  status: "Ready" | "Blocked" | "Watch";
  detail: string;
};

type SlaSignal = {
  level: "Critical" | "Warning" | "Normal";
  label: string;
  detail: string;
};

type OperationsDraft = {
  id: string;
  projectId: string;
  assignmentNumber: string;
  auditEntity: string;
  templateId: string;
  label: string;
  kind: CommunicationTemplate["kind"];
  priority: "High" | "Medium" | "Low";
  reason: string;
  subject: string;
  body: string;
};

type OperationsBrief = {
  summary: string;
  risks: string[];
  actions: string[];
  managerFocus: string[];
  aiPrompt: string;
};

type CoordinatorInsight = {
  id: string;
  projectId: string;
  assignmentNumber: string;
  auditEntity: string;
  priority: "Critical" | "High" | "Medium" | "Low";
  title: string;
  reason: string;
  recommendedAction: string;
};

type DocumentIntelligenceResult = {
  projectId: string;
  assignmentNumber: string;
  auditEntity: string;
  packageType: string;
  readiness: number;
  confidence: "High" | "Medium" | "Needs Review";
  missing: string[];
  evidence: string[];
  recommendations: string[];
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
const usersStorageKey = "audit-assignment-tracker-users-v1";
const lastExportStorageKey = "audit-assignment-tracker-last-export-v1";

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

const auditStructureOptions: AuditStructure[] = ["Solo", "Coordinated"];
const calendarSyncStatusOptions: CalendarSyncStatus[] = [
  "Not Synced",
  "Ready to Sync",
  "Synced",
  "Conflict Review",
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
    role: "Auditor",
    permissionGroup: "Auditor",
    email: "lorraine.mojica@[company-domain]",
    active: true,
    defaultVisibility: "Role Default",
  }),
  approvedPrototypeUser({
    fullName: "Walter Aviles",
    username: "walter.aviles",
    role: "Auditor",
    permissionGroup: "Auditor",
    email: "walter.aviles@[company-domain]",
    active: true,
    defaultVisibility: "Role Default",
  }),
  approvedPrototypeUser({
    fullName: "Leslie Domenech",
    username: "leslie.domenech",
    role: "Auditor",
    permissionGroup: "Auditor",
    email: "leslie.domenech@[company-domain]",
    active: true,
    defaultVisibility: "Role Default",
  }),
  approvedPrototypeUser({
    fullName: "Mark James",
    username: "mark.james",
    role: "Audit Manager",
    permissionGroup: "Audit Manager",
    email: "mark.james@[company-domain]",
    active: true,
    defaultVisibility: "Role Default",
  }),
  approvedPrototypeUser({
    fullName: "Justin Mojica",
    username: "justin.mojica",
    role: "Admin",
    permissionGroup: "Admin",
    email: "justin.mojica@[company-domain]",
    active: true,
    defaultVisibility: "Role Default",
  }),
  approvedPrototypeUser({
    fullName: "Sheilah Couture",
    username: "sheilah.couture",
    role: "Finance",
    permissionGroup: "Finance",
    email: "sheilah.couture@[company-domain]",
    active: true,
    defaultVisibility: "Role Default",
  }),
  approvedPrototypeUser({
    fullName: "Annabelle J. Crawford Mojica",
    username: "annabelle.crawford.mojica",
    role: "Read Only",
    permissionGroup: "Read Only",
    email: "annabelle.crawford.mojica@[company-domain]",
    active: true,
    defaultVisibility: "Role Default",
  }),
  approvedPrototypeUser({
    fullName: "Molly Aviles",
    username: "molly.aviles",
    role: "Auditor",
    permissionGroup: "Auditor",
    email: "molly.aviles@[company-domain]",
    active: true,
    defaultVisibility: "Role Default",
  }),
  approvedPrototypeUser({
    fullName: "Lindsie Guillermo",
    username: "lindsie.guillermo",
    role: "Auditor",
    permissionGroup: "Auditor",
    email: "lindsie.guillermo@[company-domain]",
    active: true,
    defaultVisibility: "Role Default",
  }),
];

const legacySampleProjectIds = new Set([
  "audit-001",
  "audit-002",
  "audit-003",
  "audit-004",
]);

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
  auditStructure: "Solo",
  managingAgentWorkstreams: [],
  currentStage: "Intake",
  assignmentStatus: "New",
  quoteStatus: "Not Started",
  quoteAmount: 0,
  tentativeAuditWeek: "",
  confirmedAuditDate: "",
  schedulingNotes: "",
  calendarSyncStatus: "Not Synced",
  calendarEventId: "",
  calendarEventWebLink: "",
  calendarEventLastSyncedAt: "",
  auditLocation: "",
  auditRemoteLink: "",
  auditDurationHours: 1,
  auditStartTime: "09:00",
  linkedContactId: "",
  linkedContactSource: "",
  auditType: "Remote",
  baaReceived: false,
  endorsementsReceived: false,
  premiumBdxReceived: false,
  dcaAgreementReceived: false,
  claimsBdxReceived: false,
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
  archived: false,
});

export const requiredDocuments: readonly RequiredDocument[] = [
  { key: "baaReceived", label: "BAA received" },
  { key: "endorsementsReceived", label: "Endorsements received" },
  { key: "premiumBdxReceived", label: "Premium BDX received" },
] as const;

export const dcaRequiredDocuments: readonly RequiredDocument[] = [
  { key: "dcaAgreementReceived", label: "DCA Agreement received" },
  { key: "claimsBdxReceived", label: "Claims BDX received" },
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
    "Confirm audit owner",
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
    "Complete quality check",
    "Resolve report comments",
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

function isDcaProject(project: Pick<AuditProject, "assignmentType">) {
  return project.assignmentType === "DCA";
}

export function requiredDocumentsForProject(
  project: Pick<AuditProject, "assignmentType">,
) {
  return isDcaProject(project) ? dcaRequiredDocuments : requiredDocuments;
}

function defaultWorkstreamFromProject(project: AuditProject): ManagingAgentWorkstream {
  const team = normalizeAuditTeam(project);
  return {
    id: `${project.id || `audit-${Date.now()}`}-ma-1`,
    managingAgentName:
      project.broker?.trim() ||
      project.auditEntity?.trim() ||
      (isDcaProject(project) ? "Managing agent" : "Primary workstream"),
    managingAgentCode: project.clientCoverholderCode ?? "",
    leadAuditor: primaryAuditor(project),
    supportAuditors: team
      .filter((member) => member.role === "Supporting Auditor")
      .map((member) => member.person),
    currentStage: project.currentStage ?? "Intake",
    assignmentStatus: project.assignmentStatus ?? "New",
    dueDate: project.dueDate ?? "",
    documentRequestStatus: project.documentRequestStatus ?? "Not Started",
    baaReceived: project.baaReceived ?? false,
    endorsementsReceived: project.endorsementsReceived ?? false,
    premiumBdxReceived: project.premiumBdxReceived ?? false,
    dcaAgreementReceived: project.dcaAgreementReceived ?? false,
    claimsBdxReceived: project.claimsBdxReceived ?? false,
    blockers: project.blockers ?? "",
    nextAction: project.nextAction ?? "",
    completed:
      project.assignmentStatus === "Completed" || project.currentStage === "Closed",
    waived: false,
  };
}

function normalizeManagingAgentWorkstreams(
  project: AuditProject,
): ManagingAgentWorkstream[] {
  const source =
    project.managingAgentWorkstreams?.length > 0
      ? project.managingAgentWorkstreams
      : [defaultWorkstreamFromProject(project)];
  const fallback = defaultWorkstreamFromProject(project);
  return source.map((workstream, index) => ({
    ...fallback,
    ...workstream,
    id: workstream.id || `${project.id || `audit-${Date.now()}`}-ma-${index + 1}`,
    managingAgentName:
      workstream.managingAgentName?.trim() ||
      (index === 0 ? fallback.managingAgentName : `Managing agent ${index + 1}`),
    managingAgentCode: workstream.managingAgentCode ?? "",
    leadAuditor: workstream.leadAuditor || fallback.leadAuditor,
    supportAuditors: (workstream.supportAuditors ?? fallback.supportAuditors)
      .filter((auditor) => auditor.trim())
      .filter((auditor, itemIndex, list) => list.indexOf(auditor) === itemIndex),
    currentStage: workstream.currentStage ?? fallback.currentStage,
    assignmentStatus: workstream.assignmentStatus ?? fallback.assignmentStatus,
    dueDate: workstream.dueDate ?? fallback.dueDate,
    documentRequestStatus:
      workstream.documentRequestStatus ?? fallback.documentRequestStatus,
    baaReceived: workstream.baaReceived ?? fallback.baaReceived,
    endorsementsReceived:
      workstream.endorsementsReceived ?? fallback.endorsementsReceived,
    premiumBdxReceived:
      workstream.premiumBdxReceived ?? fallback.premiumBdxReceived,
    dcaAgreementReceived:
      workstream.dcaAgreementReceived ?? fallback.dcaAgreementReceived,
    claimsBdxReceived: workstream.claimsBdxReceived ?? fallback.claimsBdxReceived,
    blockers: workstream.blockers ?? "",
    nextAction: workstream.nextAction ?? "",
    completed:
      workstream.completed ??
      (workstream.assignmentStatus === "Completed" ||
        workstream.currentStage === "Closed"),
    waived: workstream.waived ?? false,
  }));
}

function getMissingDocumentsForWorkstream(
  project: AuditProject,
  workstream: ManagingAgentWorkstream,
) {
  return requiredDocumentsForProject(project)
    .filter((doc) => !workstream[doc.key])
    .map((doc) => doc.label);
}

export function coordinatedWorkstreamSummary(project: AuditProject) {
  const workstreams = normalizeManagingAgentWorkstreams(project);
  const active = workstreams.filter(
    (workstream) => !workstream.completed && !workstream.waived,
  );
  const blocked = active.filter(
    (workstream) =>
      workstream.assignmentStatus === "Blocked" ||
      Boolean(workstream.blockers.trim()),
  );
  const missingDocs = active.filter(
    (workstream) => getMissingDocumentsForWorkstream(project, workstream).length > 0,
  );
  const dueSoon = active.filter((workstream) => daysUntil(workstream.dueDate) <= 3);
  return {
    total: workstreams.length,
    active: active.length,
    complete: workstreams.filter((workstream) => workstream.completed).length,
    waived: workstreams.filter((workstream) => workstream.waived).length,
    blocked: blocked.length,
    missingDocs: missingDocs.length,
    dueSoon: dueSoon.length,
    needsAttention: new Set([...blocked, ...missingDocs, ...dueSoon]).size,
    allResolved: active.length === 0,
  };
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
  return projectHasAuditor(project, user.fullName);
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
        "Assigned work appears here when your name is listed as lead or support on an active project.",
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
      "Create a project or import a JSON backup to start working from shared tracker storage.",
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
  const auditStructure =
    project.auditStructure ??
    ((project.managingAgentWorkstreams?.length ?? 0) > 1
      ? "Coordinated"
      : "Solo");
  const normalizedProject = {
    ...project,
    assignmentType: project.assignmentType ?? "CH",
    auditEntity: project.auditEntity ?? "",
    assignedAuditor: primaryAuditor({ ...project, auditTeam }),
    auditTeam,
    auditStructure,
    dcaAgreementReceived: project.dcaAgreementReceived ?? false,
    claimsBdxReceived: project.claimsBdxReceived ?? false,
    paymentReceived:
      project.paymentReceived ?? project.invoiceStatus === "Paid",
    labels: project.labels ?? [],
    documentRequestDate: project.documentRequestDate ?? "",
    brokerLastChasedDate: project.brokerLastChasedDate ?? "",
    brokerExpectedResponseDate: project.brokerExpectedResponseDate ?? "",
    schedulingNotes: project.schedulingNotes ?? "",
    calendarSyncStatus: project.calendarSyncStatus ?? "Not Synced",
    calendarEventId: project.calendarEventId ?? "",
    calendarEventWebLink: project.calendarEventWebLink ?? "",
    calendarEventLastSyncedAt: project.calendarEventLastSyncedAt ?? "",
    auditLocation: project.auditLocation ?? "",
    auditRemoteLink: project.auditRemoteLink ?? "",
    auditDurationHours: Number(project.auditDurationHours || 1),
    auditStartTime: project.auditStartTime ?? "09:00",
    linkedContactId: project.linkedContactId ?? "",
    linkedContactSource: project.linkedContactSource ?? "",
    checklistCompletions: project.checklistCompletions ?? {},
    comments: project.comments ?? [],
    activityEvents: project.activityEvents ?? [],
    archived: project.archived ?? false,
  };
  return {
    ...normalizedProject,
    managingAgentWorkstreams:
      normalizeManagingAgentWorkstreams(normalizedProject),
  };
}

function loadProjects(): AuditProject[] {
  const raw = localStorage.getItem(storageKey);
  if (!raw) {
    localStorage.setItem(storageKey, JSON.stringify([]));
    return [];
  }
  try {
    const projects = (JSON.parse(raw) as AuditProject[]).map(withProjectDefaults);
    const realProjects = projects.filter(
      (project) => !legacySampleProjectIds.has(project.id),
    );
    if (realProjects.length !== projects.length) {
      localStorage.setItem(storageKey, JSON.stringify(realProjects));
    }
    return realProjects;
  } catch {
    return [];
  }
}

function saveProjects(projects: AuditProject[]) {
  localStorage.setItem(storageKey, JSON.stringify(projects));
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
  return requiredDocumentsForProject(project)
    .filter((doc) => !project[doc.key])
    .map((doc) => doc.label);
}

export function computedBlockers(project: AuditProject) {
  const blockers: string[] = [...getMissingDocuments(project)];
  const workstreamSummary = coordinatedWorkstreamSummary(project);
  if (
    stages.indexOf(project.currentStage) >= stages.indexOf("Quote") &&
    project.quoteStatus !== "Accepted"
  ) {
    blockers.push("Quote not accepted");
  }
  if (
    stages.indexOf(project.currentStage) >= stages.indexOf("File Selection") &&
    isDcaProject(project) &&
    !project.claimsBdxReceived
  ) {
    blockers.push("Claims BDX required before file selection");
  }
  if (
    stages.indexOf(project.currentStage) >= stages.indexOf("File Selection") &&
    !isDcaProject(project) &&
    !project.premiumBdxReceived
  ) {
    blockers.push("Premium BDX required before file selection");
  }
  if (project.auditStructure === "Coordinated" && workstreamSummary.needsAttention > 0) {
    blockers.push(
      `${workstreamSummary.needsAttention} managing agent workstream${
        workstreamSummary.needsAttention === 1 ? "" : "s"
      } need attention`,
    );
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

function intakeRequiredIssues(project: AuditProject) {
  const issues: string[] = [];
  if (!project.assignmentNumber.trim()) issues.push("Assignment number is required.");
  if (!project.auditEntity.trim()) issues.push("Audit entity is required.");
  if (!project.clientCoverholderCode.trim()) {
    issues.push("Client / coverholder code is required.");
  }
  if (!project.broker.trim() && !isDcaProject(project)) issues.push("Broker is required.");
  if (
    isDcaProject(project) &&
    !project.broker.trim() &&
    !normalizeManagingAgentWorkstreams(project).some((workstream) =>
      workstream.managingAgentName.trim(),
    )
  ) {
    issues.push("DCA audits need a managing agent or DCA contact.");
  }
  if (!primaryAuditor(project)) issues.push("Lead auditor is required.");
  if (!project.dueDate) issues.push("Due date is required.");
  return issues;
}

function intakeWarningIssues(project: AuditProject) {
  const warnings: string[] = [];
  const stageIndex = stages.indexOf(project.currentStage);
  if (
    stageIndex >= stages.indexOf("Quote") &&
    project.quoteStatus !== "Accepted"
  ) {
    warnings.push("This project is past intake but the quote is not accepted.");
  }
  if (
    ["Sent", "Accepted"].includes(project.quoteStatus) &&
    project.quoteAmount <= 0
  ) {
    warnings.push("Quote amount should be entered once the quote is sent.");
  }
  if (
    project.documentRequestStatus === "Complete" &&
    getMissingDocuments(project).length > 0
  ) {
    warnings.push(
      `Document request is complete but still missing: ${getMissingDocuments(project).join(", ")}.`,
    );
  }
  if (
    stageIndex >= stages.indexOf("Scheduling") &&
    project.quoteStatus !== "Accepted"
  ) {
    warnings.push("Scheduling should wait until the quote is accepted.");
  }
  if (
    stageIndex >= stages.indexOf("File Selection") &&
    !isDcaProject(project) &&
    !project.premiumBdxReceived
  ) {
    warnings.push("File selection needs Premium BDX before moving forward.");
  }
  if (!project.linkedContactId) {
    warnings.push("No linked workbook contact is selected; template recipients may need manual review.");
  }
  if (project.confirmedAuditDate && project.calendarSyncStatus !== "Synced") {
    warnings.push("Confirmed audit date is set but no current Outlook invite has been sent.");
  }
  return warnings;
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
    project.assignmentType === "DCA" &&
    !project.claimsBdxReceived
  ) {
    return "Claims BDX must be received before moving to File Selection.";
  }
  if (
    targetIndex >= stages.indexOf("File Selection") &&
    project.assignmentType !== "DCA" &&
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

function projectScheduleDate(project: AuditProject) {
  return project.confirmedAuditDate || project.dueDate || "";
}

function schedulingConflictWarnings(
  project: AuditProject,
  projects: AuditProject[],
) {
  const warnings: string[] = [];
  const scheduleDate = projectScheduleDate(project);
  if (!project.confirmedAuditDate && project.tentativeAuditWeek) {
    warnings.push("Tentative week set, but confirmed audit date is still blank.");
  }
  if (!project.tentativeAuditWeek && !project.confirmedAuditDate) {
    warnings.push("No scheduling target recorded yet.");
  }
  if (project.quoteStatus !== "Accepted") {
    warnings.push("Quote is not accepted yet.");
  }
  if (project.calendarSyncStatus === "Conflict Review") {
    warnings.push("Marked for scheduling conflict review.");
  }
  if (scheduleDate) {
    const projectAuditors = assignedAuditorNames(project);
    const sameDayConflicts = projects.filter(
      (candidate) =>
        candidate.id !== project.id &&
        !candidate.archived &&
        candidate.currentStage !== "Closed" &&
        projectScheduleDate(candidate) === scheduleDate &&
        assignedAuditorNames(candidate).some((auditor) =>
          projectAuditors.includes(auditor),
        ),
    );
    if (sameDayConflicts.length > 0) {
      warnings.push(
        `Same-day auditor conflict with ${sameDayConflicts
          .slice(0, 2)
          .map((candidate) => candidate.assignmentNumber)
          .join(", ")}.`,
      );
    }
  }
  return warnings;
}

function scheduleStatus(project: AuditProject) {
  if (!project.confirmedAuditDate && !project.tentativeAuditWeek) {
    return { label: "No date", className: "muted" };
  }
  if (project.calendarEventId && project.calendarSyncStatus === "Synced") {
    return { label: "Invite sent", className: "ok" };
  }
  if (project.calendarEventId && project.calendarSyncStatus !== "Synced") {
    return { label: "Needs update", className: "warning" };
  }
  return { label: "Planned", className: "warning" };
}

function nextCalendarStatusForScheduleChange(project: AuditProject) {
  return project.calendarEventId ? "Ready to Sync" : "Not Synced";
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
    Registration: "Confirm audit ownership and prepare the quote record.",
    Quote: "Move to scheduling once the quote is accepted.",
    Scheduling:
      "Confirm the audit date, audit week, and remote or onsite format.",
    "Pre-Audit": "Complete document readiness before file selection.",
    "File Selection": "Finish sample selection and notify the audit team.",
    "Audit Fieldwork": "Complete testing, log exceptions, and prepare findings.",
    Findings:
      "Send findings follow-up and record the coverholder response date.",
    "Report Drafting": "Complete the final quality check before issue.",
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
  const projectRequiredDocuments = requiredDocumentsForProject(project);
  const requiredComplete = projectRequiredDocuments.filter((doc) => project[doc.key]).length;
  const workflowComplete = [
    project.preAuditQuestionnaireStatus === "Complete",
    project.documentRequestStatus === "Complete",
  ].filter(Boolean).length;
  const completeCount = requiredComplete + workflowComplete;
  const totalCount = projectRequiredDocuments.length + 2;
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
    baaReceived: isDcaProject(project) ? project.baaReceived : true,
    endorsementsReceived: isDcaProject(project) ? project.endorsementsReceived : true,
    premiumBdxReceived: isDcaProject(project) ? project.premiumBdxReceived : true,
    dcaAgreementReceived: isDcaProject(project) ? true : project.dcaAgreementReceived,
    claimsBdxReceived: isDcaProject(project) ? true : project.claimsBdxReceived,
    managingAgentWorkstreams: normalizeManagingAgentWorkstreams(project).map(
      (workstream) => ({
        ...workstream,
        baaReceived: isDcaProject(project) ? workstream.baaReceived : true,
        endorsementsReceived: isDcaProject(project)
          ? workstream.endorsementsReceived
          : true,
        premiumBdxReceived: isDcaProject(project)
          ? workstream.premiumBdxReceived
          : true,
        dcaAgreementReceived: isDcaProject(project)
          ? true
          : workstream.dcaAgreementReceived,
        claimsBdxReceived: isDcaProject(project)
          ? true
          : workstream.claimsBdxReceived,
        documentRequestStatus: "Complete",
      }),
    ),
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

function requiredDocumentRequestLabels(project: AuditProject) {
  return requiredDocumentsForProject(project).map((document) =>
    document.label.replace(/\s+received$/i, ""),
  );
}

function templateReceiverKind(
  project: AuditProject,
  template: CommunicationTemplate,
): TemplateReceiverKind {
  if (template.id === "invoice-note") return "Invoice contact";
  if (template.id === "findings-follow-up") return "Report contact";
  if (isDcaProject(project)) return "DCA contact";
  if (["document-request", "pre-audit-questionnaire", "quote-email"].includes(template.id)) {
    return "Coverholder contact";
  }
  return "Project contact";
}

function normalizedToken(value: string) {
  return value.toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
}

function projectContactSearchTerms(project: AuditProject) {
  return [
    project.auditEntity,
    project.clientCoverholderCode,
    project.broker,
    ...normalizeManagingAgentWorkstreams(project).flatMap((workstream) => [
      workstream.managingAgentName,
      workstream.managingAgentCode,
    ]),
  ]
    .map(normalizedToken)
    .filter((term) => term.length >= 3)
    .filter((term, index, list) => list.indexOf(term) === index);
}

function contactDirectorySearchText(contact: LinkedContact) {
  return normalizedToken(
    [
      contact.company,
      contact.coverholder,
      contact.managingAgent,
      contact.broker,
      contact.contactName,
      contact.email,
      contact.worksheetName,
      contact.workbookName,
      ...Object.values(contact.raw ?? {}),
    ].join(" "),
  );
}

function scoreContactMatch(contact: LinkedContact, project: AuditProject) {
  const searchText = contactSearchText(contact);
  return projectContactSearchTerms(project).reduce((score, term) => {
    if (searchText.includes(term)) return score + 2;
    const compactTerm = term.replace(/\s+/g, "");
    if (compactTerm.length >= 4 && searchText.replace(/\s+/g, "").includes(compactTerm)) {
      return score + 1;
    }
    return score;
  }, 0);
}

function extractEmailsFromText(value: string) {
  return Array.from(
    new Set(
      value
        .match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi)
        ?.map((email) => email.toLowerCase()) ?? [],
    ),
  );
}

function extractPhoneFromText(value: string) {
  return value.match(/(?:\+?\d[\d().\-\s]{7,}\d)/)?.[0]?.trim() ?? "";
}

function extractNameFromContactBlock(value: string) {
  return (
    value
      .split(/\r?\n|;/)
      .map((part) => part.trim())
      .find(
        (part) =>
          part &&
          !part.includes("@") &&
          !/^(title|email|tel|phone|contact|n\/a|na)$/i.test(part) &&
          !extractPhoneFromText(part),
      ) ?? ""
  );
}

function contactBlockForReceiver(contact: LinkedContact, kind: TemplateReceiverKind) {
  if (kind === "DCA contact") return contact.raw?.dcaContact ?? "";
  if (kind === "Coverholder contact") return contact.raw?.coverholderContact ?? "";
  if (kind === "Invoice contact") return contact.raw?.invoiceSubmission ?? "";
  if (kind === "Report contact") return contact.raw?.reportSubmission ?? "";
  return "";
}

function resolveTemplateReceiver(
  project: AuditProject,
  template: CommunicationTemplate,
  contacts: LinkedContact[] = [],
): TemplateReceiver {
  const kind = templateReceiverKind(project, template);
  const rankedContacts = contacts
    .map((contact) => ({ contact, score: scoreContactMatch(contact, project) }))
    .filter((item) => item.score > 0)
    .sort((a, b) => b.score - a.score);
  const matchedContact = rankedContacts[0]?.contact;
  const block = matchedContact ? contactBlockForReceiver(matchedContact, kind) : "";
  const blockEmails = extractEmailsFromText(block);
  const structuredEmails =
    matchedContact && kind === "DCA contact"
      ? matchedContact.emails?.dca ?? []
      : matchedContact && kind === "Coverholder contact"
        ? matchedContact.emails?.coverholder ?? []
        : matchedContact && kind === "Invoice contact"
          ? matchedContact.emails?.invoice ?? []
          : matchedContact && kind === "Report contact"
            ? matchedContact.emails?.report ?? []
            : [];
  const contactEmails = matchedContact
    ? [
        ...structuredEmails,
        ...extractEmailsFromText(matchedContact.email || Object.values(matchedContact.raw ?? {}).join(" ")),
      ]
    : [];
  const email = blockEmails[0] || contactEmails[0] || "";
  const name =
    extractNameFromContactBlock(block) ||
    matchedContact?.contactName ||
    matchedContact?.company ||
    project.auditEntity ||
    kind;
  const company =
    matchedContact?.company ||
    matchedContact?.coverholder ||
    matchedContact?.managingAgent ||
    project.auditEntity ||
    "Not set";
  return {
    kind,
    name,
    email,
    company,
    source: matchedContact
      ? `${matchedContact.workbookName} / ${matchedContact.worksheetName}`
      : "No linked workbook match",
    confidence: matchedContact && email ? "Matched workbook" : "Needs review",
    guidance:
      kind === "DCA contact"
        ? "DCA audits should route document and quote requests to the DCA contact."
        : kind === "Coverholder contact"
          ? "Coverholder audits should route requests to the coverholder contact."
          : kind === "Invoice contact"
            ? "Use the invoice submission contact or finance handoff address from the client instruction sheet."
            : kind === "Report contact"
              ? "Use the report submission contact when sending findings or final report follow-up."
      : "Review the linked workbook before sending.",
  };
}

function linkedContactName(contact: LinkedContact) {
  return (
    contact.company ||
    contact.coverholder ||
    contact.managingAgent ||
    contact.worksheetName ||
    contact.contactName ||
    "Linked contact"
  );
}

function linkedContactLabel(contact: LinkedContact) {
  const name = linkedContactName(contact);
  const source = contact.workbookName
    ? contact.workbookName.replace(/^Client Instructions\s*/i, "Instructions ")
    : contact.sourceLabel;
  return `${name} - ${source}`;
}

function contactEmailBuckets(contact: LinkedContact) {
  return [
    ["DCA", contact.emails?.dca ?? []],
    ["Coverholder", contact.emails?.coverholder ?? []],
    ["Report", contact.emails?.report ?? []],
    ["Invoice", contact.emails?.invoice ?? []],
  ].filter(([, values]) => Array.isArray(values) && values.length > 0) as [
    string,
    string[],
  ][];
}

type ContactFilter = "All" | "DCA" | "Coverholder" | "Report" | "Invoice" | "Missing Email";

function contactMatchesFilter(contact: LinkedContact, filter: ContactFilter) {
  if (filter === "All") return true;
  if (filter === "Missing Email") return contactEmailBuckets(contact).length === 0 && !contact.email;
  return contactEmailBuckets(contact).some(([label]) => label === filter);
}

function contactSearchText(contact: LinkedContact) {
  return [
    contact.contactName,
    contact.company,
    contact.email,
    contact.broker,
    contact.coverholder,
    contact.managingAgent,
    contact.role,
    contact.workbookName,
    contact.worksheetName,
    ...contactEmailBuckets(contact).flatMap(([, emails]) => emails),
    ...contact.specialInstructions.flatMap((instruction) => [
      instruction.label,
      instruction.value,
    ]),
  ]
    .join(" ")
    .toLowerCase();
}

function instructionValue(contact: LinkedContact, label: string) {
  const normalizedLabel = label.toLowerCase();
  return (
    contact.specialInstructions.find((instruction) =>
      instruction.label.toLowerCase().includes(normalizedLabel),
    )?.value ?? ""
  );
}

function auditTypeFromContact(contact: LinkedContact): AuditType | "" {
  const preference = instructionValue(contact, "onsite");
  if (/onsite/i.test(preference)) return "Onsite";
  if (/remote/i.test(preference)) return "Remote";
  return "";
}

function contactSchedulingNotes(contact: LinkedContact) {
  return [
    instructionValue(contact, "Onsite/Remote Preference"),
    instructionValue(contact, "Fees and Payment Terms"),
    instructionValue(contact, "Other"),
    instructionValue(contact, "Notes/Comments"),
  ]
    .filter(Boolean)
    .map((value) => value.trim())
    .filter((value, index, list) => list.indexOf(value) === index)
    .join("\n");
}

function projectWithLinkedContact(project: AuditProject, contact: LinkedContact) {
  const auditTypePreference = auditTypeFromContact(contact);
  const contactNotes = contactSchedulingNotes(contact);
  const existingNotes = project.schedulingNotes.trim();
  const mergedNotes =
    contactNotes && !existingNotes.includes(contactNotes)
      ? [existingNotes, contactNotes].filter(Boolean).join("\n\n")
      : project.schedulingNotes;
  return withProjectDefaults({
    ...project,
    linkedContactId: contact.id,
    linkedContactSource: `${contact.workbookName} / ${contact.worksheetName}`,
    auditEntity: project.auditEntity.trim() || linkedContactName(contact),
    clientCoverholderCode:
      project.clientCoverholderCode.trim() || contact.worksheetName || "",
    broker:
      project.broker.trim() ||
      contact.managingAgent ||
      contact.broker ||
      contact.company ||
      "",
    auditType: auditTypePreference || project.auditType,
    schedulingNotes: mergedNotes,
  });
}

function duplicateProjectWarnings(project: AuditProject, projects: AuditProject[]) {
  const warnings: string[] = [];
  const normalizedAssignment = project.assignmentNumber.trim().toLowerCase();
  const normalizedCode = project.clientCoverholderCode.trim().toLowerCase();
  const normalizedEntity = project.auditEntity.trim().toLowerCase();
  const activeProjects = projects.filter(
    (item) => item.id !== project.id && !item.archived && item.currentStage !== "Closed",
  );
  if (
    normalizedAssignment &&
    activeProjects.some((item) => item.assignmentNumber.trim().toLowerCase() === normalizedAssignment)
  ) {
    warnings.push("Another active project already uses this assignment number.");
  }
  if (
    normalizedCode &&
    activeProjects.some((item) => item.clientCoverholderCode.trim().toLowerCase() === normalizedCode)
  ) {
    warnings.push("Another active project already uses this client / coverholder code.");
  }
  if (
    normalizedEntity &&
    activeProjects.some((item) => item.auditEntity.trim().toLowerCase() === normalizedEntity)
  ) {
    warnings.push("A similar active project already exists for this audit entity.");
  }
  return warnings;
}

const communicationTemplates: CommunicationTemplate[] = [
  {
    id: "document-request",
    label: "Document request",
    kind: "Email",
    purpose: "Send the first document package request to the correct DCA or coverholder contact.",
    subject: (project) =>
      `Document request - ${project.assignmentNumber} - ${project.auditEntity}`,
    body: (project) =>
      `Hello,\n\nPlease provide the required audit support for ${project.auditEntity} (${project.clientCoverholderCode || "code TBD"}).\n\nRequired items:\n${requiredDocumentRequestLabels(project).map((label) => `- ${label}`).join("\n")}\n- Completed pre-audit questionnaire, if applicable\n\nRequested by: ${project.documentRequestDate || todayIso()}\nExpected response: ${project.brokerExpectedResponseDate || "TBD"}\n\nAudit team: ${formatAuditTeam(project)}\n\nThank you,`,
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
      `Hello,\n\nPlease find the audit quote details for ${project.auditEntity}.\n\nAssignment: ${project.assignmentNumber}\nAssignment type: ${project.assignmentType}\nAudit type: ${project.auditType}\nQuote status: ${project.quoteStatus}\nQuote amount: ${formatCurrency(project.quoteAmount)}\nTentative audit week: ${project.tentativeAuditWeek || "TBD"}\n\nPlease confirm acceptance or advise if any changes are required.\n\nThank you,`,
  },
  {
    id: "findings-follow-up",
    label: "Findings follow-up",
    kind: "Email",
    purpose: "Send findings and start the coverholder response cycle.",
    subject: (project) =>
      `Findings response requested - ${project.assignmentNumber} - ${project.auditEntity}`,
    body: (project) =>
      `Hello,\n\nFindings have been issued for ${project.auditEntity}.\n\nAssignment: ${project.assignmentNumber}\nFindings sent: ${project.findingsSentDate || todayIso()}\nCoverholder response received: ${project.coverholderResponseReceivedDate || "Not yet received"}\n\nPlease provide responses and supporting evidence for each finding so the audit team can continue report finalization.\n\nThank you,`,
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

function nextStage(project: AuditProject) {
  const index = stages.indexOf(project.currentStage);
  return index >= 0 ? stages[index + 1] : undefined;
}

function daysSince(dateValue: string) {
  if (!dateValue) return Number.POSITIVE_INFINITY;
  const parsed = new Date(`${dateValue}T12:00:00Z`);
  if (Number.isNaN(parsed.getTime())) return Number.POSITIVE_INFINITY;
  return Math.floor((today.getTime() - parsed.getTime()) / 86400000);
}

function workflowGates(project: AuditProject): WorkflowGate[] {
  const target = nextStage(project);
  const stageMoveBlocker = target ? canMoveToStage(project, target) : "";
  const readiness = documentReadiness(project);
  return [
    {
      label: "Next stage",
      status: stageMoveBlocker ? "Blocked" : target ? "Ready" : "Watch",
      detail: stageMoveBlocker || (target ? `Ready to move toward ${target}.` : "Project is at the final stage."),
    },
    {
      label: "Documents",
      status: readiness.percent === 100 ? "Ready" : readiness.percent >= 60 ? "Watch" : "Blocked",
      detail:
        readiness.percent === 100
          ? "Required documents are complete."
          : `${readiness.percent}% ready; missing ${readiness.missingDocuments.join(", ") || "workflow completion"}.`,
    },
    {
      label: "Quote",
      status: project.quoteStatus === "Accepted" ? "Ready" : "Blocked",
      detail:
        project.quoteStatus === "Accepted"
          ? "Quote is accepted."
          : `Quote is ${project.quoteStatus.toLowerCase()}; scheduling should wait.`,
    },
    {
      label: "Finish audit",
      status:
        project.reportStatus === "Issued" &&
        project.invoiceStatus === "Paid" &&
        project.paymentReceived
          ? "Ready"
          : stages.indexOf(project.currentStage) >= stages.indexOf("Final Submission")
            ? "Watch"
            : "Ready",
      detail:
        stages.indexOf(project.currentStage) >= stages.indexOf("Final Submission")
          ? `Report ${project.reportStatus.toLowerCase()}, invoice ${project.invoiceStatus.toLowerCase()}, payment ${project.paymentReceived ? "received" : "open"}.`
          : "Close-out is not active yet.",
    },
  ];
}

function slaSignals(project: AuditProject): SlaSignal[] {
  const signals: SlaSignal[] = [];
  const dueIn = daysUntil(project.dueDate);
  const staleDays = daysSince(project.lastUpdatedDate);
  if (dueIn < 0) {
    signals.push({
      level: "Critical",
      label: "Overdue",
      detail: `${Math.abs(dueIn)} day${Math.abs(dueIn) === 1 ? "" : "s"} overdue.`,
    });
  } else if (dueIn <= 3) {
    signals.push({
      level: "Warning",
      label: "Due soon",
      detail: `Due in ${dueIn} day${dueIn === 1 ? "" : "s"}.`,
    });
  }
  if (project.assignmentStatus === "Blocked" || computedBlockers(project).length > 0) {
    signals.push({
      level: "Critical",
      label: "Blocked",
      detail: computedBlockers(project).slice(0, 2).join("; ") || "Assignment is marked blocked.",
    });
  }
  if (project.labels.includes("Waiting on Broker")) {
    signals.push({
      level: "Warning",
      label: "Broker wait",
      detail: project.brokerLastChasedDate
        ? `Last chased ${project.brokerLastChasedDate}.`
        : "Waiting on broker with no chase date.",
    });
  }
  if (staleDays >= 7 && project.currentStage !== "Closed") {
    signals.push({
      level: "Warning",
      label: "Stale update",
      detail: `No update for ${staleDays} days.`,
    });
  }
  return signals.length
    ? signals
    : [{ level: "Normal", label: "On track", detail: "No timing alerts detected." }];
}

function workspaceFolders(project: AuditProject) {
  return [
    "01 Quote",
    "02 Planning",
    "03 Broker Documents",
    "04 File Selection",
    "05 Testing",
    "06 Findings",
    "07 Report",
    "08 Invoice",
    "09 Closeout",
  ].map((folder) => `${project.assignmentNumber || project.id}/${folder}`);
}

function recommendedDrafts(project: AuditProject): OperationsDraft[] {
  const drafts: OperationsDraft[] = [];
  const addDraft = (
    templateId: string,
    priority: OperationsDraft["priority"],
    reason: string,
  ) => {
    const template = communicationTemplates.find((item) => item.id === templateId);
    if (!template) return;
    drafts.push({
      id: `${project.id}-${template.id}`,
      projectId: project.id,
      assignmentNumber: project.assignmentNumber,
      auditEntity: project.auditEntity,
      templateId,
      label: template.label,
      kind: template.kind,
      priority,
      reason,
      subject: template.subject(project),
      body: template.body(project),
    });
  };
  if (getMissingDocuments(project).length > 0 || project.documentRequestStatus !== "Complete") {
    addDraft("document-request", "High", "Documents or request workflow are incomplete.");
  }
  if (project.preAuditQuestionnaireStatus !== "Complete") {
    addDraft("pre-audit-questionnaire", "Medium", "Pre-audit questionnaire is not complete.");
  }
  if (project.quoteStatus !== "Accepted") {
    addDraft("quote-email", "High", "Quote needs a client decision before scheduling.");
  }
  if (
    stages.indexOf(project.currentStage) >= stages.indexOf("Findings") &&
    !project.coverholderResponseReceivedDate
  ) {
    addDraft("findings-follow-up", "High", "Findings response is needed for wrap-up.");
  }
  if (
    stages.indexOf(project.currentStage) >= stages.indexOf("Final Submission") ||
    project.reportStatus === "Issued" ||
    project.invoiceStatus !== "Not Started"
  ) {
    addDraft("invoice-note", "Medium", "Finance handoff or invoice status needs tracking.");
  }
  return drafts;
}

function operationsDraftQueue(projects: AuditProject[]) {
  const priorityRank = { High: 0, Medium: 1, Low: 2 };
  return projects
    .filter((project) => project.currentStage !== "Closed")
    .flatMap(recommendedDrafts)
    .sort((a, b) => priorityRank[a.priority] - priorityRank[b.priority]);
}

function operationsBrief(projects: AuditProject[], user: PrototypeUser): OperationsBrief {
  const openProjects = projects.filter((project) => project.currentStage !== "Closed");
  const criticalProjects = openProjects.filter((project) =>
    slaSignals(project).some((signal) => signal.level === "Critical"),
  );
  const draftQueue = operationsDraftQueue(openProjects);
  const topRisks = criticalProjects.slice(0, 5).map(
    (project) =>
      `${project.assignmentNumber}: ${slaSignals(project)
        .filter((signal) => signal.level === "Critical")
        .map((signal) => signal.label)
        .join(", ")}`,
  );
  const actions = openProjects
    .flatMap((project) =>
      recommendedNextSteps(project).slice(0, 2).map((step) => `${project.assignmentNumber}: ${step}`),
    )
    .slice(0, 7);
  const managerFocus = [
    `${criticalProjects.length} critical assignment${criticalProjects.length === 1 ? "" : "s"}`,
    `${draftQueue.length} draft email/document action${draftQueue.length === 1 ? "" : "s"}`,
    `${openProjects.filter((project) => project.quoteStatus !== "Accepted").length} quote decision${openProjects.filter((project) => project.quoteStatus !== "Accepted").length === 1 ? "" : "s"} open`,
    `${openProjects.filter((project) => project.invoiceStatus !== "Paid" && stages.indexOf(project.currentStage) >= stages.indexOf("Invoice")).length} invoice follow-up${openProjects.filter((project) => project.invoiceStatus !== "Paid" && stages.indexOf(project.currentStage) >= stages.indexOf("Invoice")).length === 1 ? "" : "s"}`,
  ];
  return {
    summary: `${user.role} console: ${openProjects.length} open, ${criticalProjects.length} critical, ${draftQueue.length} communication drafts ready for review.`,
    risks: topRisks.length ? topRisks : ["No critical SLA risks in visible assignments."],
    actions: actions.length ? actions : ["Create or import real assignments to populate the operating console."],
    managerFocus,
    aiPrompt: [
      "Summarize these audit tracker priorities and recommend the next three actions.",
      `User role: ${user.role}`,
      `Open assignments: ${openProjects.length}`,
      `Risks: ${topRisks.join(" | ") || "none"}`,
      `Draft queue: ${draftQueue.slice(0, 5).map((draft) => `${draft.assignmentNumber} ${draft.label}`).join(" | ") || "none"}`,
    ].join("\n"),
  };
}

function priorityRank(priority: CoordinatorInsight["priority"]) {
  return { Critical: 0, High: 1, Medium: 2, Low: 3 }[priority];
}

function auditCoordinatorInsights(projects: AuditProject[], user: PrototypeUser) {
  const insights: CoordinatorInsight[] = [];
  const openProjects = projects.filter((project) => project.currentStage !== "Closed");
  const addInsight = (
    project: AuditProject,
    priority: CoordinatorInsight["priority"],
    title: string,
    reason: string,
    recommendedAction: string,
  ) => {
    insights.push({
      id: `${project.id}-${title}`,
      projectId: project.id,
      assignmentNumber: project.assignmentNumber || project.id,
      auditEntity: project.auditEntity || project.clientCoverholderCode || "No entity",
      priority,
      title,
      reason,
      recommendedAction,
    });
  };

  openProjects.forEach((project) => {
    const missingDocuments = getMissingDocuments(project);
    const dueIn = daysUntil(project.dueDate);
    if (dueIn < 0) {
      addInsight(
        project,
        "Critical",
        "Overdue assignment",
        `${Math.abs(dueIn)} day${Math.abs(dueIn) === 1 ? "" : "s"} overdue.`,
        "Escalate owner, reset due date, and record a current action note.",
      );
    }
    if (project.assignmentStatus === "Blocked" || computedBlockers(project).length > 0) {
      addInsight(
        project,
        "Critical",
        "Blocked workflow",
        computedBlockers(project).slice(0, 2).join("; ") || "Assignment is marked blocked.",
        "Clear or assign each blocker before moving the stage forward.",
      );
    }
    if (missingDocuments.length > 0) {
      addInsight(
        project,
        missingDocuments.length >= 2 ? "High" : "Medium",
        "Missing document evidence",
        `${missingDocuments.join(", ")} missing from the required package.`,
        "Use the document request template and update received evidence when files arrive.",
      );
    }
    if (project.quoteStatus !== "Accepted" && stages.indexOf(project.currentStage) >= stages.indexOf("Scheduling")) {
      addInsight(
        project,
        "High",
        "Scheduling ahead of quote",
        `Project is in ${project.currentStage}, but quote status is ${project.quoteStatus}.`,
        "Confirm quote acceptance before committing more audit effort.",
      );
    }
    if (project.confirmedAuditDate && project.calendarSyncStatus !== "Synced") {
      addInsight(
        project,
        "Medium",
        "Invite not sent",
        "Confirmed audit date exists, but the Outlook invite is not current.",
        "Sync or update the Outlook event once calendar permission is active.",
      );
    }
    if (!project.linkedContactId) {
      addInsight(
        project,
        "Low",
        "No linked client contact",
        "Templates may not know the best DCA, coverholder, report, or invoice recipient.",
        "Link a workbook contact before sending external communication.",
      );
    }
    if (
      user.role !== "Finance" &&
      stages.indexOf(project.currentStage) >= stages.indexOf("Invoice") &&
      project.invoiceStatus !== "Paid"
    ) {
      addInsight(
        project,
        "Medium",
        "Invoice follow-up open",
        `Invoice status is ${project.invoiceStatus}; payment received is ${project.paymentReceived ? "yes" : "no"}.`,
        "Confirm finance handoff, invoice recipient, and payment status.",
      );
    }
  });

  return insights.sort((a, b) => priorityRank(a.priority) - priorityRank(b.priority));
}

function documentIntelligence(project: AuditProject): DocumentIntelligenceResult {
  const readiness = documentReadiness(project);
  const required = requiredDocumentsForProject(project);
  const missing = getMissingDocuments(project);
  const evidence = required
    .filter((document) => project[document.key])
    .map((document) => document.label);
  const recommendations: string[] = [];
  if (missing.length > 0) {
    recommendations.push(`Request or chase ${missing.join(", ")}.`);
  }
  if (project.documentRequestStatus !== "Complete") {
    recommendations.push("Complete the document request workflow before marking the package ready.");
  }
  if (project.preAuditQuestionnaireStatus !== "Complete" && project.preAuditQuestionnaireStatus !== "Not Required") {
    recommendations.push("Resolve the pre-audit questionnaire status.");
  }
  if (project.brokerLastChasedDate && daysSince(project.brokerLastChasedDate) >= 5 && missing.length > 0) {
    recommendations.push("Broker/contact chase is stale; send a follow-up.");
  }
  if (project.auditStructure === "Coordinated") {
    const summary = coordinatedWorkstreamSummary(project);
    if (summary.needsAttention > 0) {
      recommendations.push(`${summary.needsAttention} managing-agent workstream${summary.needsAttention === 1 ? "" : "s"} still need attention.`);
    }
  }
  return {
    projectId: project.id,
    assignmentNumber: project.assignmentNumber || project.id,
    auditEntity: project.auditEntity || project.clientCoverholderCode || "No entity",
    packageType: isDcaProject(project)
      ? "DCA evidence package"
      : project.auditStructure === "Coordinated"
        ? "Coordinated managing-agent package"
        : "Coverholder audit package",
    readiness: readiness.percent,
    confidence:
      readiness.percent === 100
        ? "High"
        : readiness.percent >= 60
          ? "Medium"
          : "Needs Review",
    missing,
    evidence,
    recommendations: recommendations.length
      ? recommendations
      : ["Document evidence is ready for the next workflow stage."],
  };
}

function documentIntelligenceSummary(projects: AuditProject[]) {
  return projects
    .filter((project) => project.currentStage !== "Closed")
    .map(documentIntelligence)
    .sort((a, b) => a.readiness - b.readiness);
}

function buildOperationsReport(projects: AuditProject[], user: PrototypeUser) {
  const rows = projects.map((project) => ({
    assignmentNumber: project.assignmentNumber,
    auditEntity: project.auditEntity,
    currentStage: project.currentStage,
    status: project.assignmentStatus,
    dueDate: project.dueDate,
    dueInDays: daysUntil(project.dueDate),
    slaSignals: slaSignals(project).map((signal) => signal.label),
    blockers: computedBlockers(project),
    nextSteps: recommendedNextSteps(project),
    drafts: recommendedDrafts(project).map((draft) => draft.label),
    workspaceFolders: workspaceFolders(project),
  }));
  return {
    exportedAt: new Date().toISOString(),
    exportedBy: user.fullName,
    role: user.role,
    totals: {
      projects: projects.length,
      open: projects.filter((project) => project.currentStage !== "Closed").length,
      critical: rows.filter((row) => row.slaSignals.includes("Overdue") || row.slaSignals.includes("Blocked")).length,
      drafts: rows.reduce((sum, row) => sum + row.drafts.length, 0),
    },
    assistantBrief: operationsBrief(projects, user),
    coordinatorInsights: auditCoordinatorInsights(projects, user),
    documentIntelligence: documentIntelligenceSummary(projects),
    rows,
  };
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
    ["Lead Auditor", (project) => primaryAuditor(project)],
    ["Audit Team", (project) => formatAuditTeam(project)],
    ["Current Stage", (project) => project.currentStage],
    ["Assignment Status", (project) => project.assignmentStatus],
    ["Quote Status", (project) => project.quoteStatus],
    ["Quote Amount", (project) => project.quoteAmount],
    ["Due Date", (project) => project.dueDate],
    ["Tentative Audit Week", (project) => project.tentativeAuditWeek],
    ["Confirmed Audit Date", (project) => project.confirmedAuditDate],
    ["Schedule Status", (project) => scheduleStatus(project).label],
    ["Calendar Event Link", (project) => project.calendarEventWebLink],
    ["Audit Duration Hours", (project) => project.auditDurationHours],
    ["Audit Start Time", (project) => project.auditStartTime],
    ["Audit Location", (project) => project.auditLocation],
    ["Remote Link", (project) => project.auditRemoteLink],
    ["Linked Contact Source", (project) => project.linkedContactSource],
    ["Scheduling Notes", (project) => project.schedulingNotes],
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
  const [systemHealth, setSystemHealth] = useState<SecureSystemHealth | null>(
    null,
  );
  const [systemHealthLoading, setSystemHealthLoading] = useState(false);
  const [contactSources, setContactSources] =
    useState<LinkedContactSourcesResponse | null>(null);
  const [contactSourcesLoading, setContactSourcesLoading] = useState(false);
  const [contactSourcesError, setContactSourcesError] = useState("");
  const contactSourcesAutoLoadedFor = useRef<Set<string>>(new Set());
  const [projectStorageLoading, setProjectStorageLoading] = useState(false);
  const [selectedId, setSelectedId] = useState(projects[0]?.id ?? "");
  const [editing, setEditing] = useState<AuditProject | null>(null);
  const [filters, setFilters] = useState<Filters>(defaultFilters);
  const [message, setMessage] = useState("");
  const [viewMode, setViewMode] = useState<ViewMode>("kanban");
  const [activeSection, setActiveSection] = useState<AppSection>("dashboard");
  const [activeAdminTab, setActiveAdminTab] = useState<AdminTab>("users");
  const [durationRange, setDurationRange] = useState<DurationRange>("ytd");
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

  const refreshSystemHealth = async () => {
    if (!signedInUser) return;
    setSystemHealthLoading(true);
    try {
      setSystemHealth(await getSecureSystemHealth());
    } catch (error) {
      setMessage(error instanceof Error ? error.message : "System health check failed.");
    } finally {
      setSystemHealthLoading(false);
    }
  };

  const refreshContactSources = async () => {
    if (!signedInUser) return;
    setContactSourcesLoading(true);
    setContactSourcesError("");
    try {
      setContactSources(await getLinkedContactSources());
    } catch (error) {
      const errorMessage =
        error instanceof Error ? error.message : "Contact source refresh failed.";
      setContactSourcesError(errorMessage);
      setMessage(errorMessage);
    } finally {
      setContactSourcesLoading(false);
    }
  };

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
          redirectUri: "http://localhost:8787/api/auth/callback",
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
    if (signedInUser?.role === "Admin") {
      void refreshSystemHealth();
    } else {
      setSystemHealth(null);
      setContactSources(null);
      setContactSourcesError("");
    }
  }, [signedInUser?.role, signedInUser?.email]);

  useEffect(() => {
    if (activeSection === "admin" && signedInUser?.role !== "Admin") {
      setActiveSection("dashboard");
    }
  }, [activeSection, signedInUser?.role]);

  useEffect(() => {
    const autoLoadKey = `${signedInUser?.email || ""}:contacts`;
    if (
      activeSection === "admin" &&
      activeAdminTab === "contacts" &&
      !contactSources &&
      !contactSourcesLoading &&
      !contactSourcesAutoLoadedFor.current.has(autoLoadKey)
    ) {
      contactSourcesAutoLoadedFor.current.add(autoLoadKey);
      void refreshContactSources();
    }
  }, [
    activeSection,
    activeAdminTab,
    signedInUser?.role,
    signedInUser?.email,
    contactSources,
    contactSourcesLoading,
  ]);

  useEffect(() => {
    const autoLoadKey = `${signedInUser?.email || ""}:project-form-contacts`;
    if (
      signedInUser &&
      editing &&
      !contactSources &&
      !contactSourcesLoading &&
      !contactSourcesAutoLoadedFor.current.has(autoLoadKey)
    ) {
      contactSourcesAutoLoadedFor.current.add(autoLoadKey);
      void refreshContactSources();
    }
  }, [
    signedInUser?.email,
    editing,
    contactSources,
    contactSourcesLoading,
  ]);

  useEffect(() => {
    if (!signedInUser) return;
    let cancelled = false;
    setProjectStorageLoading(true);
    getServerProjects()
      .then((serverProjects) => {
        if (cancelled) return;
        const normalizedProjects = serverProjects.map(withProjectDefaults);
        setProjects(normalizedProjects);
        saveProjects(normalizedProjects);
        setSelectedId((currentId) =>
          normalizedProjects.some((project) => project.id === currentId && !project.archived)
            ? currentId
            : normalizedProjects.find((project) => !project.archived)?.id ?? "",
        );
      })
      .catch((error) => {
        if (cancelled) return;
        setMessage(
          error instanceof Error
            ? error.message
            : "Project storage is not available.",
        );
      })
      .finally(() => {
        if (!cancelled) setProjectStorageLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, [signedInUser?.email]);

  const pendingAccessRequests =
    secureAccess?.pendingRequests?.map(secureUserToPrototypeUser) ?? [];
  const managedAccessUsers =
    secureAccess?.managedUsers?.map(secureUserToPrototypeUser) ??
    (secureAccess?.user ? [secureUserToPrototypeUser(secureAccess.user)] : []);
  const approvedAccessUsers = managedAccessUsers.filter(
    (user) => user.accessRequestStatus === "Approved",
  );
  const auditorOptions = approvedAccessUsers
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
  const activeVisibleProjects = useMemo(
    () => visibleProjects.filter((project) => !project.archived),
    [visibleProjects],
  );
  const archivedVisibleProjects = useMemo(
    () => visibleProjects.filter((project) => project.archived),
    [visibleProjects],
  );
  const zeroLoadAuditors = auditors.filter(
    (auditor) =>
      !activeVisibleProjects.some(
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
      activeVisibleProjects.filter((project) => {
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
    [activeVisibleProjects, filters],
  );
  const selectedProject =
    activeVisibleProjects.find((project) => project.id === selectedId) ??
    activeVisibleProjects[0];
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

  const persist = (
    nextProjects: AuditProject[],
    options: { replaceAll?: boolean; successMessage?: string } = {},
  ) => {
    setProjects(nextProjects);
    saveProjects(nextProjects);
    void saveServerProjects(nextProjects, { replaceAll: options.replaceAll })
      .then((serverProjects) => {
        const normalizedProjects = serverProjects.map(withProjectDefaults);
        setProjects(normalizedProjects);
        saveProjects(normalizedProjects);
        if (options.successMessage) setMessage(options.successMessage);
      })
      .catch((error) => {
        setMessage(
          error instanceof Error
            ? error.message
            : "Project changes were saved locally but not to the server.",
        );
      });
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
      signedInUser.role === "Admin" ? approvedAccessUsers : [],
      signedInUser.fullName,
    );
    recordExport("Microsoft Lists package");
  };

  const handleExportOperationsReport = () => {
    downloadJsonFile(
      buildOperationsReport(activeVisibleProjects, signedInUser),
      `audit-operations-report-${new Date().toISOString().slice(0, 10)}.json`,
    );
    recordExport("Operations report");
  };

  const startProjectFromContact = (contact: LinkedContact) => {
    const starter = projectWithLinkedContact(blankProject(), contact);
    setEditing(starter);
    setActiveSection("assignments");
    setMessage(`Project intake started from ${linkedContactName(contact)}.`);
  };

  const clearProjectData = () => {
    requestConfirmation({
      title: "Clear project data?",
      message:
        "This removes all project records from the shared tracker so you can start testing with real audit data.",
      confirmLabel: "Clear projects",
      tone: "danger",
      onConfirm: () => {
        persist([], {
          replaceAll: true,
          successMessage:
            "Project data cleared. Add or import real audit records to continue testing.",
        });
        setSelectedId("");
        setMessage("Project data cleared. Add or import real audit records to continue testing.");
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
      message: `Importing ${file.name} replaces the current project records in shared tracker storage.`,
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

  const archiveProject = (project: AuditProject) => {
    if (!canEditProject(signedInUser, project)) {
      setMessage("Your role cannot archive this project.");
      return;
    }
    if (project.currentStage !== "Closed") {
      setMessage("Move the project to Closed before archiving it.");
      return;
    }
    const workstreamSummary = coordinatedWorkstreamSummary(project);
    if (project.auditStructure === "Coordinated" && !workstreamSummary.allResolved) {
      setMessage(
        `Resolve or waive ${workstreamSummary.active} active managing agent workstream${
          workstreamSummary.active === 1 ? "" : "s"
        } before archiving.`,
      );
      return;
    }
    const updatedProject = withProjectDefaults({
      ...project,
      archived: true,
      assignmentStatus: "Completed",
      lastUpdatedDate: new Date().toISOString().slice(0, 10),
      activityEvents: [
        ...(project.activityEvents ?? []),
        createActivityEvent(
          "stage",
          "Project archived",
          "Closed project was hidden from active operational views without deleting the record.",
          signedInUser.fullName,
        ),
      ],
    });
    persist(
      projects.map((item) => (item.id === project.id ? updatedProject : item)),
      { successMessage: `${project.assignmentNumber} archived.` },
    );
    setSelectedId(
      activeVisibleProjects.find((item) => item.id !== project.id)?.id ?? "",
    );
  };

  const restoreProject = (project: AuditProject) => {
    if (!hasFullProjectAccess(signedInUser)) {
      setMessage("Only admins and audit managers can restore archived projects.");
      return;
    }
    const updatedProject = withProjectDefaults({
      ...project,
      archived: false,
      lastUpdatedDate: new Date().toISOString().slice(0, 10),
      activityEvents: [
        ...(project.activityEvents ?? []),
        createActivityEvent(
          "stage",
          "Project restored",
          "Archived project was restored to active operational views.",
          signedInUser.fullName,
        ),
      ],
    });
    persist(
      projects.map((item) => (item.id === project.id ? updatedProject : item)),
      { successMessage: `${project.assignmentNumber} restored.` },
    );
    setSelectedId(project.id);
    setActiveSection("assignments");
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
      activityEvents: [
        ...(project.activityEvents ?? []),
        createActivityEvent(
          "field",
          "Label removed",
          `${label} was removed from the project.`,
          signedInUser.fullName,
        ),
      ],
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
      activityEvents: [
        ...(project.activityEvents ?? []),
        createActivityEvent(
          "field",
          "Comment added",
          comment.body.length > 120
            ? `${comment.body.slice(0, 117)}...`
            : comment.body,
          signedInUser.fullName,
        ),
      ],
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

  const updateProjectScheduling = (
    project: AuditProject,
    update: Partial<
      Pick<
        AuditProject,
        | "calendarSyncStatus"
        | "calendarEventId"
        | "calendarEventWebLink"
        | "calendarEventLastSyncedAt"
        | "auditLocation"
        | "auditRemoteLink"
        | "auditDurationHours"
        | "auditStartTime"
        | "schedulingNotes"
        | "confirmedAuditDate"
        | "tentativeAuditWeek"
      >
    >,
  ) => {
    if (!canEditProject(signedInUser, project)) {
      setMessage("Your role cannot update scheduling for this project.");
      return;
    }
    const updatedProject = withProjectDefaults({
      ...project,
      ...update,
      lastUpdatedDate: new Date().toISOString().slice(0, 10),
      activityEvents: [
        ...(project.activityEvents ?? []),
        createActivityEvent(
          "field",
          "Scheduling updated",
          "Scheduling notes, invite status, or audit dates were updated.",
          signedInUser.fullName,
        ),
      ],
    });
    persist(
      projects.map((item) => (item.id === project.id ? updatedProject : item)),
    );
    setSelectedId(project.id);
    setMessage(`${project.assignmentNumber} scheduling updated.`);
  };

  const syncProjectToOutlook = async (project: AuditProject) => {
    if (!canEditProject(signedInUser, project)) {
      setMessage("Your role cannot send an Outlook invite for this project.");
      return;
    }
    try {
      const teamEmails = assignedAuditorNames(project)
        .map((name) =>
          approvedAccessUsers.find((user) => user.fullName === name)?.email ?? "",
        )
        .filter(Boolean);
      const result = await createOutlookCalendarEvent(project.id, {
        durationHours: project.auditDurationHours,
        startTime: project.auditStartTime,
        confirmedAuditDate: project.confirmedAuditDate,
        location: project.auditLocation,
        remoteLink: project.auditRemoteLink,
        attendeeEmails: teamEmails,
      });
      updateProjectScheduling(project, {
        calendarSyncStatus: "Synced",
        calendarEventId: result.event.id,
        calendarEventWebLink: result.event.webLink,
        calendarEventLastSyncedAt: new Date().toISOString(),
      });
      setMessage(
        result.event.webLink
          ? `Outlook event ${result.action}: ${result.event.subject}`
          : `Outlook event ${result.action}.`,
      );
    } catch (error) {
      setMessage(
        error instanceof Error
          ? error.message
          : "Outlook invite could not be sent.",
      );
    }
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
      persist(importedProjects, {
        replaceAll: true,
        successMessage: `${importedProjects.length} projects imported from JSON.`,
      });
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

  const approveUserRequest = async (
    username: string,
    update: AccessApprovalUpdate = {},
  ) => {
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
      await approveSecureAccessRequest(user.email, update);
      await refreshSecureAccess();
      await refreshSystemHealth();
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
      await refreshSystemHealth();
      setMessage(`${user.fullName} rejected.`);
    } catch (error) {
      setMessage(error instanceof Error ? error.message : "Rejection failed.");
    }
  };

  const saveManagedUser = async (email: string, draftUser: PrototypeUser) => {
    if (signedInUser.role !== "Admin") {
      setMessage("Only admins can manage users.");
      return;
    }
    try {
      await updateSecureAccessUser(email, {
        fullName: draftUser.fullName,
        role: draftUser.role,
        defaultVisibility: draftUser.defaultVisibility,
        active: draftUser.active,
      });
      await refreshSecureAccess();
      await refreshSystemHealth();
      setMessage(`${draftUser.fullName} updated.`);
    } catch (error) {
      setMessage(error instanceof Error ? error.message : "User update failed.");
    }
  };

  return (
    <main>
      <header className="hero">
        <div className="hero-branding">
          <img src={mosaicLogoUrl} alt="Mosaic International Insurance Professionals" />
          <div>
            <p className="eyebrow">Signed in: {signedInUser.role}</p>
            <h1>Audit Assignment Tracker</h1>
            <p>
              Intake, document readiness, workload, client contacts, and close-out
              in one operating view.
            </p>
          </div>
        </div>
        <div className="hero-actions">
          {canCreateProject(signedInUser) && (
            <button onClick={() => setEditing(blankProject())}>Add project</button>
          )}
          {hasFullProjectAccess(signedInUser) && (
            <button className="secondary" onClick={clearProjectData}>
              Clear project data
            </button>
          )}
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
      {projectStorageLoading && (
        <div className="toast subtle" role="status">
          Loading shared project storage...
        </div>
      )}

      <AccessBanner
        user={signedInUser}
        visibleCount={activeVisibleProjects.length}
      />
      <AppNavigation
        activeSection={activeSection}
        setActiveSection={setActiveSection}
        currentUser={signedInUser}
        activeCount={activeVisibleProjects.length}
        archivedCount={archivedVisibleProjects.length}
        pendingRequestCount={pendingAccessRequests.length}
      />

      {activeSection === "dashboard" && (
        <>
          <Dashboard projects={activeVisibleProjects} />
          <TodaysWork projects={activeVisibleProjects} onSelect={(id) => {
            setSelectedId(id);
            setActiveSection("assignments");
          }} />
          <WorkloadCounts
            projects={activeVisibleProjects}
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
        </>
      )}

      {activeSection === "assignments" && (
        <>
          <FiltersPanel
            filters={filters}
            setFilters={setFilters}
            presets={filterPresetsForUser(signedInUser)}
            auditors={
              hasFullProjectAccess(signedInUser) || signedInUser.role === "Read Only"
                ? auditors
                : auditors.filter((auditor) =>
                    activeVisibleProjects.some((project) => projectHasAuditor(project, auditor)),
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
            {signedInUser.role === "Admin" && (
              <div className="admin-data-actions" aria-label="Admin data actions">
                <button onClick={handleExportCsv}>Export filtered CSV</button>
                <button className="secondary" onClick={handleExportJson}>
                  Export JSON backup
                </button>
                <span className="last-export">
                  Last export: {formatDateTime(lastExportedAt)}
                </span>
                <label className="import-control">
                  Import JSON
                  <input
                    type="file"
                    accept="application/json,.json"
                    onChange={(event) => {
                      const file = event.target.files?.[0];
                      if (file) queueProjectImport(file);
                      event.currentTarget.value = "";
                    }}
                  />
                </label>
              </div>
            )}
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
              onArchive={archiveProject}
              onRemoveLabel={removeProjectLabel}
              onAddComment={addProjectComment}
              onToggleChecklist={toggleChecklistItem}
              onDocumentWorkflowAction={updateProjectDocumentWorkflow}
              auditors={auditors}
              onAddSupportingAuditor={addSupportingAuditor}
              currentUser={signedInUser}
              onUpdateFinance={updateProjectFinance}
              contactSources={contactSources}
            />
          )}
          {!selectedProject && (
            <section className="panel empty-state">
              <h2>{emptyProjectState(signedInUser).title}</h2>
              <p>{emptyProjectState(signedInUser).message}</p>
            </section>
          )}
        </>
      )}

      {activeSection === "scheduling" && (
        <SchedulingCapacity
          projects={activeVisibleProjects}
          auditors={auditors}
          calendarReady={Boolean(systemHealth?.calendar?.permissionGranted)}
          onSelect={(id) => {
            setSelectedId(id);
            setActiveSection("assignments");
          }}
          onUpdateScheduling={updateProjectScheduling}
          onSyncOutlook={syncProjectToOutlook}
          currentUser={signedInUser}
        />
      )}

      {activeSection === "command" && (
        <OperationsCommandCenter
          projects={activeVisibleProjects}
          currentUser={signedInUser}
          onSelect={(id) => {
            setSelectedId(id);
            setActiveSection("assignments");
          }}
          onExportReport={handleExportOperationsReport}
        />
      )}

      {activeSection === "reports" && (
        <>
          <CycleTimeDashboard
            projects={activeVisibleProjects}
            range={durationRange}
            setRange={setDurationRange}
          />
          <WorkloadCounts
            projects={activeVisibleProjects}
            auditors={auditors}
            hiddenAuditors={effectiveHiddenWorkloadAuditors}
            toggleAuditorHidden={(auditor) =>
              setHiddenWorkloadAuditors((current) =>
                current.includes(auditor)
                  ? current.filter((item) => item !== auditor)
                  : [...current, auditor],
              )
            }
            showAllAuditors={() => {
              setHiddenWorkloadAuditors([]);
              setShownZeroLoadAuditors(zeroLoadAuditors);
            }}
          />
          <ArchivedProjectsPanel
            projects={archivedVisibleProjects}
            canRestore={hasFullProjectAccess(signedInUser)}
            onRestore={restoreProject}
          />
        </>
      )}

      {activeSection === "admin" && signedInUser.role === "Admin" && (
        <AdminWorkspace
          activeTab={activeAdminTab}
          setActiveTab={setActiveAdminTab}
          health={systemHealth}
          loading={systemHealthLoading}
          onRefreshHealth={() => void refreshSystemHealth()}
          contactSources={contactSources}
          contactSourcesLoading={contactSourcesLoading}
          contactSourcesError={contactSourcesError}
          onRefreshContactSources={() => void refreshContactSources()}
          onCreateProjectFromContact={startProjectFromContact}
          accessUsers={managedAccessUsers}
          pendingRequests={pendingAccessRequests}
          onSaveUser={saveManagedUser}
          onApproveRequest={approveUserRequest}
          onRejectRequest={rejectUserRequest}
          projects={visibleProjects}
          activeProjects={activeVisibleProjects}
          archivedProjects={archivedVisibleProjects}
          users={approvedAccessUsers}
          exportedBy={signedInUser.fullName}
          onExport={handleExportMicrosoftListsPackage}
          onClearProjects={clearProjectData}
          onRestoreProject={restoreProject}
        />
      )}
      {editing && (
        <ProjectForm
          project={editing}
          onCancel={() => setEditing(null)}
          onSave={upsertProject}
          auditorOptions={auditorOptions}
          contactSources={contactSources}
          existingProjects={activeVisibleProjects}
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
  const authStatus = new URLSearchParams(window.location.search).get("auth");
  const pendingVerification =
    access?.status === "pending-verification" ||
    access?.user?.accessRequestStatus === "Pending Verification";
  const pendingApproval =
    access?.status === "pending-approval" ||
    access?.user?.accessRequestStatus === "Pending Approval";
  const rejected =
    access?.status === "rejected" || access?.user?.accessRequestStatus === "Rejected";
  const setupRequired = access?.status === "setup-required";
  const requestRequired =
    authStatus === "request-required" ||
    access?.status === "request-required" ||
    access?.status === "not-requested";

  return (
    <main className="login-shell">
      <section className="microsoft-login-stack">
        <div className="login-company-brand">
          <img src={mosaicLogoUrl} alt="Mosaic International Insurance Professionals" />
          <span>Audit Assignment Tracker</span>
        </div>
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
          <h1>Audit Assignment Tracker</h1>
          <p className="microsoft-login-copy">
            Secure access uses your company Microsoft account. If you are new or
            unapproved, start with account approval before trying normal sign-in.
          </p>
          <div className="access-instructions">
            <strong>New users must do this first:</strong>
            <ol>
              <li>Click Request account approval.</li>
              <li>Sign in with your company Microsoft account.</li>
              <li>Enter the email verification code.</li>
              <li>Wait for an admin to approve the profile.</li>
            </ol>
          </div>
          {loading && <div className="toast">Checking secure access...</div>}
          {requestRequired && (
            <div className="secure-setup-box access-warning">
              <strong>Access request required.</strong>
              <span>
                This Microsoft account is not approved for the tracker yet.
                Use Request account approval first, then confirm the emailed code.
              </span>
            </div>
          )}
          {access?.status === "setup-required" && (
            <div className="secure-setup-box">
              <strong>Secure access server is not configured.</strong>
              <span>
                Sign-in will be enabled after Microsoft Entra and Graph email
                settings are saved in server.env and the secure server is restarted.
              </span>
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
            {setupRequired ? (
              <button type="button" className="microsoft-primary-link is-disabled" disabled>
                Sign in with Microsoft
              </button>
            ) : (
              <a className="microsoft-primary-link" href={signInUrl}>
                Sign in with Microsoft
              </a>
            )}
            <button type="button" className="secondary" onClick={onRefresh}>
              Refresh
            </button>
          </div>
          <p className="microsoft-create-line">
            New or unapproved user?{" "}
            {setupRequired ? (
              <span className="disabled-link">Request account approval</span>
            ) : (
              <a className="request-access-link" href={requestAccessUrl}>
                Request account approval
              </a>
            )}
          </p>
          {setupRequired ? (
            <span className="microsoft-help-link disabled-link">
              Start access request
            </span>
          ) : (
            <a className="microsoft-help-link" href={requestAccessUrl}>
              Start access request
            </a>
          )}
        </div>
        {setupRequired ? (
          <div className="signin-options-card is-disabled">
            <span aria-hidden="true">?</span>
            Sign-in options
          </div>
        ) : (
          <a className="signin-options-card" href={signInUrl}>
            <span aria-hidden="true">?</span>
            Sign-in options
          </a>
        )}
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

function AppNavigation({
  activeSection,
  setActiveSection,
  currentUser,
  activeCount,
  archivedCount,
  pendingRequestCount,
}: {
  activeSection: AppSection;
  setActiveSection: (section: AppSection) => void;
  currentUser: PrototypeUser;
  activeCount: number;
  archivedCount: number;
  pendingRequestCount: number;
}) {
  const items: { id: AppSection; label: string; helper: string; count?: number }[] = [
    {
      id: "dashboard",
      label: "Dashboard",
      helper: "Daily workload, due items, and auditor capacity.",
      count: activeCount,
    },
    {
      id: "assignments",
      label: "Assignments",
      helper: "Working board, filters, project detail, and intake/edit forms.",
      count: activeCount,
    },
    {
      id: "scheduling",
      label: "Scheduling",
      helper: "Audit dates, notes, Outlook invites, and conflict warnings.",
    },
    {
      id: "command",
      label: "Command center",
      helper: "SLA escalation, draft queue, and manager operating brief.",
    },
    {
      id: "reports",
      label: "Reports",
      helper: "Cycle time, workload reporting, and archived project access.",
      count: archivedCount,
    },
  ];
  if (currentUser.role === "Admin") {
    items.push({
      id: "admin",
      label: "Admin",
      helper: "User approvals, audit log, storage controls, and system health.",
      count: pendingRequestCount,
    });
  }

  return (
    <nav className="app-nav panel" aria-label="Primary tracker sections">
      {items.map((item) => (
        <button
          key={item.id}
          type="button"
          className={activeSection === item.id ? "active" : "secondary"}
          onClick={() => setActiveSection(item.id)}
          title={item.helper}
        >
          <span>{item.label}</span>
          {typeof item.count === "number" && <small>{item.count}</small>}
        </button>
      ))}
    </nav>
  );
}

function AdminWorkspace({
  activeTab,
  setActiveTab,
  health,
  loading,
  onRefreshHealth,
  contactSources,
  contactSourcesLoading,
  contactSourcesError,
  onRefreshContactSources,
  onCreateProjectFromContact,
  accessUsers,
  pendingRequests,
  onSaveUser,
  onApproveRequest,
  onRejectRequest,
  projects,
  activeProjects,
  archivedProjects,
  users,
  exportedBy,
  onExport,
  onClearProjects,
  onRestoreProject,
}: {
  activeTab: AdminTab;
  setActiveTab: (tab: AdminTab) => void;
  health: SecureSystemHealth | null;
  loading: boolean;
  onRefreshHealth: () => void;
  contactSources: LinkedContactSourcesResponse | null;
  contactSourcesLoading: boolean;
  contactSourcesError: string;
  onRefreshContactSources: () => void;
  onCreateProjectFromContact: (contact: LinkedContact) => void;
  accessUsers: PrototypeUser[];
  pendingRequests: PrototypeUser[];
  onSaveUser: (email: string, user: PrototypeUser) => void;
  onApproveRequest: (username: string, update: AccessApprovalUpdate) => void;
  onRejectRequest: (username: string) => void;
  projects: AuditProject[];
  activeProjects: AuditProject[];
  archivedProjects: AuditProject[];
  users: PrototypeUser[];
  exportedBy: string;
  onExport: () => void;
  onClearProjects: () => void;
  onRestoreProject: (project: AuditProject) => void;
}) {
  const tabs: { id: AdminTab; label: string; helper: string }[] = [
    {
      id: "users",
      label: "Users",
      helper: "Approve Microsoft accounts and control role/visibility.",
    },
    {
      id: "contacts",
      label: "Contacts",
      helper: "Preview linked OneDrive spreadsheet contacts and instructions.",
    },
    {
      id: "activity",
      label: "Audit log",
      helper: "Review recent project changes across visible records.",
    },
    {
      id: "storage",
      label: "Storage & data",
      helper: "Microsoft Lists status, backups, archive, and destructive controls.",
    },
    {
      id: "health",
      label: "System health",
      helper: "Security, Graph consent, runtime, and deployment checks.",
    },
  ];

  return (
    <section className="admin-workspace">
      <div className="panel admin-intro">
        <div>
          <p className="eyebrow dark">Admin workspace</p>
          <h2>Production controls</h2>
          <span>
            Daily controls are separated from backend checks. Storage actions below
            preserve Microsoft Lists data unless you explicitly clear or replace records.
          </span>
        </div>
      </div>
      <div className="admin-layout">
        <aside className="panel admin-tabs" aria-label="Admin sections">
          {tabs.map((tab) => (
            <button
              key={tab.id}
              type="button"
              className={activeTab === tab.id ? "active" : "secondary"}
              onClick={() => setActiveTab(tab.id)}
              title={tab.helper}
            >
              <strong>{tab.label}</strong>
              <span>{tab.helper}</span>
            </button>
          ))}
        </aside>
        <div className="admin-tab-content">
          {activeTab === "users" && (
            <UserManagementPanel
              accessUsers={accessUsers}
              pendingRequests={pendingRequests}
              onSaveUser={onSaveUser}
              onApproveRequest={onApproveRequest}
              onRejectRequest={onRejectRequest}
            />
          )}
          {activeTab === "contacts" && (
            <ContactSourcesPanel
              contactSources={contactSources}
              loading={contactSourcesLoading}
              error={contactSourcesError}
              health={health}
              onRefresh={onRefreshContactSources}
              onCreateProjectFromContact={onCreateProjectFromContact}
            />
          )}
          {activeTab === "activity" && <AdminActivityPanel projects={projects} />}
          {activeTab === "storage" && (
            <>
              <DataSafetyPanel
                activeCount={activeProjects.length}
                archivedCount={archivedProjects.length}
                onClearProjects={onClearProjects}
              />
              <CentralStoragePanel
                projects={projects}
                users={users}
                exportedBy={exportedBy}
                health={health}
                onExport={onExport}
              />
              <ArchivedProjectsPanel
                projects={archivedProjects}
                canRestore
                onRestore={onRestoreProject}
              />
            </>
          )}
          {activeTab === "health" && (
            <SystemReadinessPanel
              health={health}
              loading={loading}
              onRefresh={onRefreshHealth}
            />
          )}
        </div>
      </div>
    </section>
  );
}

function DataSafetyPanel({
  activeCount,
  archivedCount,
  onClearProjects,
}: {
  activeCount: number;
  archivedCount: number;
  onClearProjects: () => void;
}) {
  return (
    <section className="panel data-safety-panel">
      <div className="section-title">
        <div>
          <p className="eyebrow dark">Data safety</p>
          <h2>Deployments do not clear projects</h2>
          <span>
            Code updates replace the app files only. Project records stay in Microsoft Lists
            unless an admin intentionally clears data or imports a replacement JSON file.
          </span>
        </div>
        <button type="button" className="danger-button" onClick={onClearProjects}>
          Clear all projects
        </button>
      </div>
      <div className="storage-stats">
        <span>
          <strong>{activeCount}</strong>
          Active records
        </span>
        <span>
          <strong>{archivedCount}</strong>
          Archived records
        </span>
        <span>
          <strong>0</strong>
          Automatic deletes on deploy
        </span>
        <span>
          <strong>2</strong>
          Guarded destructive actions
        </span>
      </div>
    </section>
  );
}

function ArchivedProjectsPanel({
  projects,
  canRestore,
  onRestore,
}: {
  projects: AuditProject[];
  canRestore: boolean;
  onRestore: (project: AuditProject) => void;
}) {
  return (
    <section className="panel archived-projects">
      <div className="section-title">
        <div>
          <p className="eyebrow dark">Archive</p>
          <h2>Archived projects</h2>
          <span>
            Archived records are hidden from active boards and workload counts but
            remain saved for history and reporting.
          </span>
        </div>
        <span className="archive-count">{projects.length}</span>
      </div>
      {projects.length === 0 ? (
        <p className="muted-note">No archived projects yet.</p>
      ) : (
        <div className="archive-list">
          {projects.map((project) => (
            <article key={project.id} className="archive-row">
              <div>
                <strong>{project.assignmentNumber}</strong>
                <span>{project.auditEntity || project.clientCoverholderCode || "No entity"}</span>
                <small>
                  {project.currentStage} | {formatAuditTeam(project)} | updated {project.lastUpdatedDate}
                </small>
              </div>
              <button
                type="button"
                className="secondary"
                disabled={!canRestore}
                onClick={() => onRestore(project)}
              >
                Restore
              </button>
            </article>
          ))}
        </div>
      )}
    </section>
  );
}

function ContactSourcesPanel({
  contactSources,
  loading,
  error,
  health,
  onRefresh,
  onCreateProjectFromContact,
}: {
  contactSources: LinkedContactSourcesResponse | null;
  loading: boolean;
  error: string;
  health: SecureSystemHealth | null;
  onRefresh: () => void;
  onCreateProjectFromContact: (contact: LinkedContact) => void;
}) {
  const [query, setQuery] = useState("");
  const [filter, setFilter] = useState<ContactFilter>("All");
  const sourceStatus = health?.contactSources;
  const contacts = contactSources?.contacts ?? [];
  const instructionsCount = contacts.reduce(
    (total, contact) => total + contact.specialInstructions.length,
    0,
  );
  const hasRefreshed = Boolean(contactSources);
  const normalizedQuery = query.trim().toLowerCase();
  const filteredContacts = contacts.filter(
    (contact) =>
      contactMatchesFilter(contact, filter) &&
      (!normalizedQuery || contactDirectorySearchText(contact).includes(normalizedQuery)),
  );
  const visibleContacts = filteredContacts.slice(0, 40);
  const warningPreview = contactSources?.warnings.slice(0, 12) ?? [];
  return (
    <section className="panel contact-sources-panel">
      <div className="section-title">
        <div>
          <p className="eyebrow dark">Contact directory</p>
          <h2>Linked spreadsheet contacts</h2>
          <span>
            OneDrive/SharePoint workbooks stay the source of truth. Refresh pulls
            current rows and special instructions through Microsoft Graph.
          </span>
        </div>
        <button type="button" disabled={loading} onClick={onRefresh}>
          {loading ? "Refreshing..." : "Refresh contacts"}
        </button>
      </div>
      <div className="readiness-grid">
        <ReadinessCard
          label="Workbook links"
          ready={Boolean(sourceStatus?.configured)}
          detail={
            sourceStatus?.status ||
            "Add TRACKER_CONTACT_WORKBOOK_LINKS to Azure App Service settings."
          }
        />
        <ReadinessCard
          label="Graph access"
          ready={Boolean(health?.sharePoint.permissionGranted)}
          detail={
            health?.sharePoint.permissionGranted
              ? "Sites.ReadWrite.All is available for workbook reads."
              : "Waiting for Microsoft Graph SharePoint permission."
          }
        />
        <ReadinessCard
          label="Loaded contacts"
          ready={contacts.length > 0}
          detail={
            hasRefreshed
              ? `${contacts.length} contacts, ${instructionsCount} special instruction field values.`
              : "Not refreshed yet. Open this tab or use Refresh contacts to read the workbook links."
          }
        />
      </div>
      {error && (
        <div className="inline-alert">
          <strong>Refresh stopped</strong>
          <span>{error}</span>
          <small>Use Refresh contacts to retry after the workbook or Graph issue is corrected.</small>
        </div>
      )}
      {contactSources && (
        <div className="contact-source-grid">
          {contactSources.sources.map((source) => (
            <article
              key={source.id}
              className={`contact-source-card ${source.status === "ok" ? "ready" : "blocked"}`}
            >
              <span>{source.status === "ok" ? "Connected" : "Needs review"}</span>
              <strong>{source.workbookName || source.label}</strong>
              <p>
                {source.status === "ok"
                  ? `${source.worksheetCount} worksheets, ${source.rowCount} contact rows.`
                  : source.error || "Workbook could not be read."}
              </p>
            </article>
          ))}
        </div>
      )}
      {contactSources?.warnings.length ? (
        <div className="readiness-next">
          <strong>Contact parsing warnings</strong>
          <span>
            These are workbook tabs that loaded but need review. Blank template
            tabs are ignored automatically.
          </span>
          <ul>
            {warningPreview.map((warning) => (
              <li key={`${warning.sourceId}-${warning.worksheetName}-${warning.message}`}>
                {warning.worksheetName ? `${warning.worksheetName}: ` : ""}
                {warning.message}
              </li>
            ))}
          </ul>
          {contactSources.warnings.length > warningPreview.length && (
            <small>
              Showing {warningPreview.length} of {contactSources.warnings.length} warnings.
            </small>
          )}
        </div>
      ) : null}
      {contactSources && contacts.length === 0 ? (
        <p className="muted-note">
          No contacts loaded yet. If workbook links are configured, confirm the
          files have a header row and at least one contact/data row.
        </p>
      ) : null}
      {contacts.length > 0 && (
        <div className="contact-tools">
          <Input
            label="Search contacts"
            value={query}
            placeholder="Client, company, email, DCA, coverholder, instruction"
            onChange={setQuery}
          />
          <Select
            label="Receiver type"
            value={filter}
            options={["All", "DCA", "Coverholder", "Report", "Invoice", "Missing Email"]}
            placeholder="All"
            onChange={(value) => setFilter((value || "All") as ContactFilter)}
          />
          <span>
            Showing {visibleContacts.length} of {filteredContacts.length} matching contacts
          </span>
        </div>
      )}
      {visibleContacts.length > 0 && (
        <div className="contact-preview-list">
          {visibleContacts.map((contact) => (
            <article key={contact.id} className="contact-preview-card">
              <div>
                <strong>{contact.contactName || contact.company || "Unnamed contact"}</strong>
                <span>
                  {contact.email || "No email"} |{" "}
                  {contact.company || contact.coverholder || contact.managingAgent || "No company"}
                </span>
                <small>
                  {contact.workbookName} / {contact.worksheetName}
                </small>
              </div>
              <div className="contact-card-actions">
                <button
                  type="button"
                  className="secondary"
                  onClick={() => onCreateProjectFromContact(contact)}
                >
                  Start project
                </button>
              </div>
              <div className="contact-tags">
                {contact.broker && <span>Broker: {contact.broker}</span>}
                {contact.managingAgent && <span>MA: {contact.managingAgent}</span>}
                {contact.coverholder && <span>Coverholder: {contact.coverholder}</span>}
                {contact.role && <span>{contact.role}</span>}
              </div>
              {contactEmailBuckets(contact).length > 0 && (
                <div className="contact-email-grid">
                  {contactEmailBuckets(contact).map(([label, emails]) => (
                    <span key={`${contact.id}-${label}`}>
                      <strong>{label}</strong>
                      {emails.join("; ")}
                    </span>
                  ))}
                </div>
              )}
              {contact.specialInstructions.length > 0 && (
                <div className="special-instructions">
                  <strong>Special instructions</strong>
                  {contact.specialInstructions.map((instruction) => (
                    <p key={`${contact.id}-${instruction.label}`}>
                      {instruction.label}: {instruction.value}
                    </p>
                  ))}
                </div>
              )}
            </article>
          ))}
        </div>
      )}
    </section>
  );
}

function SystemReadinessPanel({
  health,
  loading,
  onRefresh,
}: {
  health: SecureSystemHealth | null;
  loading: boolean;
  onRefresh: () => void;
}) {
  const graphRoles = health?.graphApp.roles ?? [];
  const consentReady =
    graphRoles.includes("Mail.Send") && graphRoles.includes("Sites.ReadWrite.All");
  const calendarReady = Boolean(health?.calendar?.permissionGranted);
  const approvalStoreReady = Boolean(health?.approvalStore.durable);
  const projectStoreReady = Boolean(health?.projectStore?.durable);
  const configSource = health?.runtime.configSource ?? "Waiting for server check";
  const deployCommit = health?.deployment.commit
    ? health.deployment.commit.slice(0, 7)
    : "Not reported";
  return (
    <section className="panel readiness-panel">
      <div className="section-title">
        <div>
          <p className="eyebrow dark">Admin readiness</p>
          <h2>Security and rollout health</h2>
          <span>
            Server-side checks for Microsoft sign-in, email sending, SharePoint
            access, and durable approval storage.
          </span>
        </div>
        <button type="button" className="secondary" disabled={loading} onClick={onRefresh}>
          {loading ? "Checking..." : "Refresh checks"}
        </button>
      </div>
      <div className="readiness-grid">
        <ReadinessCard
          label="Secure access server"
          ready={Boolean(health?.server.configured)}
          detail={
            health?.server.configured
              ? `Redirect URI: ${health.server.redirectUri}`
              : `Missing: ${health?.server.missing.join(", ") || "health check"}`
          }
        />
        <ReadinessCard
          label="Microsoft app token"
          ready={Boolean(health?.graphApp.tokenAvailable)}
          detail={
            health?.graphApp.tokenAvailable
              ? health.graphApp.appDisplayName || "Token request succeeded"
              : health?.graphApp.error || "Waiting for server check"
          }
        />
        <ReadinessCard
          label="Graph admin consent"
          ready={consentReady}
          detail={
            consentReady
              ? "Mail.Send and Sites.ReadWrite.All are active."
              : `Missing roles: ${health?.graphApp.missingRoles.join(", ") || "not checked"}`
          }
        />
        <ReadinessCard
          label="Outlook calendar"
          ready={calendarReady}
          detail={
            calendarReady
              ? "Calendars.ReadWrite is active for Outlook invite sending."
              : "Grant Calendars.ReadWrite application consent to create Outlook project events."
          }
        />
        <ReadinessCard
          label="Approval storage"
          ready={approvalStoreReady}
          detail={
            health?.approvalStore.status ||
            "Local approval storage is active until Microsoft Lists is approved."
          }
        />
        <ReadinessCard
          label="Project storage"
          ready={projectStoreReady}
          detail={
            health?.projectStore?.status ||
            "Projects save on the app server until Microsoft Lists is configured."
          }
        />
        <ReadinessCard
          label="Runtime config source"
          ready={Boolean(health?.runtime.persistentDataEnvLoaded || health?.runtime.appSettingKeyCount)}
          detail={
            health
              ? `${configSource}. ${health.runtime.appSettingKeyCount} app setting keys, ${health.runtime.localFileKeyCount} local-file keys.`
              : "Waiting for server check"
          }
        />
        <ReadinessCard
          label="Live deployment"
          ready={Boolean(health?.deployment.commit || health?.server.publicOrigin)}
          detail={
            health?.deployment.deployedAt
              ? `${deployCommit} deployed ${formatDateTime(health.deployment.deployedAt)}`
              : `${deployCommit}. Public origin: ${health?.server.publicOrigin || "not checked"}`
          }
        />
      </div>
      {health && (
        <div className="runtime-summary">
          <span>Health checked {formatDateTime(health.generatedAt)}</span>
          <span>Node {health.runtime.nodeVersion}</span>
          <span>{health.runtime.websiteHostname || health.server.publicOrigin}</span>
        </div>
      )}
      {health && (
        <div className="readiness-next">
          <strong>Next required actions</strong>
          <ul>
            {health.recommendations.map((item) => (
              <li key={item}>{item}</li>
            ))}
          </ul>
        </div>
      )}
    </section>
  );
}

function ReadinessCard({
  label,
  ready,
  detail,
}: {
  label: string;
  ready: boolean;
  detail: string;
}) {
  return (
    <article className={`readiness-card ${ready ? "ready" : "blocked"}`}>
      <span>{ready ? "Ready" : "Action needed"}</span>
      <strong>{label}</strong>
      <p>{detail}</p>
    </article>
  );
}

function CentralStoragePanel({
  projects,
  users,
  exportedBy,
  health,
  onExport,
}: {
  projects: AuditProject[];
  users: PrototypeUser[];
  exportedBy: string;
  health: SecureSystemHealth | null;
  onExport: () => void;
}) {
  const packagePreview = useMemo(
    () =>
      buildMicrosoftListsMigrationPackage(projects, users, {
        exportedBy,
        exportedAt: "preview",
      }),
    [projects, users, exportedBy],
  );
  const approvalStoreLabel =
    health?.approvalStore.mode === "microsoft-lists"
      ? "Tracker Users list"
      : "Local server file";
  const projectStoreLabel =
    health?.projectStore?.mode === "microsoft-lists"
      ? "Audit Assignments list"
      : "App server file";
  const sharePointReady = Boolean(health?.sharePoint.permissionGranted);
  const trackerUsersReady = Boolean(
    health?.sharePoint.siteIdConfigured &&
      health?.sharePoint.trackerUsersListIdConfigured,
  );
  const projectStoreReady = Boolean(health?.projectStore?.durable);
  const durableRuntimeConfig = Boolean(
    health?.runtime.appSettingKeyCount || health?.runtime.persistentDataEnvLoaded,
  );
  const detailRows =
    packagePreview.rows.auditTeamMembers.length +
    packagePreview.rows.auditComments.length +
    packagePreview.rows.auditChecklistItems.length +
    packagePreview.rows.auditStatusHistory.length +
    packagePreview.rows.auditActivityLog.length;
  const migrationSteps = [
    {
      label: "1. Secure sign-in",
      ready: Boolean(health?.server.configured),
      detail: "Microsoft OAuth and signed server sessions control tracker access.",
    },
    {
      label: "2. Durable runtime config",
      ready: durableRuntimeConfig,
      detail:
        health?.runtime.configSource ||
        "Move secrets to Azure App Service settings or the persistent Azure data env file.",
    },
    {
      label: "3. Approval storage",
      ready: trackerUsersReady,
      detail: trackerUsersReady
        ? "Tracker Users list identifiers are configured."
        : "Create/configure the Tracker Users Microsoft List before switching approvals.",
    },
    {
      label: "4. Project storage",
      ready: projectStoreReady,
      detail: projectStoreReady
        ? "Audit Assignments list is the active project store."
        : "Create/configure the Audit Assignments Microsoft List before switching projects.",
    },
    {
      label: "5. Activity log",
      ready: detailRows > 0,
      detail: "Status history, comments, checklist rows, and activity events are export-ready.",
    },
  ];
  const listGroups = [
    {
      label: "Core project records",
      lists: ["Audit Assignments", "Audit Team Members"],
    },
    {
      label: "Audit history",
      lists: ["Audit Comments", "Audit Checklist Items", "Audit Status History", "Audit Activity Log"],
    },
    {
      label: "Access control",
      lists: ["Tracker Users"],
    },
  ];

  return (
    <section className="panel central-storage">
      <div className="section-title">
        <div>
          <p className="eyebrow dark">Data foundation</p>
          <h2>Microsoft Lists data foundation</h2>
          <span>
            Server settings control Microsoft Graph and shared project data.
            This panel shows what is live now and what still needs Microsoft
            Lists configuration.
          </span>
        </div>
        <div className="storage-actions">
          <button type="button" className="secondary" onClick={onExport}>
            Export migration package
          </button>
        </div>
      </div>
      <div className={`storage-status ${health?.approvalStore.durable ? "connected" : ""}`}>
        <strong>Access approvals: {approvalStoreLabel}</strong>
        <span>
          {health?.approvalStore.durable
            ? "New account approvals are stored in Microsoft Lists."
            : "New account approvals are stored on the app server until TRACKER_USER_STORE is switched to microsoft-lists."}
        </span>
      </div>
      <div className={`storage-status ${projectStoreReady ? "connected" : ""}`}>
        <strong>Project records: {projectStoreLabel}</strong>
        <span>
          {health?.projectStore?.status ||
            "Project records use server storage until Microsoft Lists project storage is configured."}
        </span>
      </div>
      <div className="readiness-grid">
        <ReadinessCard
          label="Graph permissions"
          ready={sharePointReady}
          detail={
            sharePointReady
              ? "Sites.ReadWrite.All is available for Microsoft Lists work."
              : "Waiting for Sites.ReadWrite.All application permission."
          }
        />
        <ReadinessCard
          label="Tracker Users list"
          ready={trackerUsersReady}
          detail={
            trackerUsersReady
              ? "Site ID and Tracker Users list ID are configured."
              : "Add TRACKER_USERS_SITE_ID and TRACKER_USERS_LIST_ID before moving approvals."
          }
        />
        <ReadinessCard
          label="Audit Assignments list"
          ready={projectStoreReady}
          detail={
            projectStoreReady
              ? "Project reads and writes are configured for Microsoft Lists."
              : "Add TRACKER_PROJECTS_SITE_ID and TRACKER_PROJECTS_LIST_ID before moving project storage."
          }
        />
        <ReadinessCard
          label="Project data package"
          ready
          detail={`${packagePreview.totals.lists} lists, ${packagePreview.totals.rows} rows ready for export.`}
        />
      </div>
      <div className="storage-stats">
        <span>
          <strong>{packagePreview.totals.lists}</strong>
          Lists defined
        </span>
        <span>
          <strong>{packagePreview.totals.assignments}</strong>
          Project records
        </span>
        <span>
          <strong>{detailRows}</strong>
          Detail rows
        </span>
        <span>
          <strong>{packagePreview.rows.trackerUsers.length}</strong>
          Directory rows
        </span>
      </div>
      <div className="storage-roadmap">
        {migrationSteps.map((step) => (
          <div key={step.label} className={step.ready ? "ready" : ""}>
            <strong>{step.label}</strong>
            <span>{step.detail}</span>
          </div>
        ))}
      </div>
      <div className="storage-schema-groups" aria-label="Microsoft Lists schema">
        {listGroups.map((group) => (
          <article key={group.label}>
            <strong>{group.label}</strong>
            <div className="storage-list-chips">
              {group.lists.map((label) => (
                <span key={label}>{label}</span>
              ))}
            </div>
          </article>
        ))}
      </div>
    </section>
  );
}

function UserManagementPanel({
  accessUsers,
  pendingRequests,
  onSaveUser,
  onApproveRequest,
  onRejectRequest,
}: {
  accessUsers: PrototypeUser[];
  pendingRequests: PrototypeUser[];
  onSaveUser: (email: string, user: PrototypeUser) => void;
  onApproveRequest: (username: string, update: AccessApprovalUpdate) => void;
  onRejectRequest: (username: string) => void;
}) {
  const approvedAccessUsers = accessUsers
    .filter((user) => user.accessRequestStatus === "Approved")
    .sort((a, b) => a.fullName.localeCompare(b.fullName));
  const firstEmail = approvedAccessUsers[0]?.email ?? "";
  const [selectedEmail, setSelectedEmail] = useState(firstEmail);
  const selectedUser =
    approvedAccessUsers.find((user) => user.email === selectedEmail) ??
    approvedAccessUsers[0];
  const [draft, setDraft] = useState<PrototypeUser>(
    selectedUser ?? secureUserToPrototypeUserPlaceholder(),
  );
  const [approvalDrafts, setApprovalDrafts] = useState<
    Record<string, AccessApprovalUpdate>
  >({});
  const waitingOnEmailCount = pendingRequests.filter(
    (user) => user.accessRequestStatus === "Pending Verification",
  ).length;
  const waitingOnAdminCount = pendingRequests.filter(
    (user) => user.accessRequestStatus === "Pending Approval",
  ).length;
  const approvedUsersSignature = approvedAccessUsers
    .map((user) =>
      [
        user.email,
        user.fullName,
        user.role,
        user.defaultVisibility,
        user.active,
      ].join(":"),
    )
    .join("|");

  const loadUser = (user: PrototypeUser) => {
    setSelectedEmail(user.email);
    setDraft(user);
  };

  useEffect(() => {
    if (!approvedAccessUsers.length) return;
    const latest =
      approvedAccessUsers.find((user) => user.email === selectedEmail) ??
      approvedAccessUsers[0];
    setSelectedEmail(latest.email);
    setDraft(latest);
  }, [approvedUsersSignature, selectedEmail]);

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

  const approvalDraftFor = (user: PrototypeUser): AccessApprovalUpdate => ({
    role: approvalDrafts[user.username]?.role ?? user.role,
    defaultVisibility:
      approvalDrafts[user.username]?.defaultVisibility ?? user.defaultVisibility,
  });

  const updateApprovalDraft = (
    user: PrototypeUser,
    update: AccessApprovalUpdate,
  ) => {
    setApprovalDrafts((current) => ({
      ...current,
      [user.username]: {
        ...approvalDraftFor(user),
        ...update,
      },
    }));
  };

  return (
    <section className="panel user-management">
      <div className="section-title">
        <div>
          <p className="eyebrow dark">Admin</p>
          <h2>Access and assignment directory</h2>
          <span>
            Approved Microsoft accounts are the directory. Pending, rejected,
            and unverified users are excluded from workload and assignment
            owner lists.
          </span>
        </div>
      </div>
      <div className="user-admin-summary">
        <span>
          <strong>{waitingOnAdminCount}</strong>
          Awaiting admin approval
        </span>
        <span>
          <strong>{waitingOnEmailCount}</strong>
          Awaiting email code
        </span>
        <span>
          <strong>{approvedAccessUsers.length}</strong>
          Approved accounts
        </span>
        <span>
          <strong>{approvedAccessUsers.filter((user) => user.active).length}</strong>
          Active in directory
        </span>
      </div>
      <div className="user-management-grid">
        <div className="access-request-queue">
          <div>
            <h3>Microsoft access requests</h3>
            <span>
              {pendingRequests.length > 0
                ? `${pendingRequests.length} request${pendingRequests.length === 1 ? "" : "s"} in progress`
                : "No account requests waiting"}
            </span>
          </div>
          {pendingRequests.length === 0 ? (
            <p className="muted-note">
              New users appear here after they sign in with Microsoft, request
              tracker access, and confirm their email code.
            </p>
          ) : (
            pendingRequests.map((user) => (
              <article className="access-request-card" key={user.username}>
                <div>
                  <strong>{user.fullName}</strong>
                  <span>{user.email}</span>
                  <small>
                    {user.accessRequestStatus === "Pending Approval"
                      ? "Email confirmed"
                      : "Waiting on email code"}
                  </small>
                </div>
                <div className="approval-controls">
                  <Select
                    label="Role on approval"
                    value={approvalDraftFor(user).role ?? "Auditor"}
                    options={userRoleOptions}
                    placeholder="Select role"
                    onChange={(value) =>
                      updateApprovalDraft(user, { role: value as UserRole })
                    }
                  />
                  <Select
                    label="Default visibility"
                    value={approvalDraftFor(user).defaultVisibility ?? "Role Default"}
                    options={projectVisibilityOptions}
                    placeholder="Select visibility"
                    onChange={(value) =>
                      updateApprovalDraft(user, {
                        defaultVisibility: value as ProjectVisibility,
                      })
                    }
                  />
                </div>
                <div className="storage-actions">
                  <button
                    type="button"
                    disabled={!canApproveAccessRequest(user)}
                    onClick={() => onApproveRequest(user.username, approvalDraftFor(user))}
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
            ))
          )}
        </div>
        <div className="user-list approved-directory-list" aria-label="Approved Microsoft account directory">
          <div className="section-subhead">
            <div>
              <h3>Approved account directory</h3>
              <span>These users can sign in and appear in assignment ownership tools.</span>
            </div>
          </div>
          {approvedAccessUsers.length === 0 ? (
            <p className="muted-note">
              Approved Microsoft users will appear here after account requests
              are confirmed and approved.
            </p>
          ) : (
            approvedAccessUsers.map((user) => (
              <button
                type="button"
                className={user.email === selectedEmail ? "selected secondary" : "secondary"}
                key={user.email}
                onClick={() => loadUser(user)}
              >
                <strong>{user.fullName}</strong>
                <span>{user.email}</span>
                <span className="user-meta">
                  <span className="user-status-badge active">{user.role}</span>
                  <span className="user-status-badge visibility">
                    {user.defaultVisibility}
                  </span>
                  <span
                    className={`user-status-badge ${
                      user.active ? "active" : "inactive"
                    }`}
                  >
                    {user.active ? "Active" : "Inactive"}
                  </span>
                </span>
              </button>
            ))
          )}
        </div>
        <form
          className="user-editor"
          onSubmit={(event) => {
            event.preventDefault();
            if (!draft.email) return;
            const preparedDraft = withUserDefaults({
              ...draft,
              permissionGroup: draft.role,
              emailVerified: true,
              accessRequestStatus: "Approved",
            });
            onSaveUser(draft.email, preparedDraft);
          }}
        >
          <div className="section-subhead">
            <div>
              <h3>Manage approved account</h3>
              <span>
                Role, visibility, and active status save to the secure access
                server and then refresh this directory.
              </span>
            </div>
          </div>
          {approvedAccessUsers.length === 0 ? (
            <p className="muted-note">
              Approve a Microsoft access request before editing user roles.
            </p>
          ) : (
            <>
              <div className="selected-account-summary">
                <span className="avatar">
                  {draft.fullName
                    .split(" ")
                    .map((part) => part[0])
                    .slice(0, 2)
                    .join("")}
                </span>
                <div>
                  <strong>{draft.fullName}</strong>
                  <span>{draft.email}</span>
                  <div className="user-meta">
                    <span className="user-status-badge active">{draft.role}</span>
                    <span className="user-status-badge visibility">
                      {draft.defaultVisibility}
                    </span>
                  </div>
                </div>
              </div>
              <div className="form-grid user-editor-grid">
                <Input
                  label="Full name"
                  value={draft.fullName}
                  onChange={(value) => updateDraft("fullName", value)}
                />
                <label>
                  Company email
                  <input value={draft.email} readOnly />
                </label>
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
                  label="Active account"
                  checked={draft.active}
                  onChange={(value) => updateDraft("active", value)}
                />
              </div>
              <div className="modal-actions">
                <button type="submit">Save approved account</button>
              </div>
            </>
          )}
        </form>
      </div>
    </section>
  );
}

function AdminActivityPanel({ projects }: { projects: AuditProject[] }) {
  const [query, setQuery] = useState("");
  const [typeFilter, setTypeFilter] = useState<"All" | ActivityItem["type"]>("All");
  const events = useMemo(
    () =>
      projects
        .flatMap((project) =>
          activityTimeline(project).map((item) => ({
            ...item,
            projectId: project.id,
            assignmentNumber: project.assignmentNumber,
            auditEntity: project.auditEntity,
          })),
        )
        .sort((a, b) => Date.parse(b.timestamp) - Date.parse(a.timestamp)),
    [projects],
  );
  const normalizedQuery = query.trim().toLowerCase();
  const filteredEvents = events.filter((event) => {
    const matchesType = typeFilter === "All" || event.type === typeFilter;
    const searchable = [
      event.assignmentNumber,
      event.auditEntity,
      event.title,
      event.detail,
      event.type,
    ]
      .join(" ")
      .toLowerCase();
    return matchesType && (!normalizedQuery || searchable.includes(normalizedQuery));
  });

  return (
    <section className="panel admin-activity-panel">
      <div className="section-title">
        <div>
          <p className="eyebrow dark">Admin audit log</p>
          <h2>Recent tracker activity</h2>
          <span>Recent project changes, comments, checklist updates, and stage movement.</span>
        </div>
      </div>
      <div className="admin-activity-controls">
        <Input
          label="Search activity"
          value={query}
          placeholder="Assignment, entity, user, or action"
          onChange={setQuery}
        />
        <Select
          label="Event type"
          value={typeFilter}
          options={["All", ...activityTypeOptions]}
          placeholder="All"
          onChange={(value) => setTypeFilter(value as "All" | ActivityItem["type"])}
        />
      </div>
      <div className="admin-activity-list">
        {filteredEvents.length === 0 ? (
          <p className="muted-note">No matching activity yet.</p>
        ) : (
          filteredEvents.slice(0, 30).map((event) => (
            <article className={`activity-item ${event.type}`} key={`${event.projectId}-${event.id}`}>
              <span className="activity-marker" />
              <div>
                <div className="activity-heading">
                  <strong>{event.title}</strong>
                  <span className={`activity-type ${event.tone ?? "muted"}`}>
                    {event.type}
                  </span>
                </div>
                <small>
                  {event.assignmentNumber} | {event.auditEntity || "No entity"} | {event.timestamp}
                </small>
                <p>{event.detail}</p>
              </div>
            </article>
          ))
        )}
      </div>
    </section>
  );
}

function secureUserToPrototypeUserPlaceholder(): PrototypeUser {
  return {
    fullName: "",
    username: "",
    role: "Auditor",
    permissionGroup: "Auditor",
    email: "",
    active: true,
    defaultVisibility: "Role Default",
    emailVerified: true,
    accessRequestStatus: "Approved",
    verificationCode: "",
    requestedAt: "",
    approvedAt: "",
    approvedBy: "",
    rejectionReason: "",
  };
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

function SchedulingCapacity({
  projects,
  auditors,
  currentUser,
  calendarReady,
  onSelect,
  onUpdateScheduling,
  onSyncOutlook,
}: {
  projects: AuditProject[];
  auditors: string[];
  currentUser: PrototypeUser;
  calendarReady: boolean;
  onSelect: (id: string) => void;
  onUpdateScheduling: (
    project: AuditProject,
    update: Partial<
      Pick<
        AuditProject,
        | "calendarSyncStatus"
        | "calendarEventId"
        | "calendarEventWebLink"
        | "calendarEventLastSyncedAt"
        | "auditLocation"
        | "auditRemoteLink"
        | "auditDurationHours"
        | "auditStartTime"
        | "schedulingNotes"
        | "confirmedAuditDate"
        | "tentativeAuditWeek"
      >
    >,
  ) => void;
  onSyncOutlook: (project: AuditProject) => void;
}) {
  const scheduledProjects = projects
    .filter((project) => project.currentStage !== "Closed")
    .slice()
    .sort((a, b) =>
      (projectScheduleDate(a) || "9999-12-31").localeCompare(
        projectScheduleDate(b) || "9999-12-31",
      ),
    );
  const unscheduled = scheduledProjects.filter(
    (project) => !projectScheduleDate(project) && !project.tentativeAuditWeek,
  );
  const conflictCount = scheduledProjects.filter(
    (project) => schedulingConflictWarnings(project, projects).length > 0,
  ).length;
  const invitesSent = scheduledProjects.filter(
    (project) => project.calendarEventId && project.calendarSyncStatus === "Synced",
  ).length;
  const capacityRows = auditors
    .map((auditor) => {
      const auditorProjects = scheduledProjects.filter((project) =>
        projectHasAuditor(project, auditor),
      );
      const conflictProjects = auditorProjects.filter(
        (project) => schedulingConflictWarnings(project, projects).length > 0,
      );
      return {
        auditor,
        total: auditorProjects.length,
        nextSeven: auditorProjects.filter(
          (project) => daysUntil(projectScheduleDate(project)) >= 0 && daysUntil(projectScheduleDate(project)) <= 7,
        ).length,
        conflicts: conflictProjects.length,
      };
    })
    .filter((row) => row.total > 0 || row.conflicts > 0)
    .sort((a, b) => b.conflicts - a.conflicts || b.nextSeven - a.nextSeven);

  return (
    <section className="scheduling-workspace">
      <article className="panel scheduling-hero">
        <div>
          <p className="eyebrow dark">Microsoft 365 scheduling readiness</p>
          <h2>Scheduling & Capacity</h2>
          <span>
            Plan audit dates, review simple warnings, and send Outlook invites when ready.
          </span>
          <p className={calendarReady ? "calendar-ready-note ready" : "calendar-ready-note blocked"}>
            {calendarReady
              ? "Outlook calendar linking is ready. Scheduled audits can be added to your calendar."
              : "Outlook calendar linking needs Microsoft Graph Calendars.ReadWrite admin consent."}
          </p>
        </div>
        <div className="scheduling-stats">
          <SummaryCard label="Scheduled" value={scheduledProjects.length - unscheduled.length} />
          <SummaryCard label="Needs date" value={unscheduled.length} tone="warning" />
          <SummaryCard label="Warnings" value={conflictCount} tone="danger" />
          <SummaryCard label="Invites sent" value={invitesSent} />
        </div>
      </article>
      <div className="scheduling-grid">
        <article className="panel audit-calendar-panel">
          <div className="section-title">
            <div>
              <h2>Audit calendar</h2>
              <span>Sorted by confirmed audit date, then due date, then unscheduled items.</span>
            </div>
          </div>
          <div className="calendar-list">
            {scheduledProjects.length === 0 ? (
              <p className="muted-note">No open projects to schedule yet.</p>
            ) : (
              scheduledProjects.map((project) => (
                <SchedulingProjectCard
                  key={project.id}
                  project={project}
                  allProjects={projects}
                  canUpdate={canEditProject(currentUser, project)}
                  calendarReady={calendarReady}
                  onSelect={onSelect}
                  onUpdateScheduling={onUpdateScheduling}
                  onSyncOutlook={onSyncOutlook}
                />
              ))
            )}
          </div>
        </article>
        <aside className="panel capacity-panel">
          <div className="section-title">
            <div>
              <h2>Capacity view</h2>
              <span>Near-term workload and conflict pressure by auditor.</span>
            </div>
          </div>
          <div className="capacity-list">
            {capacityRows.length === 0 ? (
              <p className="muted-note">No scheduled auditor workload yet.</p>
            ) : (
              capacityRows.map((row) => (
                <article className="capacity-row" key={row.auditor}>
                  <span className="avatar">
                    {row.auditor
                      .split(" ")
                      .map((part) => part[0])
                      .slice(0, 2)
                      .join("")}
                  </span>
                  <div>
                    <strong>{row.auditor}</strong>
                    <small>{row.total} open scheduled assignments</small>
                  </div>
                  <div className="capacity-metrics">
                    <span>{row.nextSeven} next 7d</span>
                    <span className={row.conflicts ? "warning" : "ok"}>
                      {row.conflicts} warnings
                    </span>
                  </div>
                </article>
              ))
            )}
          </div>
        </aside>
      </div>
    </section>
  );
}

function SchedulingProjectCard({
  project,
  allProjects,
  canUpdate,
  calendarReady,
  onSelect,
  onUpdateScheduling,
  onSyncOutlook,
}: {
  project: AuditProject;
  allProjects: AuditProject[];
  canUpdate: boolean;
  calendarReady: boolean;
  onSelect: (id: string) => void;
  onUpdateScheduling: (
    project: AuditProject,
    update: Partial<
      Pick<
        AuditProject,
        | "calendarSyncStatus"
        | "calendarEventId"
        | "calendarEventWebLink"
        | "calendarEventLastSyncedAt"
        | "auditLocation"
        | "auditRemoteLink"
        | "auditDurationHours"
        | "auditStartTime"
        | "schedulingNotes"
        | "confirmedAuditDate"
        | "tentativeAuditWeek"
      >
    >,
  ) => void;
  onSyncOutlook: (project: AuditProject) => void;
}) {
  const [isScheduling, setIsScheduling] = useState(false);
  const warnings = schedulingConflictWarnings(project, allProjects);
  const scheduleDate = projectScheduleDate(project);
  const status = scheduleStatus(project);

  return (
    <article className={`calendar-card simple ${warnings.length ? "has-warning" : ""}`}>
      <div className="calendar-date">
        <strong>{scheduleDate ? scheduleDate.slice(5) : "TBD"}</strong>
        <span>{project.confirmedAuditDate ? "Confirmed" : project.tentativeAuditWeek || "No target"}</span>
      </div>
      <div className="calendar-main">
        <div className="calendar-title-row">
          <button type="button" className="link calendar-title" onClick={() => onSelect(project.id)}>
            {project.assignmentNumber} - {project.auditEntity || "No entity"}
          </button>
          <span className={`pill ${status.className}`}>{status.label}</span>
        </div>
        <div className="calendar-meta">
          <span>{project.assignmentType}</span>
          <span>{project.auditType}</span>
          <span>{formatAuditTeam(project) || "No audit team"}</span>
          {project.auditDurationHours ? <span>{project.auditDurationHours}h</span> : null}
          {project.auditStartTime ? <span>{project.auditStartTime}</span> : null}
        </div>
        {warnings.length > 0 && (
          <ul className="calendar-warnings">
            {warnings.slice(0, 3).map((warning) => (
              <li key={warning}>{warning}</li>
            ))}
          </ul>
        )}
        <div className="schedule-card-actions">
          <button type="button" disabled={!canUpdate} onClick={() => setIsScheduling(true)}>
            {project.confirmedAuditDate || project.tentativeAuditWeek ? "Edit schedule" : "Schedule audit"}
          </button>
          {project.calendarEventWebLink && (
            <a className="secondary outlook-event-button" href={project.calendarEventWebLink} target="_blank" rel="noreferrer">
              Open invite
            </a>
          )}
        </div>
      </div>
      {isScheduling && (
        <ScheduleAuditModal
          project={project}
          allProjects={allProjects}
          calendarReady={calendarReady}
          onClose={() => setIsScheduling(false)}
          onUpdateScheduling={onUpdateScheduling}
          onSyncOutlook={onSyncOutlook}
        />
      )}
    </article>
  );
}

function ScheduleAuditModal({
  project,
  allProjects,
  calendarReady,
  onClose,
  onUpdateScheduling,
  onSyncOutlook,
}: {
  project: AuditProject;
  allProjects: AuditProject[];
  calendarReady: boolean;
  onClose: () => void;
  onUpdateScheduling: (
    project: AuditProject,
    update: Partial<
      Pick<
        AuditProject,
        | "calendarSyncStatus"
        | "auditLocation"
        | "auditRemoteLink"
        | "auditDurationHours"
        | "auditStartTime"
        | "schedulingNotes"
        | "confirmedAuditDate"
        | "tentativeAuditWeek"
      >
    >,
  ) => void;
  onSyncOutlook: (project: AuditProject) => void;
}) {
  const [notes, setNotes] = useState(project.schedulingNotes);
  const [confirmedDate, setConfirmedDate] = useState(project.confirmedAuditDate);
  const [tentativeWeek, setTentativeWeek] = useState(project.tentativeAuditWeek);
  const [auditStartTime, setAuditStartTime] = useState(project.auditStartTime || "09:00");
  const [auditLocation, setAuditLocation] = useState(project.auditLocation);
  const [auditRemoteLink, setAuditRemoteLink] = useState(project.auditRemoteLink);
  const [auditDurationHours, setAuditDurationHours] = useState(String(project.auditDurationHours || 1));
  const previewProject = withProjectDefaults({
    ...project,
    confirmedAuditDate: confirmedDate,
    tentativeAuditWeek: tentativeWeek,
    auditLocation,
    auditRemoteLink,
    auditDurationHours: Number(auditDurationHours || 1),
    auditStartTime,
    schedulingNotes: notes,
  });
  const warnings = schedulingConflictWarnings(previewProject, allProjects);
  const scheduleChanged =
    confirmedDate !== project.confirmedAuditDate ||
    tentativeWeek !== project.tentativeAuditWeek ||
    auditLocation !== project.auditLocation ||
    auditRemoteLink !== project.auditRemoteLink ||
    Number(auditDurationHours || 1) !== project.auditDurationHours ||
    auditStartTime !== (project.auditStartTime || "09:00") ||
    notes !== project.schedulingNotes;
  const update = {
    confirmedAuditDate: confirmedDate,
    tentativeAuditWeek: tentativeWeek,
    auditLocation,
    auditRemoteLink,
    auditDurationHours: Number(auditDurationHours || 1),
    auditStartTime,
    schedulingNotes: notes,
    calendarSyncStatus: scheduleChanged
      ? nextCalendarStatusForScheduleChange(project)
      : project.calendarSyncStatus,
  };
  const saveSchedule = () => {
    if (scheduleChanged) {
      onUpdateScheduling(project, update);
    }
    onClose();
  };
  const saveAndSendInvite = () => {
    const nextProject = withProjectDefaults({
      ...previewProject,
      calendarSyncStatus: update.calendarSyncStatus,
    });
    if (scheduleChanged) {
      onUpdateScheduling(project, update);
    }
    onSyncOutlook(nextProject);
    onClose();
  };

  return (
    <div className="modal-backdrop">
      <section className="modal schedule-modal">
        <div className="section-title modal-title-row">
          <div>
            <p className="eyebrow dark">Schedule audit</p>
            <h2>{project.assignmentNumber} - {project.auditEntity || "No entity"}</h2>
            <span>Save the tracker schedule first, then optionally send or update the Outlook invite.</span>
          </div>
          <button type="button" className="link" onClick={onClose}>
            Close
          </button>
        </div>
        <div className="schedule-modal-grid">
          <Input
            label="Tentative week"
            value={tentativeWeek}
            placeholder="2026-W21"
            onChange={setTentativeWeek}
          />
          <Input
            label="Confirmed date"
            type="date"
            value={confirmedDate}
            onChange={setConfirmedDate}
          />
          <Input
            label="Start time"
            type="time"
            value={auditStartTime}
            onChange={setAuditStartTime}
          />
          <Input
            label="Duration hours"
            type="number"
            value={auditDurationHours}
            onChange={setAuditDurationHours}
          />
          <Input
            label="Location"
            value={auditLocation}
            placeholder="Office, city, client site, or remote"
            onChange={setAuditLocation}
          />
          <Input
            label="Remote link"
            value={auditRemoteLink}
            placeholder="Teams/Zoom link if known"
            onChange={setAuditRemoteLink}
          />
        </div>
        <label>
          Scheduling notes
          <textarea
            value={notes}
            placeholder="Availability, travel constraints, client scheduling notes, or calendar blockers"
            onChange={(event) => setNotes(event.target.value)}
          />
        </label>
        <div className="schedule-preview-panel">
          <strong>Outlook invite preview</strong>
          <div className="meta-grid">
            <Meta label="Subject" value={`Audit: ${project.assignmentNumber || project.id} - ${project.auditEntity || "Assignment"}`} />
            <Meta label="Date" value={confirmedDate || "Add confirmed date before sending"} />
            <Meta label="Time" value={`${auditStartTime || "09:00"} for ${Number(auditDurationHours || 1)}h`} />
            <Meta label="Attendees" value={formatAuditTeam(project) || "No audit team assigned"} />
            <Meta label="Location" value={auditLocation || auditRemoteLink || "Not set"} />
          </div>
        </div>
        {warnings.length > 0 && (
          <div className="schedule-warning-panel">
            <strong>Review before sending</strong>
            <ul>
              {warnings.map((warning) => (
                <li key={warning}>{warning}</li>
              ))}
            </ul>
          </div>
        )}
        <div className="modal-actions sticky-actions">
          <button type="button" className="secondary" onClick={onClose}>
            Cancel
          </button>
          <button type="button" className="secondary" onClick={saveSchedule}>
            Save schedule
          </button>
          <button
            type="button"
            disabled={!calendarReady || !confirmedDate}
            onClick={saveAndSendInvite}
            title={
              !calendarReady
                ? "Outlook invite sending needs Calendars.ReadWrite permission."
                : !confirmedDate
                  ? "Add a confirmed audit date first."
                  : "Save the schedule and send or update the Outlook invite."
            }
          >
            {project.calendarEventId ? "Save and update invite" : "Save and send Outlook invite"}
          </button>
        </div>
      </section>
    </div>
  );
}

function OperationsCommandCenter({
  projects,
  currentUser,
  onSelect,
  onExportReport,
}: {
  projects: AuditProject[];
  currentUser: PrototypeUser;
  onSelect: (id: string) => void;
  onExportReport: () => void;
}) {
  const [copyMessage, setCopyMessage] = useState("");
  const openProjects = projects.filter((project) => project.currentStage !== "Closed");
  const criticalProjects = openProjects.filter((project) =>
    slaSignals(project).some((signal) => signal.level === "Critical"),
  );
  const draftQueue = operationsDraftQueue(openProjects);
  const brief = operationsBrief(projects, currentUser);
  const coordinatorInsights = auditCoordinatorInsights(openProjects, currentUser);
  const documentPackages = documentIntelligenceSummary(openProjects);
  const roleCards = [
    {
      label: "Manager console",
      value: criticalProjects.length,
      detail: "Critical SLA or blocker items needing review.",
    },
    {
      label: "Auditor console",
      value: openProjects.filter((project) => projectHasAuditor(project, currentUser.fullName)).length,
      detail: "Open assignments where you are lead or support.",
    },
    {
      label: "Finance console",
      value: openProjects.filter(
        (project) =>
          stages.indexOf(project.currentStage) >= stages.indexOf("Invoice") &&
          project.invoiceStatus !== "Paid",
      ).length,
      detail: "Invoice-stage assignments still not paid.",
    },
    {
      label: "Draft queue",
      value: draftQueue.length,
      detail: "Email or document drafts ready for manual review.",
    },
  ];

  const copyPrompt = async () => {
    try {
      await navigator.clipboard.writeText(brief.aiPrompt);
      setCopyMessage("Assistant brief copied.");
    } catch {
      setCopyMessage("Copy failed. Select the brief text manually.");
    }
  };

  return (
    <section className="panel operations-command">
      <div className="section-title">
        <div>
          <p className="eyebrow dark">Operating system</p>
          <h2>Workflow, SLA, automation, and reporting control</h2>
          <span>{brief.summary}</span>
        </div>
        <div className="storage-actions">
          <button type="button" className="secondary" onClick={copyPrompt}>
            Copy assistant brief
          </button>
          <button type="button" onClick={onExportReport}>
            Export ops report
          </button>
        </div>
      </div>
      {copyMessage && <p className="muted-note">{copyMessage}</p>}
      <div className="ops-role-grid">
        {roleCards.map((card) => (
          <article key={card.label}>
            <span>{card.label}</span>
            <strong>{card.value}</strong>
            <p>{card.detail}</p>
          </article>
        ))}
      </div>
      <AiAuditCoordinatorPanel insights={coordinatorInsights} onSelect={onSelect} />
      <DocumentIntelligenceOverview packages={documentPackages} onSelect={onSelect} />
      <div className="ops-grid">
        <article>
          <h3>SLA escalation</h3>
          {criticalProjects.length === 0 ? (
            <p className="muted-note">No critical visible assignments.</p>
          ) : (
            <div className="queue-list">
              {criticalProjects.slice(0, 5).map((project) => (
                <button
                  type="button"
                  className="queue-item"
                  key={project.id}
                  onClick={() => onSelect(project.id)}
                >
                  <strong>{project.assignmentNumber}</strong>
                  <span>{project.auditEntity || project.clientCoverholderCode}</span>
                  <small>{slaSignals(project).map((signal) => signal.label).join(", ")}</small>
                </button>
              ))}
            </div>
          )}
        </article>
        <article>
          <h3>Email and document queue</h3>
          {draftQueue.length === 0 ? (
            <p className="muted-note">No draft actions needed for visible assignments.</p>
          ) : (
            <div className="draft-list">
              {draftQueue.slice(0, 6).map((draft) => (
                <button
                  type="button"
                  className="draft-item"
                  key={draft.id}
                  onClick={() => onSelect(draft.projectId)}
                >
                  <strong>{draft.label}</strong>
                  <span>{draft.assignmentNumber} | {draft.priority}</span>
                  <small>{draft.reason}</small>
                </button>
              ))}
            </div>
          )}
        </article>
        <article>
          <h3>Assistant brief</h3>
          <p>{brief.summary}</p>
          <ul className="compact-list">
            {brief.managerFocus.map((item) => (
              <li key={item}>{item}</li>
            ))}
          </ul>
        </article>
      </div>
    </section>
  );
}

function AiAuditCoordinatorPanel({
  insights,
  onSelect,
}: {
  insights: CoordinatorInsight[];
  onSelect: (id: string) => void;
}) {
  const topInsights = insights.slice(0, 8);
  return (
    <article className="ai-coordinator-panel">
      <div className="section-title">
        <div>
          <p className="eyebrow dark">AI audit coordinator</p>
          <h2>Recommended operating actions</h2>
          <span>
            Ranked from project status, blockers, documents, quote stage,
            scheduling, and finance signals. External AI can be added later;
            this version is deterministic and reviewable.
          </span>
        </div>
      </div>
      {topInsights.length === 0 ? (
        <p className="muted-note">No coordinator actions detected for visible projects.</p>
      ) : (
        <div className="coordinator-grid">
          {topInsights.map((insight) => (
            <button
              type="button"
              key={insight.id}
              className={`coordinator-card ${insight.priority.toLowerCase().replace(" ", "-")}`}
              onClick={() => onSelect(insight.projectId)}
            >
              <span>{insight.priority}</span>
              <strong>{insight.title}</strong>
              <small>
                {insight.assignmentNumber} | {insight.auditEntity}
              </small>
              <p>{insight.reason}</p>
              <em>{insight.recommendedAction}</em>
            </button>
          ))}
        </div>
      )}
    </article>
  );
}

function DocumentIntelligenceOverview({
  packages,
  onSelect,
}: {
  packages: DocumentIntelligenceResult[];
  onSelect: (id: string) => void;
}) {
  const needsReview = packages.filter((item) => item.confidence !== "High");
  return (
    <article className="document-intelligence-panel">
      <div className="section-title">
        <div>
          <p className="eyebrow dark">Document intelligence</p>
          <h2>Evidence package review</h2>
          <span>
            Flags missing evidence, stale chases, DCA package requirements, and
            coordinated managing-agent gaps before audits move forward.
          </span>
        </div>
        <span className="readiness-score warning">{needsReview.length}</span>
      </div>
      {packages.length === 0 ? (
        <p className="muted-note">No open document packages to review.</p>
      ) : (
        <div className="document-intelligence-grid">
          {packages.slice(0, 6).map((item) => (
            <button
              type="button"
              key={item.projectId}
              className={`document-intelligence-card ${item.confidence.toLowerCase().replace(" ", "-")}`}
              onClick={() => onSelect(item.projectId)}
            >
              <span>{item.packageType}</span>
              <strong>{item.assignmentNumber}</strong>
              <small>{item.auditEntity}</small>
              <div className="readiness-track compact" aria-hidden="true">
                <span style={{ width: `${item.readiness}%` }} />
              </div>
              <p>{item.readiness}% ready | {item.confidence}</p>
              <em>
                {item.missing.length
                  ? `Missing: ${item.missing.join(", ")}`
                  : "No required evidence gaps detected."}
              </em>
            </button>
          ))}
        </div>
      )}
    </article>
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
  disabled = false,
}: {
  label: string;
  value: string;
  options: (string | [string, string])[];
  onChange: (value: string) => void;
  placeholder?: string;
  disabled?: boolean;
}) {
  return (
    <label>
      {label}
      <select
        value={value}
        disabled={disabled}
        onChange={(event) => onChange(event.target.value)}
      >
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
                const workstreamSummary = coordinatedWorkstreamSummary(project);
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
                    <span className="pill muted">
                      {project.auditStructure}
                      {project.auditStructure === "Coordinated"
                        ? ` | ${workstreamSummary.total} managing agents`
                        : ""}
                    </span>
                    {project.auditStructure === "Coordinated" && (
                      <span
                        className={`pill ${
                          workstreamSummary.needsAttention > 0 ? "warning" : "ok"
                        }`}
                      >
                        {workstreamSummary.needsAttention} need attention
                      </span>
                    )}
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
  onArchive,
  onRemoveLabel,
  onAddComment,
  onToggleChecklist,
  onDocumentWorkflowAction,
  auditors,
  onAddSupportingAuditor,
  currentUser,
  onUpdateFinance,
  contactSources,
}: {
  project: AuditProject;
  onEdit: () => void;
  onMove: (project: AuditProject, stage: Stage) => void;
  onArchive: (project: AuditProject) => void;
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
  contactSources: LinkedContactSourcesResponse | null;
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
          <div className="detail-actions">
            {canEdit && <button onClick={onEdit}>Edit project</button>}
            {canEdit && project.currentStage === "Closed" && !project.archived && (
              <button
                type="button"
                className="secondary"
                onClick={() => onArchive(project)}
              >
                Archive project
              </button>
            )}
          </div>
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
          <Meta label="Audit structure" value={project.auditStructure} />
          <Meta label="Audit entity" value={project.auditEntity || "Not set"} />
          <Meta
            label="Client / coverholder code"
            value={project.clientCoverholderCode}
          />
          <Meta label="Broker" value={project.broker} />
          <Meta label="Audit team" value={formatAuditTeam(project)} />
          <Meta label="Status" value={project.assignmentStatus} />
          <Meta
            label="Quote"
            value={`${project.quoteStatus} · ${project.quoteAmount.toLocaleString("en-US", { style: "currency", currency: "USD", maximumFractionDigits: 0 })}`}
          />
          <Meta
            label="Audit timing"
            value={`${project.tentativeAuditWeek || "No week"} · ${project.confirmedAuditDate || "No date"}`}
          />
          <Meta label="Schedule" value={scheduleStatus(project).label} />
          <Meta
            label="Linked contact"
            value={project.linkedContactSource || "Not linked"}
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
        {canEdit && (
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
        )}
      </article>
      <ManagingAgentWorkstreamsPanel project={project} />
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
      <ProjectDocumentIntelligence project={project} />
      <WorkflowEnginePanel project={project} />
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
      <TemplateLibrary project={project} contactSources={contactSources} />
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

function ManagingAgentWorkstreamsPanel({
  project,
}: {
  project: AuditProject;
}) {
  const workstreams = normalizeManagingAgentWorkstreams(project);
  const summary = coordinatedWorkstreamSummary(project);
  return (
    <article className="panel managing-agent-panel">
      <div className="section-title">
        <div>
          <h2>Managing agent workstreams</h2>
          <span>
            {project.auditStructure === "Coordinated"
              ? `${summary.total} managing agents tracked under this one audit card.`
              : "Solo audit workstream derived from the parent assignment."}
          </span>
        </div>
      </div>
      <div className="workstream-summary-grid">
        <Meta label="Active" value={String(summary.active)} />
        <Meta label="Complete" value={String(summary.complete)} />
        <Meta label="Waived" value={String(summary.waived)} />
        <Meta label="Need attention" value={String(summary.needsAttention)} />
      </div>
      <div className="workstream-list">
        {workstreams.map((workstream) => {
          const missing = getMissingDocumentsForWorkstream(project, workstream);
          const resolved = workstream.completed || workstream.waived;
          return (
            <div className="workstream-row" key={workstream.id}>
              <div>
                <strong>{workstream.managingAgentName}</strong>
                <span>
                  {workstream.managingAgentCode || "No code"} |{" "}
                  {workstream.leadAuditor || "No lead assigned"}
                </span>
              </div>
              <div>
                <span className={`pill ${resolved ? "ok" : "muted"}`}>
                  {workstream.waived
                    ? "Waived"
                    : workstream.completed
                      ? "Complete"
                      : workstream.assignmentStatus}
                </span>
                <span className="pill muted">{workstream.currentStage}</span>
                {missing.length > 0 && (
                  <span className="pill warning">
                    Missing {missing.length}: {missing.join(", ")}
                  </span>
                )}
                {workstream.blockers.trim() && (
                  <span className="pill danger">{workstream.blockers.trim()}</span>
                )}
              </div>
            </div>
          );
        })}
      </div>
    </article>
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

function ProjectDocumentIntelligence({ project }: { project: AuditProject }) {
  const intelligence = documentIntelligence(project);
  return (
    <article className="panel project-document-intelligence">
      <div className="section-title">
        <div>
          <p className="eyebrow dark">Document review</p>
          <h2>{intelligence.packageType}</h2>
          <span>
            Confidence: {intelligence.confidence} | {intelligence.readiness}% ready
          </span>
        </div>
      </div>
      <div className="document-intelligence-columns">
        <div>
          <h3>Received</h3>
          {intelligence.evidence.length === 0 ? (
            <p className="muted-note">No required documents have been marked received.</p>
          ) : (
            <ul className="compact-list">
              {intelligence.evidence.map((item) => (
                <li key={item}>{item}</li>
              ))}
            </ul>
          )}
        </div>
        <div>
          <h3>Still needed</h3>
          {intelligence.missing.length === 0 ? (
            <p className="muted-note">No required document gaps detected.</p>
          ) : (
            <ul className="compact-list">
              {intelligence.missing.map((item) => (
                <li key={item}>{item}</li>
              ))}
            </ul>
          )}
        </div>
        <div>
          <h3>Recommended action</h3>
          <ul className="compact-list">
            {intelligence.recommendations.map((item) => (
              <li key={item}>{item}</li>
            ))}
          </ul>
        </div>
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
        {requiredDocumentsForProject(project).map((doc) => {
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

function WorkflowEnginePanel({ project }: { project: AuditProject }) {
  const [copyMessage, setCopyMessage] = useState("");
  const gates = workflowGates(project);
  const signals = slaSignals(project);
  const drafts = recommendedDrafts(project);
  const folders = workspaceFolders(project);
  const target = nextStage(project);

  const copyWorkspacePlan = async () => {
    try {
      await navigator.clipboard.writeText(folders.join("\n"));
      setCopyMessage("Workspace folder plan copied.");
    } catch {
      setCopyMessage("Copy failed. Select the folder list manually.");
    }
  };

  const copyFirstDraft = async () => {
    const firstDraft = drafts[0];
    if (!firstDraft) return;
    try {
      await navigator.clipboard.writeText(
        `Subject: ${firstDraft.subject}\n\n${firstDraft.body}`,
      );
      setCopyMessage(`${firstDraft.label} copied.`);
    } catch {
      setCopyMessage("Copy failed. Select the draft manually.");
    }
  };

  return (
    <article className="panel workflow-engine-panel">
      <div className="section-title">
        <div>
          <p className="eyebrow dark">Next steps</p>
          <h2>What needs attention</h2>
          <span>
            {target
              ? `Next goal: move this audit toward ${target}.`
              : "This audit is at the end of the workflow."}
          </span>
        </div>
      </div>
      <div className="workflow-gates">
        {gates.map((gate) => (
          <div className={`workflow-gate ${gate.status.toLowerCase()}`} key={gate.label}>
            <span>{gate.status}</span>
            <strong>{gate.label}</strong>
            <p>{gate.detail}</p>
          </div>
        ))}
      </div>
      <div className="workflow-columns">
        <div>
          <h3>Timing alerts</h3>
          <ul className="compact-list">
            {signals.map((signal) => (
              <li key={`${signal.label}-${signal.detail}`}>
                <strong>{signal.label}:</strong> {signal.detail}
              </li>
            ))}
          </ul>
        </div>
        <div>
          <h3>Suggested folders</h3>
          <ul className="compact-list">
            {folders.slice(0, 5).map((folder) => (
              <li key={folder}>{folder}</li>
            ))}
          </ul>
          <button type="button" className="secondary" onClick={copyWorkspacePlan}>
            Copy folder plan
          </button>
        </div>
        <div>
          <h3>Suggested emails</h3>
          {drafts.length === 0 ? (
            <p className="muted-note">No email draft recommended.</p>
          ) : (
            <ul className="compact-list">
              {drafts.slice(0, 4).map((draft) => (
                <li key={draft.id}>
                  <strong>{draft.label}:</strong> {draft.reason}
                </li>
              ))}
            </ul>
          )}
          <button
            type="button"
            disabled={drafts.length === 0}
            onClick={copyFirstDraft}
          >
            Copy first email
          </button>
        </div>
      </div>
      {copyMessage && <p className="muted-note">{copyMessage}</p>}
    </article>
  );
}

function TemplateLibrary({
  project,
  contactSources,
}: {
  project: AuditProject;
  contactSources: LinkedContactSourcesResponse | null;
}) {
  const [selectedTemplateId, setSelectedTemplateId] = useState(
    communicationTemplates[0].id,
  );
  const [copyMessage, setCopyMessage] = useState("");
  const template =
    communicationTemplates.find((item) => item.id === selectedTemplateId) ??
    communicationTemplates[0];
  const subject = template.subject(project);
  const body = template.body(project);
  const receiver = resolveTemplateReceiver(
    project,
    template,
    contactSources?.contacts ?? [],
  );
  const outlookBody = `${body}\n\n[Paste above your Outlook signature]`;
  const fullDraft = [
    `To: ${receiver.email || `[${receiver.kind} email needed]`}`,
    `Receiver: ${receiver.name} (${receiver.kind})`,
    `Subject: ${subject}`,
    "",
    body,
    "",
    "Signature: paste this body above your normal Outlook signature before sending.",
  ].join("\n");

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
      <div className="template-context-grid">
        <article className={`receiver-card ${receiver.confidence === "Matched workbook" ? "ready" : "blocked"}`}>
          <span>Receiver</span>
          <strong>{receiver.name}</strong>
          <p>{receiver.email || "No email found yet"}</p>
          <small>
            {receiver.kind} | {receiver.confidence}
          </small>
          <small>{receiver.source}</small>
        </article>
        <article className="receiver-card">
          <span>Routing rule</span>
          <strong>{project.assignmentType} audit</strong>
          <p>{receiver.guidance}</p>
        </article>
        <article className="receiver-card">
          <span>Signature</span>
          <strong>Use Outlook signature</strong>
          <p>Copy the body into Outlook above your existing signature. The tracker does not replace your signature.</p>
        </article>
      </div>
      <div className="template-preview">
        <label>
          To
          <input readOnly value={receiver.email || `${receiver.kind} email not found`} />
        </label>
        <label>
          Subject
          <input readOnly value={subject} />
        </label>
        <label>
          Body
          <textarea readOnly rows={10} value={outlookBody} />
        </label>
      </div>
      <div className="template-actions">
        <button
          type="button"
          className="secondary"
          disabled={!receiver.email}
          onClick={() => void copyText("Receiver", receiver.email)}
        >
          Copy receiver
        </button>
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
        <button
          type="button"
          className="secondary"
          onClick={() => void copyText("Full draft", fullDraft)}
        >
          Copy full draft
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
        {requiredDocumentsForProject(project).map((doc) => (
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

function ManagingAgentWorkstreamEditor({
  project,
  auditorOptions,
  onChange,
}: {
  project: AuditProject;
  auditorOptions: string[];
  onChange: (workstreams: ManagingAgentWorkstream[]) => void;
}) {
  const workstreams = normalizeManagingAgentWorkstreams(project);
  const projectRequiredDocuments = requiredDocumentsForProject(project);
  const updateWorkstream = (
    id: string,
    patch: Partial<ManagingAgentWorkstream>,
  ) => {
    onChange(
      workstreams.map((workstream) =>
        workstream.id === id ? { ...workstream, ...patch } : workstream,
      ),
    );
  };
  const addWorkstream = () => {
    onChange([
      ...workstreams,
      {
        ...defaultWorkstreamFromProject(project),
        id: `${project.id}-ma-${Date.now()}`,
        managingAgentName: `Managing agent ${workstreams.length + 1}`,
        managingAgentCode: "",
        blockers: "",
        nextAction: "",
        completed: false,
        waived: false,
      },
    ]);
  };
  const removeWorkstream = (id: string) => {
    if (workstreams.length <= 1) return;
    onChange(workstreams.filter((workstream) => workstream.id !== id));
  };

  return (
    <div className="workstream-editor">
      <div className="inline-heading">
        <div>
          <strong>Managing agent workstreams</strong>
          <span>
            {project.assignmentType === "DCA"
              ? "DCA audits require a managing agent, DCA Agreement, and Claims BDX."
              : "Use this when one audit is coordinated across multiple managing agents."}
          </span>
        </div>
        <button
          type="button"
          className="secondary"
          disabled={project.auditStructure !== "Coordinated"}
          onClick={addWorkstream}
        >
          Add managing agent
        </button>
      </div>
      {workstreams.map((workstream, index) => (
        <div className="workstream-editor-row" key={workstream.id}>
          <div className="workstream-row-heading">
            <strong>
              {project.auditStructure === "Coordinated"
                ? `Managing agent ${index + 1}`
                : "Managing agent"}
            </strong>
            <button
              type="button"
              className="text-button"
              disabled={workstreams.length <= 1}
              onClick={() => removeWorkstream(workstream.id)}
            >
              Remove
            </button>
          </div>
          <div className="form-grid wizard-grid compact-form-grid">
            <Input
              label="Managing agent"
              value={workstream.managingAgentName}
              onChange={(value) =>
                updateWorkstream(workstream.id, { managingAgentName: value })
              }
            />
            <Input
              label="MA / client code"
              value={workstream.managingAgentCode}
              onChange={(value) =>
                updateWorkstream(workstream.id, { managingAgentCode: value })
              }
            />
            <Select
              label="Lead auditor"
              value={workstream.leadAuditor}
              options={auditorOptions}
              placeholder="Select lead auditor"
              onChange={(value) =>
                updateWorkstream(workstream.id, { leadAuditor: value })
              }
            />
            <Input
              label="Due date"
              type="date"
              value={workstream.dueDate}
              onChange={(value) =>
                updateWorkstream(workstream.id, { dueDate: value })
              }
            />
            <Select
              label="Stage"
              value={workstream.currentStage}
              options={stages}
              placeholder="Select stage"
              onChange={(value) =>
                updateWorkstream(workstream.id, { currentStage: value as Stage })
              }
            />
            <Select
              label="Status"
              value={workstream.assignmentStatus}
              options={assignmentStatusOptions}
              placeholder="Select status"
              onChange={(value) =>
                updateWorkstream(workstream.id, {
                  assignmentStatus: value as AssignmentStatus,
                })
              }
            />
            {projectRequiredDocuments.map((doc) => (
              <Check
                key={doc.key}
                label={doc.label}
                checked={Boolean(workstream[doc.key])}
                onChange={(value) =>
                  updateWorkstream(workstream.id, { [doc.key]: value })
                }
              />
            ))}
            <Check
              label="Workstream complete"
              checked={workstream.completed}
              onChange={(value) =>
                updateWorkstream(workstream.id, { completed: value })
              }
            />
            <Check
              label="Waived / not applicable"
              checked={workstream.waived}
              onChange={(value) =>
                updateWorkstream(workstream.id, { waived: value })
              }
            />
          </div>
          <label>
            Workstream blockers
            <textarea
              value={workstream.blockers}
              onChange={(event) =>
                updateWorkstream(workstream.id, {
                  blockers: event.target.value,
                })
              }
            />
          </label>
        </div>
      ))}
    </div>
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
  contactSources,
  existingProjects,
}: {
  project: AuditProject;
  onSave: (project: AuditProject) => void;
  onCancel: () => void;
  auditorOptions: string[];
  contactSources: LinkedContactSourcesResponse | null;
  existingProjects: AuditProject[];
}) {
  const [draft, setDraft] = useState(withProjectDefaults(project));
  const [step, setStep] = useState(0);
  const [attemptedSave, setAttemptedSave] = useState(false);
  const isNewProject = project.statusHistory.length === 0;
  const steps = [
    "Client / contact",
    "People",
    "Planning",
    "Documents & quote",
  ];
  const update = <K extends keyof AuditProject>(
    key: K,
    value: AuditProject[K],
  ) => setDraft({ ...draft, [key]: value });
  const updateAssignmentType = (value: string) => {
    setDraft(
      withProjectDefaults({
        ...draft,
        assignmentType: value as AssignmentType,
      }),
    );
  };
  const updateAuditStructure = (value: string) => {
    setDraft(
      withProjectDefaults({
        ...draft,
        auditStructure: value as AuditStructure,
      }),
    );
  };
  const linkedContacts = (contactSources?.contacts ?? [])
    .slice()
    .sort((a, b) => linkedContactName(a).localeCompare(linkedContactName(b)));
  const selectedLinkedContact = linkedContacts.find(
    (contact) => contact.id === draft.linkedContactId,
  );
  const selectLinkedContact = (contactId: string) => {
    const contact = linkedContacts.find((item) => item.id === contactId);
    if (!contact) {
      setDraft(withProjectDefaults({ ...draft, linkedContactId: "", linkedContactSource: "" }));
      return;
    }
    setDraft(projectWithLinkedContact(draft, contact));
  };
  const updateDocumentField = (key: ProjectDocumentKey, value: boolean) => {
    setDraft(
      withProjectDefaults({
        ...draft,
        [key]: value,
        managingAgentWorkstreams: normalizeManagingAgentWorkstreams(draft).map(
          (workstream) => ({ ...workstream, [key]: value }),
        ),
      }),
    );
  };
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
  const requiredIssues = intakeRequiredIssues(draft);
  const warningIssues = [
    ...intakeWarningIssues(draft),
    ...duplicateProjectWarnings(draft, existingProjects),
  ];
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
    setAttemptedSave(true);
    if (requiredIssues.length > 0) return;
    onSave(withProjectDefaults(draft));
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
      {draft.assignmentType === "DCA" && (
        <div className="guidance-strip">
          <strong>DCA setup</strong>
          <span>
            DCA audits use a DCA Agreement and Claims BDX. Use the managing
            agent or DCA contact from the linked workbook as the primary
            recipient, not the broker contact path.
          </span>
        </div>
      )}
      <div className="linked-contact-intake">
        <Select
          label="Linked client workbook contact"
          value={draft.linkedContactId}
          options={linkedContacts.map((contact) => [
            contact.id,
            linkedContactLabel(contact),
          ])}
          disabled={linkedContacts.length === 0}
          placeholder={
            contactSources
              ? linkedContacts.length > 0
                ? "Select client/contact from linked spreadsheets"
                : "No linked contacts loaded yet"
              : "Contacts loading or unavailable"
          }
          onChange={selectLinkedContact}
        />
        {selectedLinkedContact ? (
          <div className="linked-contact-preview">
            <strong>{linkedContactName(selectedLinkedContact)}</strong>
            <span>{selectedLinkedContact.email || "No email detected"}</span>
            <small>{draft.linkedContactSource}</small>
            <div className="contact-tags">
              {selectedLinkedContact.raw?.dcaContact && <span>DCA contact available</span>}
              {selectedLinkedContact.raw?.coverholderContact && <span>CH contact available</span>}
              {instructionValue(selectedLinkedContact, "Onsite/Remote Preference") && (
                <span>{instructionValue(selectedLinkedContact, "Onsite/Remote Preference")}</span>
              )}
            </div>
          </div>
        ) : (
          <p className="muted-note">
            Linking a workbook contact can prefill entity, contact source,
            scheduling notes, and onsite/remote preference from the live
            spreadsheet.
          </p>
        )}
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
          onChange={updateAssignmentType}
        />
        <Select
          label="Audit structure"
          value={draft.auditStructure}
          options={auditStructureOptions}
          placeholder="Select structure"
          onChange={updateAuditStructure}
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
      {(draft.auditStructure === "Coordinated" || draft.assignmentType === "DCA") && (
        <ManagingAgentWorkstreamEditor
          project={draft}
          auditorOptions={auditorOptions}
          onChange={(workstreams) =>
            update("managingAgentWorkstreams", workstreams)
          }
        />
      )}
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
        <Input
          label="Start time"
          type="time"
          value={draft.auditStartTime || "09:00"}
          onChange={(value) => update("auditStartTime", value)}
        />
        <Input
          label="Duration hours"
          type="number"
          value={String(draft.auditDurationHours || 1)}
          onChange={(value) => update("auditDurationHours", Number(value) || 1)}
        />
        <Input
          label="Location"
          value={draft.auditLocation}
          placeholder="Office, city, client site, or remote"
          onChange={(value) => update("auditLocation", value)}
        />
        <Input
          label="Remote link"
          value={draft.auditRemoteLink}
          placeholder="Teams/Zoom link if known"
          onChange={(value) => update("auditRemoteLink", value)}
        />
      </div>
      <label>
        Scheduling notes
        <textarea
          value={draft.schedulingNotes}
          placeholder="Availability, travel constraints, preferred dates, client scheduling notes, or calendar blockers"
          onChange={(event) => update("schedulingNotes", event.target.value)}
        />
      </label>
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
        {requiredDocumentsForProject(draft).map((doc) => (
          <Check
            key={doc.key}
            label={doc.label}
            checked={Boolean(draft[doc.key])}
            onChange={(value) => updateDocumentField(doc.key, value)}
          />
        ))}
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

  const createStepContent = [basics, people, planning, documentsQuote];

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
        {(attemptedSave || warningIssues.length > 0) && (
          <div className="intake-quality-panel">
            <strong>Intake quality check</strong>
            {requiredIssues.length > 0 && (
              <ul className="quality-errors">
                {requiredIssues.map((issue) => (
                  <li key={issue}>{issue}</li>
                ))}
              </ul>
            )}
            {warningIssues.length > 0 && (
              <ul className="quality-warnings">
                {warningIssues.map((issue) => (
                  <li key={issue}>{issue}</li>
                ))}
              </ul>
            )}
          </div>
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
          <button type="submit" disabled={attemptedSave && requiredIssues.length > 0}>
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
