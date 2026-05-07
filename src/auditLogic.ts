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
export type ProjectLabel =
  | "High Priority"
  | "Medium Priority"
  | "Low Priority"
  | "Waiting on Broker";
export type ProgressStatus =
  | "Not Started"
  | "In Progress"
  | "Complete"
  | "Not Required";
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
export type DocumentWorkflowAction =
  | "markWaitingOnBroker"
  | "recordBrokerChase"
  | "markDocumentsComplete";
export type AuditTeamRole = "Lead Auditor" | "Supporting Auditor";
export type AuditTeamMember = {
  person: string;
  role: AuditTeamRole;
};

export type LogicProject = {
  assignedAuditor: string;
  auditTeam: AuditTeamMember[];
  currentStage: Stage;
  assignmentStatus: AssignmentStatus;
  quoteStatus: QuoteStatus;
  baaReceived: boolean;
  endorsementsReceived: boolean;
  premiumBdxReceived: boolean;
  preAuditQuestionnaireStatus: ProgressStatus;
  documentRequestStatus: ProgressStatus;
  documentRequestDate: string;
  brokerLastChasedDate: string;
  brokerExpectedResponseDate: string;
  coverholderResponseReceivedDate: string;
  blockers: string;
  labels: ProjectLabel[];
  nextAction: string;
  lastUpdatedDate: string;
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

export const requiredDocuments = [
  { key: "baaReceived", label: "BAA received" },
  { key: "endorsementsReceived", label: "Endorsements received" },
  { key: "premiumBdxReceived", label: "Premium BDX received" },
] as const;

export function normalizeAuditTeam(project: LogicProject): AuditTeamMember[] {
  const seededTeam = (project.auditTeam ?? [])
    .filter((member) => member.person.trim())
    .map((member) => ({
      person: member.person.trim(),
      role:
        member.role === "Lead Auditor"
          ? ("Lead Auditor" as AuditTeamRole)
          : ("Supporting Auditor" as AuditTeamRole),
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

export function assignedAuditorNames(project: LogicProject) {
  return normalizeAuditTeam(project).map((member) => member.person);
}

export function projectHasAuditor(project: LogicProject, auditor: string) {
  return assignedAuditorNames(project).includes(auditor);
}

function addUniqueLabel(labels: ProjectLabel[], label: ProjectLabel) {
  return labels.includes(label) ? labels : [...labels, label];
}

export function getMissingDocuments(project: LogicProject) {
  return requiredDocuments
    .filter((doc) => !project[doc.key])
    .map((doc) => doc.label);
}

export function computedBlockers(project: LogicProject) {
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

export function canMoveToStage(project: LogicProject, targetStage: Stage) {
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

export function documentReadiness(project: LogicProject) {
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
  project: LogicProject,
  action: DocumentWorkflowAction,
  date: string,
): LogicProject {
  if (action === "markWaitingOnBroker") {
    return {
      ...project,
      labels: addUniqueLabel(project.labels, "Waiting on Broker"),
      assignmentStatus: project.currentStage === "Closed" ? project.assignmentStatus : "On Hold",
      documentRequestStatus:
        project.documentRequestStatus === "Not Required" ? "In Progress" : project.documentRequestStatus,
      documentRequestDate: project.documentRequestDate || date,
      brokerLastChasedDate: project.brokerLastChasedDate || date,
      nextAction: project.nextAction || "Follow up with broker on outstanding documents.",
      lastUpdatedDate: date,
    };
  }
  if (action === "recordBrokerChase") {
    return {
      ...project,
      labels: addUniqueLabel(project.labels, "Waiting on Broker"),
      assignmentStatus: project.currentStage === "Closed" ? project.assignmentStatus : "On Hold",
      documentRequestStatus:
        project.documentRequestStatus === "Not Required" ? "In Progress" : project.documentRequestStatus,
      documentRequestDate: project.documentRequestDate || date,
      brokerLastChasedDate: date,
      nextAction: `Broker chased on ${date}; await outstanding documents.`,
      lastUpdatedDate: date,
    };
  }
  return {
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
  };
}
