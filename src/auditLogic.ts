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
export type AssignmentType = "DCA" | "CH" | "MGA" | "Company Contract";
export type AuditStructure = "Solo" | "Coordinated";
export type QuoteStatus =
  | "Not Started"
  | "Drafting"
  | "Sent"
  | "Accepted"
  | "Rejected";
export type DocumentWorkflowAction =
  | "markWaitingOnBroker"
  | "recordBrokerChase"
  | "clearWaitingOnBroker"
  | "markDocumentsComplete";
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

export type StatusHistoryItem = {
  changedAt: string;
  fromStage: Stage;
  toStage: Stage;
};

export type DurationRange = "ytd" | "90d" | "7d";

export type LogicProject = {
  id?: string;
  assignmentType?: AssignmentType;
  auditStructure?: AuditStructure;
  managingAgentWorkstreams?: ManagingAgentWorkstream[];
  assignedAuditor: string;
  auditTeam: AuditTeamMember[];
  currentStage: Stage;
  assignmentStatus: AssignmentStatus;
  quoteStatus: QuoteStatus;
  baaReceived: boolean;
  endorsementsReceived: boolean;
  premiumBdxReceived: boolean;
  dcaAgreementReceived?: boolean;
  claimsBdxReceived?: boolean;
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
  dueDate?: string;
  statusHistory?: StatusHistoryItem[];
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

export const requiredDocuments: readonly RequiredDocument[] = [
  { key: "baaReceived", label: "BAA received" },
  { key: "endorsementsReceived", label: "Endorsements received" },
  { key: "premiumBdxReceived", label: "Premium BDX received" },
] as const;

export const dcaRequiredDocuments: readonly RequiredDocument[] = [
  { key: "dcaAgreementReceived", label: "DCA Agreement received" },
  { key: "claimsBdxReceived", label: "Claims BDX received" },
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

function isDcaProject(project: Pick<LogicProject, "assignmentType">) {
  return project.assignmentType === "DCA";
}

export function requiredDocumentsForProject(
  project: Pick<LogicProject, "assignmentType">,
) {
  return isDcaProject(project) ? dcaRequiredDocuments : requiredDocuments;
}

function defaultWorkstreamFromProject(project: LogicProject): ManagingAgentWorkstream {
  const team = normalizeAuditTeam(project);
  return {
    id: `${project.id || "audit"}-ma-1`,
    managingAgentName: "Primary workstream",
    managingAgentCode: "",
    leadAuditor: team.find((member) => member.role === "Lead Auditor")?.person ?? project.assignedAuditor,
    supportAuditors: team
      .filter((member) => member.role === "Supporting Auditor")
      .map((member) => member.person),
    currentStage: project.currentStage,
    assignmentStatus: project.assignmentStatus,
    dueDate: project.dueDate ?? "",
    documentRequestStatus: project.documentRequestStatus,
    baaReceived: project.baaReceived,
    endorsementsReceived: project.endorsementsReceived,
    premiumBdxReceived: project.premiumBdxReceived,
    dcaAgreementReceived: project.dcaAgreementReceived ?? false,
    claimsBdxReceived: project.claimsBdxReceived ?? false,
    blockers: project.blockers,
    nextAction: project.nextAction,
    completed:
      project.assignmentStatus === "Completed" || project.currentStage === "Closed",
    waived: false,
  };
}

export function normalizeManagingAgentWorkstreams(
  project: LogicProject,
): ManagingAgentWorkstream[] {
  const source =
    project.managingAgentWorkstreams?.length
      ? project.managingAgentWorkstreams
      : [defaultWorkstreamFromProject(project)];
  const fallback = defaultWorkstreamFromProject(project);
  return source.map((workstream, index) => ({
    ...fallback,
    ...workstream,
    id: workstream.id || `${project.id || "audit"}-ma-${index + 1}`,
    managingAgentName:
      workstream.managingAgentName?.trim() || `Managing agent ${index + 1}`,
    supportAuditors: workstream.supportAuditors ?? [],
    completed:
      workstream.completed ??
      (workstream.assignmentStatus === "Completed" ||
        workstream.currentStage === "Closed"),
    waived: workstream.waived ?? false,
  }));
}

function getMissingDocumentsForWorkstream(
  project: LogicProject,
  workstream: ManagingAgentWorkstream,
) {
  return requiredDocumentsForProject(project)
    .filter((doc) => !workstream[doc.key])
    .map((doc) => doc.label);
}

export function coordinatedWorkstreamSummary(project: LogicProject) {
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


function rangeStart(range: DurationRange, today: Date) {
  const start = new Date(today);
  if (range === "ytd") return new Date(`${today.getUTCFullYear()}-01-01T00:00:00Z`);
  start.setUTCDate(start.getUTCDate() - (range === "90d" ? 90 : 7));
  return start;
}

export function stageDurationMetrics(
  projects: LogicProject[],
  range: DurationRange,
  today = new Date("2026-05-05T12:00:00Z"),
) {
  const start = rangeStart(range, today);
  const durations = new Map<Stage, number[]>();
  stages.forEach((stage) => durations.set(stage, []));
  projects.forEach((project) => {
    const sorted = (project.statusHistory ?? [])
      .slice()
      .sort((a, b) => Date.parse(a.changedAt) - Date.parse(b.changedAt));
    sorted.forEach((item, index) => {
      const changedAt = new Date(item.changedAt);
      if (Number.isNaN(changedAt.getTime()) || changedAt < start) return;
      const prior = sorted[index - 1];
      const enteredAt = prior
        ? new Date(prior.changedAt)
        : new Date(`${project.lastUpdatedDate}T12:00:00Z`);
      const durationDays = Math.max(
        0,
        Math.round((changedAt.getTime() - enteredAt.getTime()) / 86400000),
      );
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

function addUniqueLabel(labels: ProjectLabel[], label: ProjectLabel) {
  return labels.includes(label) ? labels : [...labels, label];
}

export function getMissingDocuments(project: LogicProject) {
  return requiredDocumentsForProject(project)
    .filter((doc) => !project[doc.key])
    .map((doc) => doc.label);
}

export function computedBlockers(project: LogicProject) {
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
    project.assignmentType === "DCA" &&
    !project.claimsBdxReceived
  ) {
    blockers.push("Claims BDX required before file selection");
  }
  if (
    stages.indexOf(project.currentStage) >= stages.indexOf("File Selection") &&
    project.assignmentType !== "DCA" &&
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

export function daysUntil(
  dateValue: string | undefined,
  today = new Date("2026-05-05T12:00:00Z"),
) {
  if (!dateValue) return Number.POSITIVE_INFINITY;
  const due = new Date(`${dateValue}T12:00:00Z`);
  return Math.ceil((due.getTime() - today.getTime()) / 86400000);
}

function pushUnique(items: string[], item: string) {
  if (!items.includes(item)) items.push(item);
}

export function recommendedNextSteps(
  project: LogicProject,
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

export function documentReadiness(project: LogicProject) {
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
  if (action === "clearWaitingOnBroker") {
    return {
      ...project,
      labels: project.labels.filter((label) => label !== "Waiting on Broker"),
      assignmentStatus:
        project.currentStage === "Closed" ? project.assignmentStatus : "In Progress",
      nextAction:
        project.nextAction ||
        "Broker completed their action; review received support and continue readiness.",
      lastUpdatedDate: date,
    };
  }
  return {
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
  };
}
