import React, { FormEvent, useMemo, useState } from "react";
import { createRoot } from "react-dom/client";
import "./styles.css";

type AssignmentSource = "Email" | "DAM";
type AssignmentType = "DCA" | "CH" | "MGA" | "Company Contract";
type AuditType = "Remote" | "Onsite";
type Stage =
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
type AssignmentStatus =
  | "New"
  | "In Progress"
  | "Blocked"
  | "On Hold"
  | "Completed";
type QuoteStatus =
  | "Not Started"
  | "Drafting"
  | "Sent"
  | "Accepted"
  | "Rejected";
type ProgressStatus =
  | "Not Started"
  | "In Progress"
  | "Complete"
  | "Not Required";
type ReportStatus = "Not Started" | "Drafting" | "Review" | "Issued";
type InvoiceStatus = "Not Started" | "Prepared" | "Sent" | "Paid";
type DamSubmissionStatus =
  | "Not Required"
  | "Not Started"
  | "Submitted"
  | "Accepted";

type StatusHistoryItem = {
  id: string;
  changedAt: string;
  changedBy: string;
  fromStage: Stage;
  toStage: Stage;
  note: string;
};

type AuditProject = {
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
  statusHistory: StatusHistoryItem[];
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
  | "myAudits"
  | "blocked"
  | "dueThisWeek"
  | "awaitingDocuments";
type ViewMode = "kanban" | "table";

const stages: Stage[] = [
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

const sampleProjects: AuditProject[] = [
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
  statusHistory: [],
});

const requiredDocuments = [
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

function withProjectDefaults(project: AuditProject): AuditProject {
  return {
    ...project,
    assignmentType: project.assignmentType ?? "CH",
    auditEntity: project.auditEntity ?? "",
    paymentReceived:
      project.paymentReceived ?? project.invoiceStatus === "Paid",
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

function getMissingDocuments(project: AuditProject) {
  return requiredDocuments
    .filter((doc) => !project[doc.key])
    .map((doc) => doc.label);
}

function computedBlockers(project: AuditProject) {
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

function canMoveToStage(project: AuditProject, targetStage: Stage) {
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

function daysUntil(dateValue: string) {
  if (!dateValue) return Number.POSITIVE_INFINITY;
  const due = new Date(`${dateValue}T12:00:00Z`);
  return Math.ceil((due.getTime() - today.getTime()) / (1000 * 60 * 60 * 24));
}

function dueLabel(project: AuditProject) {
  const days = daysUntil(project.dueDate);
  if (!Number.isFinite(days))
    return { text: "No due date", className: "muted" };
  if (days < 0)
    return { text: `${Math.abs(days)}d overdue`, className: "danger" };
  if (days <= 3) return { text: `Due in ${days}d`, className: "warning" };
  return { text: `Due ${project.dueDate}`, className: "ok" };
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

function escapeCsv(value: string | number | boolean) {
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

  const auditors = [
    ...auditorOptions,
    ...projects
      .map((project) => project.assignedAuditor)
      .filter((auditor) => auditor && !auditorOptions.includes(auditor)),
  ];
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
          changedAt: new Date().toISOString().slice(0, 10),
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
      <WorkloadCounts projects={projects} auditors={auditors} />
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
}: {
  projects: AuditProject[];
  auditors: string[];
}) {
  return (
    <section className="panel">
      <div className="section-title">
        <h2>Auditor workload</h2>
        <span>Open assignments by auditor</span>
      </div>
      <div className="workload-grid">
        {auditors.map((auditor) => {
          const openCount = projects.filter(
            (project) =>
              project.assignedAuditor === auditor &&
              project.currentStage !== "Closed",
          ).length;
          const tone =
            openCount >= 4 ? "high" : openCount >= 2 ? "medium" : "low";
          return (
            <article className={`workload-card ${tone}`} key={auditor}>
              <span>{auditor}</span>
              <strong>{openCount}</strong>
              <small>{openCount === 1 ? "open audit" : "open audits"}</small>
            </article>
          );
        })}
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
  const [newAuditor, setNewAuditor] = useState("");
  const addAuditor = (event: FormEvent) => {
    event.preventDefault();
    const cleanName = newAuditor.trim();
    if (
      !cleanName ||
      auditorOptions.some(
        (auditor) => auditor.toLowerCase() === cleanName.toLowerCase(),
      )
    )
      return;
    setAuditorOptions([...auditorOptions, cleanName]);
    setNewAuditor("");
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
            Add auditor
            <input
              value={newAuditor}
              placeholder="Type auditor name"
              onChange={(event) => setNewAuditor(event.target.value)}
            />
          </label>
          <button type="submit">Add</button>
        </form>
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
}: {
  label: string;
  value: string;
  options: (string | [string, string])[];
  onChange: (value: string) => void;
}) {
  return (
    <label>
      {label}
      <select value={value} onChange={(event) => onChange(event.target.value)}>
        <option value="">All</option>
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
}: {
  project: AuditProject;
  onEdit: () => void;
  onMove: (project: AuditProject, stage: Stage) => void;
}) {
  const blockers = computedBlockers(project);
  return (
    <section className="detail-grid">
      <article className="panel detail">
        <div className="section-title">
          <h2>{project.assignmentNumber}</h2>
          <button onClick={onEdit}>Edit project</button>
        </div>
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
      <Checklist project={project} />
      <History project={project} />
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

function Checklist({ project }: { project: AuditProject }) {
  const sourceSpecific = sourceTasks(project);
  return (
    <article className="panel">
      <h2>Stage checklist</h2>
      <h3>{project.currentStage}</h3>
      <ul className="checklist">
        {checklistByStage[project.currentStage].map((item) => (
          <li key={item}>{item}</li>
        ))}
        {sourceSpecific.map((item) => (
          <li key={item} className="conditional">
            {item}
          </li>
        ))}
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
  const update = <K extends keyof AuditProject>(
    key: K,
    value: AuditProject[K],
  ) => setDraft({ ...draft, [key]: value });
  const submit = (event: FormEvent) => {
    event.preventDefault();
    onSave(draft);
  };
  return (
    <div className="modal-backdrop">
      <form className="modal" onSubmit={submit}>
        <div className="section-title">
          <h2>
            {project.statusHistory.length ? "Edit project" : "Add project"}
          </h2>
          <button type="button" className="link" onClick={onCancel}>
            Close
          </button>
        </div>
        <div className="form-grid">
          <Input
            label="Assignment number"
            value={draft.assignmentNumber}
            onChange={(value) => update("assignmentNumber", value)}
          />
          <Select
            label="Source"
            value={draft.assignmentSource}
            options={["Email", "DAM"]}
            onChange={(value) =>
              update("assignmentSource", value as AssignmentSource)
            }
          />
          <Select
            label="Assignment type"
            value={draft.assignmentType}
            options={assignmentTypeOptions}
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
          <Select
            label="Assigned auditor"
            value={draft.assignedAuditor}
            options={auditorOptions}
            onChange={(value) => update("assignedAuditor", value)}
          />
          <Input
            label="Reviewer"
            value={draft.reviewer}
            onChange={(value) => update("reviewer", value)}
          />
          <Select
            label="Current stage"
            value={draft.currentStage}
            options={stages}
            onChange={(value) => update("currentStage", value as Stage)}
          />
          <Select
            label="Assignment status"
            value={draft.assignmentStatus}
            options={["New", "In Progress", "Blocked", "On Hold", "Completed"]}
            onChange={(value) =>
              update("assignmentStatus", value as AssignmentStatus)
            }
          />
          <Select
            label="Quote status"
            value={draft.quoteStatus}
            options={[
              "Not Started",
              "Drafting",
              "Sent",
              "Accepted",
              "Rejected",
            ]}
            onChange={(value) => update("quoteStatus", value as QuoteStatus)}
          />
          <Input
            label="Quote amount"
            value={String(draft.quoteAmount)}
            type="number"
            onChange={(value) => update("quoteAmount", Number(value))}
          />
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
            onChange={(value) => update("auditType", value as AuditType)}
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
            onChange={(value) =>
              update("preAuditQuestionnaireStatus", value as ProgressStatus)
            }
          />
          <Select
            label="Document request"
            value={draft.documentRequestStatus}
            options={["Not Started", "In Progress", "Complete", "Not Required"]}
            onChange={(value) =>
              update("documentRequestStatus", value as ProgressStatus)
            }
          />
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
            onChange={(value) =>
              update("coverholderResponseReceivedDate", value)
            }
          />
          <Select
            label="Report status"
            value={draft.reportStatus}
            options={["Not Started", "Drafting", "Review", "Issued"]}
            onChange={(value) => update("reportStatus", value as ReportStatus)}
          />
          <Select
            label="Invoice status"
            value={draft.invoiceStatus}
            options={["Not Started", "Prepared", "Sent", "Paid"]}
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
              onChange={(value) =>
                update("damSubmissionStatus", value as DamSubmissionStatus)
              }
            />
          )}
          <Input
            label="Due date"
            type="date"
            value={draft.dueDate}
            onChange={(value) => update("dueDate", value)}
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
        <div className="modal-actions">
          <button type="button" className="secondary" onClick={onCancel}>
            Cancel
          </button>
          <button type="submit">Save project</button>
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

createRoot(document.getElementById("root")!).render(<App />);
