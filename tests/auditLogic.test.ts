import assert from "node:assert/strict";
import test from "node:test";
import {
  applyDocumentWorkflowAction,
  assignedAuditorNames,
  canMoveToStage,
  computedBlockers,
  documentReadiness,
  recommendedNextSteps,
  stageDurationMetrics,
  type LogicProject,
} from "../src/auditLogic.js";

const baseProject: LogicProject = {
  assignedAuditor: "Lorraine Mojica",
  auditTeam: [
    { person: "Lorraine Mojica", role: "Lead Auditor" },
    { person: "Leslie Domenech", role: "Supporting Auditor" },
  ],
  currentStage: "Quote",
  assignmentStatus: "New",
  quoteStatus: "Sent",
  baaReceived: true,
  endorsementsReceived: false,
  premiumBdxReceived: false,
  preAuditQuestionnaireStatus: "Not Started",
  documentRequestStatus: "In Progress",
  documentRequestDate: "",
  brokerLastChasedDate: "",
  brokerExpectedResponseDate: "",
  coverholderResponseReceivedDate: "",
  blockers: "",
  labels: [],
  nextAction: "",
  lastUpdatedDate: "2026-05-01",
  dueDate: "2026-05-07",
  statusHistory: [
    { changedAt: "2026-04-20", fromStage: "Intake", toStage: "Registration" },
    { changedAt: "2026-04-25", fromStage: "Registration", toStage: "Quote" },
  ],
};

test("computedBlockers includes document and quote blockers", () => {
  assert.deepEqual(computedBlockers(baseProject), [
    "Endorsements received",
    "Premium BDX received",
    "Quote not accepted",
  ]);
});

test("stage gate blocks scheduling until quote is accepted", () => {
  assert.equal(
    canMoveToStage(baseProject, "Scheduling"),
    "Quote must be accepted before moving to Scheduling.",
  );
});

test("document readiness workflow can mark waiting on broker", () => {
  const updated = applyDocumentWorkflowAction(
    baseProject,
    "markWaitingOnBroker",
    "2026-05-07",
  );

  assert.equal(updated.assignmentStatus, "On Hold");
  assert.equal(updated.documentRequestDate, "2026-05-07");
  assert.equal(updated.brokerLastChasedDate, "2026-05-07");
  assert.ok(updated.labels.includes("Waiting on Broker"));
});

test("document readiness workflow can clear waiting on broker without completing documents", () => {
  const waiting = applyDocumentWorkflowAction(
    baseProject,
    "markWaitingOnBroker",
    "2026-05-07",
  );
  const cleared = applyDocumentWorkflowAction(
    waiting,
    "clearWaitingOnBroker",
    "2026-05-08",
  );

  assert.equal(cleared.labels.includes("Waiting on Broker"), false);
  assert.equal(cleared.assignmentStatus, "In Progress");
  assert.equal(documentReadiness(cleared).percent < 100, true);
});

test("document readiness workflow completes documents and clears waiting label", () => {
  const waiting = applyDocumentWorkflowAction(
    baseProject,
    "markWaitingOnBroker",
    "2026-05-07",
  );
  const complete = applyDocumentWorkflowAction(
    waiting,
    "markDocumentsComplete",
    "2026-05-08",
  );

  assert.equal(documentReadiness(complete).percent, 100);
  assert.equal(complete.assignmentStatus, "In Progress");
  assert.equal(complete.labels.includes("Waiting on Broker"), false);
});

test("role-based audit team exposes every assigned auditor", () => {
  assert.deepEqual(assignedAuditorNames(baseProject), [
    "Lorraine Mojica",
    "Leslie Domenech",
  ]);
});

test("stage duration metrics support selectable ranges", () => {
  const metrics = stageDurationMetrics([baseProject], "ytd");
  assert.equal(metrics.some((metric) => metric.stage === "Registration"), true);
});

test("recommended next steps returns five prioritized actions", () => {
  const steps = recommendedNextSteps(baseProject);

  assert.equal(steps.length, 5);
  assert.equal(
    steps[0],
    "Chase missing documents: Endorsements received, Premium BDX received.",
  );
  assert.equal(
    steps[1],
    "Confirm quote status and capture the client decision.",
  );
});
