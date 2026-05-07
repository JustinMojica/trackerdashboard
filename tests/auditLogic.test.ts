import assert from "node:assert/strict";
import test from "node:test";
import {
  applyDocumentWorkflowAction,
  canMoveToStage,
  computedBlockers,
  documentReadiness,
  type LogicProject,
} from "../src/auditLogic.js";

const baseProject: LogicProject = {
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
