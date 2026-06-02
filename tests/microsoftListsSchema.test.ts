import assert from "node:assert/strict";
import test from "node:test";
import {
  buildMicrosoftListsMigrationPackage,
  microsoftListSchemas,
  toGraphListCreateRequest,
  type CentralAuditProject,
  type CentralPrototypeUser,
} from "../src/microsoftListsSchema.js";

const project: CentralAuditProject = {
  id: "project-1",
  assignmentNumber: "A-100",
  assignmentSource: "Email",
  assignmentType: "CH",
  auditEntity: "Example Coverholder",
  clientCoverholderCode: "CH-001",
  broker: "Example Broker",
  assignedAuditor: "Lorraine Mojica",
  auditTeam: [
    { person: "Lorraine Mojica", role: "Lead Auditor" },
    { person: "Leslie Domenech", role: "Supporting Auditor" },
  ],
  currentStage: "Quote",
  assignmentStatus: "In Progress",
  quoteStatus: "Sent",
  quoteAmount: 12500,
  tentativeAuditWeek: "2026-W22",
  confirmedAuditDate: "",
  auditType: "Remote",
  baaReceived: true,
  endorsementsReceived: false,
  premiumBdxReceived: false,
  preAuditQuestionnaireStatus: "In Progress",
  documentRequestStatus: "In Progress",
  documentRequestDate: "2026-05-01",
  brokerLastChasedDate: "2026-05-03",
  brokerExpectedResponseDate: "2026-05-07",
  fileSelectionCompleted: false,
  testingSheetCompleted: false,
  findingsSentDate: "",
  coverholderResponseReceivedDate: "",
  reportStatus: "Not Started",
  invoiceStatus: "Not Started",
  paymentReceived: false,
  damSubmissionStatus: "Not Required",
  nextAction: "Confirm quote acceptance.",
  blockers: "",
  dueDate: "2026-05-10",
  lastUpdatedDate: "2026-05-05",
  labels: ["Medium Priority"],
  checklistCompletions: {
    "Quote:Quote drafted": true,
    "Quote:Quote accepted": false,
  },
  statusHistory: [
    {
      id: "history-1",
      changedAt: "2026-05-02T14:00:00.000Z",
      changedBy: "Justin Mojica",
      fromStage: "Registration",
      toStage: "Quote",
      note: "Quote work started.",
    },
  ],
  comments: [
    {
      id: "comment-1",
      createdAt: "2026-05-04T13:00:00.000Z",
      author: "Lorraine Mojica",
      body: "Waiting on broker confirmation.",
    },
  ],
  activityEvents: [
    {
      id: "event-1",
      createdAt: "2026-05-03T15:00:00.000Z",
      actor: "Justin Mojica",
      type: "document",
      title: "Broker chase recorded",
      detail: "Broker chase recorded",
    },
  ],
};

const user: CentralPrototypeUser = {
  fullName: "Justin Mojica",
  username: "justin.mojica",
  role: "Admin",
  permissionGroup: "Admin",
  email: "justin.mojica@example.com",
  active: true,
  defaultVisibility: "All Projects",
};

test("Microsoft Lists schema can produce Graph list create payloads", () => {
  const assignmentSchema = microsoftListSchemas.find(
    (schema) => schema.key === "auditAssignments",
  );
  if (!assignmentSchema) throw new Error("Audit Assignments schema missing");

  const request = toGraphListCreateRequest(assignmentSchema);
  assert.equal(request.displayName, "Audit Assignments");
  assert.equal((request.list as { template: string }).template, "genericList");
  assert.ok(
    (request.columns as Array<{ name: string }>).some(
      (column) => column.name === "TrackerAssignmentId",
    ),
  );
});

test("migration package flattens prototype data into central list rows", () => {
  const migrationPackage = buildMicrosoftListsMigrationPackage([project], [user], {
    exportedAt: "2026-05-10T12:00:00.000Z",
    exportedBy: "Justin Mojica",
  });

  assert.equal(migrationPackage.totals.lists, 7);
  assert.equal(migrationPackage.totals.assignments, 1);
  assert.equal(migrationPackage.rows.auditAssignments.length, 1);
  assert.equal(migrationPackage.rows.auditTeamMembers.length, 2);
  assert.equal(migrationPackage.rows.auditChecklistItems.length, 2);
  assert.equal(migrationPackage.rows.auditActivityLog.length, 3);
  assert.equal(
    migrationPackage.rows.auditActivityLog[0].fields.TrackerAssignmentId,
    "project-1",
  );
});
