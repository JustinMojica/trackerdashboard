import assert from "node:assert/strict";
import test from "node:test";
import {
  hasFullMicrosoftListsConfig,
  missingMicrosoftListLabels,
  pullAssignmentRowsFromMicrosoftLists,
  pushMigrationPackageToMicrosoftLists,
  sanitizeMicrosoftListsConfig,
  type MicrosoftListsSession,
} from "../src/microsoftListsClient.js";
import {
  type MicrosoftListsMigrationPackage,
  type MicrosoftListSeedRow,
} from "../src/microsoftListsSchema.js";

const originalFetch = globalThis.fetch;

test("connection config is trimmed and reports missing list IDs", () => {
  const config = sanitizeMicrosoftListsConfig({
    siteId: "  site-id  ",
    listIds: {
      auditAssignments: " assignments ",
      trackerUsers: " users ",
    },
  });

  assert.equal(config.siteId, "site-id");
  assert.equal(config.listIds.auditAssignments, "assignments");
  assert.equal(hasFullMicrosoftListsConfig(config), false);
  assert.deepEqual(missingMicrosoftListLabels(config), [
    "Audit Team Members",
    "Audit Comments",
    "Audit Checklist Items",
    "Audit Status History",
    "Audit Activity Log",
  ]);
});

test("sync updates matching rows and creates new rows", async () => {
  const calls: Array<{ url: string; method: string; body?: unknown }> = [];
  globalThis.fetch = async (input, init = {}) => {
    const url = String(input);
    const method = init.method ?? "GET";
    const body = init.body ? JSON.parse(String(init.body)) : undefined;
    calls.push({ url, method, body });

    if (method === "GET" && url.includes("/lists/assignments/items")) {
      return jsonResponse({
        value: [
          {
            id: "10",
            fields: { TrackerAssignmentId: "project-1" },
          },
        ],
      });
    }
    if (method === "GET" && url.includes("/lists/team/items")) {
      return jsonResponse({ value: [] });
    }
    return jsonResponse({});
  };
  try {
    const session: MicrosoftListsSession = {
      siteId: "site-id",
      accessToken: "token",
      listIds: {
        auditAssignments: "assignments",
        auditTeamMembers: "team",
      },
    };
    const migrationPackage: MicrosoftListsMigrationPackage = {
      app: "audit-assignment-tracker",
      schemaVersion: 3,
      exportedAt: "2026-05-29T12:00:00.000Z",
      exportedBy: "Justin Mojica",
      storageTarget: "Microsoft Lists / SharePoint",
      schemas: [],
      graphListCreateRequests: [],
      rows: {
        auditAssignments: [
          row("auditAssignments", {
            Title: "A-100",
            TrackerAssignmentId: "project-1",
            AssignmentNumber: "A-100",
          }),
        ],
        auditTeamMembers: [
          row("auditTeamMembers", {
            Title: "A-100 - Lorraine Mojica",
            TeamMemberKey: "project-1|Lead Auditor|Lorraine Mojica",
            TrackerAssignmentId: "project-1",
            AssignmentNumber: "A-100",
            PersonName: "Lorraine Mojica",
            TeamRole: "Lead Auditor",
            ActiveOnAssignment: true,
          }),
        ],
        auditComments: [],
        auditChecklistItems: [],
        auditStatusHistory: [],
        auditActivityLog: [],
        trackerUsers: [],
      },
      totals: {
        lists: 7,
        rows: 2,
        assignments: 1,
        activityLogEvents: 0,
      },
    };

    const summary = await pushMigrationPackageToMicrosoftLists(
      session,
      migrationPackage,
    );

    assert.equal(summary.updated, 1);
    assert.equal(summary.created, 1);
    assert.equal(summary.skipped, 0);
    assert.equal(summary.errors.length, 0);
    assert.ok(
      calls.some(
        (call) =>
          call.method === "PATCH" &&
          call.url.includes("/lists/assignments/items/10/fields"),
      ),
    );
    assert.ok(
      calls.some(
        (call) =>
          call.method === "POST" && call.url.includes("/lists/team/items"),
      ),
    );
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test("pulling assignment rows follows Microsoft Graph pagination", async () => {
  const calls: string[] = [];
  globalThis.fetch = async (input) => {
    const url = String(input);
    calls.push(url);
    if (url.includes("page=2")) {
      return jsonResponse({
        value: [{ id: "2", fields: { TrackerAssignmentId: "project-2" } }],
      });
    }
    return jsonResponse({
      value: [{ id: "1", fields: { TrackerAssignmentId: "project-1" } }],
      "@odata.nextLink":
        "https://graph.microsoft.com/v1.0/sites/site-id/lists/assignments/items?expand=fields&page=2",
    });
  };
  try {
    const rows = await pullAssignmentRowsFromMicrosoftLists({
      siteId: "site-id",
      accessToken: "token",
      listIds: {
        auditAssignments: "assignments",
      },
    });

    assert.deepEqual(
      rows.map((fields) => fields.TrackerAssignmentId),
      ["project-1", "project-2"],
    );
    assert.equal(calls.length, 2);
  } finally {
    globalThis.fetch = originalFetch;
  }
});

function row(
  listKey: MicrosoftListSeedRow["listKey"],
  fields: Record<string, string | number | boolean>,
) {
  return { listKey, fields };
}

function jsonResponse(body: unknown) {
  return new Response(JSON.stringify(body), {
    status: 200,
    headers: { "Content-Type": "application/json" },
  });
}
