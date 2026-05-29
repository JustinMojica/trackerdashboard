import {
  buildMicrosoftListsMigrationPackage,
  type CentralAuditProject,
  type CentralPrototypeUser,
  type MicrosoftListSchema,
  type MicrosoftListSeedRow,
  type MicrosoftListsMigrationPackage,
} from "./microsoftListsSchema.js";

export type MicrosoftListKey = MicrosoftListSchema["key"];

export type MicrosoftListsConnectionConfig = {
  siteId: string;
  listIds: Partial<Record<MicrosoftListKey, string>>;
};

export type MicrosoftListsSession = MicrosoftListsConnectionConfig & {
  accessToken: string;
};

export type MicrosoftListsSyncSummary = {
  created: number;
  updated: number;
  skipped: number;
  errors: string[];
};

type GraphListItem = {
  id: string;
  fields?: Record<string, string | number | boolean | null>;
};

const graphBaseUrl = "https://graph.microsoft.com/v1.0";

export const microsoftListLabels: Record<MicrosoftListKey, string> = {
  auditAssignments: "Audit Assignments",
  auditTeamMembers: "Audit Team Members",
  auditComments: "Audit Comments",
  auditChecklistItems: "Audit Checklist Items",
  auditStatusHistory: "Audit Status History",
  auditActivityLog: "Audit Activity Log",
  trackerUsers: "Tracker Users",
};

export const requiredMicrosoftListKeys: MicrosoftListKey[] = [
  "auditAssignments",
  "auditTeamMembers",
  "auditComments",
  "auditChecklistItems",
  "auditStatusHistory",
  "auditActivityLog",
  "trackerUsers",
];

const uniqueFieldByList: Record<MicrosoftListKey, string> = {
  auditAssignments: "TrackerAssignmentId",
  auditTeamMembers: "TeamMemberKey",
  auditComments: "TrackerCommentId",
  auditChecklistItems: "TrackerChecklistItemId",
  auditStatusHistory: "TrackerHistoryId",
  auditActivityLog: "TrackerEventId",
  trackerUsers: "TrackerUsername",
};

export function hasMinimumMicrosoftListsConfig(
  config: MicrosoftListsConnectionConfig,
) {
  return Boolean(config.siteId && config.listIds.auditAssignments);
}

export function hasFullMicrosoftListsConfig(
  config: MicrosoftListsConnectionConfig,
) {
  return requiredMicrosoftListKeys.every((key) => Boolean(config.listIds[key]));
}

export function missingMicrosoftListLabels(config: MicrosoftListsConnectionConfig) {
  return requiredMicrosoftListKeys
    .filter((key) => !config.listIds[key])
    .map((key) => microsoftListLabels[key]);
}

export function sanitizeMicrosoftListsConfig(
  config: MicrosoftListsConnectionConfig,
): MicrosoftListsConnectionConfig {
  const listIds = Object.fromEntries(
    Object.entries(config.listIds).map(([key, value]) => [
      key,
      String(value ?? "").trim(),
    ]),
  ) as Partial<Record<MicrosoftListKey, string>>;
  return {
    siteId: config.siteId.trim(),
    listIds,
  };
}

export async function testMicrosoftListsConnection(session: MicrosoftListsSession) {
  const listId = session.listIds.auditAssignments;
  if (!session.siteId || !listId) {
    throw new Error("Site ID and Audit Assignments list ID are required.");
  }
  await graphRequest(
    session,
    `/sites/${encodeGraphSegment(session.siteId)}/lists/${encodeGraphSegment(listId)}`,
  );
}

export async function pullAssignmentRowsFromMicrosoftLists(
  session: MicrosoftListsSession,
) {
  const listId = session.listIds.auditAssignments;
  if (!session.siteId || !listId) {
    throw new Error("Site ID and Audit Assignments list ID are required.");
  }
  const items = await listItems(session, listId);
  return items.map((item) => item.fields ?? {});
}

export async function pushMigrationPackageToMicrosoftLists(
  session: MicrosoftListsSession,
  migrationPackage: MicrosoftListsMigrationPackage,
): Promise<MicrosoftListsSyncSummary> {
  const summary: MicrosoftListsSyncSummary = {
    created: 0,
    updated: 0,
    skipped: 0,
    errors: [],
  };

  for (const [listKey, rows] of Object.entries(migrationPackage.rows) as [
    MicrosoftListKey,
    MicrosoftListSeedRow[],
  ][]) {
    if (rows.length === 0) continue;

    const listId = session.listIds[listKey];
    if (!listId) {
      summary.skipped += rows.length;
      summary.errors.push(`${microsoftListLabels[listKey]} list ID is missing.`);
      continue;
    }

    const uniqueField = uniqueFieldByList[listKey];
    const existingByKey = uniqueField
      ? indexItemsByField(await listItems(session, listId), uniqueField)
      : new Map<string, GraphListItem>();

    for (const row of rows) {
      try {
        const uniqueValue = uniqueField ? String(row.fields[uniqueField] ?? "") : "";
        const existingItem = uniqueValue ? existingByKey.get(uniqueValue) : undefined;
        if (existingItem) {
          await updateListItemFields(session, listId, existingItem.id, row.fields);
          summary.updated += 1;
        } else {
          await createListItem(session, listId, row.fields);
          summary.created += 1;
        }
      } catch (error) {
        summary.errors.push(
          `${microsoftListLabels[listKey]}: ${
            error instanceof Error ? error.message : "Unknown sync error"
          }`,
        );
      }
    }
  }

  return summary;
}

export function buildSyncPackage(
  projects: CentralAuditProject[],
  users: CentralPrototypeUser[],
  exportedBy: string,
) {
  return buildMicrosoftListsMigrationPackage(projects, users, { exportedBy });
}

async function listItems(session: MicrosoftListsSession, listId: string) {
  const items: GraphListItem[] = [];
  let path =
    `/sites/${encodeGraphSegment(session.siteId)}/lists/${encodeGraphSegment(
      listId,
    )}/items?expand=fields`;

  while (path) {
    const response = await graphRequest<{
      value?: GraphListItem[];
      "@odata.nextLink"?: string;
    }>(session, path);
    items.push(...(response.value ?? []));
    path = response["@odata.nextLink"] ?? "";
  }

  return items;
}

async function createListItem(
  session: MicrosoftListsSession,
  listId: string,
  fields: Record<string, string | number | boolean>,
) {
  await graphRequest(
    session,
    `/sites/${encodeGraphSegment(session.siteId)}/lists/${encodeGraphSegment(
      listId,
    )}/items`,
    {
      method: "POST",
      body: JSON.stringify({ fields }),
    },
  );
}

async function updateListItemFields(
  session: MicrosoftListsSession,
  listId: string,
  itemId: string,
  fields: Record<string, string | number | boolean>,
) {
  await graphRequest(
    session,
    `/sites/${encodeGraphSegment(session.siteId)}/lists/${encodeGraphSegment(
      listId,
    )}/items/${encodeGraphSegment(itemId)}/fields`,
    {
      method: "PATCH",
      body: JSON.stringify(fields),
    },
  );
}

async function graphRequest<T = unknown>(
  session: MicrosoftListsSession,
  pathOrUrl: string,
  init: RequestInit = {},
): Promise<T> {
  const url = pathOrUrl.startsWith("https://")
    ? pathOrUrl
    : `${graphBaseUrl}${pathOrUrl}`;
  const response = await fetch(url, {
    ...init,
    headers: {
      Authorization: `Bearer ${session.accessToken}`,
      "Content-Type": "application/json",
      ...(init.headers ?? {}),
    },
  });
  if (!response.ok) {
    const detail = await response.text().catch(() => "");
    throw new Error(`Graph ${response.status}: ${detail || response.statusText}`);
  }
  if (response.status === 204) return undefined as T;
  return (await response.json()) as T;
}

function indexItemsByField(items: GraphListItem[], fieldName: string) {
  const map = new Map<string, GraphListItem>();
  items.forEach((item) => {
    const value = item.fields?.[fieldName];
    if (value !== undefined && value !== null && value !== "") {
      map.set(String(value), item);
    }
  });
  return map;
}

function encodeGraphSegment(value: string) {
  return encodeURIComponent(value);
}
