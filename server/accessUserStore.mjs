import { existsSync } from "node:fs";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import { dirname } from "node:path";

const graphBaseUrl = "https://graph.microsoft.com/v1.0";

export function createAccessUserStore({ mode, usersFile, graph }) {
  if (mode === "microsoft-lists") {
    return createMicrosoftListsUserStore(graph);
  }
  return createLocalUserStore(usersFile);
}

export function accessUserStoreStatus(config) {
  const mode = config.userStoreMode === "microsoft-lists" ? "microsoft-lists" : "local";
  const missing = [];
  if (mode === "microsoft-lists") {
    if (!hasConfiguredValue(config.userStoreSiteId)) missing.push("TRACKER_USERS_SITE_ID");
    if (!hasConfiguredValue(config.userStoreListId)) missing.push("TRACKER_USERS_LIST_ID");
  }
  return {
    mode,
    configured: missing.length === 0,
    missing,
  };
}

export function userToTrackerFields(user) {
  const fields = {
    Title: user.fullName || user.email,
    TrackerUsername: user.username,
    FullName: user.fullName,
    Email: user.email,
    Role: user.role,
    PermissionGroup: user.permissionGroup,
    Active: Boolean(user.active),
    EmailVerified: Boolean(user.emailVerified),
    AccessRequestStatus: user.accessRequestStatus,
    DefaultVisibility: user.defaultVisibility,
    RequestedAt: nullableDate(user.requestedAt),
    ApprovedAt: nullableDate(user.approvedAt),
    ApprovedBy: user.approvedBy || "",
    RejectionReason: user.rejectionReason || "",
    VerificationCodeHash: user.verificationCodeHash || "",
    VerificationSentAt: nullableDate(user.verificationSentAt),
  };
  return fields;
}

export function trackerFieldsToUser(fields) {
  const email = String(fields.Email || "").toLowerCase();
  const username = String(fields.TrackerUsername || usernameFromEmail(email));
  return {
    email,
    username,
    fullName: String(fields.FullName || fields.Title || username),
    role: String(fields.Role || "Auditor"),
    permissionGroup: String(fields.PermissionGroup || fields.Role || "Auditor"),
    active: Boolean(fields.Active),
    defaultVisibility: String(fields.DefaultVisibility || "Role Default"),
    emailVerified: Boolean(fields.EmailVerified),
    accessRequestStatus: String(fields.AccessRequestStatus || "Pending Verification"),
    requestedAt: stringValue(fields.RequestedAt),
    approvedAt: stringValue(fields.ApprovedAt),
    approvedBy: String(fields.ApprovedBy || ""),
    rejectionReason: String(fields.RejectionReason || ""),
    verificationCodeHash: String(fields.VerificationCodeHash || ""),
    verificationSentAt: stringValue(fields.VerificationSentAt),
  };
}

function createLocalUserStore(usersFile) {
  return {
    async load() {
      if (!existsSync(usersFile)) return { users: [] };
      const store = JSON.parse(await readFile(usersFile, "utf8"));
      return { users: Array.isArray(store.users) ? store.users : [] };
    },
    async save(store) {
      await mkdir(dirname(usersFile), { recursive: true });
      await writeFile(usersFile, `${JSON.stringify(store, null, 2)}\n`);
    },
  };
}

function createMicrosoftListsUserStore({ siteId, listId, getAccessToken }) {
  return {
    async load() {
      const items = await listTrackerUserItems({ siteId, listId, getAccessToken });
      return {
        users: items
          .map((item) => trackerFieldsToUser(item.fields ?? {}))
          .filter((user) => user.email),
      };
    },
    async save(store) {
      const items = await listTrackerUserItems({ siteId, listId, getAccessToken });
      const existingByEmail = new Map(
        items
          .map((item) => [String(item.fields?.Email || "").toLowerCase(), item])
          .filter(([email]) => email),
      );
      for (const user of store.users) {
        const fields = userToTrackerFields(user);
        const existing = existingByEmail.get(String(user.email).toLowerCase());
        if (existing) {
          await graphRequest(
            getAccessToken,
            `/sites/${encodeGraphSegment(siteId)}/lists/${encodeGraphSegment(
              listId,
            )}/items/${encodeGraphSegment(existing.id)}/fields`,
            {
              method: "PATCH",
              body: JSON.stringify(fields),
            },
          );
        } else {
          await graphRequest(
            getAccessToken,
            `/sites/${encodeGraphSegment(siteId)}/lists/${encodeGraphSegment(
              listId,
            )}/items`,
            {
              method: "POST",
              body: JSON.stringify({ fields }),
            },
          );
        }
      }
    },
  };
}

async function listTrackerUserItems({ siteId, listId, getAccessToken }) {
  const items = [];
  let path =
    `/sites/${encodeGraphSegment(siteId)}/lists/${encodeGraphSegment(
      listId,
    )}/items?expand=fields`;
  while (path) {
    const response = await graphRequest(getAccessToken, path);
    items.push(...(response.value ?? []));
    path = response["@odata.nextLink"] ?? "";
  }
  return items;
}

async function graphRequest(getAccessToken, pathOrUrl, init = {}) {
  const token = await getAccessToken();
  const url = pathOrUrl.startsWith("https://")
    ? pathOrUrl
    : `${graphBaseUrl}${pathOrUrl}`;
  const response = await fetch(url, {
    ...init,
    headers: {
      authorization: `Bearer ${token}`,
      "content-type": "application/json",
      ...(init.headers ?? {}),
    },
  });
  if (!response.ok) {
    throw new Error(`Graph ${response.status}: ${await response.text()}`);
  }
  if (response.status === 204) return {};
  return response.json();
}

function nullableDate(value) {
  const text = String(value || "").trim();
  return text ? text : null;
}

function stringValue(value) {
  return value === undefined || value === null ? "" : String(value);
}

function usernameFromEmail(email) {
  return email
    .split("@")[0]
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ".")
    .replace(/^\.+|\.+$/g, "");
}

function encodeGraphSegment(value) {
  return encodeURIComponent(value);
}

function hasConfiguredValue(value) {
  const normalized = String(value ?? "").trim().toLowerCase();
  if (!normalized) return false;
  if (normalized.startsWith("replace-with-")) return false;
  return true;
}
