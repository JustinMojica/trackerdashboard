import assert from "node:assert/strict";
import test from "node:test";
import {
  accessUserStoreStatus,
  trackerFieldsToUser,
  userToTrackerFields,
} from "../server/accessUserStore.mjs";

test("access user store reports Microsoft Lists setup requirements", () => {
  assert.deepEqual(
    accessUserStoreStatus({
      userStoreMode: "microsoft-lists",
      userStoreSiteId: "",
      userStoreListId: "",
    }),
    {
      mode: "microsoft-lists",
      configured: false,
      missing: ["TRACKER_USERS_SITE_ID", "TRACKER_USERS_LIST_ID"],
    },
  );

  assert.deepEqual(
    accessUserStoreStatus({
      userStoreMode: "local",
      userStoreSiteId: "",
      userStoreListId: "",
    }),
    {
      mode: "local",
      configured: true,
      missing: [],
    },
  );
});

test("tracker user fields preserve access approval and verification state", () => {
  const user = {
    email: "new.user@mosaic-international.com",
    username: "new.user",
    fullName: "New User",
    role: "Auditor",
    permissionGroup: "Auditor",
    active: false,
    defaultVisibility: "Role Default",
    emailVerified: false,
    accessRequestStatus: "Pending Verification",
    requestedAt: "2026-05-30T01:00:00.000Z",
    approvedAt: "",
    approvedBy: "",
    rejectionReason: "",
    verificationCodeHash: "hash-value",
    verificationSentAt: "2026-05-30T01:01:00.000Z",
  };

  const fields = userToTrackerFields(user);
  assert.equal(fields.Email, "new.user@mosaic-international.com");
  assert.equal(fields.AccessRequestStatus, "Pending Verification");
  assert.equal(fields.VerificationCodeHash, "hash-value");
  assert.equal(fields.ApprovedAt, null);

  assert.deepEqual(trackerFieldsToUser(fields), user);
});
