import assert from "node:assert/strict";
import test from "node:test";
import {
  approveAccessRequest,
  buildAccessRequestUser,
  canApproveAccessRequest,
  isValidCompanyEmail,
  verifyAccessRequestEmail,
} from "../src/accessRequests.js";

test("account requests require company-style email addresses", () => {
  assert.equal(isValidCompanyEmail("new.user@example-company.com"), true);
  assert.equal(isValidCompanyEmail("new.user@gmail.com"), false);
  assert.equal(isValidCompanyEmail("not-an-email"), false);
});

test("email verification must happen before admin approval", () => {
  const request = buildAccessRequestUser({
    fullName: "New Auditor",
    email: "new.auditor@example-company.com",
    password: "password",
    requestedAt: "2026-05-29T12:00:00.000Z",
    verificationCode: "123456",
  });

  assert.equal(request.username, "new.auditor");
  assert.equal(request.active, false);
  assert.equal(canApproveAccessRequest(request), false);

  const result = verifyAccessRequestEmail([request], request.email, "123456");
  const verifiedUser = result.users[0];

  assert.equal(result.matched, true);
  assert.equal(verifiedUser.emailVerified, true);
  assert.equal(canApproveAccessRequest(verifiedUser), true);

  const approvedUser = approveAccessRequest(
    verifiedUser,
    "Justin Mojica",
    "2026-05-29T13:00:00.000Z",
  );

  assert.equal(approvedUser.active, true);
  assert.equal(approvedUser.accessRequestStatus, "Approved");
  assert.equal(approvedUser.approvedBy, "Justin Mojica");
});
