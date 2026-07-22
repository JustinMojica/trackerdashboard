import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { createHmac } from "node:crypto";
import { once } from "node:events";
import { existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { createServer } from "node:net";
import { dirname, resolve } from "node:path";
import test from "node:test";

function base64Url(input) {
  return Buffer.from(input).toString("base64url");
}

function forgedSessionCookie(secret, session) {
  const payload = base64Url(Buffer.from(JSON.stringify(session)));
  const signature = base64Url(createHmac("sha256", secret).update(payload).digest());
  return `tracker_session=${payload}.${signature}`;
}

async function freePort() {
  const server = createServer();
  server.listen(0, "127.0.0.1");
  await once(server, "listening");
  const { port } = server.address();
  server.close();
  await once(server, "close");
  return port;
}

async function startSecureServer(env) {
  const port = await freePort();
  const child = spawn(process.execPath, ["server/secureAccessServer.mjs"], {
    cwd: process.cwd(),
    env: {
      ...process.env,
      ...env,
      PORT: String(port),
      TRACKER_PUBLIC_ORIGIN: `http://127.0.0.1:${port}`,
      TRACKER_FRONTEND_ORIGIN: `http://127.0.0.1:${port}`,
    },
    stdio: ["ignore", "pipe", "pipe"],
  });
  let output = "";
  child.stdout.on("data", (chunk) => {
    output += chunk.toString();
  });
  child.stderr.on("data", (chunk) => {
    output += chunk.toString();
  });

  const baseUrl = `http://127.0.0.1:${port}`;
  for (let attempt = 0; attempt < 100; attempt += 1) {
    if (child.exitCode !== null) {
      throw new Error(`secure server exited early with ${child.exitCode}: ${output}`);
    }
    try {
      const response = await fetch(`${baseUrl}/api/auth/config`);
      if (response.ok) return { baseUrl, child };
    } catch {
      await new Promise((resolve) => setTimeout(resolve, 25));
    }
  }
  child.kill();
  throw new Error(`secure server did not start: ${output}`);
}

async function stopSecureServer(child) {
  if (child.exitCode !== null) return;
  child.kill();
  await once(child, "exit");
}

async function withAccessUsersFile(users, run) {
  const usersFile = resolve(process.cwd(), "server", "data", "access-users.json");
  const hadFile = existsSync(usersFile);
  const prior = hadFile ? readFileSync(usersFile, "utf8") : "";
  mkdirSync(dirname(usersFile), { recursive: true });
  writeFileSync(usersFile, JSON.stringify({ users }, null, 2));
  try {
    await run(usersFile);
  } finally {
    if (hadFile) {
      writeFileSync(usersFile, prior);
    } else {
      rmSync(usersFile, { force: true });
    }
  }
}

test("admin APIs fail closed when the session secret is weak", async () => {
  const { baseUrl, child } = await startSecureServer({
    MICROSOFT_TENANT_ID: "11111111-1111-1111-1111-111111111111",
    MICROSOFT_CLIENT_ID: "22222222-2222-2222-2222-222222222222",
    MICROSOFT_CLIENT_SECRET: "configured-client-secret",
    MICROSOFT_MAIL_FROM: "tracker@example.com",
    TRACKER_SESSION_SECRET: "dev",
    TRACKER_ALLOWED_EMAIL_DOMAINS: "example.com",
    TRACKER_ADMIN_EMAILS: "admin@example.com",
    TRACKER_USER_STORE: "local",
    TRACKER_USERS_SITE_ID: "",
    TRACKER_USERS_LIST_ID: "",
  });

  try {
    const cookie = forgedSessionCookie("dev", {
      email: "admin@example.com",
      status: "approved",
    });
    const response = await fetch(`${baseUrl}/api/admin/access-requests`, {
      headers: { cookie },
    });
    const body = await response.json();

    assert.equal(response.status, 503);
    assert.equal(body.error, "setup-required");
    assert.equal(body.setup.configured, false);
    assert.deepEqual(body.setup.invalid, ["TRACKER_SESSION_SECRET"]);
  } finally {
    await stopSecureServer(child);
  }
});

test("configured admin email is repaired if stored as a non-admin", async () => {
  await withAccessUsersFile(
    [
      {
        email: "admin@example.com",
        username: "admin",
        fullName: "Admin User",
        role: "Auditor",
        permissionGroup: "Auditor",
        active: false,
        defaultVisibility: "Assigned Projects",
        emailVerified: false,
        accessRequestStatus: "Pending Approval",
        requestedAt: "2026-06-01T00:00:00.000Z",
        approvedAt: "",
        approvedBy: "",
        rejectionReason: "",
        verificationCodeHash: "stale",
        verificationSentAt: "2026-06-01T00:00:00.000Z",
      },
    ],
    async () => {
      const sessionSecret = "this-is-a-strong-test-session-secret-value";
      const { baseUrl, child } = await startSecureServer({
        MICROSOFT_TENANT_ID: "11111111-1111-1111-1111-111111111111",
        MICROSOFT_CLIENT_ID: "22222222-2222-2222-2222-222222222222",
        MICROSOFT_CLIENT_SECRET: "configured-client-secret",
        MICROSOFT_MAIL_FROM: "tracker@example.com",
        TRACKER_SESSION_SECRET: sessionSecret,
        TRACKER_ALLOWED_EMAIL_DOMAINS: "example.com",
        TRACKER_ADMIN_EMAILS: "admin@example.com",
        TRACKER_USER_STORE: "local",
        TRACKER_USERS_SITE_ID: "",
        TRACKER_USERS_LIST_ID: "",
      });

      try {
        const cookie = forgedSessionCookie(sessionSecret, {
          email: "admin@example.com",
          status: "approved",
        });
        const response = await fetch(`${baseUrl}/api/auth/me`, {
          headers: { cookie },
        });
        const body = await response.json();

        assert.equal(response.status, 200);
        assert.equal(body.authenticated, true);
        assert.equal(body.user.role, "Admin");
        assert.equal(body.user.permissionGroup, "Admin");
        assert.equal(body.user.active, true);
        assert.equal(body.user.accessRequestStatus, "Approved");
        assert.ok(Array.isArray(body.managedUsers));
      } finally {
        await stopSecureServer(child);
      }
    },
  );
});

test("expired email verification codes are rejected and cleared", async () => {
  const sessionSecret = "this-is-a-strong-test-session-secret-value";
  const email = "pending@example.com";
  const code = "123456";
  const oldSentAt = new Date(Date.now() - 11 * 60 * 1000).toISOString();
  await withAccessUsersFile(
    [
      {
        email,
        username: "pending",
        fullName: "Pending User",
        role: "Auditor",
        permissionGroup: "Auditor",
        active: false,
        defaultVisibility: "Role Default",
        emailVerified: false,
        accessRequestStatus: "Pending Verification",
        requestedAt: oldSentAt,
        approvedAt: "",
        approvedBy: "",
        rejectionReason: "",
        verificationCodeHash: createHmac("sha256", sessionSecret)
          .update(`${email}:${code}`)
          .digest("hex"),
        verificationSentAt: oldSentAt,
      },
    ],
    async (usersFile) => {
      const { baseUrl, child } = await startSecureServer({
        MICROSOFT_TENANT_ID: "11111111-1111-1111-1111-111111111111",
        MICROSOFT_CLIENT_ID: "22222222-2222-2222-2222-222222222222",
        MICROSOFT_CLIENT_SECRET: "configured-client-secret",
        MICROSOFT_MAIL_FROM: "tracker@example.com",
        TRACKER_SESSION_SECRET: sessionSecret,
        TRACKER_ALLOWED_EMAIL_DOMAINS: "example.com",
        TRACKER_ADMIN_EMAILS: "admin@example.com",
        TRACKER_USER_STORE: "local",
        TRACKER_USERS_SITE_ID: "",
        TRACKER_USERS_LIST_ID: "",
      });

      try {
        const response = await fetch(`${baseUrl}/api/access/verify-code`, {
          method: "POST",
          headers: {
            "content-type": "application/json",
            cookie: forgedSessionCookie(sessionSecret, { email, status: "pending" }),
          },
          body: JSON.stringify({ code }),
        });
        const body = await response.json();
        const stored = JSON.parse(readFileSync(usersFile, "utf8"));

        assert.equal(response.status, 400);
        assert.equal(body.error, "code_expired");
        assert.equal(stored.users[0].accessRequestStatus, "Pending Verification");
        assert.equal(stored.users[0].verificationCodeHash, "");
      } finally {
        await stopSecureServer(child);
      }
    },
  );
});

test("email verification locks after repeated invalid codes", async () => {
  const sessionSecret = "this-is-a-strong-test-session-secret-value";
  const email = "pending@example.com";
  const code = "123456";
  const sentAt = new Date().toISOString();
  await withAccessUsersFile(
    [
      {
        email,
        username: "pending",
        fullName: "Pending User",
        role: "Auditor",
        permissionGroup: "Auditor",
        active: false,
        defaultVisibility: "Role Default",
        emailVerified: false,
        accessRequestStatus: "Pending Verification",
        requestedAt: sentAt,
        approvedAt: "",
        approvedBy: "",
        rejectionReason: "",
        verificationCodeHash: createHmac("sha256", sessionSecret)
          .update(`${email}:${code}`)
          .digest("hex"),
        verificationSentAt: sentAt,
      },
    ],
    async () => {
      const { baseUrl, child } = await startSecureServer({
        MICROSOFT_TENANT_ID: "11111111-1111-1111-1111-111111111111",
        MICROSOFT_CLIENT_ID: "22222222-2222-2222-2222-222222222222",
        MICROSOFT_CLIENT_SECRET: "configured-client-secret",
        MICROSOFT_MAIL_FROM: "tracker@example.com",
        TRACKER_SESSION_SECRET: sessionSecret,
        TRACKER_ALLOWED_EMAIL_DOMAINS: "example.com",
        TRACKER_ADMIN_EMAILS: "admin@example.com",
        TRACKER_USER_STORE: "local",
        TRACKER_USERS_SITE_ID: "",
        TRACKER_USERS_LIST_ID: "",
      });

      try {
        let latestBody = null;
        let latestStatus = 0;
        for (let attempt = 0; attempt < 5; attempt += 1) {
          const response = await fetch(`${baseUrl}/api/access/verify-code`, {
            method: "POST",
            headers: {
              "content-type": "application/json",
              cookie: forgedSessionCookie(sessionSecret, { email, status: "pending" }),
            },
            body: JSON.stringify({ code: "000000" }),
          });
          latestStatus = response.status;
          latestBody = await response.json();
        }

        assert.equal(latestStatus, 429);
        assert.equal(latestBody.error, "verification_locked");
        assert.equal(latestBody.remainingAttempts, 0);
        assert.ok(latestBody.retryAfterSeconds > 0);
      } finally {
        await stopSecureServer(child);
      }
    },
  );
});
