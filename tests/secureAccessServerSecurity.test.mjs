import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { createHmac } from "node:crypto";
import { once } from "node:events";
import { createServer } from "node:net";
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
