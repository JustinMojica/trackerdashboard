import assert from "node:assert/strict";
import { createHmac } from "node:crypto";
import { spawn } from "node:child_process";
import test from "node:test";
import { fileURLToPath } from "node:url";

const serverPath = fileURLToPath(new URL("../server/secureAccessServer.mjs", import.meta.url));

function base64Url(value) {
  return Buffer.from(value).toString("base64url");
}

function forgedSessionCookie(secret, email = "admin@example.com") {
  const payload = base64Url(JSON.stringify({ email, status: "approved" }));
  const signature = createHmac("sha256", secret).update(payload).digest("base64url");
  return `tracker_session=${payload}.${signature}`;
}

function envFor(port, sessionSecret) {
  const env = {
    ...process.env,
    TRACKER_SERVER_PORT: String(port),
    MICROSOFT_TENANT_ID: "example-tenant",
    MICROSOFT_CLIENT_ID: "example-client",
    MICROSOFT_CLIENT_SECRET: "example-client-secret",
    MICROSOFT_MAIL_FROM: "tracker@example.com",
    TRACKER_ALLOWED_EMAIL_DOMAINS: "example.com",
    TRACKER_ADMIN_EMAILS: "admin@example.com",
    TRACKER_USER_STORE: "local",
  };
  if (sessionSecret === undefined) {
    env.TRACKER_SESSION_SECRET = "";
  } else {
    env.TRACKER_SESSION_SECRET = sessionSecret;
  }
  return env;
}

async function withServer(sessionSecret, run) {
  const port = 19000 + Math.floor(Math.random() * 1000);
  const server = spawn(process.execPath, [serverPath], {
    env: envFor(port, sessionSecret),
    stdio: ["ignore", "pipe", "pipe"],
  });
  let output = "";
  server.stdout.on("data", (chunk) => {
    output += chunk.toString();
  });
  server.stderr.on("data", (chunk) => {
    output += chunk.toString();
  });
  try {
    await new Promise((resolve, reject) => {
      const timeout = setTimeout(() => reject(new Error(`server did not start: ${output}`)), 5000);
      server.stdout.on("data", (chunk) => {
        if (chunk.toString().includes("Tracker secure access server listening")) {
          clearTimeout(timeout);
          resolve();
        }
      });
      server.on("exit", (code) => {
        clearTimeout(timeout);
        reject(new Error(`server exited with ${code}: ${output}`));
      });
    });
    await run(`http://127.0.0.1:${port}`);
  } finally {
    if (server.exitCode === null) {
      server.kill();
      await new Promise((resolve) => {
        const timeout = setTimeout(() => {
          server.kill("SIGKILL");
          resolve();
        }, 1000);
        server.once("exit", () => {
          clearTimeout(timeout);
          resolve();
        });
      });
    }
  }
}

test("missing session secret is not accepted for signed admin cookies", async () => {
  await withServer(undefined, async (origin) => {
    const configResponse = await fetch(`${origin}/api/auth/config`);
    const config = await configResponse.json();
    assert.equal(config.configured, false);
    assert.ok(config.missing.includes("TRACKER_SESSION_SECRET"));

    const adminResponse = await fetch(`${origin}/api/admin/access-requests`, {
      headers: { cookie: forgedSessionCookie("dev") },
    });
    assert.equal(adminResponse.status, 503);
  });
});

test("placeholder session secret is not accepted for signed admin cookies", async () => {
  await withServer("replace-with-a-long-random-secret", async (origin) => {
    const configResponse = await fetch(`${origin}/api/auth/config`);
    const config = await configResponse.json();
    assert.equal(config.configured, false);
    assert.ok(config.missing.includes("TRACKER_SESSION_SECRET"));

    const adminResponse = await fetch(`${origin}/api/admin/access-requests`, {
      headers: { cookie: forgedSessionCookie("replace-with-a-long-random-secret") },
    });
    assert.equal(adminResponse.status, 503);
  });
});
