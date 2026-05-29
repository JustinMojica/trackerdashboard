import { createHmac, createHash, randomBytes } from "node:crypto";
import { existsSync } from "node:fs";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import { createServer } from "node:http";
import { extname, join, normalize, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { createRemoteJWKSet, jwtVerify } from "jose";

const rootDir = resolve(fileURLToPath(new URL("..", import.meta.url)));
const dataDir = resolve(rootDir, "server", "data");
const usersFile = resolve(dataDir, "access-users.json");
const distDir = resolve(rootDir, "dist");
const oauthStates = new Map();
const cookieName = "tracker_session";
const oauthCookieName = "tracker_oauth_state";

const config = {
  port: Number(process.env.TRACKER_SERVER_PORT || 8787),
  frontendOrigin: process.env.TRACKER_FRONTEND_ORIGIN || "http://127.0.0.1:5173",
  publicOrigin: process.env.TRACKER_PUBLIC_ORIGIN || "http://127.0.0.1:8787",
  tenantId: process.env.MICROSOFT_TENANT_ID || "",
  clientId: process.env.MICROSOFT_CLIENT_ID || "",
  clientSecret: process.env.MICROSOFT_CLIENT_SECRET || "",
  sessionSecret: process.env.TRACKER_SESSION_SECRET || "",
  allowedDomains: splitCsv(process.env.TRACKER_ALLOWED_EMAIL_DOMAINS || ""),
  adminEmails: splitCsv(process.env.TRACKER_ADMIN_EMAILS || "").map((email) =>
    email.toLowerCase(),
  ),
  mailFrom: process.env.MICROSOFT_MAIL_FROM || "",
};

const redirectUri = `${config.publicOrigin}/api/auth/callback`;
const authority = `https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0`;
const jwks = config.tenantId
  ? createRemoteJWKSet(
      new URL(`https://login.microsoftonline.com/${config.tenantId}/discovery/v2.0/keys`),
    )
  : null;

createServer((request, response) => {
  void route(request, response).catch((error) => {
    console.error(error);
    sendJson(request, response, 500, {
      error: "server_error",
      message: "The tracker access server hit an unexpected error.",
    });
  });
}).listen(config.port, () => {
  console.log(`Tracker secure access server listening on ${config.publicOrigin}`);
});

async function route(request, response) {
  if (request.method === "OPTIONS") {
    setCors(request, response);
    response.writeHead(204);
    response.end();
    return;
  }

  const url = new URL(request.url ?? "/", config.publicOrigin);
  if (url.pathname === "/api/auth/config") {
    return sendJson(request, response, 200, accessConfigStatus());
  }
  if (url.pathname === "/api/auth/me") {
    return sendJson(request, response, 200, await currentAccessState(request));
  }
  if (url.pathname === "/api/auth/start") {
    return startMicrosoftAuth(request, response, url);
  }
  if (url.pathname === "/api/auth/callback") {
    return finishMicrosoftAuth(request, response, url);
  }
  if (url.pathname === "/api/auth/logout" && request.method === "POST") {
    clearSession(response);
    return sendJson(request, response, 200, { ok: true });
  }
  if (url.pathname === "/api/access/verify-code" && request.method === "POST") {
    return verifyAccessCode(request, response);
  }
  if (url.pathname === "/api/admin/access-requests") {
    return listAccessRequests(request, response);
  }
  const approvalMatch = url.pathname.match(
    /^\/api\/admin\/access-requests\/([^/]+)\/(approve|reject)$/,
  );
  if (approvalMatch && request.method === "POST") {
    return decideAccessRequest(request, response, approvalMatch[1], approvalMatch[2]);
  }
  return serveStatic(request, response, url);
}

function accessConfigStatus() {
  const missing = [];
  for (const [key, value] of Object.entries({
    MICROSOFT_TENANT_ID: config.tenantId,
    MICROSOFT_CLIENT_ID: config.clientId,
    MICROSOFT_CLIENT_SECRET: config.clientSecret,
    MICROSOFT_MAIL_FROM: config.mailFrom,
    TRACKER_SESSION_SECRET: config.sessionSecret,
    TRACKER_ALLOWED_EMAIL_DOMAINS: config.allowedDomains.join(","),
    TRACKER_ADMIN_EMAILS: config.adminEmails.join(","),
  })) {
    if (!value) missing.push(key);
  }
  return {
    configured: missing.length === 0,
    missing,
    redirectUri,
    frontendOrigin: config.frontendOrigin,
  };
}

async function currentAccessState(request) {
  const setup = accessConfigStatus();
  if (!setup.configured) {
    return {
      configured: false,
      authenticated: false,
      status: "setup-required",
      setup,
    };
  }

  const session = readSignedCookie(request, cookieName);
  if (!session?.email) {
    return {
      configured: true,
      authenticated: false,
      status: "not-signed-in",
      signInUrl: `${config.publicOrigin}/api/auth/start?mode=signin`,
      requestAccessUrl: `${config.publicOrigin}/api/auth/start?mode=request`,
    };
  }

  const store = await loadUserStore();
  const user = findUser(store, session.email);
  if (!user) {
    return {
      configured: true,
      authenticated: false,
      status: "not-requested",
      signInUrl: `${config.publicOrigin}/api/auth/start?mode=signin`,
      requestAccessUrl: `${config.publicOrigin}/api/auth/start?mode=request`,
    };
  }

  const state = {
    configured: true,
    authenticated: user.active && user.accessRequestStatus === "Approved",
    status: statusSlug(user.accessRequestStatus),
    user: publicUser(user),
    signInUrl: `${config.publicOrigin}/api/auth/start?mode=signin`,
    requestAccessUrl: `${config.publicOrigin}/api/auth/start?mode=request`,
  };

  if (state.authenticated && user.role === "Admin") {
    state.pendingRequests = store.users
      .filter((candidate) =>
        ["Pending Verification", "Pending Approval"].includes(
          candidate.accessRequestStatus,
        ),
      )
      .map(publicUser);
  }
  return state;
}

async function startMicrosoftAuth(_request, response, url) {
  const setup = accessConfigStatus();
  if (!setup.configured) return sendSetupError(response, setup);

  const mode = url.searchParams.get("mode") === "request" ? "request" : "signin";
  const state = base64Url(randomBytes(32));
  const nonce = base64Url(randomBytes(32));
  const codeVerifier = base64Url(randomBytes(64));
  const codeChallenge = base64Url(createHash("sha256").update(codeVerifier).digest());
  oauthStates.set(state, {
    mode,
    nonce,
    codeVerifier,
    expiresAt: Date.now() + 10 * 60 * 1000,
  });
  response.setHeader("Set-Cookie", signedCookie(oauthCookieName, { state }, 600));
  redirect(
    response,
    `${authority}/authorize?${new URLSearchParams({
      client_id: config.clientId,
      response_type: "code",
      redirect_uri: redirectUri,
      response_mode: "query",
      scope: "openid profile email User.Read",
      state,
      nonce,
      code_challenge: codeChallenge,
      code_challenge_method: "S256",
      prompt: "select_account",
    })}`,
  );
}

async function finishMicrosoftAuth(request, response, url) {
  const setup = accessConfigStatus();
  if (!setup.configured) return sendSetupError(response, setup);

  const error = url.searchParams.get("error");
  if (error) return redirect(response, `${config.frontendOrigin}/?auth=${error}`);

  const state = url.searchParams.get("state") ?? "";
  const code = url.searchParams.get("code") ?? "";
  const stateCookie = readSignedCookie(request, oauthCookieName);
  const oauthState = oauthStates.get(state);
  oauthStates.delete(state);
  response.setHeader("Set-Cookie", clearCookie(oauthCookieName));

  if (!code || !stateCookie || stateCookie.state !== state || !oauthState) {
    return redirect(response, `${config.frontendOrigin}/?auth=invalid-state`);
  }
  if (oauthState.expiresAt < Date.now()) {
    return redirect(response, `${config.frontendOrigin}/?auth=expired-state`);
  }

  const token = await exchangeCodeForToken(code, oauthState.codeVerifier);
  const claims = await validateIdToken(token.id_token, oauthState.nonce);
  const email = claimEmail(claims);
  const name = String(claims.name || email.split("@")[0]);
  if (!isAllowedEmail(email)) {
    return redirect(response, `${config.frontendOrigin}/?auth=domain-blocked`);
  }

  const store = await loadUserStore();
  seedAdmins(store);
  let user = findUser(store, email);

  if (oauthState.mode === "request") {
    user = await createOrRefreshAccessRequest(store, email, name);
    await saveUserStore(store);
    setSession(response, { email, status: "pending" });
    return redirect(response, `${config.frontendOrigin}/?auth=code-sent`);
  }

  if (!user) {
    return redirect(response, `${config.frontendOrigin}/?auth=request-required`);
  }
  if (user.accessRequestStatus !== "Approved" || !user.active) {
    setSession(response, { email, status: "pending" });
    return redirect(response, `${config.frontendOrigin}/?auth=pending`);
  }

  setSession(response, { email, status: "approved" });
  return redirect(response, `${config.frontendOrigin}/?auth=success`);
}

async function createOrRefreshAccessRequest(store, email, name) {
  let user = findUser(store, email);
  const now = new Date().toISOString();
  const code = randomCode();
  if (!user) {
    user = {
      email,
      username: usernameFromEmail(email),
      fullName: name,
      role: "Auditor",
      permissionGroup: "Auditor",
      active: false,
      defaultVisibility: "Role Default",
      emailVerified: false,
      accessRequestStatus: "Pending Verification",
      requestedAt: now,
      approvedAt: "",
      approvedBy: "",
      rejectionReason: "",
      verificationCodeHash: "",
      verificationSentAt: "",
    };
    store.users.push(user);
  }
  if (user.accessRequestStatus === "Approved") return user;
  user.fullName = user.fullName || name;
  user.emailVerified = false;
  user.active = false;
  user.accessRequestStatus = "Pending Verification";
  user.rejectionReason = "";
  user.verificationCodeHash = hashCode(email, code);
  user.verificationSentAt = now;
  await sendVerificationEmail(user, code);
  return user;
}

async function verifyAccessCode(request, response) {
  const setup = accessConfigStatus();
  if (!setup.configured) return sendJson(request, response, 503, { error: "setup-required", setup });
  const session = readSignedCookie(request, cookieName);
  if (!session?.email) {
    return sendJson(request, response, 401, { error: "not_signed_in" });
  }
  const body = await readJson(request);
  const code = String(body.code || "").trim();
  const store = await loadUserStore();
  const user = findUser(store, session.email);
  if (!user || user.accessRequestStatus !== "Pending Verification") {
    return sendJson(request, response, 400, { error: "no_pending_verification" });
  }
  if (!code || hashCode(user.email, code) !== user.verificationCodeHash) {
    return sendJson(request, response, 400, { error: "invalid_code" });
  }
  user.emailVerified = true;
  user.accessRequestStatus = "Pending Approval";
  user.verificationCodeHash = "";
  await saveUserStore(store);
  await sendAdminNotification(user).catch((error) => console.warn(error));
  return sendJson(request, response, 200, { ok: true, user: publicUser(user) });
}

async function listAccessRequests(request, response) {
  const admin = await requireAdmin(request, response);
  if (!admin) return;
  const store = await loadUserStore();
  return sendJson(request, response, 200, {
    requests: store.users
      .filter((user) =>
        ["Pending Verification", "Pending Approval"].includes(user.accessRequestStatus),
      )
      .map(publicUser),
  });
}

async function decideAccessRequest(request, response, encodedEmail, decision) {
  const admin = await requireAdmin(request, response);
  if (!admin) return;
  const email = decodeURIComponent(encodedEmail).toLowerCase();
  const body = await readJson(request).catch(() => ({}));
  const store = await loadUserStore();
  const user = findUser(store, email);
  if (!user) return sendJson(request, response, 404, { error: "not_found" });
  if (decision === "approve") {
    if (user.accessRequestStatus !== "Pending Approval" || !user.emailVerified) {
      return sendJson(request, response, 400, {
        error: "email_not_verified",
        message: "The user must confirm the email code before approval.",
      });
    }
    user.active = true;
    user.accessRequestStatus = "Approved";
    user.approvedAt = new Date().toISOString();
    user.approvedBy = admin.email;
    user.rejectionReason = "";
  } else {
    user.active = false;
    user.accessRequestStatus = "Rejected";
    user.approvedAt = new Date().toISOString();
    user.approvedBy = admin.email;
    user.rejectionReason = String(body.reason || "Request rejected by admin.");
  }
  await saveUserStore(store);
  return sendJson(request, response, 200, { ok: true, user: publicUser(user) });
}

async function requireAdmin(request, response) {
  const session = readSignedCookie(request, cookieName);
  if (!session?.email) {
    sendJson(request, response, 401, { error: "not_signed_in" });
    return null;
  }
  const store = await loadUserStore();
  const user = findUser(store, session.email);
  if (!user || !user.active || user.accessRequestStatus !== "Approved") {
    sendJson(request, response, 403, { error: "not_approved" });
    return null;
  }
  if (user.role !== "Admin") {
    sendJson(request, response, 403, { error: "admin_required" });
    return null;
  }
  return user;
}

async function exchangeCodeForToken(code, codeVerifier) {
  const response = await fetch(`${authority}/token`, {
    method: "POST",
    headers: { "content-type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: config.clientId,
      client_secret: config.clientSecret,
      grant_type: "authorization_code",
      code,
      redirect_uri: redirectUri,
      code_verifier: codeVerifier,
      scope: "openid profile email User.Read",
    }),
  });
  if (!response.ok) {
    throw new Error(`Microsoft token exchange failed: ${await response.text()}`);
  }
  return response.json();
}

async function validateIdToken(idToken, nonce) {
  if (!jwks) throw new Error("Microsoft JWKS is not configured.");
  const result = await jwtVerify(idToken, jwks, {
    audience: config.clientId,
    issuer: `https://login.microsoftonline.com/${config.tenantId}/v2.0`,
  });
  if (result.payload.nonce !== nonce) {
    throw new Error("Microsoft sign-in nonce did not match.");
  }
  return result.payload;
}

async function sendVerificationEmail(user, code) {
  const subject = "Audit Assignment Tracker verification code";
  const body = [
    `Hello ${user.fullName},`,
    "",
    "Use this verification code to finish your Audit Assignment Tracker access request:",
    "",
    code,
    "",
    "After you confirm the code, an administrator must approve your profile before you can access the tracker.",
  ].join("\n");
  await sendGraphMail(user.email, subject, body);
}

async function sendAdminNotification(user) {
  if (config.adminEmails.length === 0) return;
  await sendGraphMail(
    config.adminEmails,
    "Tracker account request ready for approval",
    `${user.fullName} (${user.email}) confirmed their email and is waiting for admin approval.`,
  );
}

async function sendGraphMail(to, subject, content) {
  const recipients = (Array.isArray(to) ? to : [to]).map((address) => ({
    emailAddress: { address },
  }));
  const token = await getGraphAppToken();
  const response = await fetch(
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(config.mailFrom)}/sendMail`,
    {
      method: "POST",
      headers: {
        authorization: `Bearer ${token}`,
        "content-type": "application/json",
      },
      body: JSON.stringify({
        message: {
          subject,
          body: { contentType: "Text", content },
          toRecipients: recipients,
        },
        saveToSentItems: true,
      }),
    },
  );
  if (!response.ok) {
    throw new Error(`Graph sendMail failed: ${response.status} ${await response.text()}`);
  }
}

async function getGraphAppToken() {
  const response = await fetch(`https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "content-type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: config.clientId,
      client_secret: config.clientSecret,
      grant_type: "client_credentials",
      scope: "https://graph.microsoft.com/.default",
    }),
  });
  if (!response.ok) {
    throw new Error(`Graph token request failed: ${await response.text()}`);
  }
  const token = await response.json();
  return token.access_token;
}

async function loadUserStore() {
  await mkdir(dataDir, { recursive: true });
  if (!existsSync(usersFile)) {
    const store = { users: [] };
    seedAdmins(store);
    await saveUserStore(store);
    return store;
  }
  const store = JSON.parse(await readFile(usersFile, "utf8"));
  if (!Array.isArray(store.users)) store.users = [];
  seedAdmins(store);
  return store;
}

async function saveUserStore(store) {
  await mkdir(dataDir, { recursive: true });
  await writeFile(usersFile, `${JSON.stringify(store, null, 2)}\n`);
}

function seedAdmins(store) {
  for (const email of config.adminEmails) {
    if (!email || findUser(store, email)) continue;
    store.users.push({
      email,
      username: usernameFromEmail(email),
      fullName: nameFromEmail(email),
      role: "Admin",
      permissionGroup: "Admin",
      active: true,
      defaultVisibility: "All Projects",
      emailVerified: true,
      accessRequestStatus: "Approved",
      requestedAt: new Date().toISOString(),
      approvedAt: new Date().toISOString(),
      approvedBy: "server-config",
      rejectionReason: "",
      verificationCodeHash: "",
      verificationSentAt: "",
    });
  }
}

async function serveStatic(request, response, url) {
  if (request.method !== "GET") {
    response.writeHead(404);
    response.end();
    return;
  }
  const requestedPath = url.pathname === "/" ? "/index.html" : url.pathname;
  const filePath = normalize(resolve(join(distDir, requestedPath)));
  const safePath = filePath.startsWith(distDir) ? filePath : join(distDir, "index.html");
  const finalPath = existsSync(safePath) ? safePath : join(distDir, "index.html");
  if (!existsSync(finalPath)) {
    response.writeHead(200, { "content-type": "text/html; charset=utf-8" });
    response.end("<h1>Tracker server running</h1><p>Run npm run build to serve the app.</p>");
    return;
  }
  response.writeHead(200, { "content-type": contentType(finalPath) });
  response.end(await readFile(finalPath));
}

function publicUser(user) {
  return {
    email: user.email,
    username: user.username,
    fullName: user.fullName,
    role: user.role,
    permissionGroup: user.permissionGroup,
    active: user.active,
    defaultVisibility: user.defaultVisibility,
    emailVerified: user.emailVerified,
    accessRequestStatus: user.accessRequestStatus,
    requestedAt: user.requestedAt,
    approvedAt: user.approvedAt,
    approvedBy: user.approvedBy,
    rejectionReason: user.rejectionReason,
  };
}

function setSession(response, session) {
  response.setHeader("Set-Cookie", signedCookie(cookieName, session, 7 * 24 * 60 * 60));
}

function clearSession(response) {
  response.setHeader("Set-Cookie", clearCookie(cookieName));
}

function signedCookie(name, value, maxAge) {
  const payload = base64Url(Buffer.from(JSON.stringify(value)));
  const signature = sign(payload);
  return `${name}=${payload}.${signature}; Path=/; HttpOnly; SameSite=Lax; Max-Age=${maxAge}`;
}

function clearCookie(name) {
  return `${name}=; Path=/; HttpOnly; SameSite=Lax; Max-Age=0`;
}

function readSignedCookie(request, name) {
  const cookies = parseCookies(request.headers.cookie || "");
  const raw = cookies[name];
  if (!raw) return null;
  const [payload, signature] = raw.split(".");
  if (!payload || !signature || sign(payload) !== signature) return null;
  try {
    return JSON.parse(Buffer.from(payload, "base64url").toString("utf8"));
  } catch {
    return null;
  }
}

function sign(value) {
  return base64Url(createHmac("sha256", config.sessionSecret || "dev").update(value).digest());
}

function hashCode(email, code) {
  return createHmac("sha256", config.sessionSecret).update(`${email}:${code}`).digest("hex");
}

function claimEmail(claims) {
  return String(claims.preferred_username || claims.email || claims.upn || "").toLowerCase();
}

function isAllowedEmail(email) {
  const domain = email.split("@")[1] || "";
  return config.allowedDomains.includes(domain.toLowerCase());
}

function findUser(store, email) {
  return store.users.find((user) => user.email.toLowerCase() === email.toLowerCase());
}

function usernameFromEmail(email) {
  return email
    .split("@")[0]
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ".")
    .replace(/^\.+|\.+$/g, "");
}

function nameFromEmail(email) {
  return usernameFromEmail(email)
    .split(".")
    .filter(Boolean)
    .map((part) => `${part[0]?.toUpperCase() ?? ""}${part.slice(1)}`)
    .join(" ");
}

function randomCode() {
  return String(Math.floor(100000 + Math.random() * 900000));
}

function base64Url(input) {
  return Buffer.from(input).toString("base64url");
}

function splitCsv(value) {
  return value
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean);
}

async function readJson(request) {
  const chunks = [];
  for await (const chunk of request) chunks.push(chunk);
  if (chunks.length === 0) return {};
  return JSON.parse(Buffer.concat(chunks).toString("utf8"));
}

function sendSetupError(response, setup) {
  response.writeHead(503, { "content-type": "text/html; charset=utf-8" });
  response.end(
    `<h1>Tracker secure access is not configured</h1><p>Missing: ${setup.missing.join(", ")}</p><p>Redirect URI to register: ${setup.redirectUri}</p>`,
  );
}

function sendJson(request, response, status, payload) {
  setCors(request, response);
  response.writeHead(status, { "content-type": "application/json; charset=utf-8" });
  response.end(JSON.stringify(payload));
}

function setCors(request, response) {
  const origin = request.headers.origin;
  if (origin === config.frontendOrigin) {
    response.setHeader("Access-Control-Allow-Origin", origin);
    response.setHeader("Access-Control-Allow-Credentials", "true");
    response.setHeader("Vary", "Origin");
  }
  response.setHeader("Access-Control-Allow-Headers", "content-type");
  response.setHeader("Access-Control-Allow-Methods", "GET,POST,OPTIONS");
}

function redirect(response, location) {
  response.writeHead(302, { location });
  response.end();
}

function statusSlug(status) {
  return status.toLowerCase().replace(/\s+/g, "-");
}

function parseCookies(header) {
  const cookies = {};
  for (const part of header.split(";")) {
    const [key, ...value] = part.trim().split("=");
    if (key) cookies[key] = value.join("=");
  }
  return cookies;
}

function contentType(filePath) {
  const types = {
    ".css": "text/css; charset=utf-8",
    ".html": "text/html; charset=utf-8",
    ".js": "text/javascript; charset=utf-8",
    ".json": "application/json; charset=utf-8",
    ".png": "image/png",
    ".svg": "image/svg+xml",
  };
  return types[extname(filePath)] || "application/octet-stream";
}
