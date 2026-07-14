import { createHmac, createHash, randomBytes } from "node:crypto";
import { existsSync, readFileSync } from "node:fs";
import { readFile } from "node:fs/promises";
import { createServer } from "node:http";
import { extname, join, normalize, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { createRemoteJWKSet, jwtVerify } from "jose";
import {
  accessUserStoreStatus,
  createAccessUserStore,
} from "./accessUserStore.mjs";
import {
  contactWorkbookStatus,
  parseWorkbookLinks,
  readLinkedContactWorkbooks,
} from "./contactWorkbookSources.mjs";
import { createProjectStore, projectStoreStatus } from "./projectStore.mjs";

const rootDir = resolve(fileURLToPath(new URL("..", import.meta.url)));
const dataDir = resolve(rootDir, "server", "data");
const usersFile = resolve(dataDir, "access-users.json");
const projectsFile = resolve(dataDir, "audit-projects.json");
const distDir = resolve(rootDir, "dist");
const deployInfoFile = resolve(rootDir, "server", "deploy-info.json");
const oauthStates = new Map();
const cookieName = "tracker_session";
const oauthCookieName = "tracker_oauth_state";
const sessionSecretMinLength = 32;
const requiredGraphRoles = ["Mail.Send", "Sites.ReadWrite.All", "Calendars.ReadWrite"];
const weakSessionSecrets = new Set([
  "dev",
  "development",
  "secret",
  "password",
  "changeme",
  "test",
]);

const localEnvironment = loadLocalEnvironment();

const inferredPublicOrigin = process.env.WEBSITE_HOSTNAME
  ? `https://${process.env.WEBSITE_HOSTNAME}`
  : "http://localhost:8787";
const publicOrigin = process.env.TRACKER_PUBLIC_ORIGIN || inferredPublicOrigin;
const listenTarget =
  process.env.PORT || process.env.TRACKER_SERVER_PORT || "8787";

const config = {
  port: /^\d+$/.test(listenTarget) ? Number(listenTarget) : listenTarget,
  frontendOrigin: process.env.TRACKER_FRONTEND_ORIGIN || publicOrigin,
  publicOrigin,
  tenantId: process.env.MICROSOFT_TENANT_ID || "",
  clientId: process.env.MICROSOFT_CLIENT_ID || "",
  clientSecret: process.env.MICROSOFT_CLIENT_SECRET || "",
  sessionSecret: process.env.TRACKER_SESSION_SECRET || "",
  allowedDomains: splitCsv(process.env.TRACKER_ALLOWED_EMAIL_DOMAINS || ""),
  adminEmails: splitCsv(process.env.TRACKER_ADMIN_EMAILS || "").map((email) =>
    email.toLowerCase(),
  ),
  mailFrom: process.env.MICROSOFT_MAIL_FROM || "",
  userStoreMode: process.env.TRACKER_USER_STORE || "local",
  userStoreSiteId: process.env.TRACKER_USERS_SITE_ID || "",
  userStoreListId: process.env.TRACKER_USERS_LIST_ID || "",
  projectStoreMode: process.env.TRACKER_PROJECT_STORE || "local",
  projectStoreSiteId: process.env.TRACKER_PROJECTS_SITE_ID || "",
  projectStoreListId: process.env.TRACKER_PROJECTS_LIST_ID || "",
  contactWorkbookLinksRaw: process.env.TRACKER_CONTACT_WORKBOOK_LINKS || "",
};

const authority = `https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0`;
const userStoreStatus = accessUserStoreStatus(config);
const currentProjectStoreStatus = projectStoreStatus(config);
const currentContactWorkbookStatus = contactWorkbookStatus(config);
const accessUserStore = createAccessUserStore({
  mode: userStoreStatus.mode,
  usersFile,
  graph: {
    siteId: config.userStoreSiteId,
    listId: config.userStoreListId,
    getAccessToken: getGraphAppToken,
  },
});
const projectStore = createProjectStore({
  mode: currentProjectStoreStatus.configured ? currentProjectStoreStatus.mode : "local",
  projectsFile,
  graph: {
    siteId: config.projectStoreSiteId,
    listId: config.projectStoreListId,
    getAccessToken: getGraphAppToken,
  },
});
const jwks = hasConfiguredValue(config.tenantId)
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
    return sendJson(request, response, 200, accessConfigStatus(request));
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
  if (url.pathname === "/api/admin/system-health") {
    return systemHealth(request, response);
  }
  if (url.pathname === "/api/contact-sources" || url.pathname === "/api/admin/contact-sources") {
    return contactSources(request, response);
  }
  if (url.pathname === "/api/calendar/project-event" && request.method === "POST") {
    return createProjectCalendarEvent(request, response);
  }
  if (url.pathname === "/api/projects" && request.method === "GET") {
    return listProjects(request, response);
  }
  if (url.pathname === "/api/projects" && request.method === "PUT") {
    return saveProjects(request, response);
  }
  const approvalMatch = url.pathname.match(
    /^\/api\/admin\/access-requests\/([^/]+)\/(approve|reject)$/,
  );
  if (approvalMatch && request.method === "POST") {
    return decideAccessRequest(request, response, approvalMatch[1], approvalMatch[2]);
  }
  const userUpdateMatch = url.pathname.match(/^\/api\/admin\/users\/([^/]+)$/);
  if (userUpdateMatch && request.method === "PATCH") {
    return updateManagedUser(request, response, userUpdateMatch[1]);
  }
  return serveStatic(request, response, url);
}

function accessConfigStatus(request) {
  const requestOrigin = request ? requestPublicOrigin(request) : config.publicOrigin;
  const frontendOrigin = request ? requestFrontendOrigin(request, requestOrigin) : config.frontendOrigin;
  const missing = [];
  const invalid = [];
  for (const [key, value] of Object.entries({
    MICROSOFT_TENANT_ID: config.tenantId,
    MICROSOFT_CLIENT_ID: config.clientId,
    MICROSOFT_CLIENT_SECRET: config.clientSecret,
    MICROSOFT_MAIL_FROM: config.mailFrom,
    TRACKER_SESSION_SECRET: config.sessionSecret,
    TRACKER_ALLOWED_EMAIL_DOMAINS: config.allowedDomains.join(","),
    TRACKER_ADMIN_EMAILS: config.adminEmails.join(","),
  })) {
    if (!hasConfiguredValue(value)) missing.push(key);
  }
  const sessionSecretStatus = validateSessionSecret(config.sessionSecret);
  if (sessionSecretStatus === "missing" && !missing.includes("TRACKER_SESSION_SECRET")) {
    missing.push("TRACKER_SESSION_SECRET");
  }
  if (sessionSecretStatus === "invalid") invalid.push("TRACKER_SESSION_SECRET");
  missing.push(...userStoreStatus.missing);
  return {
    configured: missing.length === 0 && invalid.length === 0,
    missing,
    invalid,
    redirectUri: `${requestOrigin}/api/auth/callback`,
    frontendOrigin,
    userStore: userStoreStatus,
    projectStore: currentProjectStoreStatus,
    contactSources: currentContactWorkbookStatus,
  };
}

function sendSetupRequired(request, response) {
  const setup = accessConfigStatus(request);
  if (setup.configured) return false;
  sendJson(request, response, 503, { error: "setup-required", setup });
  return true;
}

function loadLocalEnvironment() {
  const candidates = [
    resolve(rootDir, ".env"),
    resolve(rootDir, ".env.local"),
    resolve(rootDir, "server.env"),
  ];
  if (process.env.HOME) {
    candidates.push(resolve(process.env.HOME, "data", "tracker-server.env"));
  }
  const loadedFiles = [];
  const keysFromLocalFiles = new Set();
  for (const envPath of candidates) {
    if (!existsSync(envPath)) continue;
    loadedFiles.push(envPath);
    const contents = readFileSync(envPath, "utf8");
    for (const rawLine of contents.split(/\r?\n/)) {
      const line = rawLine.trim();
      if (!line || line.startsWith("#")) continue;
      const separatorIndex = line.indexOf("=");
      if (separatorIndex < 1) continue;
      const key = line.slice(0, separatorIndex).trim();
      let value = line.slice(separatorIndex + 1).trim();
      if (
        (value.startsWith('"') && value.endsWith('"')) ||
        (value.startsWith("'") && value.endsWith("'"))
      ) {
        value = value.slice(1, -1);
      }
      if (process.env[key] === undefined) {
        process.env[key] = value;
        keysFromLocalFiles.add(key);
      }
    }
  }
  return {
    loadedFiles,
    keysFromLocalFiles,
  };
}

function hasConfiguredValue(value) {
  const normalized = String(value ?? "").trim().toLowerCase();
  if (!normalized) return false;
  if (normalized === "00000000-0000-0000-0000-000000000000") return false;
  if (normalized.startsWith("replace-with-")) return false;
  return ![
    "tracker-mailbox@yourcompany.com",
    "yourcompany.com",
    "your.email@yourcompany.com",
  ].includes(normalized);
}

function hasValidSessionSecret(value) {
  const secret = String(value ?? "").trim();
  const normalized = secret.toLowerCase();
  return (
    hasConfiguredValue(secret) &&
    secret.length >= sessionSecretMinLength &&
    !weakSessionSecrets.has(normalized)
  );
}

function validateSessionSecret(value) {
  const secret = String(value ?? "").trim();
  if (!hasConfiguredValue(secret)) return "missing";
  return hasValidSessionSecret(secret) ? "valid" : "invalid";
}

async function currentAccessState(request) {
  const setup = accessConfigStatus(request);
  const publicOrigin = requestPublicOrigin(request);
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
      signInUrl: `${publicOrigin}/api/auth/start?mode=signin`,
      requestAccessUrl: `${publicOrigin}/api/auth/start?mode=request`,
    };
  }

  const store = await loadUserStore();
  const user = findUser(store, session.email);
  if (!user) {
    return {
      configured: true,
      authenticated: false,
      status: "not-requested",
      signInUrl: `${publicOrigin}/api/auth/start?mode=signin`,
      requestAccessUrl: `${publicOrigin}/api/auth/start?mode=request`,
    };
  }

  const state = {
    configured: true,
    authenticated: user.active && user.accessRequestStatus === "Approved",
    status: statusSlug(user.accessRequestStatus),
    user: publicUser(user),
    signInUrl: `${publicOrigin}/api/auth/start?mode=signin`,
    requestAccessUrl: `${publicOrigin}/api/auth/start?mode=request`,
  };

  if (state.authenticated && user.role === "Admin") {
    state.managedUsers = store.users
      .map(publicUser)
      .sort((a, b) => a.fullName.localeCompare(b.fullName));
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

async function startMicrosoftAuth(request, response, url) {
  const setup = accessConfigStatus(request);
  if (!setup.configured) return sendSetupError(response, setup);

  const publicOrigin = requestPublicOrigin(request);
  const frontendOrigin = requestFrontendOrigin(request, publicOrigin);
  const redirectUri = `${publicOrigin}/api/auth/callback`;
  const mode = url.searchParams.get("mode") === "request" ? "request" : "signin";
  const state = base64Url(randomBytes(32));
  const nonce = base64Url(randomBytes(32));
  const codeVerifier = base64Url(randomBytes(64));
  const codeChallenge = base64Url(createHash("sha256").update(codeVerifier).digest());
  oauthStates.set(state, {
    mode,
    nonce,
    codeVerifier,
    frontendOrigin,
    redirectUri,
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
  const setup = accessConfigStatus(request);
  if (!setup.configured) return sendSetupError(response, setup);

  const error = url.searchParams.get("error");
  const fallbackFrontendOrigin = requestFrontendOrigin(request, requestPublicOrigin(request));
  if (error) return redirect(response, `${fallbackFrontendOrigin}/?auth=${error}`);

  const state = url.searchParams.get("state") ?? "";
  const code = url.searchParams.get("code") ?? "";
  const stateCookie = readSignedCookie(request, oauthCookieName);
  const oauthState = oauthStates.get(state);
  const frontendOrigin = oauthState?.frontendOrigin || fallbackFrontendOrigin;
  oauthStates.delete(state);
  response.setHeader("Set-Cookie", clearCookie(oauthCookieName));

  if (!code || !stateCookie || stateCookie.state !== state || !oauthState) {
    return redirect(response, `${frontendOrigin}/?auth=invalid-state`);
  }
  if (oauthState.expiresAt < Date.now()) {
    return redirect(response, `${frontendOrigin}/?auth=expired-state`);
  }

  const token = await exchangeCodeForToken(code, oauthState.codeVerifier, oauthState.redirectUri);
  const claims = await validateIdToken(token.id_token, oauthState.nonce);
  const email = claimEmail(claims);
  const name = String(claims.name || email.split("@")[0]);
  if (!isAllowedEmail(email)) {
    return redirect(response, `${frontendOrigin}/?auth=domain-blocked`);
  }

  const store = await loadUserStore();
  seedAdmins(store);
  let user = findUser(store, email);

  if (oauthState.mode === "request") {
    user = await createOrRefreshAccessRequest(store, email, name);
    await saveUserStore(store);
    if (user.accessRequestStatus === "Approved" && user.active) {
      setSession(response, { email, status: "approved" });
      return redirect(response, `${frontendOrigin}/?auth=success`);
    }
    setSession(response, { email, status: "pending" });
    return redirect(response, `${frontendOrigin}/?auth=code-sent`);
  }

  if (!user) {
    return redirect(response, `${frontendOrigin}/?auth=request-required`);
  }
  if (user.accessRequestStatus !== "Approved" || !user.active) {
    setSession(response, { email, status: "pending" });
    return redirect(response, `${frontendOrigin}/?auth=pending`);
  }

  setSession(response, { email, status: "approved" });
  return redirect(response, `${frontendOrigin}/?auth=success`);
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
  const setup = accessConfigStatus(request);
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
    const role = normalizeRole(body.role, "Auditor");
    const defaultVisibility = normalizeVisibility(
      body.defaultVisibility,
      role === "Finance" ? "Finance Records" : "Role Default",
    );
    user.role = role;
    user.permissionGroup = role;
    user.defaultVisibility = defaultVisibility;
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

async function updateManagedUser(request, response, encodedEmail) {
  const admin = await requireAdmin(request, response);
  if (!admin) return;
  const email = decodeURIComponent(encodedEmail).toLowerCase();
  const body = await readJson(request).catch(() => ({}));
  const store = await loadUserStore();
  const user = findUser(store, email);
  if (!user) return sendJson(request, response, 404, { error: "not_found" });
  if (user.accessRequestStatus !== "Approved") {
    return sendJson(request, response, 400, {
      error: "not_approved",
      message: "Only approved Microsoft accounts can be managed here.",
    });
  }

  const role = normalizeRole(body.role, user.role || "Auditor");
  user.role = role;
  user.permissionGroup = role;
  user.defaultVisibility = normalizeVisibility(
    body.defaultVisibility,
    user.defaultVisibility || (role === "Finance" ? "Finance Records" : "Role Default"),
  );
  if (typeof body.active === "boolean") user.active = body.active;
  if (typeof body.fullName === "string" && body.fullName.trim()) {
    user.fullName = body.fullName.trim();
  }
  await saveUserStore(store);
  return sendJson(request, response, 200, { ok: true, user: publicUser(user) });
}

async function systemHealth(request, response) {
  const admin = await requireAdmin(request, response);
  if (!admin) return;
  const setup = accessConfigStatus(request);
  const graphApp = await graphAppHealth();
  const runtime = runtimeHealth();
  const deployment = deploymentInfo();
  const approvalStore = {
    mode: userStoreStatus.mode,
    configured: userStoreStatus.configured,
    missing: userStoreStatus.missing,
    durable: userStoreStatus.mode === "microsoft-lists" && userStoreStatus.configured,
    status:
      userStoreStatus.mode === "microsoft-lists" && userStoreStatus.configured
        ? "Approval records will be stored in Microsoft Lists."
        : userStoreStatus.mode === "microsoft-lists"
          ? `Approval storage is set to Microsoft Lists but missing: ${userStoreStatus.missing.join(", ")}.`
        : "Approval records are still stored on this local server.",
  };
  const projectStoreHealth = {
    mode: currentProjectStoreStatus.mode,
    configured: currentProjectStoreStatus.configured,
    missing: currentProjectStoreStatus.missing,
    durable:
      currentProjectStoreStatus.mode === "microsoft-lists" &&
      currentProjectStoreStatus.configured,
    status:
      currentProjectStoreStatus.mode === "microsoft-lists" &&
      currentProjectStoreStatus.configured
        ? "Project records will be stored in the Audit Assignments Microsoft List."
        : currentProjectStoreStatus.mode === "microsoft-lists"
          ? `Project storage is set to Microsoft Lists but missing: ${currentProjectStoreStatus.missing.join(", ")}.`
        : "Project records are still stored on this app server.",
  };
  const contactSources = {
    configured: currentContactWorkbookStatus.configured,
    count: currentContactWorkbookStatus.count,
    missing: currentContactWorkbookStatus.missing,
    status: currentContactWorkbookStatus.configured
      ? `${currentContactWorkbookStatus.count} linked contact workbook${
          currentContactWorkbookStatus.count === 1 ? "" : "s"
        } configured.`
      : "Linked contact workbooks are not configured.",
  };
  return sendJson(request, response, 200, {
    generatedAt: new Date().toISOString(),
    deployment,
    runtime,
    server: {
      configured: setup.configured,
      missing: setup.missing,
      redirectUri: setup.redirectUri,
      frontendOrigin: setup.frontendOrigin,
      publicOrigin: config.publicOrigin,
    },
    graphApp,
    mail: {
      configured: hasConfiguredValue(config.mailFrom),
      from: config.mailFrom,
      permissionGranted: graphApp.roles.includes("Mail.Send"),
    },
    sharePoint: {
      permissionGranted: graphApp.roles.includes("Sites.ReadWrite.All"),
      siteIdConfigured:
        hasConfiguredValue(config.userStoreSiteId) ||
        hasConfiguredValue(config.projectStoreSiteId),
      trackerUsersListIdConfigured: hasConfiguredValue(config.userStoreListId),
      auditAssignmentsListIdConfigured: hasConfiguredValue(config.projectStoreListId),
    },
    calendar: {
      permissionGranted: graphApp.roles.includes("Calendars.ReadWrite"),
      mode: "manual-project-event-sync",
      target: "signed-in-user-calendar",
    },
    approvalStore,
    projectStore: projectStoreHealth,
    contactSources,
    recommendations: healthRecommendations(
      setup,
      graphApp,
      approvalStore,
      projectStoreHealth,
      runtime,
      contactSources,
    ),
  });
}

function deploymentInfo() {
  if (!existsSync(deployInfoFile)) {
    return {
      source: "local",
      commit: "",
      branch: "",
      deployedAt: "",
      workflowRun: "",
    };
  }
  try {
    const parsed = JSON.parse(readFileSync(deployInfoFile, "utf8"));
    return {
      source: String(parsed.source || "github-actions"),
      commit: String(parsed.commit || ""),
      branch: String(parsed.branch || ""),
      deployedAt: String(parsed.deployedAt || ""),
      workflowRun: String(parsed.workflowRun || ""),
    };
  } catch {
    return {
      source: "unknown",
      commit: "",
      branch: "",
      deployedAt: "",
      workflowRun: "",
    };
  }
}

function runtimeHealth() {
  const requiredKeys = [
    "MICROSOFT_TENANT_ID",
    "MICROSOFT_CLIENT_ID",
    "MICROSOFT_CLIENT_SECRET",
    "MICROSOFT_MAIL_FROM",
    "TRACKER_SESSION_SECRET",
    "TRACKER_ALLOWED_EMAIL_DOMAINS",
    "TRACKER_ADMIN_EMAILS",
  ];
  const localFileKeyCount = requiredKeys.filter((key) =>
    localEnvironment.keysFromLocalFiles.has(key),
  ).length;
  const appSettingKeyCount = requiredKeys.filter(
    (key) =>
      process.env[key] !== undefined && !localEnvironment.keysFromLocalFiles.has(key),
  ).length;
  const persistentDataEnvLoaded = localEnvironment.loadedFiles.some((file) =>
    normalize(file).endsWith(normalize("data/tracker-server.env")),
  );
  return {
    nodeVersion: process.version,
    websiteHostname: process.env.WEBSITE_HOSTNAME || "",
    configSource:
      localFileKeyCount === 0
        ? "Azure App Service environment variables"
        : persistentDataEnvLoaded
          ? "Persistent Azure data env file"
          : "Local env file",
    appSettingKeyCount,
    localFileKeyCount,
    persistentDataEnvLoaded,
    loadedEnvFileCount: localEnvironment.loadedFiles.length,
  };
}

async function graphAppHealth() {
  if (
    !hasConfiguredValue(config.tenantId) ||
    !hasConfiguredValue(config.clientId) ||
    !hasConfiguredValue(config.clientSecret)
  ) {
    return {
      tokenAvailable: false,
      roles: [],
      missingRoles: requiredGraphRoles,
      error: "Microsoft tenant ID, client ID, or client secret is missing.",
    };
  }
  try {
    const token = await getGraphAppToken();
    const payload = decodeJwtPayload(token);
    const roles = Array.isArray(payload.roles) ? payload.roles.map(String) : [];
    return {
      tokenAvailable: true,
      roles,
      missingRoles: requiredGraphRoles.filter(
        (role) => !roles.includes(role),
      ),
      appDisplayName: String(payload.app_displayname || ""),
    };
  } catch (error) {
    return {
      tokenAvailable: false,
      roles: [],
      missingRoles: requiredGraphRoles,
      error: error instanceof Error ? error.message : "Graph app token failed.",
    };
  }
}

function healthRecommendations(
  setup,
  graphApp,
  approvalStore,
  projectStoreHealth,
  runtime,
  contactSources,
) {
  const recommendations = [];
  if (!setup.configured) {
    recommendations.push(`Complete missing server configuration: ${setup.missing.join(", ")}.`);
  }
  if (graphApp.missingRoles.length > 0) {
    recommendations.push(
      `Grant admin consent for Microsoft Graph application permissions: ${graphApp.missingRoles.join(", ")}.`,
    );
  }
  if (!graphApp.tokenAvailable) {
    recommendations.push("Verify the Microsoft client secret is current and has not expired.");
  }
  if (runtime.localFileKeyCount > 0) {
    recommendations.push(
      "Prefer Azure App Service Environment variables once portal or CLI access is available.",
    );
  }
  if (!approvalStore.durable) {
    recommendations.push(
      "Switch TRACKER_USER_STORE to microsoft-lists after admin consent and Tracker Users list IDs are ready.",
    );
  }
  if (!projectStoreHealth.durable) {
    recommendations.push(
      "Switch TRACKER_PROJECT_STORE to microsoft-lists after the Audit Assignments list ID is ready.",
    );
  }
  if (recommendations.length === 0) {
    recommendations.push(
      "Secure access, mail, approval storage, and project storage are ready.",
    );
  }
  return recommendations;
}

async function listProjects(request, response) {
  const user = await requireApprovedUser(request, response);
  if (!user) return;
  const store = await loadProjectStore();
  return sendJson(request, response, 200, {
    projects: store.projects.filter((project) => canViewProjectForUser(user, project)),
  });
}

async function saveProjects(request, response) {
  const user = await requireApprovedUser(request, response);
  if (!user) return;
  const body = await readJson(request).catch(() => ({}));
  const incomingProjects = Array.isArray(body.projects)
    ? body.projects.map(normalizeProjectRecord).filter((project) => project.id)
    : [];
  if (!Array.isArray(body.projects)) {
    return sendJson(request, response, 400, {
      error: "invalid_projects",
      message: "Project payload must include a projects array.",
    });
  }

  const fullReplace = body.mode === "replace-all";
  if (fullReplace && !hasFullProjectAccess(user)) {
    return sendJson(request, response, 403, {
      error: "project_admin_required",
      message: "Only admins and audit managers can replace all project records.",
    });
  }

  const store = await loadProjectStore();
  if (fullReplace) {
    store.projects = incomingProjects;
  } else {
    const existingById = new Map(store.projects.map((project) => [project.id, project]));
    for (const project of incomingProjects) {
      const existing = existingById.get(project.id);
      const validationError = validateProjectWriteForUser(user, existing, project);
      if (validationError) {
        return sendJson(request, response, 403, validationError);
      }
    }
    const incomingById = new Map(incomingProjects.map((project) => [project.id, project]));
    const merged = store.projects.map((project) =>
      incomingById.has(project.id) ? incomingById.get(project.id) : project,
    );
    for (const project of incomingProjects) {
      if (!existingById.has(project.id)) merged.unshift(project);
    }
    store.projects = merged;
  }

  await saveProjectStore(store);
  return sendJson(request, response, 200, {
    ok: true,
    projects: store.projects.filter((project) => canViewProjectForUser(user, project)),
  });
}

async function contactSources(request, response) {
  const user = await requireApprovedUser(request, response);
  if (!user) return;
  const sources = parseWorkbookLinks(config.contactWorkbookLinksRaw);
  if (sources.length === 0) {
    return sendJson(request, response, 200, {
      generatedAt: new Date().toISOString(),
      configured: false,
      sources: [],
      contacts: [],
      warnings: [
        {
          sourceId: "configuration",
          worksheetName: "",
          message: "TRACKER_CONTACT_WORKBOOK_LINKS is not configured.",
        },
      ],
    });
  }
  const result = await readLinkedContactWorkbooks({
    sources,
    getAccessToken: getGraphAppToken,
  });
  return sendJson(request, response, 200, {
    configured: true,
    ...result,
  });
}

async function createProjectCalendarEvent(request, response) {
  const user = await requireApprovedUser(request, response);
  if (!user) return;
  const body = await readJson(request).catch(() => ({}));
  const projectId = String(body.projectId || "").trim();
  if (!projectId) {
    return sendJson(request, response, 400, {
      error: "missing_project_id",
      message: "Project ID is required before creating an Outlook calendar event.",
    });
  }
  const graphApp = await graphAppHealth();
  if (!graphApp.roles.includes("Calendars.ReadWrite")) {
    return sendJson(request, response, 403, {
      error: "calendar_permission_missing",
      message: "Grant Calendars.ReadWrite application permission before Outlook calendar sync can create events.",
      missingRoles: graphApp.missingRoles,
    });
  }
  const store = await loadProjectStore();
  const project = store.projects.find((item) => item.id === projectId);
  if (!project || !canViewProjectForUser(user, project)) {
    return sendJson(request, response, 404, {
      error: "project_not_found",
      message: "The project was not found or is not visible to your account.",
    });
  }
  if (!canEditProjectForUser(user, project)) {
    return sendJson(request, response, 403, {
      error: "calendar_sync_denied",
      message: "Your role cannot create a calendar event for this project.",
    });
  }
  const startDate = String(body.confirmedAuditDate || project.confirmedAuditDate || project.dueDate || "").trim();
  if (!startDate) {
    return sendJson(request, response, 400, {
      error: "missing_schedule_date",
      message: "Add a confirmed audit date before creating an Outlook calendar event.",
    });
  }
  const durationHours = clampNumber(Number(body.durationHours || project.auditDurationHours || 1), 0.5, 12);
  const startTime = normalizeEventStartTime(body.startTime || project.auditStartTime || "09:00");
  const location = String(body.location || project.auditLocation || "").trim();
  const remoteLink = String(body.remoteLink || project.auditRemoteLink || "").trim();
  const attendeeEmails = Array.isArray(body.attendeeEmails)
    ? body.attendeeEmails.map(String).map((email) => email.trim().toLowerCase()).filter(isAllowedEmail)
    : [];
  let event;
  try {
    event = await createOutlookEventForProject(user, project, startDate, {
      durationHours,
      startTime,
      location,
      remoteLink,
      attendeeEmails,
    });
  } catch (error) {
    return sendJson(request, response, 502, {
      error: "calendar_event_failed",
      message: error instanceof Error ? error.message : "Outlook calendar event creation failed.",
    });
  }
  return sendJson(request, response, 200, {
    ok: true,
    action: event.action,
    event,
  });
}

async function loadProjectStore() {
  const store = await projectStore.load();
  return {
    projects: Array.isArray(store.projects)
      ? store.projects.map(normalizeProjectRecord).filter((project) => project.id)
      : [],
  };
}

async function saveProjectStore(store) {
  await projectStore.save({
    projects: (store.projects ?? []).map(normalizeProjectRecord).filter((project) => project.id),
  });
}

function normalizeProjectRecord(project) {
  const record = project && typeof project === "object" ? { ...project } : {};
  const auditTeam = Array.isArray(record.auditTeam) ? record.auditTeam : [];
  const assignedAuditor =
    String(record.assignedAuditor || "").trim() ||
    String(auditTeam.find((member) => member?.role === "Lead Auditor")?.person || "").trim() ||
    String(auditTeam[0]?.person || "").trim();
  return {
    ...record,
    id: String(record.id || ""),
    assignmentNumber: String(record.assignmentNumber || record.id || ""),
    assignedAuditor,
    auditTeam,
    currentStage: String(record.currentStage || "Intake"),
    schedulingNotes: String(record.schedulingNotes || ""),
    calendarSyncStatus: String(record.calendarSyncStatus || "Not Synced"),
    calendarEventId: String(record.calendarEventId || ""),
    calendarEventWebLink: String(record.calendarEventWebLink || ""),
    calendarEventLastSyncedAt: String(record.calendarEventLastSyncedAt || ""),
    auditLocation: String(record.auditLocation || ""),
    auditRemoteLink: String(record.auditRemoteLink || ""),
    auditDurationHours: Number(record.auditDurationHours || 1),
    auditStartTime: String(record.auditStartTime || "09:00"),
    linkedContactId: String(record.linkedContactId || ""),
    linkedContactSource: String(record.linkedContactSource || ""),
    invoiceStatus: String(record.invoiceStatus || "Not Started"),
    reportStatus: String(record.reportStatus || "Not Started"),
    paymentReceived: Boolean(record.paymentReceived),
    labels: Array.isArray(record.labels) ? record.labels : [],
    checklistCompletions:
      record.checklistCompletions && typeof record.checklistCompletions === "object"
        ? record.checklistCompletions
        : {},
    statusHistory: Array.isArray(record.statusHistory) ? record.statusHistory : [],
    comments: Array.isArray(record.comments) ? record.comments : [],
    activityEvents: Array.isArray(record.activityEvents) ? record.activityEvents : [],
  };
}

async function requireApprovedUser(request, response) {
  if (sendSetupRequired(request, response)) return null;
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
  return user;
}

async function requireAdmin(request, response) {
  const user = await requireApprovedUser(request, response);
  if (!user) return null;
  if (user.role !== "Admin") {
    sendJson(request, response, 403, { error: "admin_required" });
    return null;
  }
  return user;
}

async function exchangeCodeForToken(code, codeVerifier, redirectUri) {
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

async function createOutlookEventForProject(user, project, startDate, options = {}) {
  const token = await getGraphAppToken();
  const subject = `Audit: ${project.assignmentNumber || project.id} - ${project.auditEntity || "Assignment"}`;
  const startTime = normalizeEventStartTime(options.startTime || project.auditStartTime || "09:00");
  const start = `${startDate}T${startTime}:00`;
  const end = eventEndDateTime(startDate, startTime, options.durationHours || 1);
  const body = [
    `Assignment: ${project.assignmentNumber || project.id}`,
    `Audit entity: ${project.auditEntity || "Not set"}`,
    `Audit type: ${project.auditType || "Not set"}`,
    `Start time: ${startTime}`,
    `Duration: ${clampNumber(options.durationHours || 1, 0.5, 12)} hour(s)`,
    `Audit team: ${assignedAuditorNames(project).join(", ") || "Not set"}`,
    options.location ? `Location: ${options.location}` : "",
    options.remoteLink ? `Remote link: ${options.remoteLink}` : "",
    "",
    project.schedulingNotes ? `Scheduling notes:\n${project.schedulingNotes}` : "",
  ].filter(Boolean).join("\n");
  const payload = {
    subject,
    body: {
      contentType: "Text",
      content: body,
    },
    start: {
      dateTime: start,
      timeZone: "Eastern Standard Time",
    },
    end: {
      dateTime: end,
      timeZone: "Eastern Standard Time",
    },
    location: {
      displayName: options.location || options.remoteLink || "",
    },
    attendees: uniqueStrings(options.attendeeEmails || [])
      .filter((email) => email !== user.email)
      .map((email) => ({
        emailAddress: { address: email },
        type: "required",
      })),
    categories: ["Audit Assignment Tracker"],
    isReminderOn: true,
    reminderMinutesBeforeStart: 1440,
  };
  const existingEventId = String(project.calendarEventId || "").trim();
  const targetPath = existingEventId
    ? `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(user.email)}/events/${encodeURIComponent(existingEventId)}`
    : `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(user.email)}/events`;
  let response = await fetch(targetPath, {
    method: existingEventId ? "PATCH" : "POST",
    headers: {
      authorization: `Bearer ${token}`,
      "content-type": "application/json",
    },
    body: JSON.stringify(payload),
  });
  let action = existingEventId ? "updated" : "created";
  if (existingEventId && response.status === 404) {
    response = await fetch(
      `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(user.email)}/events`,
      {
        method: "POST",
        headers: {
          authorization: `Bearer ${token}`,
          "content-type": "application/json",
        },
        body: JSON.stringify(payload),
      },
    );
    action = "created";
  }
  const responsePayload = await response.json().catch(async () => ({
    raw: await response.text().catch(() => ""),
  }));
  if (!response.ok) {
    throw new Error(`Graph calendar event failed: ${response.status} ${JSON.stringify(responsePayload)}`);
  }
  return {
    action,
    id: String(responsePayload.id || ""),
    subject: String(responsePayload.subject || subject),
    webLink: String(responsePayload.webLink || ""),
    start: responsePayload.start ?? null,
    end: responsePayload.end ?? null,
  };
}

function eventEndDateTime(startDate, startTime, durationHours) {
  const [year, month, day] = String(startDate).split("-").map(Number);
  const [hour, minute] = normalizeEventStartTime(startTime).split(":").map(Number);
  const date = new Date(Date.UTC(year, month - 1, day, hour, minute, 0));
  date.setTime(date.getTime() + clampNumber(durationHours, 0.5, 12) * 60 * 60 * 1000);
  return `${String(date.getUTCFullYear()).padStart(4, "0")}-${String(date.getUTCMonth() + 1).padStart(2, "0")}-${String(date.getUTCDate()).padStart(2, "0")}T${String(date.getUTCHours()).padStart(2, "0")}:${String(date.getUTCMinutes()).padStart(2, "0")}:00`;
}

function normalizeEventStartTime(value) {
  const match = String(value || "").trim().match(/^(\d{1,2}):(\d{2})$/);
  if (!match) return "09:00";
  const hour = Math.min(Math.max(Number(match[1]), 0), 23);
  const minute = Math.min(Math.max(Number(match[2]), 0), 59);
  return `${String(hour).padStart(2, "0")}:${String(minute).padStart(2, "0")}`;
}

function clampNumber(value, min, max) {
  const number = Number.isFinite(value) ? value : min;
  return Math.min(Math.max(number, min), max);
}

function uniqueStrings(values) {
  return Array.from(new Set((values ?? []).map(String).map((value) => value.trim()).filter(Boolean)));
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
  const store = await accessUserStore.load();
  if (!Array.isArray(store.users)) store.users = [];
  if (seedAdmins(store)) {
    await saveUserStore(store);
  }
  return store;
}

async function saveUserStore(store) {
  seedAdmins(store);
  await accessUserStore.save(store);
}

function seedAdmins(store) {
  let changed = false;
  for (const email of config.adminEmails) {
    if (!email) continue;
    const existingUser = findUser(store, email);
    const adminFields = {
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
    };
    if (!existingUser) {
      store.users.push(adminFields);
      changed = true;
      continue;
    }
    const nextUser = {
      ...existingUser,
      role: "Admin",
      permissionGroup: "Admin",
      active: true,
      defaultVisibility: "All Projects",
      emailVerified: true,
      accessRequestStatus: "Approved",
      approvedAt: existingUser.approvedAt || new Date().toISOString(),
      approvedBy:
        existingUser.role === "Admin" && existingUser.approvedBy
          ? existingUser.approvedBy
          : "server-config",
      rejectionReason: "",
      verificationCodeHash: "",
      verificationSentAt: "",
    };
    if (JSON.stringify(existingUser) !== JSON.stringify(nextUser)) {
      Object.assign(existingUser, nextUser);
      changed = true;
    }
  }
  return changed;
}

function isConfiguredAdminEmail(email) {
  return config.adminEmails.includes(String(email || "").toLowerCase());
}

function enforceConfiguredAdmin(user) {
  if (!user || !isConfiguredAdminEmail(user.email)) return user;
  user.role = "Admin";
  user.permissionGroup = "Admin";
  user.active = true;
  user.defaultVisibility = "All Projects";
  user.emailVerified = true;
  user.accessRequestStatus = "Approved";
  user.rejectionReason = "";
  user.verificationCodeHash = "";
  user.verificationSentAt = "";
  if (!user.approvedAt) user.approvedAt = new Date().toISOString();
  if (!user.approvedBy || user.approvedBy !== "server-config") {
    user.approvedBy = "server-config";
  }
  return user;
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
  if (!signature) {
    throw new Error("TRACKER_SESSION_SECRET must be a strong, non-placeholder value.");
  }
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
  const expectedSignature = payload ? sign(payload) : null;
  if (!payload || !signature || !expectedSignature || expectedSignature !== signature) return null;
  try {
    return JSON.parse(Buffer.from(payload, "base64url").toString("utf8"));
  } catch {
    return null;
  }
}

function sign(value) {
  if (!hasValidSessionSecret(config.sessionSecret)) return null;
  return base64Url(createHmac("sha256", config.sessionSecret).update(value).digest());
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

function decodeJwtPayload(token) {
  try {
    const [, payload] = String(token).split(".");
    return JSON.parse(Buffer.from(payload || "", "base64url").toString("utf8"));
  } catch {
    return {};
  }
}

function normalizeRole(value, fallback) {
  const allowed = ["Admin", "Audit Manager", "Auditor", "Finance", "Read Only"];
  return allowed.includes(value) ? value : fallback;
}

function normalizeVisibility(value, fallback) {
  const allowed = ["Role Default", "All Projects", "Assigned Projects", "Finance Records"];
  return allowed.includes(value) ? value : fallback;
}

function hasFullProjectAccess(user) {
  return user.role === "Admin" || user.role === "Audit Manager";
}

function assignedAuditorNames(project) {
  const auditTeam = Array.isArray(project.auditTeam) ? project.auditTeam : [];
  const names = auditTeam
    .map((member) => String(member?.person || "").trim())
    .filter(Boolean);
  const assignedAuditor = String(project.assignedAuditor || "").trim();
  return Array.from(new Set([assignedAuditor, ...names].filter(Boolean)));
}

function projectHasAuditor(project, auditor) {
  return assignedAuditorNames(project).includes(String(auditor || "").trim());
}

function isFinanceProject(project) {
  return (
    project.currentStage === "Final Submission" ||
    project.currentStage === "Invoice" ||
    project.currentStage === "Closed" ||
    project.invoiceStatus !== "Not Started" ||
    project.reportStatus === "Issued"
  );
}

function roleDefaultVisibility(role) {
  if (role === "Finance") return "Finance Records";
  if (role === "Auditor") return "Assigned Projects";
  return "All Projects";
}

function effectiveVisibility(user) {
  return user.defaultVisibility === "Role Default"
    ? roleDefaultVisibility(user.role)
    : normalizeVisibility(user.defaultVisibility, roleDefaultVisibility(user.role));
}

function canViewProjectForUser(user, project) {
  if (!user.active) return false;
  const visibility = effectiveVisibility(user);
  if (visibility === "All Projects") return true;
  if (visibility === "Finance Records") return isFinanceProject(project);
  return projectHasAuditor(project, user.fullName);
}

const financeProjectUpdateFields = new Set([
  "invoiceStatus",
  "paymentReceived",
  "lastUpdatedDate",
  "activityEvents",
]);

const auditorProjectUpdateFields = new Set([
  "currentStage",
  "assignmentStatus",
  "baaReceived",
  "endorsementsReceived",
  "premiumBdxReceived",
  "preAuditQuestionnaireStatus",
  "documentRequestStatus",
  "documentRequestDate",
  "brokerLastChasedDate",
  "brokerExpectedResponseDate",
  "fileSelectionCompleted",
  "testingSheetCompleted",
  "findingsSentDate",
  "coverholderResponseReceivedDate",
  "tentativeAuditWeek",
  "confirmedAuditDate",
  "calendarSyncStatus",
  "calendarEventId",
  "calendarEventWebLink",
  "calendarEventLastSyncedAt",
  "auditLocation",
  "auditRemoteLink",
  "auditDurationHours",
  "auditStartTime",
  "schedulingNotes",
  "auditType",
  "nextAction",
  "blockers",
  "dueDate",
  "lastUpdatedDate",
  "labels",
  "checklistCompletions",
  "statusHistory",
  "comments",
  "activityEvents",
]);

function validateProjectWriteForUser(user, existingProject, incomingProject) {
  if (hasFullProjectAccess(user)) return null;
  if (!existingProject) {
    return {
      error: "project_create_denied",
      message: "Only admins and audit managers can create project records.",
    };
  }
  if (!canEditProjectForUser(user, existingProject)) {
    return {
      error: "project_write_denied",
      message: `You cannot save project ${existingProject.assignmentNumber || existingProject.id}.`,
    };
  }

  const allowedFields =
    user.role === "Finance"
      ? financeProjectUpdateFields
      : user.role === "Auditor"
        ? auditorProjectUpdateFields
        : new Set();
  const changedFields = changedProjectFields(existingProject, incomingProject);
  const blockedField = changedFields.find((field) => !allowedFields.has(field));
  if (blockedField) {
    return {
      error: "project_field_denied",
      message: `Your role cannot update ${blockedField} on ${
        existingProject.assignmentNumber || existingProject.id
      }.`,
    };
  }

  if (
    changedFields.includes("comments") &&
    !hasOnlyUserAppendedItems(
      existingProject.comments,
      incomingProject.comments,
      "author",
      user.fullName,
    )
  ) {
    return {
      error: "project_field_denied",
      message: `Your role can only append your own comments on ${
        existingProject.assignmentNumber || existingProject.id
      }.`,
    };
  }
  if (
    changedFields.includes("statusHistory") &&
    !hasOnlyUserAppendedItems(
      existingProject.statusHistory,
      incomingProject.statusHistory,
      "changedBy",
      user.fullName,
    )
  ) {
    return {
      error: "project_field_denied",
      message: `Your role can only append your own status history on ${
        existingProject.assignmentNumber || existingProject.id
      }.`,
    };
  }
  if (
    changedFields.includes("activityEvents") &&
    !hasOnlyUserAppendedItems(
      existingProject.activityEvents,
      incomingProject.activityEvents,
      "actor",
      user.fullName,
    )
  ) {
    return {
      error: "project_field_denied",
      message: `Your role can only append your own activity events on ${
        existingProject.assignmentNumber || existingProject.id
      }.`,
    };
  }
  return null;
}

function changedProjectFields(existingProject, incomingProject) {
  const keys = new Set([
    ...Object.keys(existingProject),
    ...Object.keys(incomingProject),
  ]);
  return Array.from(keys).filter(
    (key) => !sameProjectValue(existingProject[key], incomingProject[key]),
  );
}

function sameProjectValue(left, right) {
  return JSON.stringify(left ?? null) === JSON.stringify(right ?? null);
}

function hasOnlyUserAppendedItems(existingItems, incomingItems, userField, fullName) {
  if (!Array.isArray(existingItems) || !Array.isArray(incomingItems)) return false;
  if (incomingItems.length < existingItems.length) return false;
  for (let index = 0; index < existingItems.length; index += 1) {
    if (!sameProjectValue(existingItems[index], incomingItems[index])) return false;
  }
  return incomingItems
    .slice(existingItems.length)
    .every(
      (item) =>
        String(item?.[userField] || "").trim() === String(fullName || "").trim(),
    );
}

function canEditProjectForUser(user, project) {
  if (hasFullProjectAccess(user)) return true;
  if (user.role === "Finance") return isFinanceProject(project);
  if (user.role === "Auditor") return projectHasAuditor(project, user.fullName);
  return false;
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
  if (origin && [config.frontendOrigin, config.publicOrigin].includes(origin)) {
    response.setHeader("Access-Control-Allow-Origin", origin);
    response.setHeader("Access-Control-Allow-Credentials", "true");
    response.setHeader("Vary", "Origin");
  }
  response.setHeader("Access-Control-Allow-Headers", "content-type");
  response.setHeader("Access-Control-Allow-Methods", "GET,POST,PUT,PATCH,DELETE,OPTIONS");
}

function requestPublicOrigin(request) {
  const forwardedHost = firstForwardedHeader(request.headers["x-forwarded-host"]);
  const host = forwardedHost || request.headers.host || "";
  const forwardedProto = firstForwardedHeader(request.headers["x-forwarded-proto"]);
  const protocol =
    forwardedProto ||
    (host.includes("localhost") || host.startsWith("127.0.0.1")
      ? "http"
      : "https");
  return host ? `${protocol}://${host}` : config.publicOrigin;
}

function requestFrontendOrigin(request, publicOrigin) {
  const refererOrigin = originFromUrl(request.headers.referer || "");
  return refererOrigin || config.frontendOrigin || publicOrigin;
}

function firstForwardedHeader(value) {
  return String(value || "").split(",")[0].trim();
}

function originFromUrl(value) {
  try {
    return new URL(value).origin;
  } catch {
    return "";
  }
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
