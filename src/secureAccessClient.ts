export type SecureAccessUser = {
  email: string;
  username: string;
  fullName: string;
  role: "Admin" | "Audit Manager" | "Auditor" | "Finance" | "Read Only";
  permissionGroup: "Admin" | "Audit Manager" | "Auditor" | "Finance" | "Read Only";
  active: boolean;
  defaultVisibility: "Role Default" | "All Projects" | "Assigned Projects" | "Finance Records";
  emailVerified: boolean;
  accessRequestStatus:
    | "Approved"
    | "Pending Verification"
    | "Pending Approval"
    | "Rejected";
  requestedAt: string;
  approvedAt: string;
  approvedBy: string;
  rejectionReason: string;
};

export type SecureAccessState = {
  configured: boolean;
  authenticated: boolean;
  status:
    | "approved"
    | "domain-blocked"
    | "expired-state"
    | "invalid-state"
    | "not-requested"
    | "not-signed-in"
    | "pending"
    | "pending-approval"
    | "pending-verification"
    | "rejected"
    | "request-required"
    | "setup-required"
    | "success"
    | string;
  user?: SecureAccessUser;
  pendingRequests?: SecureAccessUser[];
  managedUsers?: SecureAccessUser[];
  signInUrl?: string;
  requestAccessUrl?: string;
  setup?: {
    configured: boolean;
    missing: string[];
    redirectUri: string;
    frontendOrigin: string;
    userStore?: {
      mode: "local" | "microsoft-lists";
      configured: boolean;
      missing: string[];
    };
    projectStore?: {
      mode: "local" | "microsoft-lists";
      configured: boolean;
      missing: string[];
    };
  };
};

export type SecureSystemHealth = {
  generatedAt: string;
  deployment: {
    source: string;
    commit: string;
    branch: string;
    deployedAt: string;
    workflowRun: string;
  };
  runtime: {
    nodeVersion: string;
    websiteHostname: string;
    configSource: string;
    appSettingKeyCount: number;
    localFileKeyCount: number;
    persistentDataEnvLoaded: boolean;
    loadedEnvFileCount: number;
  };
  server: {
    configured: boolean;
    missing: string[];
    redirectUri: string;
    frontendOrigin: string;
    publicOrigin: string;
  };
  graphApp: {
    tokenAvailable: boolean;
    roles: string[];
    missingRoles: string[];
    appDisplayName?: string;
    error?: string;
  };
  mail: {
    configured: boolean;
    from: string;
    permissionGranted: boolean;
  };
  sharePoint: {
    permissionGranted: boolean;
    siteIdConfigured: boolean;
    trackerUsersListIdConfigured: boolean;
    auditAssignmentsListIdConfigured: boolean;
  };
  approvalStore: {
    mode: "local" | "microsoft-lists";
    configured: boolean;
    missing: string[];
    durable: boolean;
    status: string;
  };
  projectStore: {
    mode: "local" | "microsoft-lists";
    configured: boolean;
    missing: string[];
    durable: boolean;
    status: string;
  };
  contactSources?: {
    configured: boolean;
    count: number;
    missing: string[];
    status: string;
  };
  recommendations: string[];
};

export type LinkedContactInstruction = {
  label: string;
  value: string;
};

export type LinkedContact = {
  id: string;
  sourceId: string;
  sourceLabel: string;
  workbookName: string;
  worksheetName: string;
  company: string;
  coverholder: string;
  managingAgent: string;
  broker: string;
  contactName: string;
  email: string;
  phone: string;
  role: string;
  specialInstructions: LinkedContactInstruction[];
  raw: Record<string, string>;
};

export type LinkedContactSource = {
  id: string;
  label: string;
  status: "ok" | "error";
  workbookName: string;
  worksheetCount: number;
  rowCount: number;
  error?: string;
};

export type LinkedContactWarning = {
  sourceId: string;
  worksheetName: string;
  message: string;
};

export type LinkedContactSourcesResponse = {
  configured: boolean;
  generatedAt: string;
  sources: LinkedContactSource[];
  contacts: LinkedContact[];
  warnings: LinkedContactWarning[];
};

export type AccessApprovalUpdate = {
  role?: SecureAccessUser["role"];
  defaultVisibility?: SecureAccessUser["defaultVisibility"];
};

export type SecureAccessUserUpdate = AccessApprovalUpdate & {
  fullName?: string;
  active?: boolean;
};

export function secureAccessApiBase() {
  if (window.location.port === "5173") return "http://localhost:8787";
  return "";
}

export function secureAccessUrl(path: string) {
  return `${secureAccessApiBase()}${path}`;
}

export async function getSecureAccessState(): Promise<SecureAccessState> {
  const response = await fetch(secureAccessUrl("/api/auth/me"), {
    credentials: "include",
  });
  if (!response.ok) {
    throw new Error("Secure access server is not available.");
  }
  return response.json();
}

export async function verifySecureAccessCode(code: string) {
  const response = await fetch(secureAccessUrl("/api/access/verify-code"), {
    method: "POST",
    credentials: "include",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ code }),
  });
  if (!response.ok) {
    const payload = await response.json().catch(() => ({}));
    throw new Error(
      payload.message || payload.error || "Verification code was not accepted.",
    );
  }
  return response.json();
}

export async function approveSecureAccessRequest(
  email: string,
  update: AccessApprovalUpdate = {},
) {
  return decideSecureAccessRequest(email, "approve", update);
}

export async function rejectSecureAccessRequest(email: string) {
  return decideSecureAccessRequest(email, "reject");
}

export async function updateSecureAccessUser(
  email: string,
  update: SecureAccessUserUpdate,
) {
  const response = await fetch(
    secureAccessUrl(`/api/admin/users/${encodeURIComponent(email)}`),
    {
      method: "PATCH",
      credentials: "include",
      headers: { "content-type": "application/json" },
      body: JSON.stringify(update),
    },
  );
  if (!response.ok) {
    const payload = await response.json().catch(() => ({}));
    throw new Error(payload.message || payload.error || "User update failed.");
  }
  return response.json();
}

export async function logoutSecureAccess() {
  await fetch(secureAccessUrl("/api/auth/logout"), {
    method: "POST",
    credentials: "include",
  });
}

export async function getSecureSystemHealth(): Promise<SecureSystemHealth> {
  const response = await fetch(secureAccessUrl("/api/admin/system-health"), {
    credentials: "include",
  });
  if (!response.ok) {
    const payload = await response.json().catch(() => ({}));
    throw new Error(payload.message || payload.error || "System health check failed.");
  }
  return response.json();
}

export async function getLinkedContactSources(): Promise<LinkedContactSourcesResponse> {
  const controller = new AbortController();
  const timeoutId = window.setTimeout(() => controller.abort(), 25000);
  try {
    const response = await fetch(secureAccessUrl("/api/admin/contact-sources"), {
      credentials: "include",
      signal: controller.signal,
    });
    if (!response.ok) {
      const payload = await response.json().catch(() => ({}));
      throw new Error(payload.message || payload.error || "Contact source refresh failed.");
    }
    return response.json();
  } catch (error) {
    if (error instanceof DOMException && error.name === "AbortError") {
      throw new Error("Contact source refresh timed out. Retry after checking workbook access.");
    }
    throw error;
  } finally {
    window.clearTimeout(timeoutId);
  }
}

async function decideSecureAccessRequest(
  email: string,
  decision: "approve" | "reject",
  update: AccessApprovalUpdate = {},
) {
  const response = await fetch(
    secureAccessUrl(
      `/api/admin/access-requests/${encodeURIComponent(email)}/${decision}`,
    ),
    {
      method: "POST",
      credentials: "include",
      headers: { "content-type": "application/json" },
      body: JSON.stringify(update),
    },
  );
  if (!response.ok) {
    const payload = await response.json().catch(() => ({}));
    throw new Error(payload.message || payload.error || "Request update failed.");
  }
  return response.json();
}
