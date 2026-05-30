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
  };
};

export type SecureSystemHealth = {
  generatedAt: string;
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
  };
  approvalStore: {
    mode: "local" | "microsoft-lists";
    configured: boolean;
    missing: string[];
    durable: boolean;
    status: string;
  };
  recommendations: string[];
};

export type AccessApprovalUpdate = {
  role?: SecureAccessUser["role"];
  defaultVisibility?: SecureAccessUser["defaultVisibility"];
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
