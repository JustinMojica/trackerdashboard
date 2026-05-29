export type AccountRequestStatus =
  | "Approved"
  | "Pending Verification"
  | "Pending Approval"
  | "Rejected";

export type AccessRequestUser = {
  fullName: string;
  username: string;
  password: string;
  role: string;
  permissionGroup: string;
  email: string;
  active: boolean;
  defaultVisibility: string;
  emailVerified: boolean;
  accessRequestStatus: AccountRequestStatus;
  verificationCode: string;
  requestedAt: string;
  approvedAt: string;
  approvedBy: string;
  rejectionReason: string;
};

const personalEmailDomains = new Set([
  "aol.com",
  "gmail.com",
  "hotmail.com",
  "icloud.com",
  "live.com",
  "outlook.com",
  "proton.me",
  "protonmail.com",
  "yahoo.com",
]);

export function normalizeCompanyEmail(email: string) {
  return email.trim().toLowerCase();
}

export function isValidCompanyEmail(email: string) {
  const normalizedEmail = normalizeCompanyEmail(email);
  const [localPart, domain, extra] = normalizedEmail.split("@");
  return Boolean(
    localPart &&
      domain &&
      !extra &&
      domain.includes(".") &&
      !personalEmailDomains.has(domain),
  );
}

export function usernameFromCompanyEmail(email: string) {
  return normalizeCompanyEmail(email)
    .split("@")[0]
    .replace(/[^a-z0-9]+/g, ".")
    .replace(/^\.+|\.+$/g, "");
}

export function createVerificationCode() {
  return Math.floor(100000 + Math.random() * 900000).toString();
}

export function buildAccessRequestUser({
  fullName,
  email,
  password,
  requestedAt,
  verificationCode = createVerificationCode(),
}: {
  fullName: string;
  email: string;
  password: string;
  requestedAt: string;
  verificationCode?: string;
}): AccessRequestUser {
  const cleanEmail = normalizeCompanyEmail(email);
  return {
    fullName: fullName.trim(),
    username: usernameFromCompanyEmail(cleanEmail),
    password,
    role: "Auditor",
    permissionGroup: "Auditor",
    email: cleanEmail,
    active: false,
    defaultVisibility: "Role Default",
    emailVerified: false,
    accessRequestStatus: "Pending Verification",
    verificationCode,
    requestedAt,
    approvedAt: "",
    approvedBy: "",
    rejectionReason: "",
  };
}

export function verifyAccessRequestEmail<T extends AccessRequestUser>(
  users: T[],
  email: string,
  verificationCode: string,
) {
  const cleanEmail = normalizeCompanyEmail(email);
  const cleanCode = verificationCode.trim();
  let matched = false;
  const nextUsers = users.map((user) => {
    if (
      user.email.toLowerCase() !== cleanEmail ||
      user.verificationCode !== cleanCode ||
      user.accessRequestStatus !== "Pending Verification"
    ) {
      return user;
    }
    matched = true;
    return {
      ...user,
      emailVerified: true,
      accessRequestStatus: "Pending Approval" as AccountRequestStatus,
      verificationCode: "",
    };
  });
  return { users: nextUsers, matched };
}

export function canApproveAccessRequest(user: AccessRequestUser) {
  return (
    user.accessRequestStatus === "Pending Approval" &&
    user.emailVerified &&
    !user.active
  );
}

export function approveAccessRequest<T extends AccessRequestUser>(
  user: T,
  approvedBy: string,
  approvedAt: string,
): T {
  return {
    ...user,
    active: true,
    emailVerified: true,
    accessRequestStatus: "Approved",
    approvedAt,
    approvedBy,
    rejectionReason: "",
    verificationCode: "",
  };
}

export function rejectAccessRequest<T extends AccessRequestUser>(
  user: T,
  rejectedBy: string,
  rejectedAt: string,
  reason = "Request rejected by admin.",
): T {
  return {
    ...user,
    active: false,
    accessRequestStatus: "Rejected",
    approvedAt: rejectedAt,
    approvedBy: rejectedBy,
    rejectionReason: reason,
    verificationCode: "",
  };
}
