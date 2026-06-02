import type { AuditProject } from "./main";
import { secureAccessUrl } from "./secureAccessClient";

export async function getServerProjects(): Promise<AuditProject[]> {
  const response = await fetch(secureAccessUrl("/api/projects"), {
    credentials: "include",
  });
  if (!response.ok) {
    const payload = await response.json().catch(() => ({}));
    throw new Error(payload.message || payload.error || "Project storage is not available.");
  }
  const payload = await response.json();
  return Array.isArray(payload.projects) ? payload.projects : [];
}

export async function saveServerProjects(
  projects: AuditProject[],
  options: { replaceAll?: boolean } = {},
): Promise<AuditProject[]> {
  const response = await fetch(secureAccessUrl("/api/projects"), {
    method: "PUT",
    credentials: "include",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({
      mode: options.replaceAll ? "replace-all" : "merge",
      projects,
    }),
  });
  if (!response.ok) {
    const payload = await response.json().catch(() => ({}));
    throw new Error(payload.message || payload.error || "Project save failed.");
  }
  const payload = await response.json();
  return Array.isArray(payload.projects) ? payload.projects : projects;
}
