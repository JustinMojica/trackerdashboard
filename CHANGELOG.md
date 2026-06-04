# Changelog

This log tracks practical production changes to the Audit Assignment Tracker.
It intentionally excludes secrets, publish-profile values, client secrets, and
other sensitive configuration values.

## 2026-06-04

### Added

- Added focused app navigation for Dashboard, Assignments, Command Center,
  Reports, and Admin so daily work is no longer mixed with backend controls.
- Added an Admin workspace with separate Users, Audit log, Storage & data, and
  System health tabs.
- Added non-destructive project archiving and restore. Archived projects are
  hidden from active boards and workload counts but remain stored for history.
- Added an Admin data-safety panel documenting that code deployments do not
  clear Microsoft Lists project records.

### Changed

- Cleaned up the sign-in screen language so Microsoft sign-in, access request,
  email code confirmation, and admin approval are clearer.
- Moved low-frequency storage/export/system controls out of the main work area
  and into Admin.
- Updated README production notes to reflect live Microsoft Lists storage and
  archive behavior.
- Updated GitHub workflows from Node 20-based maintained actions to
  `actions/checkout@v5` and `actions/setup-node@v5`.

### Verified

- `npm run check`
- `npm test`
- `npm run build`
- Local dev server returned `200` on `http://127.0.0.1:5174/`.

## 2026-06-03

### Added

- Added an operating system command center combining workflow gates, SLA
  escalation, role consoles, draft queue, reporting export, and an AI-ready
  assistant brief.
- Added per-project workflow controls for stage gates, document/quote/close-out
  gates, SLA signals, SharePoint workspace folder planning, and top draft copy.
- Added an operations report export with blockers, next actions, SLA signals,
  draft queue, and workspace folders for each visible assignment.
- Added live deployment metadata support for GitHub Actions deployments.
- Added admin health reporting for runtime configuration source, Node version,
  public host, and deployment commit metadata.
- Added a formal Microsoft Lists migration phase view in the data foundation
  panel.
- Added support for a persistent Azure data env file at
  `$HOME/data/tracker-server.env`, so live auth configuration can survive future
  Zip Deploy updates when Azure CLI/App Settings access is unavailable.
- Added a root SharePoint site setup script for creating/reusing the tracker
  Microsoft Lists and seeding initial live storage rows.

### Changed

- Updated the production readiness panel so admins can see whether runtime
  config comes from Azure App Service settings, a persistent Azure data env
  file, or a local env file.
- Tightened the data foundation language around the five real migration phases:
  secure sign-in, durable runtime config, approval storage, project storage, and
  activity logging.
- Created the seven Microsoft Lists on the root SharePoint site and switched
  live approval/project storage to Microsoft Lists.

### Verified

- The live secure access config endpoint reported `configured: true` with no
  missing values after the Kudu runtime configuration fix.
- GitHub Actions deployment succeeded for commit `541dc4b`.
- The live secure access config endpoint later reported both `userStore` and
  `projectStore` as `microsoft-lists` with no missing values.

### Follow-Up

- Rotate the Azure publish profile because it was pasted into chat during setup.
- Prefer Azure App Service Environment variables once Azure portal or Azure CLI
  access is available. The persistent Azure data env file is a practical
  fallback, not the final security posture.
- Move approvals and project records to Microsoft Lists after the Tracker Users
  and Audit Assignments list IDs are created.

## 2026-06-02

### Added

- Added GitHub Actions deployment for the live Azure App Service.
- Added secure Microsoft sign-in and request-access flow through the backend
  server.
- Added server-side approval management for Microsoft accounts.
- Added server-side project storage with optional Microsoft Lists persistence.
- Added security checks for weak, missing, or placeholder session secrets.

### Changed

- Replaced prototype username/password access with Microsoft OAuth, email-code
  verification, and admin approval.
- Updated README and deployment docs for the secure production architecture.

### Fixed

- Fixed Windows path handling in session-secret tests.
- Fixed server setup-required handling so admin APIs fail closed when critical
  auth config is missing.
