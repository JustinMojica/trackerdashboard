# Changelog

This log tracks practical production changes to the Audit Assignment Tracker.
It intentionally excludes secrets, publish-profile values, client secrets, and
other sensitive configuration values.

## 2026-06-03

### Added

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
