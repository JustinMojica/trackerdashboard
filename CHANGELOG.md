# Changelog

This log tracks practical production changes to the Audit Assignment Tracker.
It intentionally excludes secrets, publish-profile values, client secrets, and
other sensitive configuration values.

## 2026-07-14

### Added

- Added a simplified Schedule audit modal with audit date, start time,
  duration, location, remote link, notes, invite preview, and review warnings.
- Added one-click Save schedule and Save and send/update Outlook invite actions
  so users no longer need to understand calendar sync states.
- Added an Admin/Audit Manager stage override confirmation so required-item
  restrictions can be bypassed when needed while still recording the reason in
  the audit trail.
- Added an admin role access preview so admins can check visible tabs, audit
  visibility, and common permissions before approving or editing users.
- Added email verification expiry and lockout protection so signup codes expire
  after 10 minutes and repeated incorrect entries are temporarily blocked.
- Expanded the admin access preview so admins can pick a specific approved
  account and see the exact active audits currently visible to that user.

### Changed

- Reworked the Scheduling workspace to show plain statuses: No date, Planned,
  Invite sent, and Needs update.
- Removed the visible calendar sync status control from audit intake and
  scheduling cards.
- Updated Outlook invite creation to use the selected audit start time.
- Limited assignment export, JSON backup, and JSON import controls to Admin
  users only so regular users see only the daily view controls.
- Fixed the audit detail grid layout so workflow, templates, and audit trail
  panels no longer overlap.
- Simplified audit-detail language and layout by replacing technical workflow
  labels with user-facing next-step, timing-alert, document-review, folder, and
  email wording.
- Collapsed recommended next steps by default and shortened stage override
  confirmation wording.
- Collapsed attention items by default, removed extra override-helper wording, and
  explained that suggested folders are a copyable audit file-structure plan.
- Improved linked spreadsheet contact refresh by reading workbook worksheets in
  limited parallel batches, caching successful refreshes, and reusing in-flight
  refresh requests instead of stacking duplicate reads.
- Allowed Admin and Audit Manager users to archive active audits with
  confirmation instead of requiring the audit to be Closed first.
- Moved stage override guidance into a hover tip, collapsed document-review
  recommended actions by default, and replaced the signature placeholder with
  an Outlook draft action.
- Simplified the linked contact selector to show contact names only, retitled
  it to Contacts, and added search before selecting workbook contacts.
- Replaced old non-audit wording with audit/audits wording across the app UI
  and current README language.
- Replaced generic attention labels with specific attention reasons such as
  missing BAA, missing Premium BDX, missing Claims BDX, or quote not accepted.
- Added a UMR number field to Add/Edit Audit, audit detail, CSV exports, and
  the Microsoft Lists audit assignment schema.
- Added a template recipient picker that shows all detected workbook email
  options and remembers the preferred recipient for the same contact and
  routing type.
- Combined Add Audit contact search and selection into one Contacts picker,
  refreshed intake fields when a new contact is selected, exposed recipient
  options during intake, and collapsed checklist/readiness details by default.
- Added shared server recipient preferences, recently used contact shortcuts,
  contact-quality warnings, clearer create-audit shortcuts from Contacts, and
  a final Add Audit review step before saving.
- Refined the Add Audit Contacts field into a true searchable dropdown:
  users can type before selecting or open the dropdown to browse contacts.
- Removed suggested folder and suggested email cards from the audit workflow
  attention panel.
- Fixed the Add Audit Contacts dropdown so clearing a selected contact returns
  to normal search, while recent contacts remain shortcuts instead of replacing
  search results.
- Split the assignment detail page into Overview, Documents, Team & scheduling,
  and Activity tabs, with Overview as the default view for cleaner daily use.
- Added an assignment action bar and compact Overview status summary for common
  actions, attention items, readiness, schedule, quote, and next action.
- Moved archived audits into a dedicated Archive workspace that is visible
  only to Admin and Audit Manager users.
- Fixed the Add Audit Contacts dropdown so browse mode can reach the full
  loaded contact list instead of stopping at the first few A contacts.
- Consolidated backend storage, rollout health, exports, and guarded
  maintenance into a single Admin tools tab so daily admin work stays separate.
- Simplified the Command Center default view into a focused summary plus top
  operating actions, with evidence review, attention items, draft queue, and assistant
  brief available as collapsed detail sections.
- Made the template `To` field editable so users can manually enter a recipient
  when no workbook email is detected, while keeping workbook-prefilled recipients.

### Verified

- `npm run check`
- `npm test`
- `npm run build`
- Live auth config smoke check on `https://mosaic-audit-tracker-live.azurewebsites.net/api/auth/config`.

## 2026-06-22

### Added

- Added an AI Audit Coordinator panel in Command Center that ranks visible
  audit actions by attention items, due dates, document gaps, quote status, calendar
  sync, and invoice follow-up.
- Added document intelligence summaries that classify evidence package type,
  readiness confidence, missing evidence, and recommended handling at both the
  Command Center and audit-detail levels.
- Added coordinator insights and document intelligence to the exported
  operations report.

### Changed

- Disabled the linked workbook contact selector when contacts are unavailable
  so the intake wizard no longer opens an empty placeholder dropdown.

### Verified

- `npm run check`
- `npm test`
- `npm run build`
- Local endpoint smoke check on `http://127.0.0.1:5173/`.

## 2026-06-18

### Added

- Added richer Outlook scheduling controls for audit duration, location,
  remote meeting links, attendee handoff, event update/recreate behavior, and
  saved Outlook event links.
- Added contact directory search, receiver-type filtering, and a Start audit
  action that opens audit intake from a linked spreadsheet contact.
- Added intake duplicate warnings and DCA-specific intake guidance for DCA
  Agreement, Claims BDX, managing agent, and DCA contact routing.

### Changed

- Expanded audit CSV exports with Outlook event link, audit duration,
  location, and remote link fields.
- Updated the audit wizard so scheduling data can be captured before Outlook
  sync instead of only from the Scheduling workspace.

### Verified

- `npm run check`
- `npm test`
- `npm run build`
- Local browser smoke check on `http://127.0.0.1:5173/`.

## 2026-06-15

### Added

- Added manual Outlook audit event sync from the Scheduling workspace. The
  action creates an event on the signed-in user's Outlook calendar after
  `Calendars.ReadWrite` Graph application consent is granted.
- Added Mosaic branding with a transparent-background logo on the sign-in and
  signed-in tracker header.
- Added a Scheduling & Capacity workspace with audit calendar cards, schedule
  notes, calendar sync status, near-term capacity counts, and conflict warnings.
- Added linked workbook contact selection to audit intake so live spreadsheet
  contacts can prefill client/contact context, scheduling notes, and audit type
  hints.

### Changed

- Expanded linked workbook parsing to read a larger sheet area, extract DCA,
  coverholder, report, and invoice email buckets, and ignore blank placeholder
  instruction tabs.
- Removed the stale admin-only frontend guard from contact refresh so approved
  users can load linked contacts for audit intake.
- Opened linked spreadsheet contact refresh to approved users, not only admins,
  so audit intake can use contacts without exposing backend admin controls.
- Expanded CSV exports and audit detail metadata with scheduling and linked
  contact fields.
- Refined the main UI shell with a cleaner production background, stronger
  header hierarchy, clearer navigation states, and more polished action buttons.
- Improved dashboard, Kanban, workflow, admin, contact, template, and table
  surfaces with better spacing, card contrast, hover states, focus states, and
  scan-friendly status styling.
- Tightened the signed-in header copy to better explain the tracker as an
  operating view for intake, readiness, workload, contacts, and close-out.

### Verified

- `npm run check`
- `npm test`
- `npm run build`
- Local browser smoke check on `http://127.0.0.1:5173/`.

## 2026-06-06

### Added

- Added receiver-aware template previews that show the expected recipient,
  email address, routing rule, workbook source, and Outlook signature guidance.
- Added DCA-aware document request wording so DCA audits request DCA
  Agreement and Claims BDX instead of the standard coverholder document set.

### Changed

- Prevented the linked spreadsheet Contacts tab from retrying instantly after a
  failed auto-refresh; admins can still retry manually with Refresh contacts.
- Added a contact refresh timeout and visible stopped-state message so workbook
  or Graph issues do not leave the UI spinning indefinitely.
- Expanded template copy actions with receiver, subject, body, and full-draft
  copy options for cleaner Outlook handoff.

### Verified

- `npm run check`
- `npm test`
- `npm run build`

## 2026-06-05

### Added

- Added Solo vs Coordinated audit structure support so one parent audit card can
  track multiple managing-agent workstreams without duplicating Kanban cards.
- Added managing-agent workstream rollups for active, complete, waived, needing attention,
  missing-document, and needs-attention counts.
- Added DCA-specific document readiness: DCA Agreement and Claims BDX are now
  the required readiness documents for DCA assignments.
- Added an admin-only linked contact spreadsheet preview foundation for
  OneDrive/SharePoint Excel workbooks, including special-instructions columns.

### Changed

- Kept coordinated audits as one Kanban card with managing-agent summary chips
  instead of creating separate cards per managing agent.
- Updated audit detail and intake/edit screens to show managing-agent
  workstreams when an audit is Coordinated or DCA.
- Prevented coordinated audits from being archived until all managing-agent
  workstreams are complete or waived.
- Kept workbook URLs in server/Azure configuration via
  `TRACKER_CONTACT_WORKBOOK_LINKS` instead of hard-coding links into the
  frontend or GitHub.
- Updated linked workbook parsing for Mosaic client-instruction sheets, using a
  bounded cell range, request timeouts, and parallel workbook refreshes to avoid
  the Contacts tab getting stuck on large Excel files.

### Verified

- Added regression coverage for DCA document rules and ten-workstream
  coordinated audit rollups.
- Added parser coverage for linked workbook contact fields and special
  instructions.

## 2026-06-04

### Added

- Added focused app navigation for Dashboard, Assignments, Command Center,
  Reports, and Admin so daily work is no longer mixed with backend controls.
- Added an Admin workspace with separate Users, Audit log, Storage & data, and
  System health tabs.
- Added non-destructive audit archiving and restore. Archived audits are
  hidden from active boards and workload counts but remain stored for history.
- Added an Admin data-safety panel documenting that code deployments do not
  clear Microsoft Lists audit records.

### Changed

- Cleaned up the sign-in screen language so Microsoft sign-in, access request,
  email code confirmation, and admin approval are clearer.
- Moved low-frequency storage/export/system controls out of the main work area
  and into Admin.
- Updated README production notes to reflect live Microsoft Lists storage and
  archive behavior.
- Updated GitHub workflows from Node 20-based maintained actions to
  `actions/checkout@v5` and `actions/setup-node@v5`.
- Hardened configured admin handling so emails listed in
  `TRACKER_ADMIN_EMAILS` are repaired back to active Admin access if their
  Microsoft Lists user row is accidentally downgraded.
- Made the sign-in screen clearer for unapproved users by emphasizing that they
  must request account approval before normal sign-in.

### Verified

- `npm run check`
- `npm test`
- `npm run build`
- Local dev server returned `200` on `http://127.0.0.1:5174/`.
- Added regression coverage for configured admin repair.

## 2026-06-03

### Added

- Added an operating system command center combining workflow gates, SLA
  escalation, role consoles, draft queue, reporting export, and an AI-ready
  assistant brief.
- Added per-audit workflow controls for stage gates, document/quote/close-out
  gates, SLA signals, SharePoint workspace folder planning, and top draft copy.
- Added an operations report export with attention items, next actions, SLA signals,
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
  secure sign-in, durable runtime config, approval storage, audit storage, and
  activity logging.
- Created the seven Microsoft Lists on the root SharePoint site and switched
  live approval/audit storage to Microsoft Lists.

### Verified

- The live secure access config endpoint reported `configured: true` with no
  missing values after the Kudu runtime configuration fix.
- GitHub Actions deployment succeeded for commit `541dc4b`.
- The live secure access config endpoint later reported both `userStore` and
  `auditStore` as `microsoft-lists` with no missing values.

### Follow-Up

- Rotate the Azure publish profile because it was pasted into chat during setup.
- Prefer Azure App Service Environment variables once Azure portal or Azure CLI
  access is available. The persistent Azure data env file is a practical
  fallback, not the final security posture.
- Move approvals and audit records to Microsoft Lists after the Tracker Users
  and Audit Assignments list IDs are created.

## 2026-06-02

### Added

- Added GitHub Actions deployment for the live Azure App Service.
- Added secure Microsoft sign-in and request-access flow through the backend
  server.
- Added server-side approval management for Microsoft accounts.
- Added server-side audit storage with optional Microsoft Lists persistence.
- Added security checks for weak, missing, or placeholder session secrets.

### Changed

- Replaced prototype username/password access with Microsoft OAuth, email-code
  verification, and admin approval.
- Updated README and deployment docs for the secure production architecture.

### Fixed

- Fixed Windows path handling in session-secret tests.
- Fixed server setup-required handling so admin APIs fail closed when critical
  auth config is missing.
