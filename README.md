# Audit Assignment Tracker

Secure React and Node tracker for audit assignments from intake through report, invoice, and close-out.

The current app is no longer a browser-storage-only prototype. It has:

- Microsoft Entra sign-in through a backend OAuth flow.
- Email verification code workflow.
- Admin approval before tracker access.
- Signed HTTP-only session cookies.
- Role-based project visibility and edit controls.
- Server-backed project storage with Microsoft Lists persistence in production.
- Microsoft Lists schema/export support for assignments, team members, comments, checklist items, status history, activity log, and tracker users.
- GitHub Actions deployment to Azure App Service.
- Admin health reporting for auth, Graph consent, runtime config source, storage mode, and live deployment metadata.

## Live Site

Production URL:

```text
https://mosaic-audit-tracker-live.azurewebsites.net
```

Health/config endpoint:

```text
https://mosaic-audit-tracker-live.azurewebsites.net/api/auth/config
```

The site can load even when Microsoft sign-in is not fully configured. For sign-in to work, `/api/auth/config` must return `configured: true` with no missing or invalid settings.

## Repository

```text
https://github.com/JustinMojica/trackerdashboard
```

Main branch:

```text
main
```

## Local Development

Install dependencies:

```bash
npm install
```

Run the React dev app:

```bash
npm run dev
```

Run the secure backend:

```bash
npm run server
```

Build and serve the production bundle through the secure backend:

```bash
npm run secure
```

Useful local URLs:

```text
http://localhost:5173
http://localhost:8787/api/auth/config
```

## Validation

Run these before pushing changes:

```bash
npm run check
npm test
npm run build
```

Current test coverage includes workflow logic, Microsoft Lists package generation, access-request rules, role/project permissions, and server security checks for weak or missing session secrets.

## Change Log

Production changes are tracked in:

```text
CHANGELOG.md
```

Git commit history remains the source of truth for code changes. `CHANGELOG.md`
summarizes user-facing and operational changes without exposing secrets.

## Required Azure App Settings

Set these in Azure App Service > Settings > Environment variables:

```text
MICROSOFT_TENANT_ID
MICROSOFT_CLIENT_ID
MICROSOFT_CLIENT_SECRET
MICROSOFT_MAIL_FROM
TRACKER_SESSION_SECRET
TRACKER_ALLOWED_EMAIL_DOMAINS
TRACKER_ADMIN_EMAILS
TRACKER_PUBLIC_ORIGIN
TRACKER_FRONTEND_ORIGIN
TRACKER_USER_STORE
TRACKER_USERS_SITE_ID
TRACKER_USERS_LIST_ID
TRACKER_PROJECT_STORE
TRACKER_PROJECTS_SITE_ID
TRACKER_PROJECTS_LIST_ID
SCM_DO_BUILD_DURING_DEPLOYMENT
WEBSITE_NODE_DEFAULT_VERSION
```

Current production storage settings:

```text
TRACKER_USER_STORE=microsoft-lists
TRACKER_PROJECT_STORE=microsoft-lists
SCM_DO_BUILD_DURING_DEPLOYMENT=true
WEBSITE_NODE_DEFAULT_VERSION=22
```

Use a long random value for `TRACKER_SESSION_SECRET`. Do not use `dev`, `secret`, `password`, or any placeholder value.

Do not commit `server.env`, publish profiles, client secrets, or generated deployment artifacts.

If Azure CLI or portal access is unavailable, the server can also read a
persistent Azure data env file:

```text
$HOME/data/tracker-server.env
```

This fallback survives normal Zip Deploy updates better than a `server.env`
file in `site/wwwroot`, but Azure App Service Environment variables are still
the preferred long-term configuration source.

## Microsoft Entra Setup

The Entra app registration needs:

- Web redirect URI:

```text
https://mosaic-audit-tracker-live.azurewebsites.net/api/auth/callback
```

- Local redirect URI for development:

```text
http://localhost:8787/api/auth/callback
```

- Microsoft Graph application permissions:

```text
Mail.Send
Sites.ReadWrite.All
```

Admin consent must be granted for those permissions.

## GitHub to Azure Deployment

The deployment workflow is:

```text
.github/workflows/azure-app-service.yml
```

GitHub Actions requires:

- Secret: `AZURE_WEBAPP_PUBLISH_PROFILE`
- Variable: `ENABLE_AZURE_DEPLOY=true`

The workflow runs on pushes to `main` only when `ENABLE_AZURE_DEPLOY=true`.
Each deployment writes non-secret metadata to `server/deploy-info.json` so the
admin health panel can show the live commit and deployment time.

Manual deployment:

1. GitHub repo > Actions.
2. Select `Deploy Azure App Service`.
3. Run workflow on `main`.

If a publish profile was exposed, regenerate it in Azure and replace the GitHub secret.

See:

```text
docs/github-azure-deployment.md
```

## Microsoft Lists Data Foundation

Target lists:

- Audit Assignments
- Audit Team Members
- Audit Comments
- Audit Checklist Items
- Audit Status History
- Audit Activity Log
- Tracker Users

`Audit Assignments` includes `TrackerProjectJson`, allowing the backend to store the full project object in Microsoft Lists while still exposing clean reporting columns.

Current production state:

- The root SharePoint site is used for the tracker Microsoft Lists.
- `Tracker Users` stores approved Microsoft account profiles.
- `Audit Assignments` stores project records, including the full `TrackerProjectJson` payload.
- Code deployments update the application files only; they do not clear Microsoft Lists records.
- Destructive data actions are limited to explicit admin clear/import replacement flows.

The live root-site setup can be created/reused with:

```bash
node scripts/configureRootMicrosoftLists.mjs --site-url=https://mosaicint.sharepoint.com
```

Schema details:

```text
docs/microsoft-lists-schema.md
```

## Security Notes

Recent hardening includes:

- Weak or missing `TRACKER_SESSION_SECRET` fails closed.
- Admin APIs return setup-required while security-critical config is invalid.
- Project update validation prevents overposting by role.
- Opaque Codex patch auto-apply support was removed.
- Deployment workflow is gated to `main` and `ENABLE_AZURE_DEPLOY=true`.

Important operational requirement:

- Rotate the Azure publish profile after any accidental exposure.
- Rotate Microsoft client secrets if they are pasted into chat, screenshots, tickets, or docs.

## Major Features

- Secure Microsoft sign-in.
- Account request, email verification, and admin approval.
- Admin user management for approved Microsoft accounts.
- Role-based visibility: Admin, Audit Manager, Auditor, Finance, Read Only.
- Guided project intake with required-field checks.
- Assignment dashboard, Today queue, workload, cycle-time reporting, filters, Kanban/table views.
- Focused app navigation for Dashboard, Assignments, Command Center, Reports, and Admin.
- Project archive/restore so completed work can leave active views without being deleted.
- Operating system command center for workflow gates, SLA escalation, role consoles, draft queue, and AI-ready assistant brief.
- Per-project workflow controls for stage gates, workspace folder planning, and communication draft review.
- Document readiness workflow and broker follow-up actions.
- Finance update path for invoice/payment records.
- Activity timeline and admin activity review.
- Microsoft Lists migration package export.

## Recommended Next Steps

1. Regenerate the Azure publish profile and replace the GitHub secret.
2. Prefer Azure App Service Environment variables over Kudu env-file fallback once portal or CLI access is stable.
3. Test the live sign-in flow with your account.
4. Have one coworker request access, confirm the email code, and wait for admin approval.
5. Enter real test assignments and validate the intake fields.
6. Validate archive/restore on a completed test assignment before production rollout.
7. Add Power Automate only after live records and activity logging are stable.
