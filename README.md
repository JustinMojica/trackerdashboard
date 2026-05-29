# Audit Assignment Tracker Prototype

A React prototype for an auditing firm to track assignments from intake through final report, invoice, and close-out. The prototype stores data in browser local storage only and uses mock sample records with fictional client / coverholder codes.

## Run locally

```bash
npm install
npm run dev
```

Run local checks and the lightweight workflow test suite with:

```bash
npm run check
npm test
```

> The current environment may block npm registry access. The project is set up as a normal Vite + React app once dependencies can be installed.

## Lifecycle stages

The prototype uses the following fixed lifecycle stages:

1. Intake
2. Registration
3. Quote
4. Scheduling
5. Pre-Audit
6. File Selection
7. Audit Fieldwork
8. Findings
9. Report Drafting
10. Final Submission
11. Invoice
12. Closed

## SharePoint-ready data model

The `AuditProject` shape in `src/main.tsx` is intentionally flat and list-friendly so it can later map to Microsoft 365 tooling:

| Prototype field                                                                           | SharePoint column type                          | Power Platform notes                                                                                    |
| ----------------------------------------------------------------------------------------- | ----------------------------------------------- | ------------------------------------------------------------------------------------------------------- |
| `assignmentNumber`                                                                        | Single line of text                             | Unique business key and display title candidate                                                         |
| `assignmentSource`                                                                        | Choice (`Email`, `DAM`)                         | Drives conditional quote and final submission automation                                                |
| `assignmentType`                                                                          | Choice (`DCA`, `CH`, `MGA`, `Company Contract`) | Supports portfolio segmentation and reporting                                                           |
| `auditEntity`                                                                             | Single line of text                             | Captures the free-form audited entity name shown on assignment records                                  |
| `clientCoverholderCode`                                                                   | Single line of text                             | Uses non-sensitive code rather than real client name                                                    |
| `broker`                                                                                  | Single line of text or lookup                   | Can become a Broker lookup list later                                                                   |
| `auditTeam`, `assignedAuditor`, `reviewer`                                                | Multi-person / Person or Group                  | Supports lead + supporting auditor teams, legacy primary auditor compatibility, workload dashboards, and approvals |
| `currentStage`                                                                            | Choice                                          | Maps to Kanban buckets and Power BI funnel visuals                                                      |
| `assignmentStatus`, `quoteStatus`, `reportStatus`, `invoiceStatus`, `damSubmissionStatus` | Choice                                          | Suitable for Power Apps dropdowns and Power Automate conditions                                         |
| `quoteAmount`                                                                             | Currency                                        | Feeds pipeline value reporting                                                                          |
| `paymentReceived`                                                                         | Yes/No                                          | Supports invoice and payment received tracking                                                          |
| `labels`                                                                                  | Multi-choice                                    | Supports Trello-style card tags such as High Priority, Medium Priority, Low Priority, and Waiting on Broker   |
| `checklistCompletions`                                                                    | Child list                                      | Recommended as related checklist rows keyed by assignment, stage, and checklist item                    |
| Date fields                                                                               | Date and time                                   | Supports due, overdue, and SLA reporting                                                                |
| Document receipt and completion fields                                                    | Yes/No                                          | Supports missing document blockers and checklist automation                                             |
| `nextAction`, `blockers`                                                                  | Multiple lines of text                          | Manual action notes and exception management                                                            |
| `statusHistory`                                                                           | Child list                                      | Recommended as a separate Assignment Status History list related by assignment ID with date/time stamps and cycle-time reporting |
| `activityEvents`                                                                          | Child list                                      | Captures field, team, document, checklist, and stage events for a full audit activity trail             |
| `comments`                                                                                | Child list                                      | Recommended as a related comments/activity list for card notes and collaboration                        |

A future SharePoint implementation should use an **Audit Assignments** list for the main project record and an **Audit Assignment Status History** child list for the stage movement audit trail. Power Automate flows can enforce the same gate rules used in this prototype before advancing stages. The newer card comments and status history timestamps should map to child lists so users can keep a Trello-style activity trail without bloating the main assignment row.

The prototype now includes a concrete Microsoft Lists migration layer in `src/microsoftListsSchema.ts`, a live Graph client in `src/microsoftListsClient.ts`, and Microsoft Entra sign-in through `src/microsoftAuth.ts`. Admin and Audit Manager users can export a **Microsoft Lists package** from the app, or configure SharePoint site/list IDs plus a Microsoft Entra app client ID to sign in and test, sync, and load assignment rows.

- Graph-style list creation payloads for the target SharePoint/Microsoft Lists structure.
- Flattened seed rows for assignments, teams, comments, checklist rows, status history, the append-only activity log, and prototype users.
- Totals for list count, assignment count, activity-log events, and seed rows so the export can be checked before import.
- Live sync controls for testing the Audit Assignments list, pushing configured list rows, and loading assignment rows back into the browser prototype.
- Microsoft sign-in using MSAL Browser with delegated Graph scopes for `User.Read` and `Sites.ReadWrite.All`.

See `docs/microsoft-lists-schema.md` for the target list layout and implementation order.

## Update log

| Update size | Version | What changed                                                                                                                                            |
| ----------- | ------- | ------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Major       | UX-20   | Add gated account requests: company-email signup, prototype email confirmation codes, and admin approval before tracker access.                         |
| Major       | UX-19   | Replace the pasted Graph token workflow with Microsoft Entra/MSAL sign-in, token refresh, sign-out, and auth config tests.                              |
| Major       | UX-18   | Add a live Microsoft Graph connection mode, saved SharePoint/list ID settings, push/pull controls, and Graph sync tests.                                |
| Major       | UX-17   | Add a Microsoft Lists schema/migration package, central-storage readiness panel, and normalized activity-log export rows.                              |
| Medium      | UX-16   | Add Today's Work priority queues, role-based saved filter views, a clearer Clear filters action, Kanban stage counts, and due-today card badges.        |
| Medium      | UX-15   | Add reusable email/document template previews and searchable project audit trails.                                                                       |
| Minor       | UX-14   | Add confirmation prompts for reset/import actions, clearer role-based empty states, user status badges, and a last-export timestamp.                    |
| Medium      | UX-13   | Add an Admin user management panel for prototype users, roles, active status, and default project visibility.                                           |
| Major       | UX-12   | Add prototype login, role-scoped dashboards, user-aware audit activity, finance-only invoice updates, and JSON backup import/export.                    |
| Minor       | UX-11   | Add removable label chips so labels can be cleared directly from lifecycle cards and detail headers.                                                    |
| Minor       | UX-10   | Make waiting-on-broker and broker-chase workflow events idempotent and collapse legacy same-day duplicates in the audit trail.                         |
| Minor       | UX-9    | Add a clear waiting-on-broker workflow action for broker-completed follow-up that removes the label without forcing documents complete.                 |
| Minor       | UX-8    | Prevent duplicate audit trail comment entries and make document-complete workflow idempotent.                                                           |
| Minor       | UX-7    | Temporarily archive the auditor directory UI while keeping default auditor data available for filters, workload, and assignment forms.                  |
| Minor       | UX-6    | Remove saved work queues, clarify sample auditor pairings, and tighten the project detail layout to avoid empty spacing.                              |
| Minor       | UX-5    | Add five prioritized recommended next steps, clean dashboard labels, and next-action hover guidance.                                                    |
| Major       | UX-4    | Add activity event tracking, post-intake support additions, cycle-time reporting ranges, and simplified role-based workload counts.                    |
| Major       | UX-3    | Add lead/supporting audit teams, count shared assignments in workload, and clarify document readiness status colors.                                    |
| Major       | UX-2    | Add audit trail, document readiness workflow actions, zero-load workload minimization, and a Node test suite for workflow logic.                  |
| Major       | UX-1    | Redesign workload as a dashboard, add guided intake, Trello-style labels, Today's Work, interactive checklists, and stronger Office 365 data readiness. |
| Small       | Patch   | Copy changes, option-list tweaks, and small field additions that do not change the primary workflow.                                                    |

## Office 365 migration readiness

Move to Office 365 after the team validates this major UX iteration with realistic assignments. Recommended target lists:

- **Audit Assignments** for the main flat assignment record.
- **Audit Team Members** for lead/supporting assignment rows.
- **Audit Comments** for Trello-style card comments.
- **Audit Checklist Items** for per-stage checklist completion.
- **Audit Status History** for timestamped stage and status changes and stage-duration reporting.
- **Audit Activity Log** for append-only accountability across edits, comments, documents, finance, stage changes, checklist changes, and team changes.
- **Tracker Users** for the prototype role map, access-request status, email-confirmation state, and admin approval trail until Microsoft 365 identity groups replace it.

This split keeps the main assignment row Power BI friendly while letting Power Apps and Power Automate handle related activity, comments, reminders, and approvals.

### Significant next upgrades

1. **Replace prototype verification codes with real Microsoft 365 email delivery** so new users receive a confirmation link/code instead of seeing a local test code.
2. **Add a connection health/setup screen** that checks every required Microsoft List, validates permissions, and reports missing list IDs before sync.
3. **Map approved users to Microsoft 365 accounts/groups** so the current role model can move away from local test passwords.
4. **Build a Power Apps front end** for intake, stage movement, document readiness, comments, and reviewer sign-off once the SharePoint list schema is stable.
5. **Automate the repeatable follow-up work** with Power Automate flows for new intake alerts, quote approval reminders, broker chase reminders, reviewer approvals, stage-history creation, and invoice/payment notifications.
6. **Add document library integration** so BAA, endorsements, Premium BDX, testing sheets, reports, and invoice artifacts are stored against the assignment record instead of only represented as checkboxes.
7. **Publish Power BI reporting** from the SharePoint lists for workload, aging, cycle time, overdue items, quote value, document blockers, and closed-audit throughput.
8. **Define permissions and governance** before launch: auditor/reviewer roles, edit rights by stage, naming rules, required metadata, retention, and environment ownership.

### Automation candidates

- **On assignment created:** notify the lead auditor/reviewer, stamp initial status history, and create the default checklist rows.
- **On stage changed:** write a child status-history row, validate gate requirements, and notify the next owner when a blocker exists.
- **On document request date set:** generate the document-request email from a template, attach/link the document request package, send it to the broker/contact owner, and schedule follow-up reminders until required documents are marked received.
- **On pre-audit questionnaire required:** create or copy the questionnaire from a Word/Forms template, send it to the broker/coverholder, set the expected response date, and add reminders until the questionnaire status is Complete or Not Required.
- **On quote ready:** generate the quote document from assignment fields, save it to the assignment folder, send the quote email, and update quote status to Sent.
- **On quote accepted:** prompt scheduling tasks, notify the assigned audit team, and create/update the audit confirmation email.
- **On findings sent:** create the findings email from a template, send it to the recipient group, and start a coverholder-response reminder clock.
- **On final report issued:** generate the final report package checklist, create invoice tasks, draft/send the invoice email, and, if needed, add DAM submission follow-up.
- **On invoice sent or paid:** store the invoice artifact, notify finance/audit operations, update payment status, and move eligible records toward close-out after final submission and archive checks are complete.

### Teams and Microsoft 365 automation blueprint

Use SharePoint/Microsoft Lists as the system of record and Teams as the notification/action surface. A practical first release could use these automations:

| Trigger | Power Automate action | Teams / 365 outcome |
| ------- | -------------------- | ------------------- |
| Assignment item created or modified | SharePoint list trigger evaluates source, due date, assigned team, blockers, and current stage | Post a Teams message to the audit operations channel with assignment number, owner, due date, and next action |
| Stage changed | Compare previous/current stage, append a status-history row, and run the same gate checks as `canMoveToStage` | Notify the next owner in Teams and include blockers if the stage move needs attention |
| Document request sent or broker chase date updated | Generate/send the document-request email and schedule reminders until documents are received or Waiting on Broker is cleared | Send Teams reminders to the lead auditor and optional broker-facing owner |
| Pre-audit questionnaire required | Create or copy the questionnaire artifact, send the questionnaire email, set expected response date, and chase until complete | Notify the lead auditor when the questionnaire is sent, overdue, or received |
| Quote ready or quote sent | Generate the quote document/email from assignment fields, save a copy to SharePoint, and update quote status | Notify the audit team that a quote was sent and surface quote value/status in Teams |
| Quote accepted | Create scheduling/checklist tasks and update status history | Post a Teams update tagging the audit team to confirm audit week/date |
| Findings sent | Generate/send the findings email and start a response-deadline timer | Send Teams reminders until coverholder response date is populated |
| Report ready for review | Start an approval | Route reviewer approval in Teams / Power Automate before final submission |
| Invoice ready, sent, or paid | Generate the invoice document/email, store the artifact, update invoice/payment status, and evaluate close-out readiness | Notify finance/audit operations in Teams and move eligible records toward Closed |

Recommended implementation order:

1. **Lists first:** create the SharePoint lists and required columns before building flows.
2. **Templates second:** create approved Outlook/Word templates for pre-audit questionnaires, document requests, quotes, findings, reports, and invoices.
3. **Notifications third:** start with Teams posts for created/modified assignments and broker/document reminders.
4. **Email/document automation fourth:** add automated sends and generated document copies only after templates and required fields are stable.
5. **Approvals fifth:** add reviewer approvals after stage and document data is reliable.
6. **Governance sixth:** add service accounts/owners, retry rules, naming conventions, and environment ownership before production.

### Level-up plan: step, reason, practical example

| Step | Why do it | Practical example |
| ---- | --------- | ----------------- |
| 1. Move data from browser local storage to SharePoint/Microsoft Lists | Gives the team one shared source of truth, version history, permissions, and a clean trigger point for automations. | A new audit entered by operations immediately appears for every auditor instead of only in one browser profile. |
| 2. Split child data into related lists | Keeps the main assignment row fast and reportable while preserving detailed history and many-to-one records. | One assignment can have 12 comments, 8 checklist rows, 5 stage-history entries, and 4 document rows without bloating the assignment list. |
| 3. Standardize email and document templates | Prevents automation from sending inconsistent wording or incomplete files. | Pre-audit questionnaire, document request, quote, findings, report, and invoice templates all use the same assignment fields. |
| 4. Build Power Automate notifications | Reduces manual follow-up and makes the tracker actively push work to people. | When an assignment is created or modified, a Teams channel message posts the assignment number, owner, due date, blockers, and next action. |
| 5. Automate outbound emails and generated documents | Removes repetitive manual drafting while preserving a saved copy of what was sent. | A quote-ready status generates the quote PDF/Word file, stores it in SharePoint, sends the quote email, and marks Quote status as Sent. |
| 6. Add broker/document reminder flows | Prevents document requests from getting stale and standardizes chase cadence. | If Premium BDX is still missing two business days after the broker chase date, the lead auditor gets a Teams reminder. |
| 7. Add stage-change automation | Makes cycle-time metrics reliable and creates an audit trail without relying on manual notes. | Moving from Quote to Scheduling automatically writes a status-history row and notifies the scheduler to confirm audit dates. |
| 8. Add approval gates | Protects high-risk steps from moving forward without reviewer sign-off. | A report cannot move to Final Submission until the reviewer approves the draft from Teams Approvals. |
| 9. Add document library integration | Moves from checkbox-only readiness to actual evidence management. | BAA, endorsements, Premium BDX, testing sheets, draft reports, final reports, and invoices live in a linked SharePoint folder for the assignment. |
| 10. Add Power BI reporting | Turns operational data into leadership visibility. | Management can see aging by stage, open blockers, workload by auditor, quote pipeline, and average time from findings to final report. |
| 11. Add governance and roles | Prevents automation drift and protects production records. | Auditors can update checklists/comments, reviewers can approve reports, finance can update invoice status, and only admins can edit lifecycle choices. |
| 12. Pilot with real assignments | Validates the workflow before investing in more automation and avoids automating the wrong process. | Run 10 live audits through the list/app, note every exception, then adjust fields, notifications, and approvals before full rollout. |
