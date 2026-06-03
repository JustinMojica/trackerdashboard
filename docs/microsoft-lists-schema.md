# Microsoft Lists Storage Schema

This is the target central-storage structure for the audit assignment tracker. The live app now uses the backend server for Microsoft sign-in, approval records, and shared project storage. Microsoft Lists is the next durable storage target for approvals, project records, reporting rows, and the activity ledger. The exported Microsoft Lists package contains the list schemas, Graph-style list creation payloads, and seed rows.

## Target Lists

| List | Purpose | Key fields |
| ---- | ------- | ---------- |
| Audit Assignments | Main assignment record | `TrackerAssignmentId`, `AssignmentNumber`, `CurrentStage`, `AssignmentStatus`, document flags, finance fields, due dates, `TrackerProjectJson` |
| Audit Team Members | Lead/support workload rows | `TeamMemberKey`, `TrackerAssignmentId`, `PersonName`, `TeamRole`, `ActiveOnAssignment` |
| Audit Comments | Card comments | `TrackerCommentId`, `TrackerAssignmentId`, `CommentCreatedAt`, `CommentAuthor`, `CommentBody` |
| Audit Checklist Items | Per-stage checklist state | `TrackerChecklistItemId`, `TrackerAssignmentId`, `ChecklistKey`, `ChecklistStage`, `ChecklistItem`, `Completed` |
| Audit Status History | Stage movement history | `TrackerHistoryId`, `TrackerAssignmentId`, `ChangedAt`, `ChangedBy`, `FromStage`, `ToStage`, `StageNote` |
| Audit Activity Log | Append-only accountability log | `TrackerEventId`, `TrackerAssignmentId`, `OccurredAt`, `EventType`, `ActorName`, `Summary`, `Detail`, previous/new values |
| Tracker Users | Prototype role map and account access gate | `TrackerUsername`, `FullName`, `Email`, `Role`, `PermissionGroup`, `Active`, `EmailVerified`, `AccessRequestStatus`, `DefaultVisibility`, `RequestedAt`, `ApprovedAt`, `ApprovedBy`, `RejectionReason`, `VerificationCodeHash`, `VerificationSentAt` |

## Live Storage Mode

The live storage settings are server-side environment values. Do not store SharePoint list IDs, client secrets, or tenant configuration in browser storage for production.

The Microsoft Entra app registration needs:

- Platform type: Web.
- Redirect URI: `https://mosaic-audit-tracker-live.azurewebsites.net/api/auth/callback`.
- Local redirect URI for development: `http://localhost:8787/api/auth/callback`.
- Application Graph permissions for the secure backend: `Mail.Send` for verification email and `Sites.ReadWrite.All` for Microsoft Lists storage.
- Admin consent if the tenant requires consent for SharePoint list read/write scopes.

Current supported live actions:

- Sign in with Microsoft.
- Request tracker access.
- Send and verify email codes.
- Approve or reject users from the admin panel.
- Save shared project records on the app server.
- Export the Microsoft Lists migration package.
- Switch approval and project storage to Microsoft Lists after the site/list IDs are configured.

The secure account gate now uses the backend server for Microsoft OAuth, Graph email verification codes, signed HTTP-only sessions, and admin approval endpoints. A new user starts with Microsoft-hosted sign-in, requests access with that Microsoft identity, receives a verification code by email, confirms the code in the tracker, and then waits for an Admin user to approve the profile. The backend can persist approval records in `server/data/access-users.json` for local testing or in the `Tracker Users` Microsoft List when `TRACKER_USER_STORE=microsoft-lists`, `TRACKER_USERS_SITE_ID`, and `TRACKER_USERS_LIST_ID` are configured.

Project records can also move from the app server file into Microsoft Lists. Set `TRACKER_PROJECT_STORE=microsoft-lists`, `TRACKER_PROJECTS_SITE_ID`, and `TRACKER_PROJECTS_LIST_ID` after the Audit Assignments list exists. The server writes the normal reporting columns plus `TrackerProjectJson`, which preserves comments, checklist state, status history, and activity events in the main assignment record while the child lists remain available for exports and reporting.

The sync client uses stable app keys for upserts:

- `TrackerAssignmentId` for assignment rows.
- `TrackerProjectJson` for full project persistence when backend project storage uses Microsoft Lists.
- `TeamMemberKey` for team-member rows.
- `TrackerChecklistItemId` for checklist rows.
- `TrackerCommentId`, `TrackerHistoryId`, `TrackerEventId`, and `TrackerUsername` for the other child/user lists.

This keeps repeated test syncs from duplicating rows when the same assignment package is pushed more than once.

## Activity Log Rules

The activity log is intentionally separate from comments and status history. Comments and status history are operational child lists; the activity log is the accountability ledger.

Each activity row should include:

- `TrackerEventId`: stable event ID from the app.
- `TrackerAssignmentId`: stable assignment link key.
- `AssignmentNumber`: human-readable assignment number for quick filtering.
- `OccurredAt`: timestamp normalized for SharePoint date/time columns.
- `EventType`: one of `field`, `stage`, `comment`, `document`, `checklist`, `team`, `finance`, or `system`.
- `ActorName` and `ActorUsername`: who took the action.
- `Summary` and `Detail`: readable action text.
- `FieldName`, `PreviousValue`, `NewValue`: populated when the action is a field/stage change.
- `SourceList` and `SourceRecordId`: link back to the operational child row when applicable.

## Implementation Order

1. Confirm live Microsoft sign-in and admin approval work with real Mosaic accounts.
2. Create the lists from the exported `graphListCreateRequests`.
3. Import `Audit Assignments` first.
4. Import child rows in this order: team members, checklist items, comments, status history, activity log.
5. Configure `Tracker Users` with the approved-account schema.
6. Validate that each child row has a matching `TrackerAssignmentId`.
7. Set `TRACKER_USERS_SITE_ID` and `TRACKER_USERS_LIST_ID`.
8. Switch backend approval records to `TRACKER_USER_STORE=microsoft-lists`.
9. Set `TRACKER_PROJECTS_SITE_ID` and `TRACKER_PROJECTS_LIST_ID` after `TrackerProjectJson` exists on the Audit Assignments list.
10. Switch backend project storage to `TRACKER_PROJECT_STORE=microsoft-lists`.
11. Add Power Automate flows only after live list writes and activity-log entries are stable.

## Why This Split Matters

The main assignment list stays clean for dashboards and reporting. High-volume details such as comments, checklist rows, stage movements, and audit events stay in child lists. That gives Power BI cleaner tables, gives Power Automate reliable trigger points, and gives managers a defensible change history without overloading the main assignment row.
