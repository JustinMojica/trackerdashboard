# Microsoft Lists Storage Schema

This is the target central-storage structure for the audit assignment tracker. The app can still run from browser storage, but Admin and Audit Manager users can also configure a Microsoft Lists connection for live Graph test, sync, and load actions. The exported Microsoft Lists package still contains the list schemas, Graph-style list creation payloads, and seed rows.

## Target Lists

| List | Purpose | Key fields |
| ---- | ------- | ---------- |
| Audit Assignments | Main assignment record | `TrackerAssignmentId`, `AssignmentNumber`, `CurrentStage`, `AssignmentStatus`, document flags, finance fields, due dates |
| Audit Team Members | Lead/support workload rows | `TeamMemberKey`, `TrackerAssignmentId`, `PersonName`, `TeamRole`, `ActiveOnAssignment` |
| Audit Comments | Card comments | `TrackerCommentId`, `TrackerAssignmentId`, `CommentCreatedAt`, `CommentAuthor`, `CommentBody` |
| Audit Checklist Items | Per-stage checklist state | `TrackerChecklistItemId`, `TrackerAssignmentId`, `ChecklistKey`, `ChecklistStage`, `ChecklistItem`, `Completed` |
| Audit Status History | Stage movement history | `TrackerHistoryId`, `TrackerAssignmentId`, `ChangedAt`, `ChangedBy`, `FromStage`, `ToStage`, `StageNote` |
| Audit Activity Log | Append-only accountability log | `TrackerEventId`, `TrackerAssignmentId`, `OccurredAt`, `EventType`, `ActorName`, `Summary`, `Detail`, previous/new values |
| Tracker Users | Prototype role map and account access gate | `TrackerUsername`, `FullName`, `Email`, `Role`, `PermissionGroup`, `Active`, `EmailVerified`, `AccessRequestStatus`, `DefaultVisibility`, `RequestedAt`, `ApprovedAt`, `ApprovedBy`, `RejectionReason` |

## Live Connection Mode

The live connection panel stores the SharePoint site ID, list IDs, Microsoft Entra application client ID, and tenant value in browser storage. Microsoft Graph access tokens come from MSAL Browser and are session-only in React state.

The Microsoft Entra app registration needs:

- Platform type: Single-page application.
- Redirect URI: the app origin used for testing, such as `http://127.0.0.1:5173`.
- Delegated Graph permissions: `User.Read` and `Sites.ReadWrite.All`.
- Admin consent if the tenant requires consent for SharePoint list read/write scopes.

Current supported live actions:

- Sign in with Microsoft.
- Refresh the Microsoft Graph token.
- Sign out of the Microsoft session.
- Test the configured SharePoint site and Audit Assignments list.
- Push the current migration package to configured Microsoft Lists.
- Load assignment rows from the configured Audit Assignments list into the browser prototype.

The secure account gate now uses the backend server for Microsoft OAuth, Graph email verification codes, signed HTTP-only sessions, and admin approval endpoints. A new user starts with Microsoft-hosted sign-in, requests access with that Microsoft identity, receives a verification code by email, confirms the code in the tracker, and then waits for an Admin user to approve the profile. The current backend persists approval records in `server/data/access-users.json`; the next production step is moving that store into Microsoft Lists or Dataverse.

The sync client uses stable app keys for upserts:

- `TrackerAssignmentId` for assignment rows.
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

1. Create the lists from the exported `graphListCreateRequests`.
2. Import `Audit Assignments` first.
3. Import child rows in this order: team members, checklist items, comments, status history, activity log.
4. Import `Tracker Users` only for prototype testing; replace with Microsoft 365 groups before production.
5. Validate that each child row has a matching `TrackerAssignmentId`.
6. Connect the app to Microsoft Graph after the list structure is approved.
7. Validate the Microsoft Entra sign-in with a real tenant app registration.
8. Move backend approval records from local server JSON into Microsoft Lists or Dataverse.
9. Add a setup health check that tests Microsoft OAuth, Graph email delivery, and all seven configured lists.
10. Add Power Automate flows only after live list writes and activity-log entries are stable.

## Why This Split Matters

The main assignment list stays clean for dashboards and reporting. High-volume details such as comments, checklist rows, stage movements, and audit events stay in child lists. That gives Power BI cleaner tables, gives Power Automate reliable trigger points, and gives managers a defensible change history without overloading the main assignment row.
