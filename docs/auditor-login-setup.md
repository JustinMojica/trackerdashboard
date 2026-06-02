# Auditor Login And Access Setup

This setup is based on the default auditor names used in the tracker dashboard.

## Microsoft Sign-In

Users sign in with their company Microsoft account. New users must request access, confirm the emailed verification code, and wait for an admin to approve the profile.

## Permission Groups

| Group | Use For | Access Level |
| --- | --- | --- |
| Admin | System owners who manage tracker fields, templates, permissions, and process changes | Full control |
| Audit Manager | Managers who assign work, review status, approve quotes, and review outbound documents | Edit and approve |
| Auditor | Auditors who update assignment status, upload working papers, complete audit tasks, and draft findings | Edit assigned work |
| Finance | Finance users who create, send, and track invoices | Finance fields and invoice folders |
| Read Only | Leadership or stakeholders who need visibility without editing | View only |

## Internal User Setup

| Full Name | Username | Work Email | Role | Permission Group | Tracker Access | SharePoint Access | Notes |
| --- | --- | --- | --- | --- | --- | --- | --- |
| Lorraine Mojica | lorraine.mojica | lorraine.mojica@[company-domain] | Auditor | Auditor | Edit assigned rows | Assigned client folders | Default auditor directory record |
| Walter Aviles | walter.aviles | walter.aviles@[company-domain] | Auditor | Auditor | Edit assigned rows | Assigned client folders | Default auditor directory record |
| Leslie Domenech | leslie.domenech | leslie.domenech@[company-domain] | Auditor | Auditor | Edit assigned rows | Assigned client folders | Default auditor directory record |
| Mark James | mark.james | mark.james@[company-domain] | Audit Manager | Audit Manager | Edit all audit rows | All audit folders | Manager directory record |
| Justin Mojica | justin.mojica | justin.mojica@[company-domain] | Auditor / Admin | Admin | Full edit | Full control | System owner |
| Sheilah Couture | sheilah.couture | sheilah.couture@[company-domain] | Finance | Finance | Edit invoice/payment fields | Invoice folders | Finance directory record |
| Annabelle J. Crawford Mojica | annabelle.crawford.mojica | annabelle.crawford.mojica@[company-domain] | Read Only | Read Only | View only | View only | Read-only directory record; confirm preferred username/email format |
| Molly Aviles | molly.aviles | molly.aviles@[company-domain] | Auditor | Auditor | Edit assigned rows | Assigned client folders | In dashboard auditor list |
| Lindsie Guillermo | lindsie.guillermo | lindsie.guillermo@[company-domain] | Auditor | Auditor | Edit assigned rows | Assigned client folders | In dashboard auditor list |

## Recommended Tracker Fields

| Field | Purpose |
| --- | --- |
| Assigned Auditor | Shows who owns the audit work |
| Next Action Owner | Makes it clear who needs to act next |
| Last Updated By | Creates accountability for tracker changes |
| Last Updated Date | Shows whether the record is stale |
| Status | Supports reporting and later automation |
| SharePoint Folder Link | Keeps documents tied to the assignment |
| Quote Approved By | Prevents unapproved quote sends |
| Invoice Owner | Separates audit completion from billing responsibility |

## Role Rules

| Role | Tracker Access |
| --- | --- |
| Admin | View, create, edit, move stages, import JSON, reset data, and update finance fields |
| Audit Manager | View, create, edit, move stages, import JSON, reset data, and update finance fields |
| Auditor | View and update projects where they are lead or supporting auditor |
| Finance | View invoice/final-submission/closed records and update invoice/payment fields |
| Read Only | View all records without save actions |

Admins manage access and assignment names inside the dashboard. The user management panel supports:

- Microsoft access request approval
- full name, company email, role, and assignment visibility
- role changes
- active/inactive status
- default project visibility: role default, all projects, assigned projects, or finance records

## Naming Convention

Use a standard login identity format:

| Person Type | Email Format |
| --- | --- |
| Internal employees | firstname.lastname / firstname.lastname@[company-domain] |
| Shared mailbox only if needed | audits@[company-domain] |
| External clients/brokers | Do not create full internal logins yet |

## Setup Checklist

1. Create one Microsoft 365 login for each internal user.
2. Add each user to exactly one primary permission group.
3. Add secondary access only when needed, such as Finance plus Read Only.
4. Confirm each user appears consistently in the tracker name fields.
5. Create or confirm the SharePoint folder permissions.
6. Test with one auditor account before rolling out to the full team.
7. Remove any shared credentials from the process.
8. Confirm each real tester can complete Microsoft sign-in, email code confirmation, and admin approval.

## External Access Rule

Do not create full logins for brokers, coverholders, or clients at this stage. Use controlled SharePoint links, upload folders, or request forms until there is a proper portal or external-user process.
