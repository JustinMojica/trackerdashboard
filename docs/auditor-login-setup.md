# Auditor Login And Access Setup

This setup is based on the default auditor names and seeded project records in the tracker dashboard.

## Permission Groups

| Group | Use For | Access Level |
| --- | --- | --- |
| Admin | System owners who manage tracker fields, templates, permissions, and process changes | Full control |
| Audit Manager | Managers who assign work, review status, approve quotes, and review outbound documents | Edit and approve |
| Auditor | Auditors who update assignment status, upload working papers, complete audit tasks, and draft findings | Edit assigned work |
| Finance | Finance users who create, send, and track invoices | Finance fields and invoice folders |
| Read Only | Leadership or stakeholders who need visibility without editing | View only |

## Internal User Setup

| Full Name | Work Email | Role | Permission Group | Tracker Access | SharePoint Access | Notes |
| --- | --- | --- | --- | --- | --- | --- |
| Lorraine Mojica | lorraine.mojica@[company-domain] | Auditor | Auditor | Edit assigned rows | Assigned client folders | Lead auditor on seeded dashboard projects |
| Walter Aviles | walter.aviles@[company-domain] | Auditor | Auditor | Edit assigned rows | Assigned client folders | Lead/supporting auditor on seeded dashboard projects |
| Leslie Domenech | leslie.domenech@[company-domain] | Auditor | Auditor | Edit assigned rows | Assigned client folders | Supporting auditor on seeded dashboard projects |
| Mark James | mark.james@[company-domain] | Auditor | Auditor | Edit assigned rows | Assigned client folders | In dashboard auditor list |
| Justin Mojica | justin.mojica@[company-domain] | Auditor / Admin | Admin | Full edit | Full control | Lead auditor on seeded dashboard project and likely system owner |
| Sheilah Couture | sheilah.couture@[company-domain] | Auditor | Auditor | Edit assigned rows | Assigned client folders | In dashboard auditor list |
| Annabelle J. Crawford Mojica | annabelle.crawford.mojica@[company-domain] | Auditor | Auditor | Edit assigned rows | Assigned client folders | In dashboard auditor list; confirm preferred email format |
| Molly Aviles | molly.aviles@[company-domain] | Auditor | Auditor | Edit assigned rows | Assigned client folders | In dashboard auditor list |
| Lindsie Guillermo | lindsie.guillermo@[company-domain] | Auditor | Auditor | Edit assigned rows | Assigned client folders | In dashboard auditor list |

## Seeded Project Visibility

| Assignment | Audit Entity | Lead Auditor | Supporting Auditor(s) | Stage |
| --- | --- | --- | --- | --- |
| AA-2026-0142 | Northbridge Coverholder Operations | Lorraine Mojica | Walter Aviles | Quote |
| AA-2026-0148 | Harbor Specialty Program | Walter Aviles | Leslie Domenech | Pre-Audit |
| AA-2026-0155 | Summit Claims Administration | Lorraine Mojica |  | Findings |
| AA-2026-0161 | Cedar Binding Authority | Justin Mojica |  | Final Submission |

Reviewer names also appear in the seeded project records, but they are not part of the dashboard's default auditor list. Confirm whether reviewer-only users should receive Audit Manager logins.

## Recommended Tracker Fields

| Field | Purpose |
| --- | --- |
| Assigned Auditor | Shows who owns the audit work |
| Reviewer | Shows who reviews work before close-out |
| Next Action Owner | Makes it clear who needs to act next |
| Last Updated By | Creates accountability for tracker changes |
| Last Updated Date | Shows whether the record is stale |
| Status | Supports reporting and later automation |
| SharePoint Folder Link | Keeps documents tied to the assignment |
| Quote Approved By | Prevents unapproved quote sends |
| Invoice Owner | Separates audit completion from billing responsibility |

## Naming Convention

Use a standard login identity format:

| Person Type | Email Format |
| --- | --- |
| Internal employees | firstname.lastname@[company-domain] |
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

## External Access Rule

Do not create full logins for brokers, coverholders, or clients at this stage. Use controlled SharePoint links, upload folders, or request forms until there is a proper portal or external-user process.
