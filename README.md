# Audit Assignment Tracker Prototype

A React prototype for an auditing firm to track assignments from intake through final report, invoice, and close-out. The prototype stores data in browser local storage only and uses mock sample records with fictional client / coverholder codes.

## Run locally

```bash
npm install
npm run dev
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

| Prototype field                                                                           | SharePoint column type                          | Power Platform notes                                                              |
| ----------------------------------------------------------------------------------------- | ----------------------------------------------- | --------------------------------------------------------------------------------- |
| `assignmentNumber`                                                                        | Single line of text                             | Unique business key and display title candidate                                   |
| `assignmentSource`                                                                        | Choice (`Email`, `DAM`)                         | Drives conditional quote and final submission automation                          |
| `assignmentType`                                                                          | Choice (`DCA`, `CH`, `MGA`, `Company Contract`) | Supports portfolio segmentation and reporting                                     |
| `auditEntity`                                                                             | Single line of text                             | Captures the free-form audited entity name shown on assignment records            |
| `clientCoverholderCode`                                                                   | Single line of text                             | Uses non-sensitive code rather than real client name                              |
| `broker`                                                                                  | Single line of text or lookup                   | Can become a Broker lookup list later                                             |
| `assignedAuditor`, `reviewer`                                                             | Person or Group                                 | Can support workload dashboards and approvals                                     |
| `currentStage`                                                                            | Choice                                          | Maps to Kanban buckets and Power BI funnel visuals                                |
| `assignmentStatus`, `quoteStatus`, `reportStatus`, `invoiceStatus`, `damSubmissionStatus` | Choice                                          | Suitable for Power Apps dropdowns and Power Automate conditions                   |
| `quoteAmount`                                                                             | Currency                                        | Feeds pipeline value reporting                                                    |
| `paymentReceived`                                                                         | Yes/No                                          | Supports invoice and payment received tracking                                    |
| Date fields                                                                               | Date and time                                   | Supports due, overdue, and SLA reporting                                          |
| Document receipt and completion fields                                                    | Yes/No                                          | Supports missing document blockers and checklist automation                       |
| `nextAction`, `blockers`                                                                  | Multiple lines of text                          | Manual action notes and exception management                                      |
| `statusHistory`                                                                           | Child list                                      | Recommended as a separate Assignment Status History list related by assignment ID |

A future SharePoint implementation should use an **Audit Assignments** list for the main project record and an **Audit Assignment Status History** child list for the stage movement audit trail. Power Automate flows can enforce the same gate rules used in this prototype before advancing stages.
