import assert from "node:assert/strict";
import test from "node:test";
import {
  parseWorkbookLinks,
  parseWorksheetRows,
} from "../server/contactWorkbookSources.mjs";

test("workbook links can be configured as newline or comma separated values", () => {
  const links = parseWorkbookLinks("https://one.example/a.xlsx\nhttps://one.example/b.xlsx, https://one.example/c.xlsx");

  assert.equal(links.length, 3);
  assert.equal(links[0].id, "workbook-1");
  assert.equal(links[2].url, "https://one.example/c.xlsx");
});

test("worksheet parser maps contact fields and special instructions", () => {
  const parsed = parseWorksheetRows({
    source: { id: "workbook-1", label: "Workbook 1" },
    workbookName: "Client contacts.xlsx",
    worksheetName: "Contacts",
    values: [
      [
        "Client Name",
        "Managing Agent",
        "Contact Name",
        "Email",
        "Phone",
        "Special Instructions",
      ],
      [
        "Example Client",
        "Example MA",
        "Sarah Smith",
        "sarah@example.com",
        "555-0100",
        "Always CC compliance before requesting Claims BDX.",
      ],
    ],
  });

  assert.equal(parsed.contacts.length, 1);
  assert.equal(parsed.contacts[0].company, "Example Client");
  assert.equal(parsed.contacts[0].managingAgent, "Example MA");
  assert.equal(parsed.contacts[0].email, "sarah@example.com");
  assert.deepEqual(parsed.contacts[0].specialInstructions, [
    {
      label: "Special Instructions",
      value: "Always CC compliance before requesting Claims BDX.",
    },
  ]);
});

test("worksheet parser supports client instruction sheet layout", () => {
  const parsed = parseWorksheetRows({
    source: { id: "workbook-1", label: "Workbook 1" },
    workbookName: "Client Instructions (A-C).xlsx",
    worksheetName: "Ascot",
    values: [
      [
        "Client Name/Address/Tel#",
        "DCA Primary Contact/Title/Email/Tel#",
        "COVER HOLDER Primary Contact/Title/Email/Tel#",
        "Audit Scope",
        "Deliverables",
        "TOBA",
      ],
      [
        "Ascot Group\n20 Fenchurch Street",
        "Danny Sambridge\nEmail: danny.sambridge@ascotgroup.com",
        "Daisy Murphy\nEmail: daisy.murphy@ascotgroup.com",
        "LMA",
        "Copy AscotDAAudit@ascotgroup.com on future correspondence.",
        "None provided",
      ],
      [],
      [
        "Report Submission Email",
        "Invoice Submission Email",
        "Onsite/Remote Preference",
        "Fees and Payment Terms",
        "Other",
        "Notes/Comments",
      ],
      [
        "AscotDAAudit@ascotgroup.com",
        "AscotDAAudit@ascotgroup.com",
        "Remote preferred",
        "",
        "",
        "Use the centralized inbox.",
      ],
    ],
  });

  assert.equal(parsed.contacts.length, 1);
  assert.equal(parsed.contacts[0].company, "Ascot Group");
  assert.equal(parsed.contacts[0].managingAgent, "Ascot Group");
  assert.equal(parsed.contacts[0].contactName, "Danny Sambridge");
  assert.equal(
    parsed.contacts[0].email,
    "danny.sambridge@ascotgroup.com; daisy.murphy@ascotgroup.com; AscotDAAudit@ascotgroup.com",
  );
  assert.equal(
    parsed.contacts[0].specialInstructions.some(
      (instruction) => instruction.label === "Deliverables",
    ),
    true,
  );
  assert.equal(
    parsed.contacts[0].specialInstructions.some(
      (instruction) => instruction.label === "Notes/Comments",
    ),
    true,
  );
});
