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
