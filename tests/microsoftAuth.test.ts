import assert from "node:assert/strict";
import test from "node:test";
import {
  hasMicrosoftAuthConfig,
  microsoftAuthScopeLabel,
  sanitizeMicrosoftAuthConfig,
} from "../src/microsoftAuth.js";

test("Microsoft auth config trims saved Entra values", () => {
  const config = sanitizeMicrosoftAuthConfig({
    clientId: "  client-id  ",
    tenantId: "  example.com  ",
  });

  assert.equal(config.clientId, "client-id");
  assert.equal(config.tenantId, "example.com");
  assert.equal(hasMicrosoftAuthConfig(config), true);
});

test("Microsoft auth requires a client ID", () => {
  assert.equal(
    hasMicrosoftAuthConfig({ clientId: "", tenantId: "example.com" }),
    false,
  );
  assert.ok(microsoftAuthScopeLabel().includes("Sites.ReadWrite.All"));
});
