import { describe, expect, it } from "vitest";
import { errorEnvelope, successEnvelope } from "./envelope.js";
import { getRequiredErrorCodes, listErrors } from "./errors.js";
import { validateSchema } from "./schemas.js";

describe("errors and envelope", () => {
  it("contains all required error catalog codes", () => {
    const codes = listErrors().map((entry) => entry.code);
    expect(codes).toEqual(getRequiredErrorCodes());
    expect(codes).toContain("SECURITY_ZIP_BOMB_DETECTED");
    expect(codes).toContain("FEATURE_HIDDEN_FROM_AGENT");
  });

  it("creates schema-valid success envelopes", () => {
    const envelope = successEnvelope({ value: "ok" }, { command: "inspect", capabilitiesHash: "sha256:abc" });

    const validation = validateSchema("officegen.envelope@1.2", envelope);

    expect(envelope.ok).toBe(true);
    expect(validation.ok).toBe(true);
  });

  it("creates schema-valid error envelopes with recovery commands", () => {
    const envelope = errorEnvelope("FEATURE_DISABLED", {
      command: "design",
      availableCommands: ["inspect", "view", "edit"],
      nextSuggestedCommands: ["officegen capabilities --agent --json"],
      errorOptions: { feature: "design", command: "design capture" }
    });

    const validation = validateSchema("officegen.envelope@1.2", envelope);

    expect(envelope.ok).toBe(false);
    expect(envelope.error.feature).toBe("design");
    expect(envelope.availableCommands).toContain("inspect");
    expect(validation.ok).toBe(true);
  });

  it("rejects envelopes that do not match their ok discriminator", () => {
    expect(
      validateSchema("officegen.envelope@1.2", {
        schema: "officegen.envelope@1.2",
        ok: true,
        cliVersion: "1.2.0",
        pathsRedacted: true,
        warnings: [],
        diagnostics: [],
        artifacts: [],
        nextSuggestedCommands: []
      }).ok
    ).toBe(false);

    expect(
      validateSchema("officegen.envelope@1.2", {
        schema: "officegen.envelope@1.2",
        ok: false,
        cliVersion: "1.2.0",
        pathsRedacted: true,
        warnings: [],
        diagnostics: [],
        artifacts: [],
        nextSuggestedCommands: [],
        availableCommands: []
      }).ok
    ).toBe(false);
  });

  it("rejects error envelopes missing required error payload fields", () => {
    const validation = validateSchema("officegen.envelope@1.2", {
      schema: "officegen.envelope@1.2",
      ok: false,
      cliVersion: "1.2.0",
      pathsRedacted: true,
      error: { code: "FEATURE_DISABLED", message: "disabled" },
      warnings: [],
      diagnostics: [],
      artifacts: [],
      availableCommands: [],
      nextSuggestedCommands: []
    });

    expect(validation.ok).toBe(false);
    if (!validation.ok) {
      expect(validation.errors.some((error) => error.instancePath === "/error" && error.keyword === "required")).toBe(true);
    }
  });
});
