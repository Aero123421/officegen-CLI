import { describe, expect, it } from "vitest";
import { getBuiltinConfig } from "./config.js";
import { redactJson, redactPathsInText, redactSecretsInText } from "./redaction.js";

describe("redaction", () => {
  it("redacts secret-like tokens in text", () => {
    const result = redactSecretsInText("Authorization: Bearer abcdefghijklmnopqrstuvwxyz");

    expect(result.value).toContain("<redacted:secret-like-token>");
    expect(result.redactions).toHaveLength(1);
  });

  it("redacts project paths in text and JSON", () => {
    const config = getBuiltinConfig("substrate");
    config.paths.projectRoot = "D:\\codebase\\tool\\Officegen-CLI";
    config.paths.userConfigDir = "C:\\Users\\someone\\.officegen";
    const text = "Wrote D:\\codebase\\tool\\Officegen-CLI\\.officegen\\outputs\\final.pptx";
    const textResult = redactPathsInText(text, config);

    expect(textResult.value).toContain("<project>/.officegen/outputs/final.pptx");

    const jsonResult = redactJson(
      {
        path: "D:\\codebase\\tool\\Officegen-CLI\\deck.pptx",
        token: "Bearer abcdefghijklmnopqrstuvwxyz"
      },
      config
    );
    expect(jsonResult.value.path).toBe("<project>/deck.pptx");
    expect(jsonResult.value.token).toBe("Bearer <redacted:secret-like-token>");
    expect(jsonResult.redactions.length).toBeGreaterThanOrEqual(2);
  });
});
