import { describe, expect, it } from "vitest";
import { getBuiltinConfig } from "./config.js";
import { redactJson, redactPathsInText, redactSecretsInText } from "./redaction.js";

describe("redaction", () => {
  it("redacts secret-like tokens in text", () => {
    const result = redactSecretsInText("Authorization: Bearer abcdefghijklmnopqrstuvwxyz");

    expect(result.value).toContain("<redacted:secret-like-token>");
    expect(result.redactions).toHaveLength(1);
  });

  it("redacts secret key-value pairs without leaking values", () => {
    const result = redactSecretsInText(
      'OPENAI_API_KEY=sk-proj-abcdefghijklmnop client_secret = "supersecretvalue" url=https://example.test/cb?api_key=querysecretvalue'
    );

    expect(result.value).not.toContain("sk-proj-abcdefghijklmnop");
    expect(result.value).not.toContain("supersecretvalue");
    expect(result.value).not.toContain("querysecretvalue");
    expect(result.value).toContain("OPENAI_API_KEY=<redacted:secret-like-token>");
    expect(result.redactions.length).toBeGreaterThanOrEqual(3);
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

  it("redacts Windows, POSIX, UNC, home, and configured roots independent of host OS", () => {
    const config = getBuiltinConfig("substrate");
    config.paths.projectRoot = "D:\\codebase\\tool\\Officegen-CLI";
    config.paths.userConfigDir = "C:\\Users\\someone\\.officegen";

    const result = redactPathsInText(
      [
        "project=D:/codebase/tool/Officegen-CLI/src/index.ts",
        "unix=/srv/build/private/report.docx",
        "unc=\\\\fileserver\\share\\team\\budget.xlsx",
        "home=~/Desktop/private.txt",
        "userConfig=C:\\Users\\someone\\.officegen\\config.json"
      ].join(" "),
      config
    );

    expect(result.value).toContain("<project>/src/index.ts");
    expect(result.value).toContain("<userConfig>/config.json");
    expect(result.value).not.toContain("D:/codebase/tool/Officegen-CLI");
    expect(result.value).not.toContain("/srv/build/private");
    expect(result.value).not.toContain("\\\\fileserver\\share");
    expect(result.value).not.toContain("~/Desktop");
    expect(result.value).not.toContain("C:\\Users\\someone");
    expect(result.redactions.length).toBeGreaterThanOrEqual(5);
  });

  it("does not redact paths inside SVG/XML payload strings", () => {
    const config = getBuiltinConfig("substrate");
    config.paths.projectRoot = "D:\\codebase\\tool\\Officegen-CLI";
    const svg = '<svg xmlns="http://www.w3.org/2000/svg"><path d="M10/20 L30/40"/><text>D:/codebase/tool/Officegen-CLI</text></svg>';

    const result = redactJson({ svg, outPath: "D:\\codebase\\tool\\Officegen-CLI\\out.svg" }, config);

    expect(result.value.svg).toBe(svg);
    expect(result.value.outPath).toBe("<project>/out.svg");
  });

  it("does not corrupt relative OOXML relationship targets while redacting absolute POSIX paths", () => {
    const config = getBuiltinConfig("substrate");
    const result = redactPathsInText("../media/missing.png and /srv/build/private/report.docx", config);

    expect(result.value).toContain("../media/missing.png");
    expect(result.value).toContain("<absolutePath>");
  });

  it("still redacts ordinary content and generic markup strings", () => {
    const config = getBuiltinConfig("substrate");
    config.paths.projectRoot = "D:\\codebase\\tool\\Officegen-CLI";

    const result = redactJson({
      content: "D:\\codebase\\tool\\Officegen-CLI\\secret.txt",
      htmlSnippet: "<div>D:\\codebase\\tool\\Officegen-CLI\\secret.txt</div>"
    }, config);

    expect(result.value.content).toBe("<project>/secret.txt");
    expect(result.value.htmlSnippet).toContain("<project>/secret.txt");
  });
});
