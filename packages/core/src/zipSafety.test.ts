import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { getBuiltinConfig } from "./config.js";
import { inspectZipSafety } from "./zipSafety.js";

describe("zip safety", () => {
  it("detects traversal, XML entity, external relationship, macro, and size limit risks", async () => {
    const config = getBuiltinConfig("substrate");
    config.security.untrustedInput.maxZipExpandedBytes = 10;
    const zip = new JSZip();
    zip.file("../evil.txt", "escape");
    zip.file("[Content_Types].xml", "<!DOCTYPE x [<!ENTITY e SYSTEM 'file:///etc/passwd'>]><Types />");
    zip.file("_rels/.rels", '<Relationships><Relationship TargetMode="External" Target="https://example.com/a.png" /></Relationships>');
    zip.file("ppt/vbaProject.bin", "macro");

    const buffer = await zip.generateAsync({ type: "nodebuffer" });
    const report = await inspectZipSafety(buffer, config);
    const codes = report.warnings.map((item) => item.code);

    expect(report.ok).toBe(false);
    expect(report.hasMacros).toBe(true);
    expect(codes).toEqual(
      expect.arrayContaining([
        "ZIP_PATH_TRAVERSAL",
        "ZIP_XML_ENTITY_DENIED",
        "ZIP_EXTERNAL_RELATIONSHIP",
        "ZIP_MACRO_DETECTED",
        "ZIP_EXPANDED_BYTES_EXCEEDED"
      ])
    );
  });
});
