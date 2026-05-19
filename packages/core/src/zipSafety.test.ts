import JSZip from "jszip";
import { describe, expect, it, vi } from "vitest";
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

  it("does not expand oversized XML parts as strings", async () => {
    const config = getBuiltinConfig("substrate");
    config.security.untrustedInput.maxSingleXmlPartBytes = 32;

    const prototypeZip = new JSZip();
    prototypeZip.file("__prototype.xml", "<root />");
    const xmlPrototype = Object.getPrototypeOf(prototypeZip.files["__prototype.xml"]);
    const asyncSpy = vi.spyOn(xmlPrototype, "async");

    const zip = new JSZip();
    zip.file("word/document.xml", `<root>${"x".repeat(128)}</root>`);
    const buffer = await zip.generateAsync({ type: "nodebuffer" });

    try {
      const report = await inspectZipSafety(buffer, config);
      const codes = report.warnings.map((item) => item.code);

      expect(report.ok).toBe(false);
      expect(codes).toContain("ZIP_XML_PART_TOO_LARGE");
      expect(asyncSpy).not.toHaveBeenCalledWith("string");
    } finally {
      asyncSpy.mockRestore();
    }
  });
});
