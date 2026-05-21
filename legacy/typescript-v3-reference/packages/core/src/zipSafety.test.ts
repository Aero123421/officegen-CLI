import JSZip from "jszip";
import { describe, expect, it, vi } from "vitest";
import { getBuiltinConfig } from "./config.js";
import { inspectZipSafety } from "./zipSafety.js";
import type { OfficegenConfig, ZipSafetyWarning } from "./types.js";

const LOCAL_FILE_HEADER_SIGNATURE = 0x04034b50;
const CENTRAL_DIRECTORY_SIGNATURE = 0x02014b50;

describe("zip safety", () => {
  it("rejects traversal and aggregate size metadata before loading with JSZip", async () => {
    const config = getBuiltinConfig("substrate");
    config.security.untrustedInput.maxZipExpandedBytes = 10;
    const zip = new JSZip();
    zip.file("../evil.txt", "escape");
    zip.file("ppt/vbaProject.bin", "macro");

    const buffer = await zip.generateAsync({ type: "nodebuffer" });
    const report = await inspectWithoutJsZipLoad(buffer, config);
    const reportCodes = codes(report.warnings);

    expect(report.ok).toBe(false);
    expect(report.hasMacros).toBe(true);
    expect(reportCodes).toEqual(expect.arrayContaining(["ZIP_PATH_TRAVERSAL", "ZIP_MACRO_DETECTED", "ZIP_EXPANDED_BYTES_EXCEEDED"]));
  });

  it("detects XML entity and external relationship risks after metadata passes", async () => {
    const config = getBuiltinConfig("substrate");
    const zip = new JSZip();
    zip.file("[Content_Types].xml", "<!DOCTYPE x [<!ENTITY e SYSTEM 'file:///etc/passwd'>]><Types />");
    zip.file("_rels/.rels", '<Relationships><Relationship TargetMode="External" Target="https://example.com/a.png" /></Relationships>');

    const buffer = await zip.generateAsync({ type: "nodebuffer" });
    const loadSpy = vi.spyOn(JSZip, "loadAsync");
    try {
      const report = await inspectZipSafety(buffer, config);
      const reportCodes = codes(report.warnings);

      expect(loadSpy).toHaveBeenCalledTimes(1);
      expect(report.ok).toBe(false);
      expect(reportCodes).toEqual(expect.arrayContaining(["ZIP_XML_ENTITY_DENIED", "ZIP_EXTERNAL_RELATIONSHIP"]));
    } finally {
      loadSpy.mockRestore();
    }
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
      const report = await inspectWithoutJsZipLoad(buffer, config);
      const reportCodes = codes(report.warnings);

      expect(report.ok).toBe(false);
      expect(reportCodes).toContain("ZIP_XML_PART_TOO_LARGE");
      expect(asyncSpy).not.toHaveBeenCalledWith("string");
    } finally {
      asyncSpy.mockRestore();
    }
  });

  it("rejects duplicate central directory names before loading with JSZip", async () => {
    const config = getBuiltinConfig("substrate");
    const zip = new JSZip();
    zip.file("a.txt", "one");
    zip.file("b.txt", "two");

    const buffer = await zip.generateAsync({ type: "nodebuffer" });
    const duplicated = replaceAscii(buffer, "b.txt", "a.txt");
    const report = await inspectWithoutJsZipLoad(duplicated, config);

    expect(report.ok).toBe(false);
    expect(codes(report.warnings)).toContain("ZIP_DUPLICATE_ENTRY");
  });

  it("rejects encrypted entry flags before loading with JSZip", async () => {
    const config = getBuiltinConfig("substrate");
    const zip = new JSZip();
    zip.file("doc.txt", "secret");

    const buffer = await zip.generateAsync({ type: "nodebuffer" });
    const encrypted = withZipFlags(buffer, 0x0001);
    const report = await inspectWithoutJsZipLoad(encrypted, config);

    expect(report.ok).toBe(false);
    expect(codes(report.warnings)).toContain("ZIP_ENCRYPTED_ENTRY");
  });

  it("rejects zip64 size markers before loading with JSZip", async () => {
    const config = getBuiltinConfig("substrate");
    const zip = new JSZip();
    zip.file("doc.txt", "content");

    const buffer = await zip.generateAsync({ type: "nodebuffer" });
    const zip64 = withCentralDirectoryUInt32(buffer, 24, 0xffffffff);
    const report = await inspectWithoutJsZipLoad(zip64, config);

    expect(report.ok).toBe(false);
    expect(codes(report.warnings)).toContain("ZIP64_UNSUPPORTED");
  });

  it("rejects too many entries before loading with JSZip", async () => {
    const config = getBuiltinConfig("substrate");
    config.security.untrustedInput.maxZipEntries = 1;
    const zip = new JSZip();
    zip.file("one.txt", "one");
    zip.file("two.txt", "two");

    const buffer = await zip.generateAsync({ type: "nodebuffer" });
    const report = await inspectWithoutJsZipLoad(buffer, config);

    expect(report.ok).toBe(false);
    expect(codes(report.warnings)).toContain("ZIP_ENTRY_LIMIT_EXCEEDED");
  });
});

async function inspectWithoutJsZipLoad(input: Buffer | Uint8Array, config: OfficegenConfig) {
  const loadSpy = vi.spyOn(JSZip, "loadAsync");
  try {
    const report = await inspectZipSafety(input, config);
    expect(loadSpy).not.toHaveBeenCalled();
    return report;
  } finally {
    loadSpy.mockRestore();
  }
}

function codes(warnings: ZipSafetyWarning[]): ZipSafetyWarning["code"][] {
  return warnings.map((item) => item.code);
}

function replaceAscii(input: Uint8Array, from: string, to: string): Uint8Array {
  if (from.length !== to.length) throw new Error("Replacement must preserve zip offsets.");
  const output = Uint8Array.from(input);
  const fromBytes = Buffer.from(from, "ascii");
  const toBytes = Buffer.from(to, "ascii");
  for (let index = 0; index <= output.byteLength - fromBytes.byteLength; index += 1) {
    if (fromBytes.every((byte, offset) => output[index + offset] === byte)) {
      output.set(toBytes, index);
    }
  }
  return output;
}

function withZipFlags(input: Uint8Array, flags: number): Uint8Array {
  const output = Uint8Array.from(input);
  const view = new DataView(output.buffer, output.byteOffset, output.byteLength);
  for (const offset of findSignatures(view, LOCAL_FILE_HEADER_SIGNATURE)) {
    view.setUint16(offset + 6, view.getUint16(offset + 6, true) | flags, true);
  }
  for (const offset of findSignatures(view, CENTRAL_DIRECTORY_SIGNATURE)) {
    view.setUint16(offset + 8, view.getUint16(offset + 8, true) | flags, true);
  }
  return output;
}

function withCentralDirectoryUInt32(input: Uint8Array, fieldOffset: number, value: number): Uint8Array {
  const output = Uint8Array.from(input);
  const view = new DataView(output.buffer, output.byteOffset, output.byteLength);
  const [centralDirectoryOffset] = findSignatures(view, CENTRAL_DIRECTORY_SIGNATURE);
  if (centralDirectoryOffset === undefined) throw new Error("Missing central directory.");
  view.setUint32(centralDirectoryOffset + fieldOffset, value, true);
  return output;
}

function findSignatures(view: DataView, signature: number): number[] {
  const offsets: number[] = [];
  for (let offset = 0; offset <= view.byteLength - 4; offset += 1) {
    if (view.getUint32(offset, true) === signature) offsets.push(offset);
  }
  return offsets;
}
