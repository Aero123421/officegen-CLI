import { mkdtemp, stat } from "node:fs/promises";
import { join } from "node:path";
import { tmpdir } from "node:os";
import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { edit } from "../src/edit.js";

describe("Office XML edit integration", () => {
  it("rolls back earlier part edits when an atomic batch fails later", async () => {
    const zip = new JSZip();
    zip.file("ppt/slides/slide1.xml", "<p:sld><p:sp><p:txBody><a:p><a:r><a:t>Alpha</a:t></a:r></a:p></p:txBody></p:sp></p:sld>");
    const bytes = await zip.generateAsync({ type: "uint8array" });

    const result = await edit(
      { data: bytes, format: "pptx" },
      [
        { op: "replaceText", from: "Alpha", to: "Beta" },
        { op: "setText", selector: { stableObjectId: "pptx:missing:shape:0001" }, text: "Should not survive" }
      ],
      { atomic: true, validateFirst: false }
    );

    expect(result).toMatchObject({
      changed: false,
      applied: 0,
      rolledBack: true,
      errors: [expect.objectContaining({ operationIndex: 1, reason: "not-found" })]
    });
    expect(result.opResults?.[0]).toMatchObject({ operationIndex: 0, applied: true });
    expect(result.caveats.join("\n")).toContain("EDIT_TRANSACTION_ROLLBACK: restored");
    expect(result.bytes).toBeUndefined();
  });

  it("records unprocessed operations as skipped-after-error when an atomic batch aborts", async () => {
    const zip = new JSZip();
    zip.file("ppt/slides/slide1.xml", "<p:sld><p:sp><p:txBody><a:p><a:r><a:t>Alpha</a:t></a:r></a:p></p:txBody></p:sp></p:sld>");
    const bytes = await zip.generateAsync({ type: "uint8array" });

    const result = await edit(
      { data: bytes, format: "pptx" },
      [
        { op: "replaceText", from: "Alpha", to: "Beta" },
        { op: "setText", selector: { stableObjectId: "pptx:missing:shape:0001" }, text: "Should not survive" },
        { op: "replaceText", from: "Gamma", to: "Delta" }
      ],
      { atomic: true, validateFirst: false }
    );

    expect(result).toMatchObject({
      changed: false,
      applied: 0,
      skipped: 2,
      rolledBack: true
    });
    expect(result.opResults?.[1]).toMatchObject({ operationIndex: 1, applied: false, reason: "not-found" });
    expect(result.opResults?.[2]).toMatchObject({ operationIndex: 2, applied: false, reason: "skipped-after-error" });
  });

  it("does not fall back to whole-XML replacement for attribute-only matches", async () => {
    const zip = new JSZip();
    zip.file("ppt/slides/slide1.xml", "<p:sld><p:cNvPr name=\"Alpha\"/><p:sp><p:txBody><a:p><a:r><a:t>Visible</a:t></a:r></a:p></p:txBody></p:sp></p:sld>");
    const bytes = await zip.generateAsync({ type: "uint8array" });

    const result = await edit(
      { data: bytes, format: "pptx" },
      [{ op: "replaceText", from: "Alpha", to: "Beta" }],
      { validateFirst: false }
    );

    expect(result.changed).toBe(false);
    expect(result.opResults?.[0]).toMatchObject({ applied: false, reason: "not-found" });
    expect(result.bytes).toBeUndefined();
  });

  it("updates matching text runs while preserving matching attributes", async () => {
    const zip = new JSZip();
    zip.file("ppt/slides/slide1.xml", "<p:sld><p:cNvPr name=\"Alpha\"/><p:sp><p:txBody><a:p><a:r><a:t>Alpha</a:t></a:r></a:p></p:txBody></p:sp></p:sld>");
    const bytes = await zip.generateAsync({ type: "uint8array" });

    const result = await edit(
      { data: bytes, format: "pptx" },
      [{ op: "replaceText", from: "Alpha", to: "Beta" }],
      { validateFirst: false }
    );

    const edited = await JSZip.loadAsync(result.bytes as Uint8Array);
    const xml = await edited.file("ppt/slides/slide1.xml")?.async("string");
    expect(result.changed).toBe(true);
    expect(xml).toContain("name=\"Alpha\"");
    expect(xml).toContain("<a:t>Beta</a:t>");
  });

  it("does not expose partial PPTX bytes when a later required op is not found", async () => {
    const zip = new JSZip();
    zip.file("ppt/slides/slide1.xml", "<p:sld><p:sp><p:txBody><a:p><a:r><a:t>Alpha</a:t></a:r></a:p></p:txBody></p:sp></p:sld>");
    const bytes = await zip.generateAsync({ type: "uint8array" });

    const result = await edit(
      { data: bytes, format: "pptx" },
      [
        { op: "replaceText", from: "Alpha", to: "Beta" },
        { op: "setText", selector: { stableObjectId: "pptx:missing:shape:0001" }, text: "Should not write" }
      ],
      { atomic: false, continueOnError: true, validateFirst: false }
    );

    expect(result.changed).toBe(false);
    expect(result.applied).toBe(0);
    expect(result.bytes).toBeUndefined();
    expect(result.errors?.[0]).toMatchObject({ operationIndex: 1, reason: "not-found" });
    expect(result.opResults?.[0]).toMatchObject({ operationIndex: 0, applied: true });
  });

  it("allows partial PPTX bytes only when allowPartial is explicit", async () => {
    const zip = new JSZip();
    zip.file("ppt/slides/slide1.xml", "<p:sld><p:sp><p:txBody><a:p><a:r><a:t>Alpha</a:t></a:r></a:p></p:txBody></p:sp></p:sld>");
    const bytes = await zip.generateAsync({ type: "uint8array" });

    const result = await edit(
      { data: bytes, format: "pptx" },
      [
        { op: "replaceText", from: "Alpha", to: "Beta" },
        { op: "setText", selector: { stableObjectId: "pptx:missing:shape:0001" }, text: "Should not write" }
      ],
      { atomic: false, continueOnError: true, validateFirst: false, allowPartial: true }
    );
    const edited = await JSZip.loadAsync(result.bytes as Uint8Array);
    const xml = await edited.file("ppt/slides/slide1.xml")?.async("string");

    expect(result.changed).toBe(true);
    expect(result.applied).toBe(1);
    expect(result.partial).toBe(true);
    expect(result.errors?.[0]).toMatchObject({ operationIndex: 1, reason: "not-found" });
    expect(xml).toContain("<a:t>Beta</a:t>");
  });

  it("does not expose bytes or write out files for dry-run Office XML edits", async () => {
    const zip = new JSZip();
    zip.file("ppt/slides/slide1.xml", "<p:sld><p:sp><p:txBody><a:p><a:r><a:t>Alpha</a:t></a:r></a:p></p:txBody></p:sp></p:sld>");
    const bytes = await zip.generateAsync({ type: "uint8array" });
    const dir = await mkdtemp(join(tmpdir(), "officegen-edit-"));
    const out = join(dir, "edited.pptx");

    const result = await edit(
      { data: bytes, format: "pptx" },
      [{ op: "replaceText", from: "Alpha", to: "Beta" }],
      { dryRun: true, out, validateFirst: false }
    );

    await expect(stat(out)).rejects.toMatchObject({ code: "ENOENT" });
    expect(result.changed).toBe(true);
    expect(result.applied).toBe(1);
    expect(result.out).toBeUndefined();
    expect(result.bytes).toBeUndefined();
    expect(result.patchPlan).toMatchObject({
      schema: "officegen.patchPlan@2",
      wouldWrite: false,
      operations: [expect.objectContaining({ operationIndex: 0, op: "replaceText", wouldApply: true })],
      expectedChangedParts: ["ppt/slides/slide1.xml"]
    });
    expect(result.patchPlan?.touchedParts[0]).toMatchObject({
      path: "ppt/slides/slide1.xml",
      change: "modified",
      beforeSha256: expect.stringMatching(/^sha256:/),
      afterSha256: expect.stringMatching(/^sha256:/),
      sourceFingerprint: expect.objectContaining({ algorithm: "sha256", path: "ppt/slides/slide1.xml" })
    });
    expect(result.patchPlan?.sourceFingerprint).toMatchObject({ algorithm: "sha256", byteLength: bytes.byteLength });
  });

  it("returns blocked patch-plan evidence for stale dry-run hashes", async () => {
    const zip = new JSZip();
    zip.file("ppt/slides/slide1.xml", "<p:sld><p:sp><p:txBody><a:p><a:r><a:t>Alpha</a:t></a:r></a:p></p:txBody></p:sp></p:sld>");
    const bytes = await zip.generateAsync({ type: "uint8array" });

    const result = await edit(
      { data: bytes, format: "pptx" },
      [{ op: "replaceText", from: "Alpha", to: "Beta" }],
      { dryRun: true, expectedInputSha256: "sha256:0000", validateFirst: true }
    );

    expect(result.changed).toBe(false);
    expect(result.patchPlan).toMatchObject({
      schema: "officegen.patchPlan@2",
      wouldWrite: false,
      blocked: [expect.objectContaining({
        reason: "stale-plan",
        evidence: expect.objectContaining({ field: "expectedInputSha256", wouldWrite: false })
      })]
    });
    expect(result.errors?.[0].message).toContain("blocked before write");
  });
});
