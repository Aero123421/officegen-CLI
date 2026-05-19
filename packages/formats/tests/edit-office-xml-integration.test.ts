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
  });
});
