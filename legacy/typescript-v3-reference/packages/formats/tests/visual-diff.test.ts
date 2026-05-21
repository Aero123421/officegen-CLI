import { mkdtemp, rm, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { createCanvas } from "@napi-rs/canvas";
import { getBuiltinConfig } from "@officegen/core";
import { describe, expect, it } from "vitest";
import { comparePngPixels, compareRasterPixels } from "../src/visualDiff.js";

describe("native visual diff foundations", () => {
  it("compares raw raster buffers with changed pixels, bounding boxes, and thresholds", () => {
    const before = rgba(2, 2);
    const after = rgba(2, 2);
    after[4 * 3] = 5;

    expect(compareRasterPixels({ width: 2, height: 2, data: before }, { width: 2, height: 2, data: before })).toMatchObject({
      status: "compared",
      changed: false,
      changedPixels: 0,
      boundingBox: undefined
    });
    expect(compareRasterPixels({ width: 2, height: 2, data: before }, { width: 2, height: 2, data: after }, { threshold: 5 })).toMatchObject({
      changed: false,
      changedPixels: 0
    });
    expect(compareRasterPixels({ width: 2, height: 2, data: before }, { width: 2, height: 2, data: after }, { threshold: 4 })).toMatchObject({
      changed: true,
      changedPixels: 1,
      boundingBox: { x: 1, y: 1, width: 1, height: 1 }
    });
  });

  it("decodes small PNG buffers for deterministic pixel diffs", async () => {
    const before = pngWithPixel();
    const after = pngWithPixel({ x: 2, y: 0, color: "#ff0000" });
    const diff = await comparePngPixels(before, after);

    expect(diff.status).toBe("compared");
    expect(diff.changed).toBe(true);
    expect(diff.changedPixels).toBe(1);
    expect(diff.boundingBox).toEqual({ x: 2, y: 0, width: 1, height: 1 });
  });

  it("uses raster pixel diff for PDF visual comparisons", async () => {
    const { diffDocuments, render } = await import("../src/index.js");
    const before = await render({ title: "PDF", sections: [{ title: "Page", body: "Alpha" }] }, { target: "pdf" });
    const after = await render({ title: "PDF", sections: [{ title: "Page", body: "Beta" }] }, { target: "pdf" });
    const diff = await diffDocuments({ data: before.bytes, format: "pdf" }, { data: after.bytes, format: "pdf" }, { visual: true });

    expect(diff.visual?.status).toBe("compared");
    expect(diff.visual?.kind).toBe("raster-pixel");
    expect(diff.visual?.pageScores[0]?.pixelDiff?.changedPixels).toBeGreaterThan(0);
    expect(diff.summary.visualRegressionScore).toBeGreaterThan(0);
  });

  it("returns blocked native visual diff results when the renderer is unavailable", async () => {
    const { diffDocuments, render } = await import("../src/index.js");
    const rendered = await render({ title: "Native", slides: [{ title: "Same", body: "Body" }] }, { target: "pptx" });
    const config = getBuiltinConfig("substrate");
    const diff = await diffDocuments(
      { data: rendered.bytes, format: "pptx" },
      { data: rendered.bytes, format: "pptx" },
      { visual: true, native: true, config }
    );

    expect(diff.visual?.status).toBe("blocked");
    expect(diff.visual?.message).toMatch(/Native|renderer|disabled|requires|LibreOffice|Office/i);
    expect(diff.summary.visualRegressionScore).toBeUndefined();
  });

  it("does not pass verify when native renderer verification cannot run", async () => {
    const { render, verify } = await import("../src/index.js");
    const rendered = await render({ title: "Verify Native", slides: [{ title: "Ready", body: "Body" }] }, { target: "pptx" });
    const dir = await mkdtemp(path.join(os.tmpdir(), "officegen-verify-native-test-"));
    try {
      const input = path.join(dir, "input.pptx");
      await writeFile(input, rendered.bytes);
      const result = await verify(input, { native: true, config: getBuiltinConfig("substrate"), visual: true });

      expect(result.readiness).toBe("blocked");
      expect(result.nativeRenderer?.ok).toBe(false);
      expect(result.blockingIssues.join("\n")).toContain("NATIVE_RENDERER_BLOCKED");
      expect(result.visualDiff).toMatchObject({ status: "skipped", expectedDiffOnly: false });
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });
});

function rgba(width: number, height: number): Uint8Array {
  const data = new Uint8Array(width * height * 4);
  for (let index = 3; index < data.length; index += 4) data[index] = 255;
  return data;
}

function pngWithPixel(pixel?: { x: number; y: number; color: string }): Uint8Array {
  const canvas = createCanvas(3, 2);
  const context = canvas.getContext("2d");
  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, 3, 2);
  if (pixel) {
    context.fillStyle = pixel.color;
    context.fillRect(pixel.x, pixel.y, 1, 1);
  }
  return new Uint8Array(canvas.encodeSync("png"));
}
