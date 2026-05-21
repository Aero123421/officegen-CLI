import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { diffDocuments, inspect, render } from "@officegen/formats";

import { applyLayoutConstraints } from "./layout.js";

describe("layout apply PPTX geometry", () => {
  it("mutates the selected object bounds and reports the same geometry change in diff", async () => {
    const dir = await mkdtemp(path.join(os.tmpdir(), "officegen-layout-"));
    try {
      const before = await render({ title: "Before", slides: [{ title: "Slide", body: "Alpha" }] }, { target: "pptx" });
      const beforeBytes = before.bytes!;
      const beforePath = path.join(dir, "before.pptx");
      const afterPath = path.join(dir, "after.pptx");
      await writeFile(beforePath, beforeBytes);

      const beforeInspect = await inspect({ data: beforeBytes, format: "pptx" });
      const target = beforeInspect.objectMap.find((entry) => entry.kind === "shape" && entry.text === "Alpha");
      const title = beforeInspect.objectMap.find((entry) => entry.kind === "shape" && entry.text === "Slide");
      expect(target?.bounds).toBeDefined();
      expect(title?.bounds).toBeDefined();

      const result = await applyLayoutConstraints({
        boxes: [{
          id: target!.stableObjectId,
          x: target!.bounds!.x,
          y: target!.bounds!.y,
          width: target!.bounds!.width,
          height: target!.bounds!.height
        }],
        constraints: [{ id: target!.stableObjectId, minWidth: target!.bounds!.width + 50 }],
        targetPath: beforePath,
        outputPath: afterPath
      });

      expect((result as any).editResult.changed).toBe(true);
      expect((result as any).editResult.applied).toBe(1);

      const afterBytes = await readFile(afterPath);
      const afterInspect = await inspect({ data: afterBytes, format: "pptx" });
      expect(afterInspect.objectMap.find((entry) => entry.stableObjectId === target!.stableObjectId)?.bounds?.width).toBe(target!.bounds!.width + 50);
      expect(afterInspect.objectMap.find((entry) => entry.stableObjectId === title!.stableObjectId)?.bounds).toEqual(title!.bounds);

      const diff = await diffDocuments({ data: beforeBytes, format: "pptx" }, { data: afterBytes, format: "pptx" });
      expect(diff.changed).toBe(true);
      expect(diff.summary.changedGeometryObjects).toBe(1);
      expect(diff.semantic.changedGeometry[0]?.stableObjectId).toBe(target!.stableObjectId);
      expect(diff.semantic.changedGeometry[0]?.delta.width).toBe(50);
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });
});
