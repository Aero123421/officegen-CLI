import JSZip from "jszip";
import { describe, expect, it, vi } from "vitest";
import { loadZip, normalizeInput } from "./shared.js";

describe("shared zip loading", () => {
  it("rejects unsafe zip metadata before JSZip load", async () => {
    const zip = new JSZip();
    zip.file("../evil.txt", "escape");
    const bytes = await zip.generateAsync({ type: "uint8array" });
    const input = await normalizeInput(bytes);
    const loadSpy = vi.spyOn(JSZip, "loadAsync");

    try {
      await expect(loadZip(input)).rejects.toThrow(/ZIP_PATH_TRAVERSAL/);
      expect(loadSpy).not.toHaveBeenCalled();
    } finally {
      loadSpy.mockRestore();
    }
  });
});
