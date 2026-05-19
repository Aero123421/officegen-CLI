import { appendFile, mkdtemp, readFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { describe, expect, it, vi } from "vitest";
import { appendTrace } from "./run.js";
import type { RunFolder } from "./types.js";

vi.mock("node:fs/promises", async (importOriginal) => {
  const actual = await importOriginal<typeof import("node:fs/promises")>();
  return {
    ...actual,
    appendFile: vi.fn(actual.appendFile)
  };
});

function runFolder(root: string): RunFolder {
  return {
    runId: "trace-test",
    root,
    inputDir: path.join(root, "input"),
    irDir: path.join(root, "ir"),
    opsDir: path.join(root, "ops"),
    viewsDir: path.join(root, "views"),
    diagnosticsDir: path.join(root, "diagnostics"),
    outputDir: path.join(root, "output"),
    backupDir: path.join(root, "backup"),
    logsDir: path.join(root, "logs"),
    tracePath: path.join(root, "trace.jsonl"),
    runJsonPath: path.join(root, "run.json"),
    manifestPath: path.join(root, "manifest.json")
  };
}

describe("run trace append", () => {
  it("uses appendFile for large JSONL traces instead of read-all-write", async () => {
    const root = await mkdtemp(path.join(tmpdir(), "officegen-trace-"));
    const folder = runFolder(root);
    const appendFileSpy = vi.mocked(appendFile);
    appendFileSpy.mockClear();

    for (let index = 0; index < 1000; index += 1) {
      await appendTrace(folder, { event: "step", index });
    }

    expect(appendFileSpy).toHaveBeenCalledTimes(1000);
    const lines = (await readFile(folder.tracePath, "utf8")).trimEnd().split("\n");
    expect(lines).toHaveLength(1000);
    expect(JSON.parse(lines[0] as string)).toEqual({ event: "step", index: 0 });
    expect(JSON.parse(lines[999] as string)).toEqual({ event: "step", index: 999 });
  });
});
