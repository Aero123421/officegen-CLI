#!/usr/bin/env node
import { existsSync } from "node:fs";
import { mkdir, readFile, readdir, writeFile } from "node:fs/promises";
import { spawnSync } from "node:child_process";
import path from "node:path";
import {
  DEFAULT_BENCHMARK_MANIFEST_PATH,
  assertBenchmarkRelativePath,
  resolveBenchmarkDocumentPath,
  resolveBenchmarkManifestPath,
  resolveBenchmarkStorageRoot
} from "./benchmark-policy.mjs";

const manifestPath = resolveBenchmarkManifestPath(process.cwd(), process.argv[2] ?? DEFAULT_BENCHMARK_MANIFEST_PATH);
const manifest = JSON.parse(await readFile(manifestPath, "utf8"));
const root = resolveBenchmarkStorageRoot(process.cwd(), manifest.storageRoot);
const outRoot = ".officegen/benchmark-results";
const timeoutMs = Number(process.env.OFFICEGEN_BENCHMARK_TIMEOUT_MS ?? 45000);
await mkdir(outRoot, { recursive: true });
const files = existsSync(root) ? await readdir(root) : [];
const rows = [];

for (const [index, doc] of (manifest.documents ?? []).entries()) {
  const relative = doc.path ?? doc.fileName;
  if (relative) {
    const input = resolveBenchmarkDocumentPath(root, String(relative), `documents[${index}].path`);
    if (!existsSync(input)) {
      rows.push({ id: doc.id, ok: false, skipped: true, reason: "not fetched" });
      continue;
    }
    rows.push(await reviewDocument(doc.id, input, outRoot));
    continue;
  }
  assertBenchmarkRelativePath(`${doc.id}.${doc.kind ?? "bin"}`, `documents[${index}].id`);
  const file = files.find((name) => name.startsWith(`${doc.id}.`));
  if (!file) {
    rows.push({ id: doc.id, ok: false, skipped: true, reason: "not fetched" });
    continue;
  }
  rows.push(await reviewDocument(doc.id, resolveBenchmarkDocumentPath(root, file, `documents[${index}].path`), outRoot));
}

const summary = {
  schema: "officegen.benchmark-review.result@2.2",
  generatedAt: new Date().toISOString(),
  corpusRoot: root,
  results: rows
};
await writeJson(path.join(outRoot, "summary.json"), summary);
await writeFile(path.join(outRoot, "summary.md"), [
  "# officegen benchmark review",
  "",
  "| id | inspect | verify | note |",
  "| --- | --- | --- | --- |",
  ...rows.map((row) => `| ${row.id} | ${row.inspectOk === true ? "ok" : row.skipped ? "skipped" : "fail"} | ${row.verifyOk === true ? "ok" : row.skipped ? "skipped" : "fail"} | ${row.reason ?? row.inspectReason ?? row.verifyReason ?? ""} |`)
].join("\n"), "utf8");
console.log(JSON.stringify(summary, null, 2));

async function reviewDocument(id, input, outRoot) {
  const inspect = runOfficegen(["inspect", input, "--depth", "summary", "--agent", "--json", "--json-budget-bytes", "160000"]);
  const verify = runOfficegen(["verify", input, "--visual", "--agent", "--json", "--json-budget-bytes", "160000"]);
  const reportId = String(id).replace(/[\\/]/g, "_");
  return {
    id,
    path: input,
    inspectOk: inspect.ok,
    verifyOk: verify.ok,
    inspectTimedOut: inspect.timedOut,
    verifyTimedOut: verify.timedOut,
    inspectReason: inspect.reason,
    verifyReason: verify.reason,
    inspectPath: await writeJson(path.join(outRoot, `${reportId}.inspect.json`), inspect),
    verifyPath: await writeJson(path.join(outRoot, `${reportId}.verify.json`), verify)
  };
}

function runOfficegen(args) {
  const bin = path.join("packages", "cli", "dist", "main.js");
  const result = spawnSync(process.execPath, [bin, ...args], {
    encoding: "utf8",
    timeout: timeoutMs,
    maxBuffer: 20 * 1024 * 1024
  });
  let parsed;
  try {
    parsed = JSON.parse(result.stdout);
  } catch {
    parsed = { ok: false, stdout: result.stdout, stderr: result.stderr };
  }
  const timedOut = result.error?.code === "ETIMEDOUT";
  return {
    ok: result.status === 0 && parsed.ok !== false,
    status: result.status,
    timedOut,
    reason: timedOut ? `timed out after ${timeoutMs}ms` : result.error?.message,
    envelope: parsed
  };
}

async function writeJson(file, value) {
  await writeFile(file, `${JSON.stringify(value, null, 2)}\n`, "utf8");
  return file;
}
