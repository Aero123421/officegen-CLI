#!/usr/bin/env node
import assert from "node:assert/strict";
import os from "node:os";
import path from "node:path";
import { mkdtemp, rm } from "node:fs/promises";
import {
  BENCHMARK_MANIFEST_PATH_DENIED,
  resolveBenchmarkDocumentPath,
  resolveBenchmarkManifestPath,
  resolveBenchmarkStorageRoot
} from "./benchmark-policy.mjs";

const cwd = await mkdtemp(path.join(os.tmpdir(), "officegen-benchmark-policy-"));
try {
  const storageRoot = resolveBenchmarkStorageRoot(cwd, ".officegen/benchmark-corpus");
  assert.equal(storageRoot, path.join(cwd, ".officegen", "benchmark-corpus"));
  assert.equal(resolveBenchmarkManifestPath(cwd, "benchmarks/office-corpus/manifest.json"), path.join(cwd, "benchmarks", "office-corpus", "manifest.json"));
  assert.equal(resolveBenchmarkDocumentPath(storageRoot, "samples/deck.pptx"), path.join(storageRoot, "samples", "deck.pptx"));

  for (const [field, fn] of [
    ["manifest absolute", () => resolveBenchmarkManifestPath(cwd, path.join(cwd, "manifest.json"))],
    ["storageRoot traversal", () => resolveBenchmarkStorageRoot(cwd, "../outside")],
    ["document absolute", () => resolveBenchmarkDocumentPath(storageRoot, path.join(cwd, "secret.pdf"))],
    ["document traversal", () => resolveBenchmarkDocumentPath(storageRoot, "../../secret.pdf")]
  ]) {
    assertBenchmarkDenied(field, fn);
  }
} finally {
  await rm(cwd, { recursive: true, force: true });
}

console.log("benchmark:local-fixtures ok");

function assertBenchmarkDenied(label, fn) {
  assert.throws(fn, (error) => {
    assert.equal(error.code, BENCHMARK_MANIFEST_PATH_DENIED, label);
    return true;
  });
}
