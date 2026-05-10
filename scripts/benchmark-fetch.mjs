#!/usr/bin/env node
import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import crypto from "node:crypto";
import {
  DEFAULT_BENCHMARK_MANIFEST_PATH,
  resolveBenchmarkDocumentPath,
  resolveBenchmarkManifestPath,
  resolveBenchmarkStorageRoot
} from "./benchmark-policy.mjs";

const manifestPath = resolveBenchmarkManifestPath(process.cwd(), process.argv[2] ?? DEFAULT_BENCHMARK_MANIFEST_PATH);
const manifest = JSON.parse(await readFile(manifestPath, "utf8"));
const root = resolveBenchmarkStorageRoot(process.cwd(), manifest.storageRoot);
await mkdir(root, { recursive: true });
const results = [];

for (const [index, doc] of (manifest.documents ?? []).entries()) {
  const ext = path.extname(new URL(doc.url).pathname) || `.${doc.kind}`;
  const relative = String(doc.path ?? doc.fileName ?? `${doc.id}${ext}`);
  const out = resolveBenchmarkDocumentPath(root, relative, `documents[${index}].path`);
  await mkdir(path.dirname(out), { recursive: true });
  const urls = [doc.url, doc.fallbackUrl].filter(Boolean);
  let lastError = "";
  for (const url of urls) {
    try {
      const res = await fetch(url, { redirect: "follow" });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const bytes = Buffer.from(await res.arrayBuffer());
      await writeFile(out, bytes);
      results.push({
        id: doc.id,
        ok: true,
        url,
        path: out,
        byteLength: bytes.length,
        sha256: crypto.createHash("sha256").update(bytes).digest("hex")
      });
      lastError = "";
      break;
    } catch (error) {
      lastError = error instanceof Error ? error.message : String(error);
    }
  }
  if (lastError) results.push({ id: doc.id, ok: false, reason: lastError, sourcePage: doc.sourcePage, fallbackUrl: doc.fallbackUrl });
}

const output = {
  schema: "officegen.benchmark-fetch.result@2.2",
  generatedAt: new Date().toISOString(),
  root,
  results
};
await writeFile(path.join(root, "fetch-manifest.json"), `${JSON.stringify(output, null, 2)}\n`, "utf8");
console.log(JSON.stringify(output, null, 2));
