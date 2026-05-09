#!/usr/bin/env node
import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import crypto from "node:crypto";

const manifest = JSON.parse(await readFile("benchmarks/office-corpus/manifest.json", "utf8"));
const root = manifest.storageRoot ?? ".officegen/benchmark-corpus";
await mkdir(root, { recursive: true });
const results = [];

for (const doc of manifest.documents ?? []) {
  const ext = path.extname(new URL(doc.url).pathname) || `.${doc.kind}`;
  const out = path.join(root, `${doc.id}${ext}`);
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
