#!/usr/bin/env node
import { createHash } from "node:crypto";
import { createReadStream, existsSync, readFileSync } from "node:fs";
import path from "node:path";

const version = (valueArg("--version") ?? readPackageVersion()).replace(/^v/, "");
const distDir = path.resolve(valueArg("--dist-dir") ?? "dist");
const includeInstallers = process.argv.includes("--include-installers");
const targets = (valueArg("--targets") ?? process.env.OFFICEGEN_NATIVE_TARGETS ?? [
  "x86_64-unknown-linux-gnu",
  "aarch64-unknown-linux-gnu",
  "x86_64-apple-darwin",
  "aarch64-apple-darwin",
  "x86_64-pc-windows-msvc",
  "aarch64-pc-windows-msvc"
].join(","))
  .split(",")
  .map((target) => target.trim())
  .filter(Boolean);

const expectedAssets = targets.map((target) => {
  const archiveExt = target.includes("windows") ? "zip" : "tar.gz";
  return `officegen-v${version}-${target}.${archiveExt}`;
});

if (includeInstallers) {
  expectedAssets.push("install.sh", "install.ps1");
}

const failures = [];
for (const asset of expectedAssets) {
  const assetPath = path.join(distDir, asset);
  const checksumPath = `${assetPath}.sha256`;
  if (!existsSync(assetPath)) {
    failures.push(`missing release asset: ${relative(assetPath)}`);
    continue;
  }
  if (!existsSync(checksumPath)) {
    failures.push(`missing checksum: ${relative(checksumPath)}`);
    continue;
  }
  const expectedDigest = readFileSync(checksumPath, "utf8").trim().split(/\s+/)[0]?.toLowerCase();
  const actualDigest = await sha256(assetPath);
  if (!/^[0-9a-f]{64}$/.test(expectedDigest ?? "")) {
    failures.push(`invalid checksum file: ${relative(checksumPath)}`);
  } else if (actualDigest !== expectedDigest) {
    failures.push(`checksum mismatch: ${relative(assetPath)}`);
  }
}

if (failures.length) {
  for (const failure of failures) console.error(failure);
  process.exit(1);
}

console.log(`officegen native release asset check passed for v${version}`);

function valueArg(name) {
  const index = process.argv.indexOf(name);
  return index >= 0 ? process.argv[index + 1] : undefined;
}

function readPackageVersion() {
  const manifest = JSON.parse(readFileSync("package.json", "utf8"));
  return manifest.version;
}

function relative(filePath) {
  return path.relative(process.cwd(), filePath).replace(/\\/g, "/");
}

function sha256(filePath) {
  return new Promise((resolve, reject) => {
    const hash = createHash("sha256");
    const stream = createReadStream(filePath);
    stream.on("data", (chunk) => hash.update(chunk));
    stream.on("error", reject);
    stream.on("end", () => resolve(hash.digest("hex")));
  });
}
