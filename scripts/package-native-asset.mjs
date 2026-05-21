#!/usr/bin/env node
import { createHash } from "node:crypto";
import { createReadStream } from "node:fs";
import { copyFile, mkdir, mkdtemp, rm, writeFile } from "node:fs/promises";
import { spawnSync } from "node:child_process";
import os from "node:os";
import path from "node:path";

const bin = requiredArg("--bin");
const target = requiredArg("--target");
const version = requiredArg("--version").replace(/^v/, "");
const outDir = path.resolve(valueArg("--out-dir") ?? "dist/native");
const exe = target.includes("windows") ? ".exe" : "";
const archiveExt = target.includes("windows") ? "zip" : "tar.gz";
const assetName = `officegen-v${version}-${target}.${archiveExt}`;
const archivePath = path.join(outDir, assetName);
const temp = await mkdtemp(path.join(os.tmpdir(), "officegen-native-package-"));

try {
  const root = path.join(temp, `officegen-v${version}-${target}`);
  await mkdir(root, { recursive: true });
  await copyFile(path.resolve(bin), path.join(root, `officegen${exe}`));
  await writeFile(path.join(root, "manifest.json"), `${JSON.stringify({
    name: "officegen",
    version,
    target,
    binary: `officegen${exe}`
  }, null, 2)}\n`, "utf8");
  await mkdir(outDir, { recursive: true });

  if (archiveExt === "zip") {
    run("powershell.exe", ["-NoProfile", "-Command", `Compress-Archive -Path '${escapePwsh(root)}\\*' -DestinationPath '${escapePwsh(archivePath)}' -Force`]);
  } else {
    run("tar", ["-czf", archivePath, "-C", temp, path.basename(root)]);
  }

  const digest = await sha256(archivePath);
  await writeFile(`${archivePath}.sha256`, `${digest}  ${assetName}\n`, "utf8");
  console.log(JSON.stringify({ ok: true, asset: archivePath, checksum: `${archivePath}.sha256` }, null, 2));
} finally {
  await rm(temp, { recursive: true, force: true });
}

function valueArg(name) {
  const index = process.argv.indexOf(name);
  return index >= 0 ? process.argv[index + 1] : undefined;
}

function requiredArg(name) {
  const value = valueArg(name);
  if (!value) {
    console.error(`missing required argument: ${name}`);
    process.exit(2);
  }
  return value;
}

function run(command, args) {
  const result = spawnSync(command, args, { stdio: "inherit", shell: false });
  if (result.status !== 0) process.exit(result.status ?? 1);
}

function escapePwsh(value) {
  return String(value).replace(/'/g, "''");
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
