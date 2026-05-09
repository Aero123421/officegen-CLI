import { mkdtemp, readFile, rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { execFileSync } from "node:child_process";

const npm = process.platform === "win32" ? "npm.cmd" : "npm";
const npxOfficegen = (prefix) =>
  process.platform === "win32"
    ? path.join(prefix, "node_modules", ".bin", "officegen.cmd")
    : path.join(prefix, "node_modules", ".bin", "officegen");

function run(command, args, options = {}) {
  const file = process.platform === "win32" ? process.env.ComSpec ?? "cmd.exe" : command;
  const finalArgs = process.platform === "win32"
    ? ["/d", "/s", "/c", [command, ...args].map(windowsQuote).join(" ")]
    : args;
  return execFileSync(file, finalArgs, {
    stdio: options.capture ? ["ignore", "pipe", "inherit"] : "inherit",
    encoding: "utf8",
    ...options
  });
}

function windowsQuote(value) {
  const text = String(value);
  if (!/[()\s"%^&|<>]/.test(text)) return text;
  return `"${text.replace(/"/g, '""')}"`;
}

const prefix = await mkdtemp(path.join(os.tmpdir(), "officegen-pack-smoke-"));
let tarball;

try {
  tarball = run(npm, ["pack", "--silent"], { capture: true }).trim().split(/\r?\n/).at(-1);
  if (!tarball) throw new Error("npm pack did not return a tarball path.");

  run(npm, ["install", "--prefix", prefix, path.resolve(tarball), "--ignore-scripts", "--no-audit", "--no-fund"]);

  const version = run(npxOfficegen(prefix), ["--version"], { capture: true }).trim();
  const manifest = JSON.parse(await readFile(path.join(prefix, "node_modules", "officegen", "package.json"), "utf8"));
  const capabilityEnvelope = JSON.parse(run(npxOfficegen(prefix), ["capabilities", "--json"], { capture: true }));

  if (version !== manifest.version) {
    throw new Error(`officegen --version (${version}) did not match package version (${manifest.version}).`);
  }
  if (manifest.name !== "officegen") {
    throw new Error(`packed package name must be officegen, got ${manifest.name}.`);
  }
  if (!capabilityEnvelope.ok || capabilityEnvelope.result?.schema !== "officegen.capabilities@1.2") {
    throw new Error("installed CLI did not return a valid capabilities envelope.");
  }

  console.log(JSON.stringify({
    ok: true,
    tarball,
    packageName: manifest.name,
    version,
    command: "officegen capabilities --json"
  }, null, 2));
} finally {
  if (process.platform === "win32" && process.env.CI) {
    console.warn(`Skipping recursive cleanup on Windows CI: ${prefix}`);
  } else {
    await rm(prefix, { recursive: true, force: true });
  }
  if (tarball) await rm(path.resolve(tarball), { force: true });
}
