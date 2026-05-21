#!/usr/bin/env node
import { spawnSync } from "node:child_process";
import { existsSync } from "node:fs";
import path from "node:path";

const args = process.argv.slice(2);
if (args.length === 0) {
  console.error("Usage: node scripts/run-cargo.mjs <cargo-args...>");
  process.exit(2);
}

const cargo = resolveCargo();
const result = spawnSync(cargo, args, { stdio: "inherit", shell: false });
if (result.error) {
  console.error(result.error.message);
  process.exit(1);
}
process.exit(result.status ?? 1);

function resolveCargo() {
  const exe = process.platform === "win32" ? "cargo.exe" : "cargo";
  const pathEntries = (process.env.PATH ?? "").split(path.delimiter).filter(Boolean);
  for (const entry of pathEntries) {
    const candidate = path.join(entry, exe);
    if (existsSync(candidate)) return candidate;
  }
  const home = process.env.CARGO_HOME
    ?? (process.platform === "win32"
      ? path.join(process.env.USERPROFILE ?? "", ".cargo")
      : path.join(process.env.HOME ?? "", ".cargo"));
  const candidate = path.join(home, "bin", exe);
  if (existsSync(candidate)) return candidate;
  return exe;
}
