#!/usr/bin/env node
import { access, mkdir, rm, stat } from "node:fs/promises";
import { constants } from "node:fs";
import { spawn } from "node:child_process";
import path from "node:path";

const root = process.cwd();
const entry = path.join(root, "packages", "cli", "dist", "main.js");
const outDir = path.join(root, ".officegen", "binary-smoke");

await access(entry, constants.R_OK);
await mkdir(outDir, { recursive: true });

const result = {
  schema: "officegen.binary-smoke@1.2",
  entry,
  nodeRuntime: process.version,
  strategies: []
};

const bun = await commandVersion("bun", ["--version"]);
if (bun.ok) {
  const output = path.join(outDir, process.platform === "win32" ? "officegen.exe" : "officegen");
  const compiled = await run("bun", ["build", entry, "--compile", "--outfile", output], { timeoutMs: 120000 });
  const exists = compiled.ok ? await stat(output).then((s) => s.isFile()).catch(() => false) : false;
  const version = exists ? await run(output, ["--version"], { timeoutMs: 30000 }) : { ok: false, stderr: "compiled binary missing" };
  const capabilities = exists ? await run(output, ["capabilities", "--json"], { timeoutMs: 30000 }) : { ok: false, stderr: "compiled binary missing" };
  result.strategies.push({
    id: "bun-build-compile",
    available: true,
    ok: compiled.ok && exists && version.ok && capabilities.ok,
    output: exists ? output : undefined,
    detail: compiled.ok
      ? `Compiled and executed with bun build --compile. versionOk=${version.ok}; capabilitiesOk=${capabilities.ok}.`
      : compiled.stderr || compiled.error
  });
} else {
  result.strategies.push({
    id: "bun-build-compile",
    available: false,
    ok: false,
    detail: "bun is not installed or not on PATH."
  });
}

result.strategies.push({
  id: "node-entry",
  available: true,
  ok: (await run(process.execPath, [entry, "--version"])).ok,
  detail: "The distributable npm entry runs under the bundled Node runtime."
});

result.singleExecutableVerified = result.strategies.some((strategy) => strategy.id === "bun-build-compile" && strategy.ok);

console.log(JSON.stringify(result, null, 2));
if (!result.strategies.some((strategy) => strategy.ok)) process.exitCode = 1;

await rm(outDir, { recursive: true, force: true });

async function commandVersion(command, args) {
  return run(command, args, { timeoutMs: 10000 });
}

async function run(command, args, options = {}) {
  return new Promise((resolve) => {
    const child = spawn(command, args, { windowsHide: true });
    let stdout = "";
    let stderr = "";
    const timeout = setTimeout(() => {
      child.kill();
      resolve({ ok: false, stdout, stderr, error: "timeout" });
    }, options.timeoutMs ?? 30000);
    child.stdout.on("data", (chunk) => {
      stdout += String(chunk);
    });
    child.stderr.on("data", (chunk) => {
      stderr += String(chunk);
    });
    child.on("error", (error) => {
      clearTimeout(timeout);
      resolve({ ok: false, stdout, stderr, error: error.message });
    });
    child.on("exit", (code) => {
      clearTimeout(timeout);
      resolve({ ok: code === 0, stdout, stderr, code });
    });
  });
}
