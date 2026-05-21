#!/usr/bin/env node
import { spawnSync } from "node:child_process";
import { existsSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import os from "node:os";
import path from "node:path";

const bin = valueArg("--bin") ?? process.env.OFFICEGEN_NATIVE_BIN;
const expectedVersion = valueArg("--expected-version") ?? process.env.OFFICEGEN_EXPECTED_VERSION;

if (!bin) {
  console.error("usage: node scripts/native-release-smoke.mjs --bin <path> [--expected-version x.y.z]");
  process.exit(2);
}

const binPath = path.resolve(bin);
if (!existsSync(binPath)) {
  console.error(`native binary does not exist: ${binPath}`);
  process.exit(1);
}

for (const args of [["--version"], ["--help"], ["capabilities", "--agent", "--json"]]) {
  const result = run(binPath, args);
  if (args[0] === "--version" && expectedVersion && result.stdout.trim() !== expectedVersion) {
    console.error(`officegen --version returned ${result.stdout.trim()}, expected ${expectedVersion}.`);
    process.exit(1);
  }
  if (args[0] === "capabilities") {
    const envelope = JSON.parse(result.stdout);
    if (!envelope.ok || envelope.result?.schema !== "officegen.capabilities@1.2") {
      console.error("native binary capabilities smoke did not emit a valid capabilities envelope.");
      process.exit(1);
    }
  }
}

const temp = mkdtempSync(path.join(os.tmpdir(), "officegen-native-smoke-"));
try {
  const ir = path.join(temp, "smoke.ir.json");
  const out = path.join(temp, "smoke.docx");
  writeFileSync(ir, JSON.stringify({
    schema: "officegen.ir.document@1.2",
    title: "Native smoke",
    targets: ["docx"],
    sections: [{ blocks: [{ type: "paragraph", text: "Rust release smoke" }] }]
  }));
  run(binPath, ["render", "smoke.ir.json", "--target", "docx", "--out", "smoke.docx", "--agent", "--strict-json"], temp);
  const inspect = run(binPath, ["inspect", "smoke.docx", "--agent", "--strict-json"], temp);
  const envelope = JSON.parse(inspect.stdout);
  if (!envelope.ok || envelope.result?.format !== "docx") {
    console.error("native binary render/inspect smoke failed.");
    process.exit(1);
  }
  const unsupported = runAllowFailure(binPath, ["definitely-unknown", "--agent", "--strict-json"], temp);
  if (unsupported.status === 0) {
    console.error("native binary returned success for an unknown command.");
    process.exit(1);
  }
  const failure = JSON.parse(unsupported.stdout);
  if (failure.ok !== false || failure.error?.code !== "UNKNOWN_COMMAND") {
    console.error("native binary did not emit structured unknown-command failure.");
    process.exit(1);
  }
} finally {
  rmSync(temp, { recursive: true, force: true });
}

console.log(`officegen native binary smoke passed for ${binPath}`);

function valueArg(name) {
  const index = process.argv.indexOf(name);
  return index >= 0 ? process.argv[index + 1] : undefined;
}

function run(command, args, cwd = process.cwd()) {
  const result = runAllowFailure(command, args, cwd);
  if (result.status !== 0) {
    console.error(result.stdout);
    console.error(result.stderr);
    process.exit(result.status ?? 1);
  }
  return result;
}

function runAllowFailure(command, args, cwd = process.cwd()) {
  const result = process.platform === "win32"
    ? spawnSync("cmd.exe", ["/d", "/c", command, ...args], { encoding: "utf8", cwd })
    : spawnSync(command, args, { encoding: "utf8", shell: false, cwd });
  if (result.error) throw result.error;
  return result;
}
