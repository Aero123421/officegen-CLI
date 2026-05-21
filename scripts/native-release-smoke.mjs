#!/usr/bin/env node
import { spawnSync } from "node:child_process";
import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import os from "node:os";
import path from "node:path";

const bin = valueArg("--bin") ?? process.env.OFFICEGEN_NATIVE_BIN ?? defaultNativeBin();
const expectedVersion = resolveExpectedVersion(valueArg("--expected-version") ?? process.env.OFFICEGEN_EXPECTED_VERSION);

if (!bin) {
  console.error("usage: node scripts/native-release-smoke.mjs [--bin <path>] [--expected-version x.y.z|current]");
  process.exit(2);
}

const binPath = path.resolve(bin);
if (!existsSync(binPath)) {
  console.error(`native binary does not exist: ${binPath}`);
  process.exit(1);
}

for (const args of [["--version"], ["--help"], ["capabilities", "--agent", "--json"], ["help", "workflow", "inspect-edit-verify", "--agent", "--strict-json"]]) {
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
    if (envelope.runtime?.nodeRequired !== false || envelope.result?.nodeRequired !== false) {
      console.error("native binary capabilities did not report nodeRequired=false.");
      process.exit(1);
    }
    const capabilityText = JSON.stringify(envelope.result);
    if (capabilityText.includes("mcp serve")) {
      console.error("native binary capabilities exposed mcp serve.");
      process.exit(1);
    }
    if (/\bmcp\b/i.test(capabilityText) || /\bplugin\b/i.test(capabilityText)) {
      console.error("native binary agent capabilities exposed removed management runtime wording.");
      process.exit(1);
    }
  }
  if (args[0] === "help") {
    const envelope = JSON.parse(result.stdout);
    const helpText = JSON.stringify(envelope.result);
    if (!envelope.ok || envelope.result?.schema !== "officegen.help@1.2") {
      console.error("native binary workflow help smoke did not emit a valid help envelope.");
      process.exit(1);
    }
    if (/\bmcp\b/i.test(helpText) || /\bplugin\b/i.test(helpText)) {
      console.error("native binary agent workflow help exposed removed management runtime wording.");
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
  const workflow = path.join(temp, "workflow.json");
  writeFileSync(workflow, JSON.stringify({
    schema: "officegen.workflow@2.0",
    version: "2.0",
    outputRoot: "workflow-out",
    steps: [
      {
        id: "scaffold",
        command: "scaffold",
        args: ["--kind", "docx", "--title", "Workflow smoke", "--out", "workflow.ir.json"]
      },
      {
        id: "render",
        command: "render",
        args: ["workflow-out/workflow.ir.json", "--target", "docx", "--out", "workflow.docx"]
      }
    ]
  }));
  const workflowRun = run(binPath, ["run", "workflow.json", "--agent", "--strict-json"], temp);
  const workflowEnvelope = JSON.parse(workflowRun.stdout);
  if (!workflowEnvelope.ok || workflowEnvelope.result?.schema !== "officegen.workflow.run.result@2.0") {
    console.error("native binary workflow run smoke failed.");
    process.exit(1);
  }
  for (const artifact of ["manifest.json", "trace.json", "summary.json", "workflow.docx"]) {
    if (!existsSync(path.join(temp, "workflow-out", artifact))) {
      console.error(`native binary workflow smoke did not write ${artifact}.`);
      process.exit(1);
    }
  }
  const rasterOut = path.join(temp, "raster-view");
  const raster = runAllowFailure(binPath, ["view", "smoke.docx", "--format", "png", "--out", "raster-view", "--agent", "--strict-json"], temp);
  if (raster.status === 0) {
    console.error("native binary returned success for portable PNG raster preview.");
    process.exit(1);
  }
  const rasterEnvelope = JSON.parse(raster.stdout);
  if (rasterEnvelope.ok !== false || rasterEnvelope.error?.code !== "FEATURE_NOT_IMPLEMENTED") {
    console.error("native binary did not fail closed for unsupported PNG raster preview.");
    process.exit(1);
  }
  if (existsSync(path.join(rasterOut, "page-001.png"))) {
    console.error("native binary wrote a placeholder PNG despite failing closed.");
    process.exit(1);
  }
  const mcp = runAllowFailure(binPath, ["mcp", "serve", "--agent", "--strict-json"], temp);
  if (mcp.status === 0) {
    console.error("native binary returned success for mcp serve.");
    process.exit(1);
  }
  const mcpEnvelope = JSON.parse(mcp.stdout);
  if (mcpEnvelope.ok !== false || mcpEnvelope.error?.code !== "FEATURE_REMOVED_FROM_SCOPE") {
    console.error("native binary did not report FEATURE_REMOVED_FROM_SCOPE for mcp serve.");
    process.exit(1);
  }
  const plugin = runAllowFailure(binPath, ["plugin", "install", "--agent", "--strict-json"], temp);
  if (plugin.status === 0 || JSON.parse(plugin.stdout).error?.code !== "FEATURE_REMOVED_FROM_SCOPE") {
    console.error("native binary did not fail closed for plugin install.");
    process.exit(1);
  }
  const agent = runAllowFailure(binPath, ["agent", "install", "--agent", "--strict-json"], temp);
  if (agent.status === 0 || JSON.parse(agent.stdout).error?.code !== "FEATURE_NOT_IMPLEMENTED") {
    console.error("native binary did not fail closed for agent install.");
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

function defaultNativeBin() {
  return path.join("target", "release", process.platform === "win32" ? "officegen.exe" : "officegen");
}

function resolveExpectedVersion(value) {
  if (!value) return undefined;
  if (value === "current") {
    return JSON.parse(readFileSync("package.json", "utf8")).version;
  }
  return value.replace(/^v/, "");
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
