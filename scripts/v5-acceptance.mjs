#!/usr/bin/env node
import { createHash } from "node:crypto";
import { spawnSync } from "node:child_process";
import { existsSync, readFileSync, statSync } from "node:fs";
import { mkdir, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import JSZip from "jszip";

const root = process.cwd();
const packageJson = JSON.parse(readFileSync(path.join(root, "package.json"), "utf8"));
const expectedVersion = resolveExpectedVersion(valueArg("--expected-version"));
const bin = path.resolve(valueArg("--bin") ?? process.env.OFFICEGEN_NATIVE_BIN ?? defaultNativeBin());
const keep = process.argv.includes("--keep");
const release = expectedVersion ?? packageJson.version;
const releaseTag = `v${release}`;
const evidenceDir = path.join(root, ".officegen", "acceptance", releaseTag);
const workDir = await mkdtemp(path.join(os.tmpdir(), "officegen-v5-acceptance-"));
const evidence = {
  schema: "officegen.v5.acceptance.evidence@1",
  release: releaseTag,
  generatedAt: new Date().toISOString(),
  binary: bin,
  expectedVersion: expectedVersion ?? null,
  workDir,
  checks: []
};

try {
  if (!existsSync(bin)) fail("A1", `native binary does not exist: ${bin}`);

  await mkdir(evidenceDir, { recursive: true });
  await seedFixtures(workDir);

  await check("A1", "native binary is Node-free and versioned", async () => {
    const version = run(bin, ["--version"]).stdout.trim();
    if (expectedVersion && version !== expectedVersion) {
      throw new Error(`--version returned ${version}, expected ${expectedVersion}`);
    }
    const caps = runJson(bin, ["capabilities", "--agent", "--strict-json"]);
    assertEnvelope(caps, "capabilities");
    if (caps.runtime?.nodeRequired !== false || caps.result?.nodeRequired !== false) {
      throw new Error("capabilities must report nodeRequired=false");
    }
    if (caps.result?.runtime !== "rust-native") throw new Error("capabilities must report rust-native runtime");
    return { version, capabilitiesHash: caps.capabilitiesHash };
  });

  await check("A2", "curl/irm installer story is preserved without Node runtime claims", async () => {
    const sh = readText("install/install.sh");
    const ps1 = readText("install/install.ps1");
    const readme = readText("README.md");
    for (const token of ["curl -fsSL", "install.sh", "Node.js is not required"]) {
      if (!readme.includes(token)) throw new Error(`README.md is missing ${token}`);
    }
    for (const token of ["irm ", "install.ps1", "GitHub Release"]) {
      if (!readme.includes(token)) throw new Error(`README.md is missing ${token}`);
    }
    for (const token of ["releases/download", ".sha256", "OFFICEGEN_VERSION"]) {
      if (!sh.includes(token)) throw new Error(`install/install.sh is missing ${token}`);
      if (!ps1.includes(token)) throw new Error(`install/install.ps1 is missing ${token}`);
    }
    return { installers: ["install/install.sh", "install/install.ps1"] };
  });

  await check("A3", "agent docs and agent JSON stay scoped to built-in CLI commands", async () => {
    const caps = runJson(bin, ["capabilities", "--agent", "--strict-json"]);
    const help = runJson(bin, ["help", "workflow", "inspect-edit-verify", "--agent", "--strict-json"]);
    const readme = readText("README.md");
    assertNoManagementRuntimeSales("capabilities", caps.result);
    assertNoManagementRuntimeSales("help workflow", help.result);
    const agentSection = section(readme, "## Agent Workflow", "## Command Map");
    assertNoManagementRuntimeSales("README Agent Workflow", agentSection);
    return { agentCommandCount: caps.result?.agentCommands?.length ?? 0 };
  });

  await check("A4", "strict JSON envelopes and failure classes are stable", async () => {
    const success = runJson(bin, ["doctor", "--agent", "--strict-json"]);
    assertEnvelope(success, "doctor");
    const unknown = runJsonAllowFailure(bin, ["definitely-unknown", "--agent", "--strict-json"]);
    if (unknown.status === 0) throw new Error("unknown command returned success");
    assertEnvelope(unknown.json, "definitely-unknown", false);
    if (unknown.json.error?.code !== "UNKNOWN_COMMAND") {
      throw new Error(`unknown command code was ${unknown.json.error?.code}`);
    }
    if (unknown.json.failureClass !== "usage") throw new Error("unknown command must be a usage failure");
    return { unknownCommandCode: unknown.json.error?.code };
  });

  await check("A5", "schema list/get/validate are machine executable", async () => {
    const list = runJson(bin, ["schema", "list", "--agent", "--strict-json"]);
    assertEnvelope(list, "schema list");
    const schemaId = list.result?.schemas?.find((entry) => entry.id === "officegen.ir.document@1.2")?.id;
    if (!schemaId) throw new Error("schema list did not include officegen.ir.document@1.2");
    const get = runJson(bin, ["schema", "get", schemaId, "--agent", "--strict-json"]);
    if (get.result?.id !== schemaId) throw new Error("schema get returned the wrong id");
    const valid = runJson(bin, ["schema", "validate", "document.ir.json", "--schema", schemaId, "--agent", "--strict-json"], workDir);
    if (valid.result?.ok !== true) throw new Error("valid IR did not pass schema validate");
    const invalid = runJsonAllowFailure(bin, ["schema", "validate", "invalid.ir.json", "--schema", schemaId, "--agent", "--strict-json"], workDir);
    if (invalid.status === 0 || invalid.json.result?.ok !== false || invalid.json.error?.code !== "SCHEMA_INVALID") {
      throw new Error("invalid IR did not fail closed with SCHEMA_INVALID");
    }
    return { schemaId };
  });

  await check("A6", "render smoke creates Office/PDF plus chart and diagram SVG artifacts", async () => {
    const artifacts = [];
    for (const target of ["pptx", "docx", "xlsx", "pdf"]) {
      const out = `rendered.${target}`;
      const rendered = runJson(bin, ["render", "document.ir.json", "--target", target, "--out", out, "--agent", "--strict-json"], workDir);
      assertArtifact(path.join(workDir, out));
      if (rendered.result?.target !== target) throw new Error(`render target mismatch for ${target}`);
      artifacts.push(out);
    }
    const chart = runJson(bin, ["chart", "render", "chart.example.json", "--out", "chart.svg", "--agent", "--strict-json"], workDir);
    const diagram = runJson(bin, ["diagram", "render", "process.example.mmd", "--out", "process.svg", "--agent", "--strict-json"], workDir);
    if (!chart.result?.svg?.includes("<svg") || !diagram.result?.svg?.includes("<svg")) {
      throw new Error("chart/diagram render did not return SVG");
    }
    assertArtifact(path.join(workDir, "chart.svg"));
    assertArtifact(path.join(workDir, "process.svg"));
    return { artifacts: [...artifacts, "chart.svg", "process.svg"] };
  });

  await check("A7", "inspect covers PPTX, DOCX, XLSX, and PDF", async () => {
    const formats = {};
    for (const target of ["pptx", "docx", "xlsx", "pdf"]) {
      const inspected = runJson(bin, ["inspect", `rendered.${target}`, "--depth", "summary", "--agent", "--strict-json"], workDir);
      assertEnvelope(inspected, "inspect");
      if (inspected.result?.format !== target) throw new Error(`inspect format mismatch for ${target}`);
      formats[target] = inspected.result?.trusted?.summary ?? {};
      await writeEvidenceJson(`inspect-${target}.json`, inspected);
    }
    return { formats };
  });

  await check("A8", "edit supports dry-run, explicit output, and no accidental in-place mutation", async () => {
    const beforeHash = await sha256(path.join(workDir, "rendered.docx"));
    const dryRun = runJson(bin, ["edit", "rendered.docx", "--ops", "edit-ops.docx.json", "--out", "dry-run.docx", "--dry-run", "--resolve-selectors", "--agent", "--strict-json"], workDir);
    if (dryRun.result?.dryRun !== true || dryRun.result?.applied !== 1) {
      throw new Error("dry-run edit did not resolve exactly one operation");
    }
    const afterDryRunHash = await sha256(path.join(workDir, "rendered.docx"));
    if (afterDryRunHash !== beforeHash) throw new Error("dry-run mutated the input file");
    const applied = runJson(bin, ["edit", "rendered.docx", "--ops", "edit-ops.docx.json", "--out", "edited.docx", "--agent", "--strict-json"], workDir);
    if (!applied.ok || applied.result?.applied !== 1) throw new Error("edit apply failed");
    assertArtifact(path.join(workDir, "edited.docx"));
    const originalHash = await sha256(path.join(workDir, "rendered.docx"));
    if (originalHash !== beforeHash) throw new Error("explicit-output edit mutated the input file");
    return { dryRunApplied: dryRun.result.applied, output: "edited.docx" };
  });

  await check("A9", "view, verify, and diff report portable readiness truthfully", async () => {
    const view = runJson(bin, ["view", "edited.docx", "--format", "svg", "--out", "view-docx", "--agent", "--strict-json"], workDir);
    assertArtifact(path.join(workDir, "view-docx", "page-001.svg"));
    if (view.result?.format !== "svg") throw new Error("view did not use SVG format");
    const png = runJsonAllowFailure(bin, ["view", "edited.docx", "--format", "png", "--out", "view-png", "--agent", "--strict-json"], workDir);
    if (png.status === 0 || png.json.error?.code !== "FEATURE_NOT_IMPLEMENTED") {
      throw new Error("portable PNG preview must fail closed");
    }
    const verify = runJson(bin, ["verify", "edited.docx", "--visual", "--agent", "--strict-json"], workDir);
    if (verify.result?.status !== "pass") throw new Error("verify did not pass edited DOCX");
    const warnings = JSON.stringify(verify.warnings ?? verify.result?.warnings ?? []);
    if (!warnings.includes("NATIVE_PROOF_NOT_RUN")) throw new Error("verify must disclose native proof was not run");
    const diff = runJson(bin, ["diff", "rendered.docx", "edited.docx", "--visual", "--agent", "--strict-json"], workDir);
    if (diff.result?.changed !== true || diff.result?.summary?.textChanged !== true) {
      throw new Error("diff did not report the edited DOCX change");
    }
    return { view: "view-docx/page-001.svg", verifyStatus: verify.result.status, diffChanged: diff.result.changed };
  });

  await check("A10", "path and ZIP safety fail closed", async () => {
    const traversal = runJsonAllowFailure(bin, ["render", "document.ir.json", "--target", "docx", "--out", "../outside.docx", "--agent", "--strict-json"], workDir);
    if (traversal.status === 0 || traversal.json.error?.code !== "SECURITY_PATH_OUTSIDE_ROOT") {
      throw new Error("output traversal did not fail with SECURITY_PATH_OUTSIDE_ROOT");
    }
    await writeUnsafeZip(path.join(workDir, "bad.docx"));
    const badInspect = runJsonAllowFailure(bin, ["inspect", "bad.docx", "--agent", "--strict-json"], workDir);
    if (badInspect.status === 0 || badInspect.json.error?.code !== "SECURITY_ZIP_UNSAFE") {
      throw new Error("unsafe ZIP inspect did not fail with SECURITY_ZIP_UNSAFE");
    }
    const badVerify = runJsonAllowFailure(bin, ["verify", "bad.docx", "--agent", "--strict-json"], workDir);
    if (badVerify.status === 0 || badVerify.json.result?.status !== "fail") {
      throw new Error("unsafe ZIP verify did not return blocked failure evidence");
    }
    return { traversalCode: traversal.json.error.code, zipCode: badInspect.json.error.code };
  });

  await check("A11", "native release asset naming and smoke gates are wired", async () => {
    const docs = readText("docs/planning/v5.0.0-acceptance-matrix.md");
    const pkg = JSON.parse(readText("package.json"));
    for (const script of ["v5:acceptance", "native:smoke", "installer:smoke", "release:gate"]) {
      if (!pkg.scripts?.[script]) throw new Error(`package.json is missing script ${script}`);
    }
    if (pkg.bin?.officegen || pkg.scripts?.officegen) {
      throw new Error("package.json must not expose an npm officegen CLI path");
    }
    for (const target of ["x86_64-unknown-linux-gnu", "aarch64-apple-darwin", "x86_64-pc-windows-msvc"]) {
      if (!docs.includes(`officegen-v${release}-${target}`) && !docs.includes(`officegen-v5.0.0-${target}`)) {
        throw new Error(`v5 matrix is missing ${target} asset`);
      }
    }
    if (!docs.includes(`npm run native:assets:check -- --version ${release} --include-installers`) && !docs.includes("npm run native:assets:check -- --version 5.0.0 --include-installers")) {
      throw new Error("v5 matrix is missing native asset check gate");
    }
    return { scripts: ["v5:acceptance", "release:gate"] };
  });

  await check("A12", `publish gate checklist is explicit before ${releaseTag} tag/release`, async () => {
    const gatePath = existsSync(path.join(root, `docs/reviews/${releaseTag}-release-gates.md`))
      ? `docs/reviews/${releaseTag}-release-gates.md`
      : "docs/reviews/v5.0.0-release-gates.md";
    const gateDoc = readText(gatePath);
    const required = [
      `npm run version:bump -- ${release}`,
      "npm run version:check",
      "npm run installer:smoke",
      `npm run v5:acceptance -- --bin target/release/officegen --expected-version ${release}`,
      `npm run v5:acceptance -- --bin target/release/officegen.exe --expected-version ${release}`,
      "cargo fmt --check",
      "cargo test --locked",
      "cargo build --release --locked",
      "npm run typecheck",
      "npm test",
      "npm run build"
    ];
    for (const command of required) {
      if (!gateDoc.includes(command)) throw new Error(`release gate doc is missing: ${command}`);
    }
    if (!gateDoc.includes(`Do not create the ${releaseTag} tag`) && !gateDoc.includes("Do not create the v5.0.0 tag")) {
      throw new Error("release gate doc must explicitly block tagging before gates pass");
    }
    return { requiredCommands: required.length };
  });

  evidence.summary = {
    ok: true,
    passed: evidence.checks.filter((entry) => entry.status === "pass").length,
    failed: 0
  };
  await writeEvidenceJson("manifest.json", evidence);
  console.log(`officegen v5 acceptance passed (${evidence.summary.passed} checks)`);
  console.log(`evidence: ${path.join(evidenceDir, "manifest.json")}`);
} catch (error) {
  evidence.summary = {
    ok: false,
    passed: evidence.checks.filter((entry) => entry.status === "pass").length,
    failed: evidence.checks.filter((entry) => entry.status === "fail").length || 1
  };
  await mkdir(evidenceDir, { recursive: true });
  await writeEvidenceJson("manifest.json", evidence);
  console.error(error instanceof Error ? error.message : String(error));
  console.error(`evidence: ${path.join(evidenceDir, "manifest.json")}`);
  process.exit(1);
} finally {
  if (!keep) await rm(workDir, { recursive: true, force: true });
}

async function check(id, title, fn) {
  const startedAt = new Date().toISOString();
  try {
    const details = await fn();
    evidence.checks.push({ id, title, status: "pass", startedAt, finishedAt: new Date().toISOString(), details });
  } catch (error) {
    evidence.checks.push({
      id,
      title,
      status: "fail",
      startedAt,
      finishedAt: new Date().toISOString(),
      error: error instanceof Error ? error.message : String(error)
    });
    throw new Error(`${id} failed: ${error instanceof Error ? error.message : String(error)}`);
  }
}

function run(command, args, cwd = root) {
  const result = process.platform === "win32"
    ? spawnSync("cmd.exe", ["/d", "/c", command, ...args], { cwd, encoding: "utf8" })
    : spawnSync(command, args, { cwd, encoding: "utf8", shell: false });
  if (result.error) throw result.error;
  if (result.status !== 0) {
    throw new Error(`command failed (${result.status}): ${command} ${args.join(" ")}\n${result.stdout}\n${result.stderr}`);
  }
  return result;
}

function runJson(command, args, cwd = root) {
  const result = run(command, args, cwd);
  return JSON.parse(result.stdout);
}

function runJsonAllowFailure(command, args, cwd = root) {
  const result = process.platform === "win32"
    ? spawnSync("cmd.exe", ["/d", "/c", command, ...args], { cwd, encoding: "utf8" })
    : spawnSync(command, args, { cwd, encoding: "utf8", shell: false });
  if (result.error) throw result.error;
  let json;
  try {
    json = JSON.parse(result.stdout);
  } catch (error) {
    throw new Error(`expected JSON stdout from failing command, got:\n${result.stdout}\n${result.stderr}`);
  }
  return { status: result.status, json };
}

function assertEnvelope(envelope, command, expectedOk = true) {
  if (envelope.schema !== "officegen.envelope@1.2") throw new Error(`${command} returned wrong envelope schema`);
  if (envelope.command !== command) throw new Error(`${command} returned command ${envelope.command}`);
  if (envelope.ok !== expectedOk) throw new Error(`${command} ok was ${envelope.ok}, expected ${expectedOk}`);
  if (envelope.pathsRedacted !== true) throw new Error(`${command} did not report pathsRedacted=true`);
  if (!String(envelope.capabilitiesHash ?? "").startsWith("sha256:")) throw new Error(`${command} did not include capabilitiesHash`);
}

function assertArtifact(filePath) {
  if (!existsSync(filePath)) throw new Error(`expected artifact missing: ${filePath}`);
  if (statSync(filePath).size <= 0) throw new Error(`expected artifact is empty: ${filePath}`);
}

function assertNoManagementRuntimeSales(label, value) {
  const text = typeof value === "string" ? value : JSON.stringify(value);
  for (const pattern of [/\bmcp\b/i, /\bplugin\b/i]) {
    if (pattern.test(text)) throw new Error(`${label} contains ${pattern}`);
  }
}

function section(text, start, end) {
  const startIndex = text.indexOf(start);
  if (startIndex < 0) throw new Error(`missing section ${start}`);
  const endIndex = text.indexOf(end, startIndex + start.length);
  return endIndex < 0 ? text.slice(startIndex) : text.slice(startIndex, endIndex);
}

async function seedFixtures(dir) {
  await writeFile(path.join(dir, "document.ir.json"), `${JSON.stringify({
    schema: "officegen.ir.document@1.2",
    title: "v5 acceptance",
    targets: ["pptx", "docx", "xlsx", "pdf"],
    sections: [{ title: "Acceptance", blocks: [{ type: "paragraph", text: "Replace this acceptance text" }] }]
  }, null, 2)}\n`);
  await writeFile(path.join(dir, "invalid.ir.json"), `${JSON.stringify({ schema: "officegen.ir.document@1.2" }, null, 2)}\n`);
  await writeFile(path.join(dir, "edit-ops.docx.json"), `${JSON.stringify({
    schema: "officegen.edit.ops@1.2",
    operations: [{ op: "docx.setText", selector: { contains: "Replace this acceptance text" }, text: "v5 acceptance edit applied" }]
  }, null, 2)}\n`);
  await writeFile(path.join(dir, "chart.example.json"), `${JSON.stringify({
    title: "v5 acceptance chart",
    data: { labels: ["A1", "A2", "A3"], values: [3, 5, 8] },
    encoding: { x: "labels", y: "values" }
  }, null, 2)}\n`);
  await writeFile(path.join(dir, "process.example.mmd"), "graph TD; A[Inspect] --> B[Edit]; B --> C[Verify]\n");
}

async function writeUnsafeZip(filePath) {
  const zip = new JSZip();
  zip.file("../evil.xml", "<evil/>");
  const buffer = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
  await writeFile(filePath, buffer);
}

async function writeEvidenceJson(name, value) {
  await mkdir(evidenceDir, { recursive: true });
  await writeFile(path.join(evidenceDir, name), `${JSON.stringify(value, null, 2)}\n`, "utf8");
}

async function sha256(filePath) {
  const hash = createHash("sha256");
  hash.update(await readFile(filePath));
  return `sha256:${hash.digest("hex")}`;
}

function fail(id, message) {
  evidence.checks.push({ id, title: "preflight", status: "fail", error: message });
  throw new Error(`${id} failed: ${message}`);
}

function valueArg(name) {
  const index = process.argv.indexOf(name);
  return index >= 0 ? process.argv[index + 1] : undefined;
}

function resolveExpectedVersion(value) {
  if (!value) return undefined;
  if (value === "current") return packageJson.version;
  return value.replace(/^v/, "");
}

function defaultNativeBin() {
  return path.join("target", "release", process.platform === "win32" ? "officegen.exe" : "officegen");
}

function readText(relativePath) {
  return readFileSync(path.join(root, relativePath), "utf8");
}
