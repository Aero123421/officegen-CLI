#!/usr/bin/env node
import { createHash } from "node:crypto";
import { readFileSync, statSync } from "node:fs";
import path from "node:path";
import process from "node:process";

const root = process.cwd();
const args = process.argv.slice(2);
const options = {
  gate: "visibility",
  json: false,
  suite: path.join(root, "goal", "acceptance-suite.perfect-spec-v2.json"),
  matrix: path.join(root, "docs", "reviews", "v3.1.0-remediation-matrix.md"),
  v31EvidenceMatrix: path.join(root, "goal", "v3.1.0-evidence-matrix.json"),
  evidenceManifest: path.join(root, ".officegen", "acceptance", "perfect-spec", "manifest.json"),
  fixtureEvidence: path.join(root, ".officegen", "acceptance", "perfect-spec", "fixtures.json"),
  cliParity: path.join(root, ".officegen", "acceptance", "perfect-spec", "cli-parity.json"),
  nativeRenderer: path.join(root, ".officegen", "acceptance", "perfect-spec", "native-renderer.json"),
  postTagSmoke: path.join(root, ".officegen", "acceptance", "perfect-spec", "post-tag-smoke.json")
};

for (const arg of args) {
  if (arg === "--json") options.json = true;
  else if (arg.startsWith("--gate=")) options.gate = arg.slice("--gate=".length);
  else if (arg.startsWith("--suite=")) options.suite = path.resolve(root, arg.slice("--suite=".length));
  else if (arg.startsWith("--matrix=")) options.matrix = path.resolve(root, arg.slice("--matrix=".length));
  else if (arg.startsWith("--v31-evidence-matrix=")) options.v31EvidenceMatrix = path.resolve(root, arg.slice("--v31-evidence-matrix=".length));
  else if (arg.startsWith("--evidence-manifest=")) options.evidenceManifest = path.resolve(root, arg.slice("--evidence-manifest=".length));
  else if (arg.startsWith("--fixture-evidence=")) options.fixtureEvidence = path.resolve(root, arg.slice("--fixture-evidence=".length));
  else if (arg.startsWith("--cli-parity=")) options.cliParity = path.resolve(root, arg.slice("--cli-parity=".length));
  else if (arg.startsWith("--native-renderer=")) options.nativeRenderer = path.resolve(root, arg.slice("--native-renderer=".length));
  else if (arg.startsWith("--post-tag-smoke=")) options.postTagSmoke = path.resolve(root, arg.slice("--post-tag-smoke=".length));
  else usage(`unknown argument: ${arg}`);
}

if (!["visibility", "publish"].includes(options.gate)) usage(`invalid --gate: ${options.gate}`);

const failures = [];
const warnings = [];

const suite = readJson(options.suite);
const matrix = readMatrix(options.matrix);
const report = {
  gate: options.gate,
  suite: relative(options.suite),
  matrix: relative(options.matrix),
  release: suite.release,
  level: suite.level,
  totals: {
    tests: 0,
    ready: 0,
    pending: 0,
    blocked: 0,
    remediationRows: matrix.rows.size,
    remediationCovered: 0,
    capabilityTruthfulnessChecks: 0
  },
  priorities: {},
  nonReady: [],
  blocked: [],
  remediationMissing: [],
  capabilityTruthfulness: []
};

checkSuiteShape();
checkPriorityCoverage();
checkRemediationCoverage();
checkCapabilityTruthfulness();
checkFixtureEvidenceGate();
checkCliParityGate();
checkNativeRendererGate();
checkV31RuntimeEvidenceGate();
checkPublishGate();

if (options.json) {
  console.log(JSON.stringify({ ok: failures.length === 0, failures, warnings, report }, null, 2));
} else {
  for (const warning of warnings) console.warn(`perfect-spec:check warning: ${warning}`);
  if (failures.length) {
    console.error("perfect-spec:check failed:");
    for (const failure of failures) console.error(`- ${failure}`);
  } else {
    console.log(
      `perfect-spec:check ok (${report.totals.tests} tests; ready ${report.totals.ready}, pending ${report.totals.pending}, blocked ${report.totals.blocked}; remediation ${report.totals.remediationCovered}/${report.totals.remediationRows}; gate ${options.gate})`
    );
    printVisibleNonReady();
  }
}

process.exit(failures.length ? 1 : 0);

function checkSuiteShape() {
  if (suite.schema !== "officegen.perfect-spec.acceptance-suite@2.0") failures.push("suite schema marker is invalid");
  if (suite.level !== "L7") failures.push("suite level must be L7");
  if (!suite.suiteId || typeof suite.suiteId !== "string") failures.push("suiteId is required");
  if (!/^\d+\.\d+\.\d+$/.test(suite.release ?? "")) failures.push("release must be semver-like x.y.z");
  if (!Array.isArray(suite.tests) || suite.tests.length < 8) failures.push("suite needs at least 8 tests");
  if (!Array.isArray(suite.capabilityTruthfulnessChecks) || suite.capabilityTruthfulnessChecks.length < 3) {
    failures.push("suite needs at least 3 capabilityTruthfulnessChecks");
  }

  const ids = new Set();
  const allowedPriorities = new Set(["P0", "P1", "P2"]);
  const allowedStatuses = new Set(["ready", "pending", "blocked"]);
  const allowedKinds = new Set(["contract", "cli", "e2e", "smoke", "security", "quality", "packaging", "benchmark", "manual"]);

  for (const test of suite.tests ?? []) {
    report.totals.tests += 1;
    report.totals[test.status] = (report.totals[test.status] ?? 0) + 1;

    if (!/^L7-A[0-9]{3}$/.test(test.id ?? "")) failures.push(`${test.id ?? "<missing>"} has invalid L7 test id`);
    if (ids.has(test.id)) failures.push(`duplicate test id: ${test.id}`);
    ids.add(test.id);
    if (!test.title || typeof test.title !== "string") failures.push(`${test.id} is missing title`);
    if (!allowedPriorities.has(test.priority)) failures.push(`${test.id} has invalid priority: ${test.priority}`);
    if (!allowedStatuses.has(test.status)) failures.push(`${test.id} has invalid status: ${test.status}`);
    if (!allowedKinds.has(test.kind)) failures.push(`${test.id} has invalid kind: ${test.kind}`);
    if (typeof test.blocking !== "boolean") failures.push(`${test.id} must declare blocking boolean`);
    if (!Array.isArray(test.remediationIds) || test.remediationIds.length === 0) failures.push(`${test.id} is missing remediationIds`);
    if (!Array.isArray(test.capabilityTruthfulnessIds)) failures.push(`${test.id} is missing capabilityTruthfulnessIds`);
    if (!Array.isArray(test.steps) || test.steps.length === 0) failures.push(`${test.id} is missing steps`);
    if (!Array.isArray(test.passCriteria) || test.passCriteria.length === 0) failures.push(`${test.id} is missing passCriteria`);
    if (!Array.isArray(test.evidence) || test.evidence.length === 0) failures.push(`${test.id} is missing evidence`);

    if (test.status === "pending" && (!test.statusReason || !test.nextAction)) {
      failures.push(`${test.id} is pending but missing statusReason or nextAction`);
    }
    if (test.status === "blocked" && (!test.statusReason || !test.blockedBy || !test.nextAction)) {
      failures.push(`${test.id} is blocked but missing statusReason, blockedBy, or nextAction`);
    }

    report.priorities[test.priority] = (report.priorities[test.priority] ?? 0) + 1;
    if (test.status === "pending") report.nonReady.push({ id: test.id, priority: test.priority, status: test.status, nextAction: test.nextAction });
    if (test.status === "blocked") {
      const item = { id: test.id, priority: test.priority, status: test.status, blockedBy: test.blockedBy, nextAction: test.nextAction };
      report.nonReady.push(item);
      report.blocked.push(item);
    }
  }
}

function checkPriorityCoverage() {
  const minimums = suite.minimumPriorityCounts ?? {};
  for (const priority of ["P0", "P1", "P2"]) {
    const actual = report.priorities[priority] ?? 0;
    const minimum = minimums[priority] ?? 1;
    if (actual < minimum) failures.push(`expected at least ${minimum} ${priority} tests, found ${actual}`);
  }
}

function checkRemediationCoverage() {
  const known = matrix.rows;
  const covered = new Set();
  const exclusions = new Map((suite.coverageExclusions ?? []).map((entry) => [entry.remediationId, entry]));

  for (const test of suite.tests ?? []) {
    for (const id of test.remediationIds ?? []) {
      if (!known.has(id)) failures.push(`${test.id} references unknown remediation id: ${id}`);
      covered.add(id);
    }
  }

  for (const check of suite.capabilityTruthfulnessChecks ?? []) {
    for (const id of check.remediationIds ?? []) {
      if (!known.has(id)) failures.push(`${check.id} references unknown remediation id: ${id}`);
    }
  }

  for (const id of known.keys()) {
    if (covered.has(id)) continue;
    const exclusion = exclusions.get(id);
    if (exclusion?.reason) continue;
    report.remediationMissing.push(id);
    failures.push(`remediation row is not covered by L7 acceptance suite: ${id}`);
  }

  report.totals.remediationCovered = [...known.keys()].filter((id) => covered.has(id) || exclusions.has(id)).length;
}

function checkCapabilityTruthfulness() {
  const ids = new Set();
  const referenced = new Set((suite.tests ?? []).flatMap((test) => test.capabilityTruthfulnessIds ?? []));

  for (const check of suite.capabilityTruthfulnessChecks ?? []) {
    report.totals.capabilityTruthfulnessChecks += 1;
    if (!/^CT-[A-Z0-9-]+$/.test(check.id ?? "")) failures.push(`${check.id ?? "<missing>"} has invalid capability truthfulness id`);
    if (ids.has(check.id)) failures.push(`duplicate capability truthfulness id: ${check.id}`);
    ids.add(check.id);
    if (!referenced.has(check.id)) warnings.push(`${check.id} is not referenced by any acceptance test`);
    if (!Array.isArray(check.requiredTerms) || check.requiredTerms.length === 0) failures.push(`${check.id} is missing requiredTerms`);
    if (!Array.isArray(check.remediationIds) || check.remediationIds.length === 0) failures.push(`${check.id} is missing remediationIds`);
    if (!check.source) {
      failures.push(`${check.id} is missing source`);
      continue;
    }

    const sourcePath = path.resolve(root, check.source);
    let source = "";
    try {
      source = readFileSync(sourcePath, "utf8");
    } catch (error) {
      failures.push(`${check.id} cannot read ${check.source}: ${error.message}`);
      continue;
    }

    const missingTerms = [];
    for (const term of check.requiredTerms ?? []) {
      if (!source.includes(term)) missingTerms.push(term);
    }
    if (missingTerms.length) {
      failures.push(`${check.id} missing required truthfulness terms in ${check.source}: ${missingTerms.join(", ")}`);
    }
    report.capabilityTruthfulness.push({
      id: check.id,
      source: check.source,
      requiredTerms: check.requiredTerms?.length ?? 0,
      missingTerms
    });
  }

  for (const test of suite.tests ?? []) {
    for (const id of test.capabilityTruthfulnessIds ?? []) {
      if (!ids.has(id)) failures.push(`${test.id} references unknown capability truthfulness check: ${id}`);
    }
  }
}

function checkFixtureEvidenceGate() {
  const fixtureTest = (suite.tests ?? []).find((test) => test.id === "L7-A007");
  if (!fixtureTest || fixtureTest.status !== "ready") return;

  const evidence = readJson(options.fixtureEvidence);
  const evidencePath = relative(options.fixtureEvidence);
  if (!evidence || Object.keys(evidence).length === 0) return;

  if (evidence.schema !== "officegen.perfect-spec.fixture-evidence@1.0") {
    failures.push(`${evidencePath} has invalid fixture evidence schema`);
  }
  if (evidence.suiteId !== suite.suiteId) failures.push(`${evidencePath} suiteId does not match suite`);
  if (evidence.release !== suite.release) failures.push(`${evidencePath} release does not match suite`);
  if (evidence.level !== suite.level) failures.push(`${evidencePath} level does not match suite`);
  if (evidence.acceptanceId !== "L7-A007") failures.push(`${evidencePath} acceptanceId must be L7-A007`);
  if (!evidence.generatedAt || Number.isNaN(Date.parse(evidence.generatedAt))) {
    failures.push(`${evidencePath} generatedAt must be an ISO timestamp`);
  }

  const requiredFormats = ["pptx", "docx", "xlsx", "pdf"];
  const declaredFormats = new Set(Array.isArray(evidence.minimumFormats) ? evidence.minimumFormats : []);
  for (const format of requiredFormats) {
    if (!declaredFormats.has(format)) failures.push(`${evidencePath} minimumFormats missing ${format}`);
  }

  const entries = Array.isArray(evidence.entries) ? evidence.entries : [];
  if (entries.length === 0) failures.push(`${evidencePath} has no fixture evidence entries`);
  const entriesById = new Map(entries.map((entry) => [entry.id, entry]));
  const formats = evidence.formats && typeof evidence.formats === "object" ? evidence.formats : {};

  for (const format of requiredFormats) {
    const formatSummary = formats[format];
    if (!formatSummary || typeof formatSummary !== "object") {
      failures.push(`${evidencePath} missing ${format} format coverage summary`);
      continue;
    }
    if (formatSummary.status !== "covered") failures.push(`${evidencePath} ${format} status must be covered`);
    if (!Number.isInteger(formatSummary.count) || formatSummary.count < 1) {
      failures.push(`${evidencePath} ${format} must have at least one fixture evidence entry`);
    }
    const formatEntryIds = Array.isArray(formatSummary.entries) ? formatSummary.entries : [];
    if (formatEntryIds.length < 1) failures.push(`${evidencePath} ${format} entries list is empty`);
    for (const id of formatEntryIds) {
      const entry = entriesById.get(id);
      if (!entry) {
        failures.push(`${evidencePath} ${format} references unknown fixture entry ${id}`);
      } else if (entry.format !== format) {
        failures.push(`${evidencePath} ${id} is listed under ${format} but declares ${entry.format}`);
      }
    }
  }

  for (const entry of entries) {
    const prefix = `${evidencePath} entry ${entry.id ?? "<missing>"}`;
    if (!entry.id || typeof entry.id !== "string") failures.push(`${prefix} is missing id`);
    if (!requiredFormats.includes(entry.format)) failures.push(`${prefix} has invalid format ${entry.format}`);
    if (!entry.testName || typeof entry.testName !== "string") failures.push(`${prefix} is missing testName`);
    if (!entry.command || typeof entry.command !== "string") failures.push(`${prefix} is missing command`);
    if (!entry.targetFeature || typeof entry.targetFeature !== "string") failures.push(`${prefix} is missing targetFeature`);
    if (!entry.status || typeof entry.status !== "string") failures.push(`${prefix} is missing status`);
    if (!entry.source || typeof entry.source !== "object") {
      failures.push(`${prefix} is missing source`);
      continue;
    }
    if (!entry.source.path || typeof entry.source.path !== "string") {
      failures.push(`${prefix} source is missing path`);
      continue;
    }
    if (!/^[a-f0-9]{64}$/.test(entry.source.sha256 ?? "")) failures.push(`${prefix} source has invalid sha256`);
    if (!Number.isInteger(entry.source.bytes) || entry.source.bytes < 0) failures.push(`${prefix} source has invalid bytes`);

    const sourcePath = path.resolve(root, entry.source.path);
    let bytes = 0;
    let actualHash = "";
    try {
      bytes = statSync(sourcePath).size;
      actualHash = sha256File(sourcePath);
    } catch (error) {
      failures.push(`${entry.source.path} cannot be read for fixture evidence integrity: ${error.message}`);
      continue;
    }
    if (bytes !== entry.source.bytes) failures.push(`${entry.source.path} bytes mismatch in ${evidencePath}: expected ${entry.source.bytes}, found ${bytes}`);
    if (actualHash !== entry.source.sha256) failures.push(`${entry.source.path} sha256 mismatch in ${evidencePath}`);

    const generatedArtifact = asRecord(entry.generatedArtifact);
    if (!generatedArtifact || Object.keys(generatedArtifact).length === 0) {
      failures.push(`${prefix} is missing generatedArtifact descriptor evidence`);
      continue;
    }
    checkFixtureGeneratedArtifact(prefix, entry, generatedArtifact);
  }
}

function checkFixtureGeneratedArtifact(prefix, entry, artifact) {
  if (artifact.kind !== "fixture-descriptor") failures.push(`${prefix} generatedArtifact.kind must be fixture-descriptor`);
  if (!artifact.path || typeof artifact.path !== "string") {
    failures.push(`${prefix} generatedArtifact.path is required`);
    return;
  }
  if (!/^[a-f0-9]{64}$/.test(artifact.sha256 ?? "")) failures.push(`${prefix} generatedArtifact has invalid sha256`);
  if (!Number.isInteger(artifact.bytes) || artifact.bytes <= 0) failures.push(`${prefix} generatedArtifact has invalid bytes`);

  const artifactPath = path.resolve(root, artifact.path);
  let actualBytes = 0;
  let actualHash = "";
  let descriptor;
  try {
    actualBytes = statSync(artifactPath).size;
    actualHash = sha256File(artifactPath);
    descriptor = JSON.parse(readFileSync(artifactPath, "utf8"));
  } catch (error) {
    failures.push(`${artifact.path} cannot be read for generated fixture evidence integrity: ${error.message}`);
    return;
  }
  if (actualBytes !== artifact.bytes) failures.push(`${artifact.path} bytes mismatch: expected ${artifact.bytes}, found ${actualBytes}`);
  if (actualHash !== artifact.sha256) failures.push(`${artifact.path} sha256 mismatch`);

  const descriptorRecord = asRecord(descriptor);
  if (descriptorRecord.schema !== "officegen.perfect-spec.fixture-descriptor@1.0") {
    failures.push(`${artifact.path} has invalid fixture descriptor schema`);
  }
  if (descriptorRecord.id !== entry.id) failures.push(`${artifact.path} descriptor id does not match ${entry.id}`);
  if (descriptorRecord.format !== entry.format) failures.push(`${artifact.path} descriptor format does not match ${entry.format}`);
  if (asRecord(descriptorRecord.source).sha256 !== entry.source.sha256) {
    failures.push(`${artifact.path} descriptor source sha256 does not match fixture entry`);
  }
}

function checkCliParityGate() {
  const parityTest = (suite.tests ?? []).find((test) => test.id === "L7-A011");
  if (!parityTest || parityTest.status !== "ready") return;

  const artifact = readJson(options.cliParity);
  const artifactPath = relative(options.cliParity);
  if (!artifact || Object.keys(artifact).length === 0) return;

  if (artifact.schema !== "officegen.cli-parity@1.0") failures.push(`${artifactPath} has invalid CLI parity schema`);
  if (!artifact.generatedAt || Number.isNaN(Date.parse(artifact.generatedAt))) {
    failures.push(`${artifactPath} generatedAt must be an ISO timestamp`);
  }
  if (artifact.summary?.ok !== true) failures.push(`${artifactPath} summary.ok must be true`);

  const commands = Array.isArray(artifact.commands) ? artifact.commands : [];
  if (commands.length < 10) failures.push(`${artifactPath} must include at least 10 command transcripts`);
  const commandsById = new Map(commands.map((command) => [command.id, command]));
  const requiredCommands = new Map([
    ["version-human", { kind: "human", resultSchemaId: null }],
    ["capabilities-human", { kind: "human", resultSchemaId: null }],
    ["capabilities-agent-json", { kind: "agent-json", resultSchemaId: "officegen.capabilities@1.2" }],
    ["schema-list-human", { kind: "human", resultSchemaId: null }],
    ["schema-list-agent-json", { kind: "agent-json", resultSchemaId: "officegen.schema.list@1.2" }],
    ["render-pptx-human", { kind: "human", resultSchemaId: null }],
    ["render-pptx-agent-json", { kind: "agent-json", resultSchemaId: "officegen.render.result@1.2" }],
    ["inspect-pptx-human", { kind: "human", resultSchemaId: null }],
    ["inspect-pptx-agent-json", { kind: "agent-json", resultSchemaId: "officegen.inspect.result@1.2" }],
    ["render-pdf-agent-json", { kind: "agent-json", resultSchemaId: "officegen.render.result@1.2" }],
    ["view-pdf-human", { kind: "human", resultSchemaId: null }],
    ["view-pdf-agent-json", { kind: "agent-json", resultSchemaId: "officegen.view.result@1.2" }]
  ]);

  for (const [id, expected] of requiredCommands) {
    const command = commandsById.get(id);
    if (!command) {
      failures.push(`${artifactPath} missing required command transcript ${id}`);
      continue;
    }
    checkCliParityCommand(artifactPath, command, expected);
  }

  const assertions = Array.isArray(artifact.parityAssertions) ? artifact.parityAssertions : [];
  const requiredAssertions = new Set(["capabilities", "schema-list", "render-pptx", "inspect-pptx", "view-pdf"]);
  for (const id of requiredAssertions) {
    const assertion = assertions.find((entry) => entry.id === id);
    if (!assertion) failures.push(`${artifactPath} missing parity assertion ${id}`);
    else if (assertion.ok !== true) failures.push(`${artifactPath} parity assertion ${id} is not ok`);
  }

  const limitationParity = Array.isArray(artifact.limitationParity) ? artifact.limitationParity : [];
  for (const term of [
    "featureContracts",
    "formatCapabilities",
    "runtimeProfiles",
    "specProfile",
    "knownLimitations",
    "unsupportedNow",
    "SmartArt creation and full SmartArt editing are unsupported",
    "Full-fidelity Office/PDF editing"
  ]) {
    const entry = limitationParity.find((item) => item.term === term);
    if (!entry?.present) failures.push(`${artifactPath} missing limitation parity term: ${term}`);
  }
}

function checkCliParityCommand(artifactPath, command, expected) {
  const prefix = `${artifactPath} command ${command.id ?? "<missing>"}`;
  if (!command.id || typeof command.id !== "string") failures.push(`${prefix} is missing id`);
  if (!command.command || typeof command.command !== "string") failures.push(`${prefix} is missing command`);
  if (command.kind !== expected.kind) failures.push(`${prefix} kind must be ${expected.kind}`);
  if (!Number.isInteger(command.exitCode)) failures.push(`${prefix} exitCode must be an integer`);
  if (command.exitCode !== 0) failures.push(`${prefix} exitCode must be 0`);
  checkDigestRecord(`${prefix} stdout`, command.stdout, { requireNonEmpty: true });
  if (command.stderr) checkDigestRecord(`${prefix} stderr`, command.stderr, { requireNonEmpty: false });

  const parsed = command.parsed && typeof command.parsed === "object" ? command.parsed : undefined;
  if (!parsed) {
    failures.push(`${prefix} is missing parsed status`);
    return;
  }
  if (typeof parsed.ok !== "boolean") failures.push(`${prefix} parsed.ok must be boolean`);
  if (parsed.ok !== true) failures.push(`${prefix} parsed.ok must be true`);
  if (!Object.prototype.hasOwnProperty.call(parsed, "schemaId")) failures.push(`${prefix} parsed.schemaId is required`);
  if (!Object.prototype.hasOwnProperty.call(parsed, "resultSchemaId")) failures.push(`${prefix} parsed.resultSchemaId is required`);

  if (expected.kind === "agent-json") {
    if (parsed.schemaId !== "officegen.envelope@1.2") failures.push(`${prefix} parsed.schemaId must be officegen.envelope@1.2`);
    if (parsed.resultSchemaId !== expected.resultSchemaId) failures.push(`${prefix} parsed.resultSchemaId must be ${expected.resultSchemaId}`);
    if (parsed.envelopeOk !== true) failures.push(`${prefix} parsed.envelopeOk must be true`);
  } else {
    if (parsed.schemaId !== null) failures.push(`${prefix} text transcript parsed.schemaId must be null`);
  }
}

function checkNativeRendererGate() {
  const nativeTest = (suite.tests ?? []).find((test) => test.id === "L7-A008");
  if (!nativeTest || nativeTest.status !== "ready") return;

  const artifactPath = relative(options.nativeRenderer);
  if (!fileExists(options.nativeRenderer)) {
    const message = `${artifactPath} is required when L7-A008 is ready; run npm run perfect-spec:evidence to generate it`;
    if (options.gate === "publish") failures.push(message);
    else warnings.push(message);
    return;
  }

  const artifact = readJson(options.nativeRenderer);
  if (!artifact || Object.keys(artifact).length === 0) return;
  checkNativeRendererArtifact(artifactPath, artifact);
}

function checkNativeRendererArtifact(artifactPath, artifact) {
  if (artifact.schema !== "officegen.perfect-spec.native-renderer-evidence@1.0") {
    failures.push(`${artifactPath} has invalid native renderer evidence schema`);
  }
  if (artifact.l7AcceptanceId !== "L7-A008") failures.push(`${artifactPath} l7AcceptanceId must be L7-A008`);
  if (!artifact.generatedAt || Number.isNaN(Date.parse(artifact.generatedAt))) {
    failures.push(`${artifactPath} generatedAt must be an ISO timestamp`);
  }

  const decision = asRecord(artifact.decision);
  if (decision.status !== "ready") failures.push(`${artifactPath} decision.status must be ready`);
  if (decision.nativeExecutionRequired !== false) failures.push(`${artifactPath} decision.nativeExecutionRequired must be false`);
  if (decision.publishBlocker !== false) failures.push(`${artifactPath} decision.publishBlocker must be false`);

  const commands = asRecord(artifact.commands);
  const doctorCommand = asRecord(commands.doctor);
  if (doctorCommand.exitCode !== 0) failures.push(`${artifactPath} doctor command must exit 0`);
  const doctorEnvelope = asRecord(doctorCommand.stdout);
  if (doctorEnvelope.ok !== true) failures.push(`${artifactPath} doctor envelope ok must be true`);
  if (asRecord(doctorEnvelope.result).schema !== "officegen.renderer.doctor@2.2") {
    failures.push(`${artifactPath} doctor result schema must be officegen.renderer.doctor@2.2`);
  }

  const environment = asRecord(artifact.environment);
  const libreOffice = asRecord(environment.libreOffice);
  if (libreOffice.checked !== true) failures.push(`${artifactPath} must record LibreOffice availability check`);
  if (typeof libreOffice.available !== "boolean") failures.push(`${artifactPath} libreOffice.available must be boolean`);
  const windowsOfficeCom = asRecord(environment.windowsOfficeCom);
  if (windowsOfficeCom.checked !== true) failures.push(`${artifactPath} must record Windows Office COM availability check`);
  if (typeof windowsOfficeCom.available !== "boolean") failures.push(`${artifactPath} windowsOfficeCom.available must be boolean`);
  if (!Array.isArray(windowsOfficeCom.renderers)) failures.push(`${artifactPath} windowsOfficeCom.renderers must be an array`);

  const defaultPolicy = asRecord(artifact.defaultPolicy);
  if (defaultPolicy.externalProcess !== "deny") failures.push(`${artifactPath} defaultPolicy.externalProcess must be deny`);
  if (defaultPolicy.nativeConversionBlockedByDefault !== true) {
    failures.push(`${artifactPath} defaultPolicy.nativeConversionBlockedByDefault must be true`);
  }
  const blockedExpectation = asRecord(defaultPolicy.blockedExpectation);
  if (blockedExpectation.expectedErrorCode !== "SECURITY_EXTERNAL_PROCESS_DENIED") {
    failures.push(`${artifactPath} blockedExpectation.expectedErrorCode must be SECURITY_EXTERNAL_PROCESS_DENIED`);
  }
  if (blockedExpectation.observedErrorCode !== "SECURITY_EXTERNAL_PROCESS_DENIED") {
    failures.push(`${artifactPath} blockedExpectation.observedErrorCode must be SECURITY_EXTERNAL_PROCESS_DENIED`);
  }

  const optionality = asRecord(artifact.nativeGateOptionality);
  if (optionality.nativeExecutionRequiredForThisEvidence !== false) {
    failures.push(`${artifactPath} nativeGateOptionality.nativeExecutionRequiredForThisEvidence must be false`);
  }
  if (optionality.nonNativeReleaseGateIndependent !== true) {
    failures.push(`${artifactPath} nativeGateOptionality.nonNativeReleaseGateIndependent must be true`);
  }
  if (optionality.doctorIsSafeDiscovery !== true) failures.push(`${artifactPath} nativeGateOptionality.doctorIsSafeDiscovery must be true`);
  const requirements = Array.isArray(optionality.requiredForNativeSuccess) ? optionality.requiredForNativeSuccess : [];
  for (const requiredTerm of ["externalProcess=allow", "renderers=enabled"]) {
    if (!requirements.some((requirement) => String(requirement).includes(requiredTerm))) {
      failures.push(`${artifactPath} requiredForNativeSuccess must mention ${requiredTerm}`);
    }
  }
  const releaseRunnerRequirements = Array.isArray(optionality.releaseRunnerRequirements) ? optionality.releaseRunnerRequirements : [];
  if (releaseRunnerRequirements.length < 2) failures.push(`${artifactPath} must document tag/release runner native requirements`);
}

function checkDigestRecord(prefix, record, { requireNonEmpty }) {
  if (!record || typeof record !== "object") {
    failures.push(`${prefix} digest record is missing`);
    return;
  }
  if (!record.path || typeof record.path !== "string") {
    failures.push(`${prefix} path is required`);
    return;
  }
  if (!/^[a-f0-9]{64}$/.test(record.sha256 ?? "")) failures.push(`${prefix} has invalid sha256`);
  if (!Number.isInteger(record.bytes) || record.bytes < 0) failures.push(`${prefix} has invalid bytes`);
  if (requireNonEmpty && record.bytes <= 0) failures.push(`${prefix} must be non-empty`);

  const file = path.resolve(root, record.path);
  let actualBytes = 0;
  let actualHash = "";
  try {
    actualBytes = statSync(file).size;
    actualHash = sha256File(file);
  } catch (error) {
    failures.push(`${record.path} cannot be read for CLI parity integrity: ${error.message}`);
    return;
  }
  if (actualBytes !== record.bytes) failures.push(`${record.path} bytes mismatch: expected ${record.bytes}, found ${actualBytes}`);
  if (actualHash !== record.sha256) failures.push(`${record.path} sha256 mismatch`);
}

function checkV31RuntimeEvidenceGate() {
  if (suite.release !== "3.1.0") return;
  const matrixPath = relative(options.v31EvidenceMatrix);
  const artifact = readJson(options.v31EvidenceMatrix);
  if (!artifact || Object.keys(artifact).length === 0) return;

  if (artifact.schema !== "officegen.v3.1.0.evidence-matrix@1.0") {
    failures.push(`${matrixPath} has invalid v3.1.0 evidence matrix schema`);
  }
  if (artifact.release !== "3.1.0") failures.push(`${matrixPath} release must be 3.1.0`);

  const phase0 = asRecord(artifact.phase0);
  if (phase0.priority !== "P0") failures.push(`${matrixPath} phase0.priority must be P0`);
  if (phase0.status !== "complete") failures.push(`${matrixPath} phase0.status must be complete`);
  if (phase0.acceptanceId !== "L7-A013") failures.push(`${matrixPath} phase0.acceptanceId must be L7-A013`);

  checkV31CapabilityProfiles(matrixPath, artifact);

  const runtimeProjection = asRecord(artifact.runtimeProjection);
  if (runtimeProjection.version !== "runtime-v2") failures.push(`${matrixPath} runtimeProjection.version must be runtime-v2`);
  if (runtimeProjection.status !== "visible") failures.push(`${matrixPath} runtimeProjection.status must be visible`);
  if (runtimeProjection.currentProfileId !== "current-limited-v3.1") failures.push(`${matrixPath} runtimeProjection.currentProfileId must be current-limited-v3.1`);
  if (runtimeProjection.support !== "supported-current") failures.push(`${matrixPath} runtimeProjection.support must be supported-current`);
  if (runtimeProjection.acceptanceId !== "L7-A014") failures.push(`${matrixPath} runtimeProjection.acceptanceId must be L7-A014`);

  const requiredPhases = Array.isArray(runtimeProjection.requiredPhases) ? runtimeProjection.requiredPhases : [];
  for (const phase of ["inspect", "select", "plan", "dry-run", "edit", "verify", "diff", "repair", "report"]) {
    if (!requiredPhases.includes(phase)) failures.push(`${matrixPath} runtimeProjection.requiredPhases missing ${phase}`);
  }
  const requiredTerms = Array.isArray(runtimeProjection.requiredManifestTerms) ? runtimeProjection.requiredManifestTerms : [];
  for (const term of ["skeleton-evidence", "does not execute complete autonomous repair"]) {
    if (!requiredTerms.includes(term)) failures.push(`${matrixPath} runtimeProjection.requiredManifestTerms missing ${term}`);
  }

  const phase0Test = (suite.tests ?? []).find((test) => test.id === "L7-A013");
  const runtimeTest = (suite.tests ?? []).find((test) => test.id === "L7-A014");
  if (!phase0Test || phase0Test.status !== "ready" || phase0Test.blocking !== true) {
    failures.push("L7-A013 must be ready and blocking for v3.1.0 Phase 0 P0 completion");
  }
  if (!runtimeTest || runtimeTest.status !== "ready" || runtimeTest.blocking !== true) {
    failures.push("L7-A014 must be ready and blocking for v3.1.0 runtime-v2 office-agent projection");
  }
}

function checkV31CapabilityProfiles(matrixPath, artifact) {
  const capabilityProfiles = asRecord(artifact.capabilityProfiles);
  if (capabilityProfiles.currentProfileId !== "current-limited-v3.1") {
    failures.push(`${matrixPath} capabilityProfiles.currentProfileId must be current-limited-v3.1`);
  }
  if (capabilityProfiles.targetProfileId !== "perfect-runtime-target") {
    failures.push(`${matrixPath} capabilityProfiles.targetProfileId must be perfect-runtime-target`);
  }
  if (!String(capabilityProfiles.truthfulnessPolicy ?? "").includes("current-limited-v3.1")) {
    failures.push(`${matrixPath} capabilityProfiles.truthfulnessPolicy must mention current-limited-v3.1`);
  }

  const profileEvidence = Array.isArray(capabilityProfiles.profileEvidence) ? capabilityProfiles.profileEvidence : [];
  const evidenceById = new Map(profileEvidence.map((entry) => [entry.id, entry]));
  const runtimeV2 = asRecord(evidenceById.get("runtime-v2-projections"));
  if (runtimeV2.currentStatus !== "supported") failures.push(`${matrixPath} runtime-v2-projections currentStatus must be supported`);
  if (runtimeV2.targetStatus !== "supported") failures.push(`${matrixPath} runtime-v2-projections targetStatus must be supported`);

  const remainingTargetGaps = Array.isArray(artifact.remainingTargetGaps) ? artifact.remainingTargetGaps : [];
  if (remainingTargetGaps.length < 3) failures.push(`${matrixPath} remainingTargetGaps must list current-vs-target gaps`);
  const gapsById = new Map(remainingTargetGaps.map((entry) => [entry.id, entry]));
  for (const id of ["smartart-editing", "pdf-true-redaction"]) {
    const gap = asRecord(gapsById.get(id));
    if (!gap || Object.keys(gap).length === 0) {
      failures.push(`${matrixPath} remainingTargetGaps missing ${id}`);
      continue;
    }
    if (gap.currentProfile !== "current-limited-v3.1") failures.push(`${matrixPath} ${id} currentProfile must be current-limited-v3.1`);
    if (gap.targetProfile !== "perfect-runtime-target") failures.push(`${matrixPath} ${id} targetProfile must be perfect-runtime-target`);
    if (gap.currentStatus !== "unsupported") failures.push(`${matrixPath} ${id} currentStatus must be unsupported`);
    if (gap.targetStatus !== "target-only") failures.push(`${matrixPath} ${id} targetStatus must be target-only`);
    if (gap.requiredForPerfectRuntime !== true) failures.push(`${matrixPath} ${id} requiredForPerfectRuntime must be true`);
  }
}

function checkPublishGate() {
  if (options.gate !== "publish") return;
  for (const test of suite.tests ?? []) {
    if (test.blocking && test.status !== "ready") {
      failures.push(`${test.id} is blocking and ${test.status}; publish gate requires ready`);
    }
  }

  const evidenceTest = (suite.tests ?? []).find((test) => test.id === "L7-A005");
  if (!evidenceTest) {
    failures.push("L7-A005 evidence bundle test is missing; publish gate requires it");
  } else if (evidenceTest.status !== "ready") {
    failures.push(`L7-A005 is ${evidenceTest.status}; publish gate requires ready evidence bundle manifest`);
  }

  checkPostTagSmokeGate();
  checkEvidenceManifest();
  checkV31PublishEvidence();
}

function checkV31PublishEvidence() {
  if (suite.release !== "3.1.0") return;
  const manifest = readJson(options.evidenceManifest);
  const artifacts = Array.isArray(manifest.artifacts) ? manifest.artifacts : [];
  const evidenceArtifact = artifacts.find((artifact) => artifact.role === "v3.1.0-evidence-matrix");
  if (!evidenceArtifact) {
    failures.push(`${relative(options.evidenceManifest)} is missing v3.1.0-evidence-matrix artifact`);
    return;
  }
  const artifactPath = path.resolve(root, evidenceArtifact.path);
  const generated = readJson(artifactPath);
  if (generated.schema !== "officegen.perfect-spec.v3.1.0-evidence-matrix@1.0") {
    failures.push(`${evidenceArtifact.path} has invalid generated v3.1.0 evidence matrix schema`);
  }
  if (generated.release !== "3.1.0") failures.push(`${evidenceArtifact.path} release must be 3.1.0`);
  if (asRecord(generated.phase0).status !== "complete") failures.push(`${evidenceArtifact.path} phase0.status must be complete`);
  if (asRecord(generated.runtimeProjection).version !== "runtime-v2") failures.push(`${evidenceArtifact.path} runtimeProjection.version must be runtime-v2`);
  if (asRecord(generated.capabilityProfiles).currentProfileId !== "current-limited-v3.1") {
    failures.push(`${evidenceArtifact.path} capabilityProfiles.currentProfileId must be current-limited-v3.1`);
  }
  const remainingTargetGaps = Array.isArray(generated.remainingTargetGaps) ? generated.remainingTargetGaps : [];
  if (!remainingTargetGaps.some((gap) => gap.id === "pdf-true-redaction" && gap.currentStatus === "unsupported" && gap.targetStatus === "target-only")) {
    failures.push(`${evidenceArtifact.path} remainingTargetGaps must preserve pdf-true-redaction as current unsupported / target-only`);
  }
  if (asRecord(generated.publishGateConnection).requiresPostTagSmoke !== true) {
    failures.push(`${evidenceArtifact.path} publishGateConnection.requiresPostTagSmoke must be true`);
  }
}

function checkPostTagSmokeGate() {
  const postTagTest = (suite.tests ?? []).find((test) => test.id === "L7-A009");
  if (!postTagTest) {
    failures.push("L7-A009 post-tag install smoke test is missing; publish gate requires explicit post-tag evidence");
    return;
  }

  if (postTagTest.status === "blocked") {
    failures.push("L7-A009 is blocked; publish gate requires post-tag install evidence before release completion");
  }

  const artifactPath = relative(options.postTagSmoke);
  if (!fileExists(options.postTagSmoke)) {
    failures.push(`${artifactPath} is required for L7-A009 before release completion; run npm run perfect-spec:post-tag-smoke`);
    return;
  }

  const artifact = readJson(options.postTagSmoke);
  if (!artifact || Object.keys(artifact).length === 0) return;

  if (artifact.schema !== "officegen.perfect-spec.post-tag-smoke@1.0") {
    failures.push(`${artifactPath} has invalid post-tag smoke schema`);
  }
  if (artifact.acceptanceId !== "L7-A009") failures.push(`${artifactPath} acceptanceId must be L7-A009`);
  if (!artifact.generatedAt || Number.isNaN(Date.parse(artifact.generatedAt))) {
    failures.push(`${artifactPath} generatedAt must be an ISO timestamp`);
  }
  if (artifact.ok !== true) failures.push(`${artifactPath} ok must be true`);

  const checks = Array.isArray(artifact.checks) ? artifact.checks : [];
  const checksById = new Map(checks.map((check) => [check.id, check]));
  for (const id of ["github-install-tag-smoke", "github-install-remote-smoke"]) {
    const check = checksById.get(id);
    if (!check) {
      failures.push(`${artifactPath} missing ${id} check`);
      continue;
    }
    if (check.exitCode !== 0) failures.push(`${artifactPath} ${id} exitCode must be 0`);
    checkDigestRecord(`${artifactPath} ${id} log`, check.log, { requireNonEmpty: true });
  }
}

function checkEvidenceManifest() {
  const manifest = readJson(options.evidenceManifest);
  const manifestPath = relative(options.evidenceManifest);
  if (!manifest || Object.keys(manifest).length === 0) return;

  if (manifest.schema !== "officegen.perfect-spec.evidence-manifest@1.0") {
    failures.push(`${manifestPath} has invalid evidence manifest schema`);
  }
  if (manifest.suiteId !== suite.suiteId) failures.push(`${manifestPath} suiteId does not match suite`);
  if (manifest.release !== suite.release) failures.push(`${manifestPath} release does not match suite`);
  if (manifest.level !== suite.level) failures.push(`${manifestPath} level does not match suite`);
  if (!manifest.generatedAt || Number.isNaN(Date.parse(manifest.generatedAt))) {
    failures.push(`${manifestPath} generatedAt must be an ISO timestamp`);
  }

  const suiteIds = (suite.tests ?? []).map((test) => test.id);
  const manifestTests = Array.isArray(manifest.tests) ? manifest.tests : [];
  const manifestIds = manifestTests.map((test) => test.id);
  const manifestL7Ids = Array.isArray(manifest.l7Ids) ? manifest.l7Ids : [];

  if (manifestTests.length !== suiteIds.length) {
    failures.push(`${manifestPath} must enumerate all ${suiteIds.length} suite tests, found ${manifestTests.length}`);
  }

  for (const id of suiteIds) {
    if (!/^L7-A[0-9]{3}$/.test(id)) failures.push(`${id} is not an L7 acceptance id`);
    if (!manifestIds.includes(id)) failures.push(`${manifestPath} missing test entry for ${id}`);
    if (!manifestL7Ids.includes(id)) failures.push(`${manifestPath} missing l7Ids entry for ${id}`);
  }

  for (const id of manifestIds) {
    if (!suiteIds.includes(id)) failures.push(`${manifestPath} includes unknown test entry ${id}`);
  }

  const artifacts = Array.isArray(manifest.artifacts) ? manifest.artifacts : [];
  if (artifacts.length === 0) failures.push(`${manifestPath} has no artifacts`);
  const requiredArtifactRoles = ["summary", "events", "check-output"];
  const nativeTest = (suite.tests ?? []).find((test) => test.id === "L7-A008");
  const postTagTest = (suite.tests ?? []).find((test) => test.id === "L7-A009");
  if (nativeTest?.status === "ready") requiredArtifactRoles.push("native-renderer");
  if (postTagTest?.status === "pending" || postTagTest?.status === "ready") requiredArtifactRoles.push("post-tag-smoke");
  if (suite.release === "3.1.0") requiredArtifactRoles.push("v3.1.0-evidence-matrix");
  for (const role of requiredArtifactRoles) {
    if (!artifacts.some((artifact) => artifact.role === role)) {
      failures.push(`${manifestPath} is missing ${role} artifact`);
    }
  }

  for (const artifact of artifacts) {
    if (!artifact.role || typeof artifact.role !== "string") failures.push(`${manifestPath} has artifact without role`);
    if (!artifact.path || typeof artifact.path !== "string") {
      failures.push(`${manifestPath} has artifact without path`);
      continue;
    }
    if (!/^[a-f0-9]{64}$/.test(artifact.sha256 ?? "")) failures.push(`${artifact.path} has invalid sha256 in ${manifestPath}`);
    if (!Number.isInteger(artifact.bytes) || artifact.bytes < 0) failures.push(`${artifact.path} has invalid bytes in ${manifestPath}`);

    const artifactPath = path.resolve(root, artifact.path);
    let bytes = 0;
    let actualHash = "";
    try {
      bytes = statSync(artifactPath).size;
      actualHash = sha256File(artifactPath);
    } catch (error) {
      failures.push(`${artifact.path} cannot be read for evidence integrity: ${error.message}`);
      continue;
    }
    if (bytes !== artifact.bytes) failures.push(`${artifact.path} bytes mismatch in ${manifestPath}: expected ${artifact.bytes}, found ${bytes}`);
    if (actualHash !== artifact.sha256) failures.push(`${artifact.path} sha256 mismatch in ${manifestPath}`);
  }

  const eventsArtifact = artifacts.find((artifact) => artifact.role === "events");
  if (eventsArtifact?.path) checkEventsJsonl(eventsArtifact.path, suiteIds);
  const nativeArtifact = artifacts.find((artifact) => artifact.role === "native-renderer");
  if (nativeTest?.status === "ready" && nativeArtifact?.path) {
    checkNativeRendererArtifact(nativeArtifact.path, readJson(path.resolve(root, nativeArtifact.path)));
  }
}

function checkEventsJsonl(file, suiteIds) {
  let text = "";
  try {
    text = readFileSync(path.resolve(root, file), "utf8");
  } catch {
    return;
  }

  const declared = new Set();
  for (const [index, line] of text.split(/\r?\n/).entries()) {
    if (!line.trim()) continue;
    let event;
    try {
      event = JSON.parse(line);
    } catch (error) {
      failures.push(`${file}:${index + 1} is not valid JSONL: ${error.message}`);
      continue;
    }
    if (event.type === "test.declared" && event.id) declared.add(event.id);
  }

  for (const id of suiteIds) {
    if (!declared.has(id)) failures.push(`${file} missing test.declared event for ${id}`);
  }
}

function readJson(file) {
  try {
    return JSON.parse(readFileSync(file, "utf8"));
  } catch (error) {
    failures.push(`${relative(file)} is not parseable JSON: ${error.message}`);
    return {};
  }
}

function fileExists(file) {
  try {
    statSync(file);
    return true;
  } catch {
    return false;
  }
}

function asRecord(value) {
  return value && typeof value === "object" && !Array.isArray(value) ? value : {};
}

function sha256File(file) {
  return createHash("sha256").update(readFileSync(file)).digest("hex");
}

function readMatrix(file) {
  const rows = new Map();
  let text = "";
  try {
    text = readFileSync(file, "utf8");
  } catch (error) {
    failures.push(`${relative(file)} cannot be read: ${error.message}`);
  }

  for (const line of text.split(/\r?\n/)) {
    if (!line.startsWith("| V")) continue;
    const cells = line.split("|").map((cell) => cell.trim()).filter(Boolean);
    const id = cells[0];
    if (!/^V[0-9]+-[A-Z][0-9]{3}$/.test(id)) continue;
    rows.set(id, {
      id,
      area: cells[1],
      finding: cells[2],
      status: cells[3],
      evidence: cells[4]
    });
  }

  if (rows.size === 0) failures.push(`${relative(file)} has no remediation matrix rows`);
  return { rows };
}

function printVisibleNonReady() {
  if (report.nonReady.length === 0) return;
  console.log("perfect-spec:check visible non-ready items:");
  for (const item of report.nonReady) {
    const detail = item.blockedBy ? ` blockedBy=${item.blockedBy}` : "";
    console.log(`- ${item.id} ${item.priority} ${item.status}:${detail} next=${item.nextAction}`);
  }
}

function relative(file) {
  return path.relative(root, file).replaceAll(path.sep, "/");
}

function usage(message) {
  console.error(`perfect-spec:check: ${message}`);
  console.error("usage: node scripts/check-perfect-spec.mjs [--gate=visibility|publish] [--json] [--suite=path] [--matrix=path] [--v31-evidence-matrix=path] [--evidence-manifest=path] [--fixture-evidence=path] [--cli-parity=path] [--native-renderer=path] [--post-tag-smoke=path]");
  process.exit(2);
}
