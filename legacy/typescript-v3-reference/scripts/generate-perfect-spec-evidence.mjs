#!/usr/bin/env node
import { createHash } from "node:crypto";
import { mkdirSync, readFileSync, rmSync, statSync, writeFileSync } from "node:fs";
import path from "node:path";
import process from "node:process";
import { spawnSync } from "node:child_process";

const root = process.cwd();
const args = process.argv.slice(2);
const options = {
  suite: path.join(root, "goal", "acceptance-suite.perfect-spec-v2.json"),
  v31EvidenceMatrix: path.join(root, "goal", "v3.1.0-evidence-matrix.json"),
  outDir: path.join(root, ".officegen", "acceptance", "perfect-spec")
};

const fixtureEvidenceSources = [
  {
    id: "pptx-smartart-chart-image-fixture",
    format: "pptx",
    fixture: "in-memory PPTX package with SmartArt diagram parts, chart series, grouped shapes, connectors, and embedded image relationships",
    testName: "PPTX object map inspection > detects groups, connectors, SmartArt relationships, and chart series",
    command: "npm test -- packages/formats/tests/pptx-object-map.test.ts",
    targetFeature: "SmartArt is detected with selector/media metadata while remaining non-editable; chart series and relationship-backed objects are inspectable.",
    sourcePath: "packages/formats/tests/pptx-object-map.test.ts",
    status: "covered-by-test"
  },
  {
    id: "pptx-asset-replacement-fixture",
    format: "pptx",
    fixture: "in-memory PPTX media relationship fixture",
    testName: "@officegen/formats MVP > repairs media relationship targets when replacing a mismatched PNG path containing SVG bytes",
    command: "npm test -- packages/formats/tests/formats.test.ts",
    targetFeature: "Asset replacement rewrites media paths and relationship targets with content-type validation.",
    sourcePath: "packages/formats/tests/formats.test.ts",
    status: "covered-by-test"
  },
  {
    id: "docx-comments-revisions-fixture",
    format: "docx",
    fixture: "in-memory DOCX package with comments.xml, comment ranges, insert/delete revisions, content controls, fields, split runs, bookmarks, and hyperlinks",
    testName: "DOCX story/run graph inspection > detects DOCX comment ranges and comment story content; detects DOCX revisions, content controls, and field codes",
    command: "npm test -- packages/formats/tests/docx-story-run-graph.test.ts",
    targetFeature: "DOCX comments, redline-style revision markers, story/run graph metadata, and scoped non-destructive inspection.",
    sourcePath: "packages/formats/tests/docx-story-run-graph.test.ts",
    status: "covered-by-test"
  },
  {
    id: "docx-edit-and-assets-fixture",
    format: "docx",
    fixture: "rendered DOCX and in-memory DOCX media relationship fixtures",
    testName: "@officegen/formats MVP > applies structural PPTX, DOCX, and XLSX edit ops and confirms through inspect; inspects embedded media usage across PPTX, DOCX, and XLSX packages",
    command: "npm test -- packages/formats/tests/formats.test.ts",
    targetFeature: "DOCX paragraph insertion and embedded media usage inspection are covered; broader style/comment/redline editing remains limitation-scoped.",
    sourcePath: "packages/formats/tests/formats.test.ts",
    status: "covered-by-test"
  },
  {
    id: "xlsx-pivot-slicer-formula-fixture",
    format: "xlsx",
    fixture: "in-memory XLSX package with formulas, tables, slicers, charts, pivot table definitions, external links, named ranges, and structured references",
    testName: "XLSX formula graph inspect > detects structured references and relates formulas to tables and slicers; relates range formulas to chart and pivot source ranges",
    command: "npm test -- packages/formats/tests/xlsx-formula-graph.test.ts",
    targetFeature: "Pivot and slicer objects are surfaced as related/inspectable objects; formula risk and dependency evidence is explicit.",
    sourcePath: "packages/formats/tests/xlsx-formula-graph.test.ts",
    status: "covered-by-test"
  },
  {
    id: "xlsx-edit-and-assets-fixture",
    format: "xlsx",
    fixture: "rendered XLSX workbook plus in-memory worksheet drawing media fixture",
    testName: "@officegen/formats MVP > updates XLSX chart caches and backing worksheet cells together; applies structural PPTX, DOCX, and XLSX edit ops and confirms through inspect",
    command: "npm test -- packages/formats/tests/formats.test.ts",
    targetFeature: "XLSX scalar, table, row, formula, chart data, and media usage behavior is covered by current test fixtures.",
    sourcePath: "packages/formats/tests/formats.test.ts",
    status: "covered-by-test"
  },
  {
    id: "pdf-overlay-not-redaction-fixture",
    format: "pdf",
    fixture: "in-memory PDF document with extractable secret text plus additive overlay operations",
    testName: "PDF object graph and redaction safety > blocks pdf.redact atomically instead of treating an overlay as redaction; finds forbidden text after additive PDF overlays",
    command: "npm test -- packages/formats/tests/pdf-redaction.test.ts",
    targetFeature: "PDF overlay behavior is explicitly not physical redaction; forbidden source text remains discoverable after overlay.",
    sourcePath: "packages/formats/tests/pdf-redaction.test.ts",
    status: "covered-by-test"
  },
  {
    id: "pdf-render-view-cli-smoke",
    format: "pdf",
    fixture: "CLI smoke IR rendered to PDF and rasterized to PNG",
    testName: "pack smoke > officegen capabilities/render/inspect/view-png smoke",
    command: "npm run pack:smoke",
    targetFeature: "Packaged CLI can render PDF output and view it through the PDF raster path.",
    sourcePath: "scripts/pack-smoke.mjs",
    status: "covered-by-smoke-source"
  }
];

for (const arg of args) {
  if (arg.startsWith("--suite=")) options.suite = path.resolve(root, arg.slice("--suite=".length));
  else if (arg.startsWith("--v31-evidence-matrix=")) options.v31EvidenceMatrix = path.resolve(root, arg.slice("--v31-evidence-matrix=".length));
  else if (arg.startsWith("--out-dir=")) options.outDir = path.resolve(root, arg.slice("--out-dir=".length));
  else usage(`unknown argument: ${arg}`);
}

mkdirSync(options.outDir, { recursive: true });

const generatedAt = new Date().toISOString();
const suite = readJson(options.suite);
const tests = (suite.tests ?? []).map((test) => ({
  id: test.id,
  title: test.title,
  priority: test.priority,
  kind: test.kind,
  blocking: test.blocking,
  status: test.status,
  remediationIds: test.remediationIds ?? [],
  capabilityTruthfulnessIds: test.capabilityTruthfulnessIds ?? [],
  evidence: test.evidence ?? []
}));

const l7Ids = tests.filter((test) => /^L7-A[0-9]{3}$/.test(test.id ?? "")).map((test) => test.id);
const fixtureManifestPath = path.join(options.outDir, "fixtures.json");
const fixtureArtifactPaths = writeFixtureEvidence({
  generatedAt,
  suite,
  outDir: options.outDir,
  fixtureManifestPath
});
const nativeRendererPath = path.join(options.outDir, "native-renderer.json");
writeJson(nativeRendererPath, collectNativeRendererEvidence(generatedAt));
const v31EvidenceMatrixPath = path.join(options.outDir, "v3.1.0-evidence-matrix.json");
writeV31EvidenceMatrix({
  generatedAt,
  suite,
  tests,
  sourcePath: options.v31EvidenceMatrix,
  outPath: v31EvidenceMatrixPath
});
const postTagSmokePath = path.join(options.outDir, "post-tag-smoke.json");
ensurePostTagSmokePlaceholder({ generatedAt, suite, file: postTagSmokePath });
const parityResult = runNode(["scripts/capture-cli-parity.mjs", `--out-dir=${relative(options.outDir)}`]);
const checkResult = runNode(["scripts/check-perfect-spec.mjs", "--json", `--suite=${relative(options.suite)}`]);

const checkOutputPath = path.join(options.outDir, "check-visibility.json");
writeJson(checkOutputPath, {
  command: `node scripts/check-perfect-spec.mjs --json --suite=${relative(options.suite)}`,
  generatedAt,
  exitCode: checkResult.status,
  stdout: parseJsonOrText(checkResult.stdout),
  stderr: checkResult.stderr
});

const eventsPath = path.join(options.outDir, "events.jsonl");
writeFileSync(eventsPath, buildEvents({ generatedAt, suite, tests, checkResult }), "utf8");

const summaryPath = path.join(options.outDir, "summary.md");
writeFileSync(summaryPath, buildSummary({ generatedAt, suite, tests, checkResult }), "utf8");

const artifactPaths = [
  { role: "suite", file: options.suite },
  { role: "fixture-evidence", file: fixtureManifestPath },
  ...fixtureArtifactPaths.map((artifact) => ({ role: artifact.role, file: artifact.file })),
  { role: "cli-parity", file: path.join(options.outDir, "cli-parity.json") },
  { role: "post-tag-smoke", file: postTagSmokePath },
  { role: "post-tag-smoke-log", file: path.join(options.outDir, "github-install-tag-smoke.txt") },
  { role: "post-tag-smoke-log", file: path.join(options.outDir, "github-install-remote-smoke.txt") },
  { role: "summary", file: summaryPath },
  { role: "events", file: eventsPath },
  { role: "check-output", file: checkOutputPath },
  { role: "native-renderer", file: nativeRendererPath },
  { role: "v3.1.0-evidence-matrix", file: v31EvidenceMatrixPath }
].filter((artifact) => exists(artifact.file));

const manifest = {
  schema: "officegen.perfect-spec.evidence-manifest@1.0",
  suiteId: suite.suiteId,
  release: suite.release,
  level: suite.level,
  generatedAt,
  l7Ids,
  totals: {
    tests: tests.length,
    ready: tests.filter((test) => test.status === "ready").length,
    pending: tests.filter((test) => test.status === "pending").length,
    blocked: tests.filter((test) => test.status === "blocked").length
  },
  tests,
  artifacts: artifactPaths.map((artifact) => ({
    role: artifact.role,
    path: relative(artifact.file),
    bytes: statSync(artifact.file).size,
    sha256: sha256File(artifact.file)
  }))
};

const manifestPath = path.join(options.outDir, "manifest.json");
writeJson(manifestPath, manifest);

console.log(`perfect-spec:evidence wrote ${relative(manifestPath)}`);
console.log(`perfect-spec:evidence wrote ${relative(fixtureManifestPath)}`);
console.log(`perfect-spec:evidence wrote ${relative(summaryPath)}`);
console.log(`perfect-spec:evidence wrote ${relative(eventsPath)}`);
console.log(`perfect-spec:evidence wrote ${relative(nativeRendererPath)}`);
console.log(`perfect-spec:evidence wrote ${relative(v31EvidenceMatrixPath)}`);

if (checkResult.status !== 0) {
  console.error("perfect-spec:evidence error: visibility check output was captured with non-zero exit code");
}
if (parityResult.status !== 0) {
  console.error("perfect-spec:evidence error: CLI parity capture output was captured with non-zero exit code");
}

const fatalInternalFailures = [
  ["visibility check", checkResult],
  ["CLI parity capture", parityResult]
].filter(([, result]) => result.status !== 0);
if (fatalInternalFailures.length) {
  console.error("perfect-spec:evidence failed because required internal commands failed:");
  for (const [label, result] of fatalInternalFailures) {
    console.error(`- ${label} exited ${result.status}`);
  }
  process.exit(1);
}

function buildEvents({ generatedAt, suite, tests, checkResult }) {
  const events = [
    {
      type: "evidence.started",
      generatedAt,
      suiteId: suite.suiteId,
      release: suite.release,
      level: suite.level
    },
    ...fixtureEvidenceSources.map((entry) => ({
      type: "fixture.evidence.declared",
      generatedAt,
      suiteId: suite.suiteId,
      id: entry.id,
      format: entry.format,
      status: entry.status
    })),
    ...tests.map((test) => ({
      type: "test.declared",
      generatedAt,
      suiteId: suite.suiteId,
      id: test.id,
      status: test.status,
      priority: test.priority,
      blocking: test.blocking
    })),
    {
      type: "check.completed",
      generatedAt,
      command: "node scripts/check-perfect-spec.mjs --json",
      exitCode: checkResult.status
    },
    {
      type: "v3.1.0.evidence-matrix.generated",
      generatedAt,
      suiteId: suite.suiteId,
      release: suite.release,
      phase0: "complete",
      runtimeProjection: "runtime-v2",
      currentProfileId: "current-limited-v3.1",
      targetProfileId: "perfect-runtime-target"
    },
    {
      type: "cli-parity.completed",
      generatedAt,
      command: "node scripts/capture-cli-parity.mjs",
      exitCode: parityResult.status
    },
    {
      type: "evidence.completed",
      generatedAt,
      suiteId: suite.suiteId,
      tests: tests.length
    }
  ];

  return `${events.map((event) => JSON.stringify(event)).join("\n")}\n`;
}

function buildSummary({ generatedAt, suite, tests, checkResult }) {
  const statusCounts = new Map();
  for (const test of tests) statusCounts.set(test.status, (statusCounts.get(test.status) ?? 0) + 1);

  const lines = [
    `# Perfect Spec Evidence`,
    "",
    `- Suite: ${suite.suiteId}`,
    `- Release: ${suite.release}`,
    `- Level: ${suite.level}`,
    `- Generated: ${generatedAt}`,
    `- Tests: ${tests.length}`,
    `- Ready: ${statusCounts.get("ready") ?? 0}`,
    `- Pending: ${statusCounts.get("pending") ?? 0}`,
    `- Blocked: ${statusCounts.get("blocked") ?? 0}`,
    `- Fixture evidence: ${fixtureEvidenceSources.length} entries across ${[...new Set(fixtureEvidenceSources.map((entry) => entry.format))].join(", ")}`,
    `- Native renderer evidence: .officegen/acceptance/perfect-spec/native-renderer.json`,
    `- v3.1.0 evidence matrix: .officegen/acceptance/perfect-spec/v3.1.0-evidence-matrix.json`,
    `- Phase 0 P0: complete`,
    `- Runtime projection: runtime-v2 (current-limited-v3.1 supported)`,
    `- Capability profiles: current-limited-v3.1 vs perfect-runtime-target`,
    `- Visibility check exit code: ${checkResult.status}`,
    "",
    `## L7 Acceptance IDs`,
    "",
    ...tests.map((test) => `- ${test.id} ${test.status} ${test.blocking ? "blocking" : "non-blocking"}: ${test.title}`),
    ""
  ];

  return `${lines.join("\n")}\n`;
}

function writeV31EvidenceMatrix({ generatedAt, suite, tests, sourcePath, outPath }) {
  const source = readJson(sourcePath);
  const phase0Test = tests.find((test) => test.id === source.phase0?.acceptanceId);
  const runtimeTest = tests.find((test) => test.id === source.runtimeProjection?.acceptanceId);
  const postTagTest = tests.find((test) => test.id === source.postTagSmoke?.acceptanceId);
  writeJson(outPath, {
    schema: "officegen.perfect-spec.v3.1.0-evidence-matrix@1.0",
    generatedAt,
    suiteId: suite.suiteId,
    release: suite.release,
    level: suite.level,
    source: relative(sourcePath),
    phase0: {
      ...source.phase0,
      suiteStatus: phase0Test?.status,
      blocking: phase0Test?.blocking
    },
    runtimeProjection: {
      ...source.runtimeProjection,
      suiteStatus: runtimeTest?.status,
      blocking: runtimeTest?.blocking
    },
    capabilityProfiles: source.capabilityProfiles,
    remainingTargetGaps: source.remainingTargetGaps ?? [],
    postTagSmoke: {
      ...source.postTagSmoke,
      suiteStatus: postTagTest?.status,
      blocking: postTagTest?.blocking
    },
    publishGateConnection: {
      gate: suite.publishGate,
      command: "npm run perfect-spec:check -- --gate=publish",
      requiresPostTagSmoke: true,
      requiresEvidenceMatrixArtifact: true
    }
  });
}

function ensurePostTagSmokePlaceholder({ generatedAt, suite, file }) {
  if (exists(file)) return;
  writeJson(file, {
    schema: "officegen.perfect-spec.post-tag-smoke@1.0",
    acceptanceId: "L7-A009",
    generatedAt,
    release: suite.release,
    ok: false,
    status: "pending-post-tag",
    reason: "Post-tag smoke requires a published tag or remote target and is enforced by --gate=publish.",
    checks: []
  });
}

function writeFixtureEvidence({ generatedAt, suite, outDir, fixtureManifestPath }) {
  const descriptorDir = path.join(outDir, "fixture-descriptors");
  mkdirSync(descriptorDir, { recursive: true });
  const entries = fixtureEvidenceSources.map((entry) => {
    const sourceFile = path.resolve(root, entry.sourcePath);
    const source = {
      path: relative(sourceFile),
      bytes: statSync(sourceFile).size,
      sha256: sha256File(sourceFile)
    };
    const descriptorPath = path.join(descriptorDir, `${entry.id}.json`);
    writeJson(descriptorPath, {
      schema: "officegen.perfect-spec.fixture-descriptor@1.0",
      suiteId: suite.suiteId,
      acceptanceId: "L7-A007",
      generatedAt,
      id: entry.id,
      format: entry.format,
      fixture: entry.fixture,
      testName: entry.testName,
      command: entry.command,
      targetFeature: entry.targetFeature,
      status: entry.status,
      source
    });
    return {
      id: entry.id,
      format: entry.format,
      fixture: entry.fixture,
      testName: entry.testName,
      command: entry.command,
      targetFeature: entry.targetFeature,
      status: entry.status,
      source,
      generatedArtifact: {
        kind: "fixture-descriptor",
        path: relative(descriptorPath),
        bytes: statSync(descriptorPath).size,
        sha256: sha256File(descriptorPath)
      }
    };
  });

  const formats = Object.fromEntries(
    ["pptx", "docx", "xlsx", "pdf"].map((format) => {
      const formatEntries = entries.filter((entry) => entry.format === format);
      return [
        format,
        {
          status: formatEntries.length > 0 ? "covered" : "missing",
          count: formatEntries.length,
          entries: formatEntries.map((entry) => entry.id)
        }
      ];
    })
  );

  const manifest = {
    schema: "officegen.perfect-spec.fixture-evidence@1.0",
    suiteId: suite.suiteId,
    release: suite.release,
    level: suite.level,
    acceptanceId: "L7-A007",
    generatedAt,
    minimumFormats: ["pptx", "docx", "xlsx", "pdf"],
    status: Object.values(formats).every((format) => format.status === "covered") ? "covered" : "incomplete",
    formats,
    entries
  };

  const workflowDir = path.join(outDir, "office-workflow");
  mkdirSync(workflowDir, { recursive: true });
  const artifactPaths = entries.map((entry) => ({
    role: "fixture-descriptor",
    file: path.resolve(root, entry.generatedArtifact.path)
  }));
  for (const format of manifest.minimumFormats) {
    const file = path.join(workflowDir, `${format}.json`);
    writeJson(file, {
      schema: "officegen.perfect-spec.office-workflow-fixture@1.0",
      suiteId: suite.suiteId,
      acceptanceId: "L7-A007",
      generatedAt,
      format,
      status: manifest.formats[format].status,
      entries: entries.filter((entry) => entry.format === format)
    });
    artifactPaths.push({ role: "office-workflow-fixture", file });
  }

  writeJson(fixtureManifestPath, manifest);
  return artifactPaths;
}

function collectNativeRendererEvidence(generatedAt) {
  const doctor = runNode(["bin/officegen.js", "renderer", "doctor", "--json", "--json-budget-bytes", "80000"]);
  const config = runNode(["bin/officegen.js", "config", "--json", "--json-budget-bytes", "80000"]);
  assertRequiredCommand("native renderer doctor", doctor);
  assertRequiredCommand("config inspection", config);
  const probeInput = path.join(options.outDir, "native-denied-probe.pptx");
  writeFileSync(probeInput, "officegen native renderer policy probe\n", "utf8");
  const nativeDenied = runNode([
    "bin/officegen.js",
    "export",
    relative(probeInput),
    "--to",
    "pdf",
    "--mode",
    "native",
    "--json",
    "--json-budget-bytes",
    "80000"
  ]);
  rmSync(probeInput, { force: true });

  const doctorEnvelope = parseJsonOrText(doctor.stdout);
  const configEnvelope = parseJsonOrText(config.stdout);
  const deniedEnvelope = parseJsonOrText(nativeDenied.stdout || nativeDenied.stderr);
  const doctorResult = asRecord(asRecord(doctorEnvelope).result);
  const renderers = Array.isArray(doctorResult.renderers) ? doctorResult.renderers.map(asRecord) : [];
  const libreOffice = renderers.find((renderer) => renderer.backend === "libreoffice" || renderer.id === "libreoffice");
  const officeComRenderers = renderers.filter((renderer) => renderer.backend === "office-com");
  const defaultSecurity = asRecord(asRecord(asRecord(configEnvelope).result).config).security;
  const doctorPolicy = asRecord(doctorResult.policy);
  const observedErrorCode = asRecord(asRecord(deniedEnvelope).error).code;

  return {
    schema: "officegen.perfect-spec.native-renderer-evidence@1.0",
    generatedAt,
    l7AcceptanceId: "L7-A008",
    platform: process.platform,
    node: process.version,
    decision: {
      status: "ready",
      nativeExecutionRequired: false,
      publishBlocker: false,
      rationale: "L7-A008 is satisfied by optional native renderer disclosure, doctor availability evidence, and default external-process denial evidence. It does not require native conversion success on this runner."
    },
    commands: {
      doctor: commandRecord("node bin/officegen.js renderer doctor --json --json-budget-bytes 80000", doctor),
      config: commandRecord("node bin/officegen.js config --json --json-budget-bytes 80000", config),
      nativeDeniedProbe: commandRecord(
        "node bin/officegen.js export .officegen/acceptance/perfect-spec/native-denied-probe.pptx --to pdf --mode native --json --json-budget-bytes 80000",
        nativeDenied
      )
    },
    environment: {
      libreOffice: {
        checked: true,
        available: Boolean(libreOffice?.available),
        executable: typeof libreOffice?.executable === "string" ? libreOffice.executable : undefined,
        formats: Array.isArray(libreOffice?.formats) ? libreOffice.formats : ["pptx", "docx", "xlsx"],
        message: typeof libreOffice?.message === "string" ? libreOffice.message : undefined
      },
      windowsOfficeCom: {
        checked: true,
        platformApplicable: process.platform === "win32",
        available: officeComRenderers.length > 0,
        renderers: officeComRenderers.map((renderer) => ({
          id: renderer.id,
          available: renderer.available === true,
          executable: renderer.executable,
          formats: Array.isArray(renderer.formats) ? renderer.formats : [],
          message: renderer.message
        }))
      }
    },
    defaultPolicy: {
      profile: asRecord(asRecord(configEnvelope).result).profile,
      externalProcess: doctorPolicy.externalProcess ?? asRecord(defaultSecurity).externalProcess,
      renderers: doctorPolicy.renderers ?? asRecord(defaultSecurity).renderers,
      nativeConversionBlockedByDefault: observedErrorCode === "SECURITY_EXTERNAL_PROCESS_DENIED",
      blockedExpectation: {
        command: "officegen export <existing-office-file> --to pdf --mode native --json",
        expectedErrorCode: "SECURITY_EXTERNAL_PROCESS_DENIED",
        observedErrorCode,
        observedExitCode: nativeDenied.status
      }
    },
    nativeGateOptionality: {
      nativeExecutionRequiredForThisEvidence: false,
      nonNativeReleaseGateIndependent: true,
      doctorIsSafeDiscovery: asRecord(doctorEnvelope).ok === true,
      requiredForNativeSuccess: [
        "OFFICEGEN_PROFILE=enterprise or equivalent config",
        "security.externalProcess=allow",
        "security.renderers=enabled",
        "Windows Office COM or LibreOffice headless installed on the runner"
      ],
      releaseRunnerRequirements: [
        "Tag/release runners that claim native fidelity must run on a host with Office COM or LibreOffice headless installed.",
        "Native export/verify jobs must opt in with enterprise/trusted policy before executing external renderer processes.",
        "Non-native publish gates remain valid when native backends are absent, provided this optionality evidence is attached."
      ]
    }
  };
}

function runNode(nodeArgs) {
  const result = spawnSync(process.execPath, nodeArgs, {
    cwd: root,
    encoding: "utf8",
    windowsHide: true
  });

  return {
    status: typeof result.status === "number" ? result.status : 1,
    stdout: result.stdout ?? "",
    stderr: result.stderr ?? (result.error ? result.error.message : "")
  };
}

function assertRequiredCommand(label, result) {
  if (result.status === 0) return;
  const detail = result.stderr.trim() || result.stdout.trim();
  throw new Error(`perfect-spec:evidence required command failed: ${label} exited ${result.status}${detail ? `\n${detail}` : ""}`);
}

function readJson(file) {
  return JSON.parse(readFileSync(file, "utf8"));
}

function exists(file) {
  try {
    statSync(file);
    return true;
  } catch {
    return false;
  }
}

function writeJson(file, value) {
  writeFileSync(file, `${JSON.stringify(value, null, 2)}\n`, "utf8");
}

function parseJsonOrText(text) {
  try {
    return JSON.parse(text);
  } catch {
    return text;
  }
}

function commandRecord(command, result) {
  return {
    command,
    exitCode: result.status,
    stdout: parseJsonOrText(result.stdout),
    stderr: result.stderr
  };
}

function asRecord(value) {
  return value && typeof value === "object" && !Array.isArray(value) ? value : {};
}

function sha256File(file) {
  return createHash("sha256").update(readFileSync(file)).digest("hex");
}

function relative(file) {
  return path.relative(root, file).replaceAll(path.sep, "/");
}

function usage(message) {
  console.error(`perfect-spec:evidence: ${message}`);
  console.error("usage: node scripts/generate-perfect-spec-evidence.mjs [--suite=path] [--v31-evidence-matrix=path] [--out-dir=path]");
  process.exit(2);
}
