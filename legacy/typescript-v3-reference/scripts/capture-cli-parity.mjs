#!/usr/bin/env node
import { createHash } from "node:crypto";
import { mkdirSync, readFileSync, statSync, writeFileSync } from "node:fs";
import path from "node:path";
import process from "node:process";
import { spawnSync } from "node:child_process";

const root = process.cwd();
const args = process.argv.slice(2);
const options = {
  outDir: path.join(root, ".officegen", "acceptance", "perfect-spec")
};

for (const arg of args) {
  if (arg.startsWith("--out-dir=")) options.outDir = path.resolve(root, arg.slice("--out-dir=".length));
  else usage(`unknown argument: ${arg}`);
}

const generatedAt = new Date().toISOString();
const runStamp = generatedAt.replace(/[:.]/g, "-");
const parityDir = path.join(options.outDir, "parity");
const workDir = path.join(options.outDir, "parity-work", runStamp);
const cliBin = path.join(root, "bin", "officegen.js");
const packageJson = readJson(path.join(root, "package.json"));

mkdirSync(parityDir, { recursive: true });
mkdirSync(workDir, { recursive: true });

writeJson(path.join(workDir, "deck.ir.json"), {
  schema: "officegen.ir.document@1.2",
  title: "CLI parity smoke",
  targets: ["pptx", "pdf"],
  sections: [
    {
      title: "CLI parity smoke",
      blocks: [{ type: "table", rows: [{ metric: "ok", value: "true" }] }]
    }
  ]
});

const commandSpecs = [
  {
    id: "version-human",
    kind: "human",
    args: ["--version"],
    cwd: root,
    parse: "version"
  },
  {
    id: "capabilities-human",
    kind: "human",
    args: ["capabilities"],
    cwd: root,
    parse: "human-ok"
  },
  {
    id: "capabilities-agent-json",
    kind: "agent-json",
    args: ["capabilities", "--agent", "--json"],
    cwd: root,
    expectResultSchema: "officegen.capabilities@1.2"
  },
  {
    id: "schema-list-human",
    kind: "human",
    args: ["schema", "list"],
    cwd: root,
    parse: "human-ok"
  },
  {
    id: "schema-list-agent-json",
    kind: "agent-json",
    args: ["schema", "list", "--agent", "--json"],
    cwd: root,
    expectResultSchema: "officegen.schema.list@1.2"
  },
  {
    id: "render-pptx-human",
    kind: "human",
    args: ["render", "deck.ir.json", "--target", "pptx", "--out", "deck-human.pptx"],
    cwd: workDir,
    parse: "human-ok"
  },
  {
    id: "render-pptx-agent-json",
    kind: "agent-json",
    args: ["render", "deck.ir.json", "--target", "pptx", "--out", "deck-agent.pptx", "--agent", "--json"],
    cwd: workDir,
    expectResultSchema: "officegen.render.result@1.2"
  },
  {
    id: "inspect-pptx-human",
    kind: "human",
    args: ["inspect", "deck-agent.pptx", "--depth", "summary"],
    cwd: workDir,
    parse: "human-ok"
  },
  {
    id: "inspect-pptx-agent-json",
    kind: "agent-json",
    args: ["inspect", "deck-agent.pptx", "--depth", "summary", "--agent", "--json"],
    cwd: workDir,
    expectResultSchema: "officegen.inspect.result@1.2"
  },
  {
    id: "render-pdf-agent-json",
    kind: "agent-json",
    args: ["render", "deck.ir.json", "--target", "pdf", "--out", "deck-agent.pdf", "--agent", "--json"],
    cwd: workDir,
    expectResultSchema: "officegen.render.result@1.2"
  },
  {
    id: "view-pdf-human",
    kind: "human",
    args: ["view", "deck-agent.pdf", "--format", "png", "--out", "view-human-png"],
    cwd: workDir,
    parse: "human-ok"
  },
  {
    id: "view-pdf-agent-json",
    kind: "agent-json",
    args: ["view", "deck-agent.pdf", "--format", "png", "--out", "view-agent-png", "--agent", "--json"],
    cwd: workDir,
    expectResultSchema: "officegen.view.result@1.2"
  }
];

const commands = commandSpecs.map((spec) => runCommand(spec));
const parityAssertions = buildParityAssertions(commands);
const capabilityTerms = [
  "featureContracts",
  "formatCapabilities",
  "runtimeProfiles",
  "specProfile",
  "current-limited-v3.1",
  "perfect-runtime-target",
  "knownLimitations",
  "unsupportedNow",
  "SmartArt creation and full SmartArt editing are unsupported",
  "Full-fidelity Office/PDF editing"
];
const capabilitiesAgent = commands.find((command) => command.id === "capabilities-agent-json");
const capabilityText = capabilitiesAgent ? readFileSync(path.resolve(root, capabilitiesAgent.stdout.path), "utf8") : "";
const limitationParity = capabilityTerms.map((term) => ({ term, present: capabilityText.includes(term) }));

const artifact = {
  schema: "officegen.cli-parity@1.0",
  generatedAt,
  cli: {
    executable: "node bin/officegen.js",
    packageName: packageJson.name,
    packageVersion: packageJson.version
  },
  workDir: relative(workDir),
  commands,
  parityAssertions,
  limitationParity,
  summary: {
    ok:
      commands.every((command) => command.exitCode === 0 && command.parsed.ok === true) &&
      parityAssertions.every((assertion) => assertion.ok) &&
      limitationParity.every((entry) => entry.present),
    commandCount: commands.length,
    agentJsonCommandCount: commands.filter((command) => command.kind === "agent-json").length,
    humanCommandCount: commands.filter((command) => command.kind === "human").length
  }
};

const artifactPath = path.join(options.outDir, "cli-parity.json");
writeJson(artifactPath, artifact);

console.log(`perfect-spec:cli-parity wrote ${relative(artifactPath)}`);
console.log(`perfect-spec:cli-parity recorded ${commands.length} command transcripts in ${relative(parityDir)}`);

if (!artifact.summary.ok) process.exit(1);

function runCommand(spec) {
  const result = spawnSync(process.execPath, [cliBin, ...spec.args], {
    cwd: spec.cwd,
    encoding: "utf8",
    windowsHide: true,
    maxBuffer: 8 * 1024 * 1024
  });

  const stdout = result.stdout ?? "";
  const stderr = result.stderr ?? (result.error ? result.error.message : "");
  const stdoutPath = path.join(parityDir, `${runStamp}.${spec.id}.stdout.txt`);
  const stderrPath = path.join(parityDir, `${runStamp}.${spec.id}.stderr.txt`);
  writeFileSync(stdoutPath, stdout, "utf8");
  writeFileSync(stderrPath, stderr, "utf8");

  const parsed = parseOutput(spec, stdout, typeof result.status === "number" ? result.status : 1);
  return {
    id: spec.id,
    kind: spec.kind,
    command: commandLabel(spec.args),
    cwd: relative(spec.cwd),
    exitCode: typeof result.status === "number" ? result.status : 1,
    stdout: digestRecord(stdoutPath),
    stderr: digestRecord(stderrPath),
    parsed
  };
}

function parseOutput(spec, stdout, exitCode) {
  if (spec.parse === "version") {
    const version = stdout.trim();
    return {
      format: "text",
      schemaId: null,
      resultSchemaId: null,
      ok: exitCode === 0 && version === packageJson.version && /^\d+\.\d+\.\d+$/.test(version),
      status: version
    };
  }

  if (spec.kind === "agent-json") {
    try {
      const value = JSON.parse(stdout);
      const resultSchemaId = typeof value?.result?.schema === "string" ? value.result.schema : null;
      const ok = exitCode === 0 && value?.schema === "officegen.envelope@1.2" && value?.ok === true && resultSchemaId === spec.expectResultSchema;
      return {
        format: "json",
        schemaId: typeof value?.schema === "string" ? value.schema : null,
        resultSchemaId,
        ok,
        status: value?.readiness ?? null,
        envelopeOk: value?.ok === true,
        executionOk: value?.executionOk === true,
        objectiveOk: value?.objectiveOk === true,
        command: value?.command ?? null
      };
    } catch (error) {
      return {
        format: "json",
        schemaId: null,
        resultSchemaId: null,
        ok: false,
        status: "parse_error",
        error: error.message
      };
    }
  }

  return {
    format: "text",
    schemaId: null,
    resultSchemaId: null,
    ok: exitCode === 0 && (spec.parse === "human-ok" ? stdout.trim().length > 0 : /completed/i.test(stdout)),
    status: exitCode === 0 ? "completed" : "failed"
  };
}

function buildParityAssertions(commandRecords) {
  const pairs = [
    ["capabilities", "capabilities-human", "capabilities-agent-json"],
    ["schema-list", "schema-list-human", "schema-list-agent-json"],
    ["render-pptx", "render-pptx-human", "render-pptx-agent-json"],
    ["inspect-pptx", "inspect-pptx-human", "inspect-pptx-agent-json"],
    ["view-pdf", "view-pdf-human", "view-pdf-agent-json"]
  ];
  return pairs.map(([id, humanId, agentId]) => {
    const human = commandRecords.find((command) => command.id === humanId);
    const agent = commandRecords.find((command) => command.id === agentId);
    const ok = Boolean(human && agent && human.exitCode === 0 && agent.exitCode === 0 && human.parsed.ok && agent.parsed.ok);
    return {
      id,
      humanCommandId: humanId,
      agentCommandId: agentId,
      ok,
      reason: ok ? "Human transcript exit status is not stronger than the agent JSON envelope status." : "Human/agent command pair did not both complete with parsed ok status."
    };
  });
}

function commandLabel(commandArgs) {
  return ["officegen", ...commandArgs].map(quoteCommandValue).join(" ");
}

function quoteCommandValue(value) {
  const text = String(value);
  if (!/[\s"]/u.test(text)) return text;
  return `"${text.replaceAll("\"", "\\\"")}"`;
}

function digestRecord(file) {
  return {
    path: relative(file),
    bytes: statSync(file).size,
    sha256: sha256File(file)
  };
}

function readJson(file) {
  return JSON.parse(readFileSync(file, "utf8"));
}

function writeJson(file, value) {
  writeFileSync(file, `${JSON.stringify(value, null, 2)}\n`, "utf8");
}

function sha256File(file) {
  return createHash("sha256").update(readFileSync(file)).digest("hex");
}

function relative(file) {
  return path.relative(root, file).replaceAll(path.sep, "/");
}

function usage(message) {
  console.error(`perfect-spec:cli-parity: ${message}`);
  console.error("usage: node scripts/capture-cli-parity.mjs [--out-dir=path]");
  process.exit(2);
}
