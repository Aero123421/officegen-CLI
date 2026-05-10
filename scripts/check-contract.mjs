import { readFile } from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import ts from "typescript";

const root = process.cwd();
const failures = [];

const capabilitiesPath = path.join(root, "packages/core/src/capabilities.ts");
const metadataPath = path.join(root, "packages/cli/src/shared/metadata.ts");
const registerPath = path.join(root, "packages/cli/src/commands/register.ts");
const contextPath = path.join(root, "packages/cli/src/shared/context.ts");
const envelopePath = path.join(root, "packages/cli/src/shared/envelope.ts");
const payloadsPath = path.join(root, "packages/cli/src/commands/payloads.ts");

const [capabilitiesSource, metadataSource, registerSource, contextSource, envelopeSource, payloadsSource] = await Promise.all([
  readFile(capabilitiesPath, "utf8"),
  readFile(metadataPath, "utf8"),
  readFile(registerPath, "utf8"),
  readFile(contextPath, "utf8"),
  readFile(envelopePath, "utf8"),
  readFile(payloadsPath, "utf8")
]);

const commandSpecs = readCommandSpecs(capabilitiesSource, capabilitiesPath);
const metadata = readCommandMetadata(metadataSource, metadataPath);
const leafPayloads = readObjectKeys(registerSource, registerPath, "leafPayloads");
const groupPayloads = readObjectKeys(registerSource, registerPath, "groupPayloads");
const specialRegistrars = new Set(["config", "schema", "errors"]);

compareFeatureCommands(commandSpecs, metadata);
checkCommanderCoverage(metadata, leafPayloads, groupPayloads, specialRegistrars);
checkRendererDoctor(commandSpecs, metadata, payloadsSource, contextSource);
checkSchemaFetchAlias(commandSpecs, metadata, registerSource, contextSource, envelopeSource);
checkWiredExposure(metadata, leafPayloads, groupPayloads, specialRegistrars, payloadsSource, registerSource);

if (failures.length) {
  console.error("contract:check failed:");
  for (const failure of failures) console.error(`- ${failure}`);
  process.exit(1);
}

console.log(`contract:check ok (${metadata.size} command groups, ${[...metadata.values()].flat().length} commands)`);

function compareFeatureCommands(core, cli) {
  const coreFeatures = [...core.keys()].sort();
  const cliFeatures = [...cli.keys()].sort();
  compareArrays("core COMMAND_SPECS features", coreFeatures, cliFeatures);

  for (const feature of new Set([...coreFeatures, ...cliFeatures])) {
    compareArrays(`commands for ${feature}`, core.get(feature) ?? [], cli.get(feature) ?? []);
  }
}

function checkCommanderCoverage(cli, leaf, group, special) {
  for (const feature of cli.keys()) {
    if (special.has(feature) || leaf.has(feature) || group.has(feature)) continue;
    failures.push(`COMMAND_METADATA feature "${feature}" is not covered by leafPayloads, groupPayloads, or a special registrar.`);
  }
}

function checkRendererDoctor(core, cli, payloads, context) {
  for (const [label, source] of [["COMMAND_SPECS", core], ["COMMAND_METADATA", cli]]) {
    if (!source.get("renderer")?.includes("renderer doctor")) {
      failures.push(`${label} is missing renderer doctor.`);
    }
  }
  if (!/subcommand\s*===\s*"doctor"[\s\S]*nativeRendererDoctor/.test(payloads)) {
    failures.push("rendererPayload does not route renderer doctor to nativeRendererDoctor.");
  }
  if (!/entry\.feature\s*===\s*"renderer"[\s\S]*secondCommandToken\(context\.argv\)\s*===\s*"doctor"/.test(context)) {
    failures.push("gateTopLevelCommand no longer permits renderer doctor as disabled-safe discovery.");
  }
}

function checkSchemaFetchAlias(core, cli, register, context, envelope) {
  for (const [label, source] of [["COMMAND_SPECS", core], ["COMMAND_METADATA", cli]]) {
    if (!source.get("schema")?.includes("schema fetch")) {
      failures.push(`${label} is missing schema fetch alias.`);
    }
  }
  if (!/baseCommand\("fetch"[\s\S]*schemaGetPayload/.test(register)) {
    failures.push("schema fetch is not registered as an alias for schema get.");
  }
  if (!/entry\.commandGroup\s*===\s*"schema"[\s\S]*second\s*===\s*"fetch"[\s\S]*schema get/.test(context)) {
    failures.push("schema fetch does not have a schema get gate suggestion.");
  }
  if (!/command\.startsWith\("schema fetch"\)[\s\S]*schema get/.test(envelope)) {
    failures.push("schema fetch does not have a schema get repair suggestion.");
  }
}

function checkWiredExposure(cli, leaf, group, special, payloads, register) {
  const uncovered = [...cli.keys()].filter((feature) => !special.has(feature) && !leaf.has(feature) && !group.has(feature));
  if (uncovered.length) {
    failures.push(`metadata features would fall back to wired groupPayload: ${uncovered.join(", ")}`);
  }

  const wiredInRegister = locationsFor(register, /status\s*:\s*"wired"/g);
  for (const line of wiredInRegister) failures.push(`registered command emits wired placeholder at packages/cli/src/commands/register.ts:${line}`);

  for (const line of locationsFor(payloads, /status\s*:\s*"wired"/g)) {
    const owner = nearestFunctionName(payloads, line);
    if (owner !== "wiredPayload" && owner !== "groupPayload") {
      failures.push(`agent-visible payload may emit wired placeholder in ${owner ?? "unknown function"} at packages/cli/src/commands/payloads.ts:${line}`);
    }
  }

  if (/wiredPayload\(/.test(register.replace(/wiredPayload\s*,?/, ""))) {
    failures.push("register.ts directly wires a command through wiredPayload.");
  }
}

function readCommandSpecs(source, fileName) {
  const sf = ts.createSourceFile(fileName, source, ts.ScriptTarget.Latest, true, ts.ScriptKind.TS);
  const out = new Map();
  visit(sf, (node) => {
    if (!ts.isVariableDeclaration(node) || node.name.getText(sf) !== "COMMAND_SPECS") return;
    const array = unwrapArray(node.initializer);
    if (!array) return;
    for (const element of array.elements) {
      if (!ts.isCallExpression(element) || element.expression.getText(sf) !== "spec") continue;
      const feature = stringLiteral(element.arguments[0]);
      const commands = stringArray(element.arguments[1]);
      if (feature) out.set(feature, commands);
    }
  });
  return out;
}

function readCommandMetadata(source, fileName) {
  const sf = ts.createSourceFile(fileName, source, ts.ScriptTarget.Latest, true, ts.ScriptKind.TS);
  const out = new Map();
  visit(sf, (node) => {
    if (!ts.isVariableDeclaration(node) || node.name.getText(sf) !== "COMMAND_METADATA") return;
    const array = unwrapArray(node.initializer);
    if (!array) return;
    for (const element of array.elements) {
      if (!ts.isCallExpression(element)) continue;
      const callee = element.expression.getText(sf);
      if (!["meta", "core", "optional"].includes(callee)) continue;
      const feature = stringLiteral(element.arguments[0]);
      const commands = stringArray(element.arguments[2]);
      if (feature) out.set(feature, commands);
    }
  });
  return out;
}

function readObjectKeys(source, fileName, variableName) {
  const sf = ts.createSourceFile(fileName, source, ts.ScriptTarget.Latest, true, ts.ScriptKind.TS);
  const out = new Set();
  visit(sf, (node) => {
    if (!ts.isVariableDeclaration(node) || node.name.getText(sf) !== variableName) return;
    const object = unwrapObject(node.initializer);
    if (!object) return;
    for (const property of object.properties) {
      if (ts.isPropertyAssignment(property) && ts.isIdentifier(property.name)) out.add(property.name.text);
    }
  });
  return out;
}

function unwrapArray(node) {
  if (!node) return undefined;
  if (ts.isArrayLiteralExpression(node)) return node;
  if (ts.isAsExpression(node) || ts.isSatisfiesExpression(node)) return unwrapArray(node.expression);
  return undefined;
}

function unwrapObject(node) {
  if (!node) return undefined;
  if (ts.isObjectLiteralExpression(node)) return node;
  if (ts.isAsExpression(node) || ts.isSatisfiesExpression(node)) return unwrapObject(node.expression);
  return undefined;
}

function stringArray(node) {
  const array = unwrapArray(node);
  return array ? array.elements.map(stringLiteral).filter(Boolean) : [];
}

function stringLiteral(node) {
  return node && ts.isStringLiteralLike(node) ? node.text : undefined;
}

function visit(node, callback) {
  callback(node);
  ts.forEachChild(node, (child) => visit(child, callback));
}

function compareArrays(label, actual, expected) {
  if (actual.length === expected.length && actual.every((value, index) => value === expected[index])) return;
  failures.push(`${label} drifted. core=[${actual.join(", ")}] cli=[${expected.join(", ")}]`);
}

function locationsFor(source, pattern) {
  const lines = [];
  for (const match of source.matchAll(pattern)) {
    lines.push(source.slice(0, match.index).split(/\r?\n/).length);
  }
  return lines;
}

function nearestFunctionName(source, line) {
  const prefix = source.split(/\r?\n/).slice(0, line).join("\n");
  const matches = [...prefix.matchAll(/function\s+([A-Za-z0-9_]+)\s*\(/g)];
  return matches.at(-1)?.[1];
}
