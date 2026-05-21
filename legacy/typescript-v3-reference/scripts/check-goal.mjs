#!/usr/bin/env node
import { readFileSync, readdirSync, statSync } from "node:fs";
import path from "node:path";
import process from "node:process";

const root = process.cwd();
const goalDir = path.join(root, "goal");
const contractPath = path.join(goalDir, "office-operation-os-v2.6.0.goal.json");
const suitePath = path.join(goalDir, "acceptance-suite.v2.6.0.json");
const contractSchemaPath = path.join(goalDir, "schemas", "goal-contract.schema.json");
const suiteSchemaPath = path.join(goalDir, "schemas", "acceptance-suite.schema.json");
const failures = [];

const contract = readJson(contractPath);
const suite = readJson(suitePath);
const contractSchema = readJson(contractSchemaPath);
const suiteSchema = readJson(suiteSchemaPath);

checkNoDraftMarkers(goalDir);
checkSchemaIds();
checkRelease();
checkContractShape();
checkSuiteShape();
checkCoverage();

if (failures.length) {
  console.error("goal:check failed:");
  for (const failure of failures) console.error(`- ${failure}`);
  process.exit(1);
}

const requirementCount = contract.domains.reduce((sum, domain) => sum + domain.requirements.length, 0);
console.log(`goal:check ok (${contract.domains.length} domains, ${requirementCount} requirements, ${suite.tests.length} acceptance tests)`);

function readJson(file) {
  try {
    return JSON.parse(readFileSync(file, "utf8"));
  } catch (error) {
    failures.push(`${relative(file)} is not parseable JSON: ${error.message}`);
    return {};
  }
}

function checkNoDraftMarkers(dir) {
  for (const file of filesUnder(dir)) {
    const text = readFileSync(file, "utf8");
    const marker = text.match(/\b(TODO|TBD|unclassified|未分類)\b/i);
    if (marker) failures.push(`${relative(file)} contains draft marker "${marker[0]}"`);
  }
}

function checkSchemaIds() {
  if (contract.schema !== "officegen.goal-contract@1.0") failures.push("goal contract schema marker is invalid");
  if (suite.schema !== "officegen.acceptance-suite@1.0") failures.push("acceptance suite schema marker is invalid");
  if (contractSchema.$id !== "https://officegen.dev/schemas/goal-contract.schema.json") failures.push("goal contract schema $id drifted");
  if (suiteSchema.$id !== "https://officegen.dev/schemas/acceptance-suite.schema.json") failures.push("acceptance suite schema $id drifted");
}

function checkRelease() {
  if (contract.release?.version !== "2.6.0") failures.push("goal contract release.version must be 2.6.0");
  if (contract.release?.targetTag !== "v2.6.0") failures.push("goal contract release.targetTag must be v2.6.0");
  if (suite.release !== contract.release?.version) failures.push("acceptance suite release must match goal contract release.version");
  if (suite.contract !== path.basename(contractPath)) failures.push("acceptance suite contract filename does not match goal contract file");
}

function checkContractShape() {
  if (!Array.isArray(contract.definitionOfDone) || contract.definitionOfDone.length < 5) failures.push("goal contract needs at least 5 definitionOfDone entries");
  if (!Array.isArray(contract.principles) || contract.principles.length < 5) failures.push("goal contract needs at least 5 principles");
  if (!Array.isArray(contract.domains) || contract.domains.length < 8) failures.push("goal contract needs at least 8 domains");
  if (!Array.isArray(contract.requiredEvidence) || contract.requiredEvidence.length < 5) failures.push("goal contract needs at least 5 requiredEvidence entries");

  const domainIds = new Set();
  const requirementIds = new Set();
  for (const domain of contract.domains ?? []) {
    if (domainIds.has(domain.id)) failures.push(`duplicate domain id: ${domain.id}`);
    domainIds.add(domain.id);
    if (!Array.isArray(domain.requirements) || domain.requirements.length < 2) failures.push(`${domain.id} needs at least 2 requirements`);
    for (const requirement of domain.requirements ?? []) {
      if (requirementIds.has(requirement.id)) failures.push(`duplicate requirement id: ${requirement.id}`);
      requirementIds.add(requirement.id);
      if (!["MUST", "SHOULD", "MAY"].includes(requirement.level)) failures.push(`${requirement.id} has invalid level`);
      if (!Array.isArray(requirement.evidence) || requirement.evidence.length === 0) failures.push(`${requirement.id} is missing evidence`);
      if (!Array.isArray(requirement.ownerSurface) || requirement.ownerSurface.length === 0) failures.push(`${requirement.id} is missing ownerSurface`);
      if (!Array.isArray(requirement.acceptanceRefs) || requirement.acceptanceRefs.length === 0) failures.push(`${requirement.id} is missing acceptanceRefs`);
    }
  }
}

function checkSuiteShape() {
  if (!Array.isArray(suite.tests) || suite.tests.length < 12) failures.push("acceptance suite needs at least 12 tests");
  const ids = new Set();
  for (const test of suite.tests ?? []) {
    if (ids.has(test.id)) failures.push(`duplicate acceptance test id: ${test.id}`);
    ids.add(test.id);
    if (!Array.isArray(test.requirementIds) || test.requirementIds.length === 0) failures.push(`${test.id} is missing requirementIds`);
    if (!Array.isArray(test.steps) || test.steps.length === 0) failures.push(`${test.id} is missing steps`);
    if (!Array.isArray(test.passCriteria) || test.passCriteria.length === 0) failures.push(`${test.id} is missing passCriteria`);
    if (!Array.isArray(test.evidence) || test.evidence.length === 0) failures.push(`${test.id} is missing evidence`);
    if (typeof test.blocking !== "boolean") failures.push(`${test.id} must declare blocking boolean`);
  }
}

function checkCoverage() {
  const requirements = new Map();
  for (const domain of contract.domains ?? []) {
    for (const requirement of domain.requirements ?? []) requirements.set(requirement.id, requirement);
  }

  const tests = new Map((suite.tests ?? []).map((test) => [test.id, test]));
  const coveredRequirements = new Set();
  const blockingCoveredMust = new Set();

  for (const test of suite.tests ?? []) {
    for (const requirementId of test.requirementIds ?? []) {
      if (!requirements.has(requirementId)) failures.push(`${test.id} references unknown requirement ${requirementId}`);
      coveredRequirements.add(requirementId);
      if (test.blocking) blockingCoveredMust.add(requirementId);
    }
  }

  for (const requirement of requirements.values()) {
    for (const ref of requirement.acceptanceRefs ?? []) {
      if (!tests.has(ref)) failures.push(`${requirement.id} references unknown acceptance test ${ref}`);
    }
    if (!coveredRequirements.has(requirement.id)) failures.push(`${requirement.id} is not covered by any acceptance test`);
    if (requirement.level === "MUST" && !blockingCoveredMust.has(requirement.id)) {
      failures.push(`${requirement.id} is MUST but lacks blocking acceptance coverage`);
    }
  }
}

function filesUnder(dir) {
  const out = [];
  for (const entry of readdirSync(dir)) {
    const full = path.join(dir, entry);
    const stat = statSync(full);
    if (stat.isDirectory()) out.push(...filesUnder(full));
    else if (stat.isFile()) out.push(full);
  }
  return out;
}

function relative(file) {
  return path.relative(root, file).replaceAll(path.sep, "/");
}
