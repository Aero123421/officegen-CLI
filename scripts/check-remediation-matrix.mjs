#!/usr/bin/env node
import { readFileSync } from "node:fs";

const version = JSON.parse(readFileSync("package.json", "utf8")).version;
const majorMinor = version.split(".").slice(0, 2).join(".");
const idPrefix = `V${majorMinor.replace(".", "")}-`;
const file = `docs/reviews/v${version}-remediation-matrix.md`;
const allowed = new Set(["fixed", "limited-disclosed", "test-covered", "deferred-v3", "not-applicable"]);
const requiredSections = [
  "Review 1: Agent Contract / False Success",
  "Review 2: Install / Release / Packaging",
  "Review 3: Asset / Office Workflow",
  "Review 4: UX / Critique / Traceability"
];

const text = readFileSync(file, "utf8");
const failures = [];
const ids = new Set();
let matrixRows = 0;

for (const section of requiredSections) {
  if (!text.includes(section)) failures.push(`missing mandatory review section: ${section}`);
}

for (const line of text.split(/\r?\n/)) {
  if (!line.startsWith(`| ${idPrefix}`)) continue;
  matrixRows += 1;
  const cells = line.split("|").map((cell) => cell.trim()).filter(Boolean);
  const id = cells[0];
  const status = cells[3];
  if (ids.has(id)) failures.push(`duplicate remediation id: ${id}`);
  ids.add(id);
  if (!allowed.has(status)) failures.push(`${id} has invalid status "${status}"`);
  if (/\bunclassified\b|TODO|TBD|未分類/i.test(line)) failures.push(`${id} is unclassified/TODO`);
  if (!cells[4] || cells[4].length < 12) failures.push(`${id} is missing evidence`);
}

if (matrixRows < 20) failures.push(`expected at least 20 v${version} remediation rows, found ${matrixRows}`);
for (const status of allowed) {
  if (!text.includes(`| ${status} |`)) failures.push(`matrix is missing status coverage: ${status}`);
}

if (failures.length) {
  console.error("remediation:check failed:");
  for (const failure of failures) console.error(`- ${failure}`);
  process.exit(1);
}

console.log(`remediation:check ok (${matrixRows} rows, ${ids.size} unique ids)`);
