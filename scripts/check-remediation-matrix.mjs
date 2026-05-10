#!/usr/bin/env node
import { readFileSync } from "node:fs";

const file = "docs/reviews/v2.4.0-remediation-matrix.md";
const allowed = new Set(["fixed", "limited-disclosed", "test-covered", "deferred-v3", "not-applicable"]);
const requiredSections = [
  "Review 1: Agent Contract / False Success",
  "Review 2: Security / Path / Report / Native",
  "Review 3: OOXML / Office Quality",
  "Review 4: Architecture / Release / Traceability"
];

const text = readFileSync(file, "utf8");
const failures = [];
const ids = new Set();
let matrixRows = 0;

for (const section of requiredSections) {
  if (!text.includes(section)) failures.push(`missing mandatory review section: ${section}`);
}

for (const line of text.split(/\r?\n/)) {
  if (!line.startsWith("| V24-")) continue;
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

if (matrixRows < 40) failures.push(`expected at least 40 v2.4 remediation rows, found ${matrixRows}`);
for (const status of allowed) {
  if (!text.includes(`| ${status} |`)) failures.push(`matrix is missing status coverage: ${status}`);
}

if (failures.length) {
  console.error("remediation:check failed:");
  for (const failure of failures) console.error(`- ${failure}`);
  process.exit(1);
}

console.log(`remediation:check ok (${matrixRows} rows, ${ids.size} unique ids)`);
