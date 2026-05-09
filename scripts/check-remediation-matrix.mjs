#!/usr/bin/env node
import { readFileSync } from "node:fs";

const file = "docs/reviews/v2.2.0-remediation-matrix.md";
const text = readFileSync(file, "utf8");
const bad = text
  .split(/\r?\n/)
  .filter((line) => /^\|/.test(line))
  .filter((line) => /\bunclassified\b|TODO|TBD|未分類/.test(line));
if (bad.length) {
  console.error(`Remediation matrix has unclassified/TODO rows:\n${bad.join("\n")}`);
  process.exit(1);
}
for (const status of ["implemented", "limited-but-disclosed", "optional-gated", "benchmark-covered"]) {
  if (!text.includes(status)) {
    console.error(`Remediation matrix is missing status: ${status}`);
    process.exit(1);
  }
}
console.log("Remediation matrix check passed.");
