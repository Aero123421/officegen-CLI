#!/usr/bin/env node
import { readdirSync, readFileSync } from "node:fs";

const files = readdirSync("docs/reviews")
  .filter((file) => /remediation-matrix\.md$/.test(file))
  .map((file) => `docs/reviews/${file}`);
for (const file of files) {
  const text = readFileSync(file, "utf8");
  const bad = text
    .split(/\r?\n/)
    .filter((line) => /^\|/.test(line))
    .filter((line) => /\bunclassified\b|TODO|TBD|未分類/.test(line));
  if (bad.length) {
    console.error(`${file} has unclassified/TODO rows:\n${bad.join("\n")}`);
    process.exit(1);
  }
  for (const status of ["implemented", "limited-but-disclosed", "optional-gated", "benchmark-covered"]) {
    if (!text.includes(status)) {
      console.error(`${file} is missing status: ${status}`);
      process.exit(1);
    }
  }
}
console.log("Remediation matrix check passed.");
