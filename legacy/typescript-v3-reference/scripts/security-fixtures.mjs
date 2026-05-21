#!/usr/bin/env node
import { spawnSync } from "node:child_process";
import process from "node:process";

const args = [
  "node_modules/vitest/vitest.mjs",
  "run",
  "packages/cli/test/program.test.ts",
  "packages/formats/tests/ooxml-validator.test.ts",
  "packages/formats/tests/formats.test.ts",
  "-t",
  "benchmark|outside|OOXML|risky|native renderer timeout|format-specific"
];

const result = spawnSync(process.execPath, args, {
  cwd: process.cwd(),
  stdio: "inherit"
});

if (result.error) {
  console.error(result.error);
  process.exit(1);
}

process.exit(result.status ?? 1);
