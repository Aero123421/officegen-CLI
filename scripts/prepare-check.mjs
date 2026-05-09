#!/usr/bin/env node
import { existsSync } from "node:fs";
import { spawnSync } from "node:child_process";

const requiredDist = [
  "packages/core/dist/index.js",
  "packages/formats/dist/index.js",
  "packages/optional/dist/index.js",
  "packages/cli/dist/main.js"
];

if (requiredDist.every((file) => existsSync(file))) {
  process.exit(0);
}

if (!existsSync("node_modules/typescript/bin/tsc")) {
  console.error([
    "officegen prepare: dist artifacts are missing and TypeScript is not installed.",
    "This checkout cannot be installed directly from GitHub without built dist files.",
    "Run `npm install && npm run build`, or install a release tarball that contains dist artifacts."
  ].join("\n"));
  process.exit(1);
}

const build = spawnSync(process.execPath, [
  "node_modules/typescript/bin/tsc",
  "-b",
  "packages/core",
  "packages/formats",
  "packages/optional",
  "packages/cli"
], { stdio: "inherit", shell: false });
if (build.status !== 0) process.exit(build.status ?? 1);

const patch = spawnSync(process.execPath, ["scripts/patch-dist-imports.mjs"], { stdio: "inherit", shell: false });
process.exit(patch.status ?? 0);
