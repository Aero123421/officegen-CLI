#!/usr/bin/env node
import { createHash } from "node:crypto";
import { mkdirSync, readFileSync, statSync, writeFileSync } from "node:fs";
import path from "node:path";
import process from "node:process";
import { spawnSync } from "node:child_process";

const root = process.cwd();
const outDir = path.join(root, ".officegen", "acceptance", "perfect-spec");
const npm = process.platform === "win32" ? "npm.cmd" : "npm";
const generatedAt = new Date().toISOString();

mkdirSync(outDir, { recursive: true });

const checks = [
  {
    id: "github-install-tag-smoke",
    command: [npm, "run", "github-install:tag-smoke"],
    logFile: path.join(outDir, "github-install-tag-smoke.txt")
  },
  {
    id: "github-install-remote-smoke",
    command: [npm, "run", "github-install:remote-smoke"],
    logFile: path.join(outDir, "github-install-remote-smoke.txt")
  }
];

const results = checks.map((check) => {
  const [command, ...args] = check.command;
  const result = spawnSync(command, args, {
    cwd: root,
    encoding: "utf8",
    shell: process.platform === "win32",
    windowsHide: true
  });
  const exitCode = typeof result.status === "number" ? result.status : 1;
  const stdout = result.stdout ?? "";
  const stderr = result.stderr ?? (result.error ? result.error.message : "");
  writeFileSync(
    check.logFile,
    [
      `# ${check.id}`,
      `generatedAt=${generatedAt}`,
      `command=${[command, ...args].join(" ")}`,
      `exitCode=${exitCode}`,
      "",
      "## stdout",
      stdout.trimEnd(),
      "",
      "## stderr",
      stderr.trimEnd(),
      ""
    ].join("\n"),
    "utf8"
  );

  return {
    id: check.id,
    command: [command, ...args].join(" "),
    exitCode,
    stdoutBytes: Buffer.byteLength(stdout),
    stderrBytes: Buffer.byteLength(stderr),
    log: artifact(check.logFile)
  };
});

const manifestPath = path.join(outDir, "post-tag-smoke.json");
const ok = results.every((result) => result.exitCode === 0);
writeFileSync(
  manifestPath,
  `${JSON.stringify({
    schema: "officegen.perfect-spec.post-tag-smoke@1.0",
    acceptanceId: "L7-A009",
    generatedAt,
    ok,
    checks: results
  }, null, 2)}\n`,
  "utf8"
);

console.log(`perfect-spec:post-tag-smoke wrote ${relative(manifestPath)}`);
for (const result of results) console.log(`perfect-spec:post-tag-smoke wrote ${result.log.path}`);

if (!ok) {
  console.error("perfect-spec:post-tag-smoke failed:");
  for (const result of results.filter((entry) => entry.exitCode !== 0)) {
    console.error(`- ${result.id} exited ${result.exitCode}; see ${result.log.path}`);
  }
  process.exit(1);
}

function artifact(file) {
  return {
    path: relative(file),
    bytes: statSync(file).size,
    sha256: createHash("sha256").update(readFileSync(file)).digest("hex")
  };
}

function relative(file) {
  return path.relative(root, file).replaceAll(path.sep, "/");
}
