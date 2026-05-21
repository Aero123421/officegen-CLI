#!/usr/bin/env node
import { existsSync, realpathSync } from "node:fs";
import { mkdir, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { spawnSync } from "node:child_process";
import os from "node:os";
import path from "node:path";
import { pathToFileURL } from "node:url";

const temp = await mkdtemp(path.join(os.tmpdir(), "officegen-github-install-"));
const cwd = process.cwd();
const remote = process.argv.includes("--remote");
const head = process.argv.includes("--head");
const specArg = valueArg("--spec");
const ref = valueArg("--ref");
const expectedVersionArg = valueArg("--expected-version");
const rootPackage = JSON.parse(await readFile(path.join(cwd, "package.json"), "utf8"));
const repository = process.env.GITHUB_REPOSITORY ?? "Aero123421/officegen-CLI";
const defaultSpec = specArg ?? process.env.OFFICEGEN_GITHUB_INSTALL_SPEC ?? (ref
  ? `github:${repository}#${ref}`
  : remote
  ? `github:${repository}#v${rootPackage.version}`
  : head
    ? `github:${repository}`
  : pathToFileURL(cwd).href);
const expectedVersion = resolveExpectedVersion({
  expectedVersionArg,
  rootVersion: rootPackage.version,
  spec: defaultSpec,
  remote
});
try {
  const npmCli = process.env.npm_execpath;
  const installArgs = ["install", defaultSpec, "--prefix", temp, "--no-audit", "--no-fund", "--force", "--prefer-online"];
  const install = npmCli
    ? spawnSync(process.execPath, [npmCli, ...installArgs], { stdio: "inherit", shell: false })
    : spawnSync("npm", installArgs, {
    stdio: "inherit",
    shell: process.platform === "win32"
      });
  if (install.status !== 0) process.exit(install.status ?? 1);
  const packageRoot = path.join(temp, "node_modules", "officegen");
  const cliMain = path.join(packageRoot, "packages", "cli", "dist", "main.js");
  const coreMain = path.join(packageRoot, "packages", "core", "dist", "index.js");
  if (!existsSync(cliMain) || !existsSync(coreMain)) {
    console.error("Installed officegen package is missing runtime dist artifacts.");
    console.error(JSON.stringify({
      packageRoot,
      realPackageRoot: existsSync(packageRoot) ? realpathSync(packageRoot) : undefined,
      cliMain,
      cliMainExists: existsSync(cliMain),
      coreMain,
      coreMainExists: existsSync(coreMain)
    }, null, 2));
    process.exit(1);
  }
  if (expectedVersion) {
    const installedManifest = JSON.parse(await readFile(path.join(packageRoot, "package.json"), "utf8"));
    if (installedManifest.version !== expectedVersion) {
      console.error(`installed package version is ${installedManifest.version}, expected ${expectedVersion}.`);
      process.exit(1);
    }
  }

  const bin = process.platform === "win32"
    ? path.join(temp, "node_modules", ".bin", "officegen.cmd")
    : path.join(temp, "node_modules", ".bin", "officegen");
  for (const args of [
    ["--version"],
    ["--help"],
    ["capabilities", "--agent", "--json", "--json-budget-bytes", "80000"],
    ["schema", "list", "--agent", "--json", "--json-budget-bytes", "80000"]
  ]) {
    const result = process.platform === "win32"
      ? spawnSync("cmd.exe", ["/d", "/c", bin, ...args], { encoding: "utf8" })
      : spawnSync(bin, args, { encoding: "utf8", shell: false });
    if (result.status !== 0) {
      console.error(result.stdout);
      console.error(result.stderr);
      process.exit(result.status ?? 1);
    }
    if (args[0] === "--version" && expectedVersion && result.stdout.trim() !== expectedVersion) {
      console.error(`version smoke returned ${result.stdout.trim()}, expected ${expectedVersion ?? rootPackage.version}.`);
      process.exit(1);
    }
    if (args[0] === "capabilities") {
      const envelope = JSON.parse(result.stdout);
      if (!envelope.ok || envelope.result?.schema !== "officegen.capabilities@1.2" || !envelope.capabilitiesHash) {
        console.error("capabilities smoke did not emit a valid capabilities envelope.");
        process.exit(1);
      }
    }
    if (args[0] === "schema") {
      const envelope = JSON.parse(result.stdout);
      if (!envelope.ok || envelope.result?.schema !== "officegen.schema.list@1.2") {
        console.error("schema list smoke did not emit a valid schema envelope.");
        process.exit(1);
      }
    }
  }

  const smokeDir = path.join(temp, "smoke-work");
  await mkdir(smokeDir, { recursive: true });
  await writeFile(path.join(smokeDir, "deck.ir.json"), `${JSON.stringify({
    schema: "officegen.ir.document@1.2",
    title: "GitHub install smoke",
    targets: ["pptx"],
    sections: [{ title: "Install smoke", blocks: [{ type: "table", rows: [{ metric: "ok", value: "true" }] }] }]
  })}\n`, "utf8");
  const render = runInstalled(bin, ["render", "deck.ir.json", "--target", "pptx", "--out", "deck.pptx", "--json"], smokeDir);
  if (!render.ok || render.result?.target !== "pptx") {
    console.error("render smoke failed.");
    process.exit(1);
  }
  const inspected = runInstalled(bin, ["inspect", "deck.pptx", "--depth", "summary", "--agent", "--json"], smokeDir);
  if (!inspected.ok || inspected.result?.trusted?.summary?.slides !== 1) {
    console.error("inspect smoke failed.");
    process.exit(1);
  }
  console.log(`officegen github-install smoke passed for ${defaultSpec}`);
} finally {
  await rm(temp, { recursive: true, force: true });
}

function valueArg(name) {
  const index = process.argv.indexOf(name);
  return index >= 0 ? process.argv[index + 1] : undefined;
}

function resolveExpectedVersion({ expectedVersionArg, rootVersion, spec, remote }) {
  if (expectedVersionArg === "current") return rootVersion;
  if (expectedVersionArg) return expectedVersionArg;
  if (spec === pathToFileURL(cwd).href) return rootVersion;

  const tagVersion = versionFromTagSpec(spec);
  if (tagVersion) return tagVersion;
  if (remote) {
    console.error("Remote/tag GitHub install smoke requires an expected version.");
    console.error("Use --expected-version <x.y.z>, --expected-version current, or a git install spec/ref ending in #v<x.y.z>.");
    process.exit(2);
  }
  return undefined;
}

function versionFromTagSpec(spec) {
  return /(?:#|\/tree\/)v(\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?)$/.exec(spec)?.[1];
}

function runInstalled(bin, args, cwd) {
  const result = process.platform === "win32"
    ? spawnSync("cmd.exe", ["/d", "/c", bin, ...args], { encoding: "utf8", cwd })
    : spawnSync(bin, args, { encoding: "utf8", shell: false, cwd });
  if (result.status !== 0) {
    console.error(result.stdout);
    console.error(result.stderr);
    process.exit(result.status ?? 1);
  }
  return JSON.parse(result.stdout);
}
