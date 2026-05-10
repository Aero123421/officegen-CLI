#!/usr/bin/env node
import { existsSync, realpathSync } from "node:fs";
import { mkdir, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { spawnSync } from "node:child_process";
import os from "node:os";
import path from "node:path";

const temp = await mkdtemp(path.join(os.tmpdir(), "officegen-release-tarball-"));
const cwd = process.cwd();
const rootPackage = JSON.parse(await readFile(path.join(cwd, "package.json"), "utf8"));
let spec = process.env.OFFICEGEN_RELEASE_TARBALL_SPEC
  ?? process.argv.find((arg) => arg.endsWith(".tgz") || arg.startsWith("http"))
  ?? path.join(cwd, `officegen-v${rootPackage.version}.tgz`);
let generatedTarball;

try {
  if (!spec.startsWith("http") && !existsSync(spec)) {
    const npmCli = process.env.npm_execpath;
    const pack = npmCli
      ? spawnSync(process.execPath, [npmCli, "pack", "--silent"], { encoding: "utf8", shell: false })
      : spawnSync("npm", ["pack", "--silent"], { encoding: "utf8", shell: process.platform === "win32" });
    if (pack.status !== 0) {
      console.error(pack.stdout);
      console.error(pack.stderr);
      process.exit(pack.status ?? 1);
    }
    generatedTarball = pack.stdout.trim().split(/\r?\n/).at(-1);
    spec = path.resolve(generatedTarball);
  }
  const npmCli = process.env.npm_execpath;
  const installArgs = ["install", "-g", spec, "--prefix", temp, "--ignore-scripts", "--no-audit", "--no-fund"];
  const install = npmCli
    ? spawnSync(process.execPath, [npmCli, ...installArgs], { stdio: "inherit", shell: false })
    : spawnSync("npm", installArgs, { stdio: "inherit", shell: process.platform === "win32" });
  if (install.status !== 0) process.exit(install.status ?? 1);

  const packageRoot = process.platform === "win32"
    ? path.join(temp, "node_modules", "officegen")
    : path.join(temp, "lib", "node_modules", "officegen");
  const cliMain = path.join(packageRoot, "packages", "cli", "dist", "main.js");
  if (!existsSync(cliMain)) {
    console.error("Installed release tarball is missing CLI runtime.");
    console.error(JSON.stringify({ packageRoot, realPackageRoot: existsSync(packageRoot) ? realpathSync(packageRoot) : undefined, cliMain }, null, 2));
    process.exit(1);
  }

  const bin = process.platform === "win32" ? path.join(temp, "officegen.cmd") : path.join(temp, "bin", "officegen");
  for (const args of [["--version"], ["--help"], ["capabilities", "--agent", "--json"], ["schema", "list", "--agent", "--json"]]) {
    const result = process.platform === "win32"
      ? spawnSync("cmd.exe", ["/d", "/c", bin, ...args], { encoding: "utf8" })
      : spawnSync(bin, args, { encoding: "utf8", shell: false });
    if (result.status !== 0) {
      console.error(result.stdout);
      console.error(result.stderr);
      process.exit(result.status ?? 1);
    }
    if (args[0] === "--version" && result.stdout.trim() !== rootPackage.version) {
      throw new Error(`officegen --version returned ${result.stdout.trim()}, expected ${rootPackage.version}.`);
    }
  }
  const smokeDir = path.join(temp, "smoke-work");
  await mkdir(smokeDir, { recursive: true });
  await writeFile(path.join(smokeDir, "deck.ir.json"), `${JSON.stringify({
    schema: "officegen.ir.document@1.2",
    title: "Release tarball smoke",
    targets: ["pptx"],
    sections: [{ title: "Release tarball smoke", blocks: [{ type: "table", rows: [{ metric: "ok", value: "true" }] }] }]
  })}\n`, "utf8");
  const render = runInstalled(bin, ["render", "deck.ir.json", "--target", "pptx", "--out", "deck.pptx", "--agent", "--json"], smokeDir);
  if (!render.ok || render.result?.target !== "pptx") {
    throw new Error("Installed release tarball could not render a PPTX smoke artifact.");
  }
  const inspected = runInstalled(bin, ["inspect", "deck.pptx", "--depth", "summary", "--agent", "--json"], smokeDir);
  if (!inspected.ok || inspected.result?.trusted?.summary?.slides !== 1) {
    throw new Error("Installed release tarball could not inspect the PPTX smoke artifact.");
  }
  console.log(`officegen release tarball smoke passed for ${spec}`);
} finally {
  await rm(temp, { recursive: true, force: true });
  if (generatedTarball) await rm(path.resolve(generatedTarball), { force: true });
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
