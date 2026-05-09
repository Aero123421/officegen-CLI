#!/usr/bin/env node
import { existsSync, realpathSync } from "node:fs";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import { spawnSync } from "node:child_process";
import os from "node:os";
import path from "node:path";
import { pathToFileURL } from "node:url";

const temp = await mkdtemp(path.join(os.tmpdir(), "officegen-github-install-"));
const cwd = process.cwd();
const remote = process.argv.includes("--remote");
const rootPackage = JSON.parse(await readFile(path.join(cwd, "package.json"), "utf8"));
const defaultSpec = process.env.OFFICEGEN_GITHUB_INSTALL_SPEC ?? (remote
  ? `github:Aero123421/officegen-CLI#v${rootPackage.version}`
  : pathToFileURL(cwd).href);
try {
  const npmCli = process.env.npm_execpath;
  const install = npmCli
    ? spawnSync(process.execPath, [npmCli, "install", "-g", defaultSpec, "--prefix", temp], { stdio: "inherit", shell: false })
    : spawnSync("npm", ["install", "-g", defaultSpec, "--prefix", temp], {
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

  const bin = process.platform === "win32"
    ? path.join(temp, "officegen.cmd")
    : path.join(temp, "bin", "officegen");
  for (const args of [["--version"], ["capabilities", "--agent", "--json", "--json-budget-bytes", "80000"]]) {
    const result = process.platform === "win32"
      ? spawnSync("cmd.exe", ["/d", "/c", bin, ...args], { encoding: "utf8" })
      : spawnSync(bin, args, { encoding: "utf8", shell: false });
    if (result.status !== 0) {
      console.error(result.stdout);
      console.error(result.stderr);
      process.exit(result.status ?? 1);
    }
    if (args[0] === "capabilities" && !result.stdout.includes("capabilitiesHash")) {
      console.error("capabilities smoke did not emit capabilitiesHash.");
      process.exit(1);
    }
  }
  console.log(`officegen github-install smoke passed for ${defaultSpec}`);
} finally {
  await rm(temp, { recursive: true, force: true });
}
