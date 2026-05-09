#!/usr/bin/env node
import { mkdtemp, rm } from "node:fs/promises";
import { spawnSync } from "node:child_process";
import os from "node:os";
import path from "node:path";
import { pathToFileURL } from "node:url";

const temp = await mkdtemp(path.join(os.tmpdir(), "officegen-github-install-"));
const cwd = process.cwd();
const defaultSpec = process.env.OFFICEGEN_GITHUB_INSTALL_SPEC ?? pathToFileURL(cwd).href;
try {
  const npmCli = process.env.npm_execpath;
  const install = npmCli
    ? spawnSync(process.execPath, [npmCli, "install", "-g", defaultSpec, "--prefix", temp], { stdio: "inherit", shell: false })
    : spawnSync("npm", ["install", "-g", defaultSpec, "--prefix", temp], {
    stdio: "inherit",
    shell: process.platform === "win32"
      });
  if (install.status !== 0) process.exit(install.status ?? 1);

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
