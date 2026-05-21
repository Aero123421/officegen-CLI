#!/usr/bin/env node
import { existsSync, readFileSync } from "node:fs";
import { spawnSync } from "node:child_process";
import path from "node:path";

const files = ["install/install.sh", "install/install.ps1"];
for (const file of files) {
  if (!existsSync(file)) {
    console.error(`missing installer: ${file}`);
    process.exit(1);
  }
}

const sh = readFileSync("install/install.sh", "utf8");
for (const required of ["OFFICEGEN_VERSION", "OFFICEGEN_INSTALL_DIR", "unknown-linux-gnu", "aarch64", "x86_64", ".sha256", "releases/download", "PATH collision", "officegen --version"]) {
  if (!sh.includes(required)) {
    console.error(`install/install.sh does not include ${required}`);
    process.exit(1);
  }
}
runOptional("bash", ["-n", process.platform === "win32" ? "install/install.sh" : path.resolve("install/install.sh")]);

const ps1 = readFileSync("install/install.ps1", "utf8");
for (const required of ["OFFICEGEN_VERSION", "OFFICEGEN_INSTALL_DIR", "x86_64-pc-windows-msvc", "aarch64-pc-windows-msvc", "Get-FileHash", "Expand-Archive", "releases/download", "Set-OfficegenPathPrecedence", "Test-OfficegenCommandResolution", "stale shim"]) {
  if (!ps1.includes(required)) {
    console.error(`install/install.ps1 does not include ${required}`);
    process.exit(1);
  }
}
runOptional("powershell.exe", ["-NoProfile", "-Command", `$tokens=$null; $errors=$null; [System.Management.Automation.Language.Parser]::ParseFile('${path.resolve("install/install.ps1").replace(/'/g, "''")}', [ref]$tokens, [ref]$errors) > $null; if ($errors.Count) { $errors | ForEach-Object { Write-Error $_ }; exit 1 }`]);

console.log("officegen installer smoke passed");

function runOptional(command, args) {
  const result = spawnSync(command, args, { stdio: "inherit", shell: false });
  if (result.error?.code === "ENOENT") return;
  if (result.status !== 0) process.exit(result.status ?? 1);
}
