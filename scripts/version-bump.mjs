import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";

const managedPackages = [
  "package.json",
  "packages/cli/package.json",
  "packages/core/package.json",
  "packages/formats/package.json",
  "packages/optional/package.json"
];
const lockPath = "package-lock.json";
const coreTypesPath = "packages/core/src/types.ts";
const coreDistTypesPath = "packages/core/dist/types.js";
const coreDistDtsPath = "packages/core/dist/types.d.ts";
const readmePath = "README.md";
const cargoManifestPath = "Cargo.toml";
const cargoLockPath = "Cargo.lock";

const args = process.argv.slice(2);
const checkOnly = args.includes("--check");
const bumpArg = args.find((arg) => !arg.startsWith("--"));

const rootPackage = await readJson("package.json");
const currentVersion = rootPackage.version;
const nextVersion = checkOnly ? currentVersion : resolveNextVersion(currentVersion, bumpArg);

const issues = [];
await assertCurrentVersions(currentVersion, issues);
if (checkOnly) {
  if (issues.length > 0) {
    for (const issue of issues) console.error(issue);
    process.exitCode = 1;
  } else {
    console.log(`All managed Officegen versions are synchronized at ${currentVersion}.`);
  }
} else {
  if (!bumpArg) usage();
  await updatePackageManifests(nextVersion);
  await updatePackageLock(nextVersion);
  await updateCargoManifestIfExists(nextVersion);
  await updateCargoLockIfExists(nextVersion);
  await updateCoreVersionConstant(currentVersion, nextVersion);
  await updateReadme(currentVersion, nextVersion);
  console.log(`Updated Officegen version: ${currentVersion} -> ${nextVersion}`);
}

async function assertCurrentVersions(expectedVersion, foundIssues) {
  for (const packagePath of managedPackages) {
    const manifest = await readJson(packagePath);
    if (manifest.version !== expectedVersion) {
      foundIssues.push(`${packagePath}: expected ${expectedVersion}, found ${manifest.version}`);
    }
  }

  const lock = await readJson(lockPath);
  const lockVersionPaths = ["", "packages/cli", "packages/core", "packages/formats", "packages/optional"];
  if (lock.version !== expectedVersion) foundIssues.push(`${lockPath}: root version expected ${expectedVersion}, found ${lock.version}`);
  for (const packagePath of lockVersionPaths) {
    const entry = lock.packages?.[packagePath];
    if (entry?.version !== expectedVersion) {
      foundIssues.push(`${lockPath} packages.${packagePath || "<root>"}: expected ${expectedVersion}, found ${entry?.version}`);
    }
  }

  const coreTypes = await readText(coreTypesPath);
  if (!coreTypes.includes(`OFFICEGEN_CLI_VERSION = "${expectedVersion}"`)) {
    foundIssues.push(`${coreTypesPath}: OFFICEGEN_CLI_VERSION is not ${expectedVersion}`);
  }
  await assertVersionText(coreDistTypesPath, expectedVersion, foundIssues, "OFFICEGEN_CLI_VERSION");
  await assertVersionText(coreDistDtsPath, expectedVersion, foundIssues, "OFFICEGEN_CLI_VERSION");
  await assertCargoManifestVersion(expectedVersion, foundIssues);
  await assertCargoLockVersion(expectedVersion, foundIssues);
  await assertReadmeReleaseRefs(expectedVersion, foundIssues);
}

async function updatePackageManifests(version) {
  for (const packagePath of managedPackages) {
    const manifest = await readJson(packagePath);
    manifest.version = version;
    await writeJson(packagePath, manifest);
  }
}

async function updatePackageLock(version) {
  const lock = await readJson(lockPath);
  lock.version = version;
  for (const packagePath of ["", "packages/cli", "packages/core", "packages/formats", "packages/optional"]) {
    if (lock.packages?.[packagePath]) lock.packages[packagePath].version = version;
  }
  await writeJson(lockPath, lock);
}

async function updateCargoManifestIfExists(version) {
  let text;
  try {
    text = await readText(cargoManifestPath);
  } catch (error) {
    if (error?.code === "ENOENT") return;
    throw error;
  }
  const after = replaceCargoPackageVersion(text, version);
  await writeFile(path.resolve(cargoManifestPath), after, "utf8");
}

async function updateCargoLockIfExists(version) {
  let text;
  try {
    text = await readText(cargoLockPath);
  } catch (error) {
    if (error?.code === "ENOENT") return;
    throw error;
  }
  const after = replaceCargoLockPackageVersion(text, version);
  await writeFile(path.resolve(cargoLockPath), after, "utf8");
}

async function updateCoreVersionConstant(currentVersion, version) {
  await replaceCoreVersionConstant(coreTypesPath, version);
  await replaceCoreVersionConstantIfExists(coreDistTypesPath, version);
  await replaceCoreVersionConstantIfExists(coreDistDtsPath, version);
}

async function updateReadme(currentVersion, version) {
  const before = await readText(readmePath);
  const after = before
    .replaceAll(currentVersion, version)
    .replace(/releases\/download\/v\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?/g, `releases/download/v${version}`)
    .replace(/officegen-v\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?\.tgz/g, `officegen-v${version}.tgz`);
  await writeFile(readmePath, after, "utf8");
}

function resolveNextVersion(currentVersion, bump) {
  if (!bump) usage();
  if (/^\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?$/.test(bump)) return bump;
  const match = /^(\d+)\.(\d+)\.(\d+)$/.exec(currentVersion);
  if (!match) throw new Error(`Current version is not a simple semver version: ${currentVersion}`);
  const major = Number(match[1]);
  const minor = Number(match[2]);
  const patch = Number(match[3]);
  if (bump === "patch") return `${major}.${minor}.${patch + 1}`;
  if (bump === "minor") return `${major}.${minor + 1}.0`;
  if (bump === "major") return `${major + 1}.0.0`;
  throw new Error(`Unknown version bump "${bump}". Use patch, minor, major, or an explicit x.y.z version.`);
}

async function replaceVersionIfExists(filePath, currentVersion, version) {
  try {
    await replaceVersion(filePath, currentVersion, version);
  } catch (error) {
    if (error?.code !== "ENOENT") throw error;
  }
}

async function replaceVersion(filePath, currentVersion, version) {
  const before = await readText(filePath);
  const after = before.replaceAll(currentVersion, version);
  await writeFile(filePath, after, "utf8");
}

async function replaceCoreVersionConstantIfExists(filePath, version) {
  try {
    await replaceCoreVersionConstant(filePath, version);
  } catch (error) {
    if (error?.code !== "ENOENT") throw error;
  }
}

async function replaceCoreVersionConstant(filePath, version) {
  const before = await readText(filePath);
  const after = before.replace(/(OFFICEGEN_CLI_VERSION\s*=\s*")[^"]+(")/, `$1${version}$2`);
  if (after === before && !before.includes(`OFFICEGEN_CLI_VERSION = "${version}"`)) {
    throw new Error(`${filePath}: could not find OFFICEGEN_CLI_VERSION`);
  }
  await writeFile(filePath, after, "utf8");
}

async function readJson(filePath) {
  return JSON.parse(await readText(filePath));
}

async function readText(filePath) {
  return readFile(path.resolve(filePath), "utf8");
}

async function assertVersionText(filePath, expectedVersion, foundIssues, label) {
  try {
    const text = await readText(filePath);
    if (!text.includes(expectedVersion)) foundIssues.push(`${filePath}: ${label} does not include ${expectedVersion}`);
  } catch (error) {
    if (error?.code !== "ENOENT") throw error;
  }
}

async function assertReadmeReleaseRefs(expectedVersion, foundIssues) {
  const patterns = [
    { label: "GitHub release download tag", regex: /github\.com\/[^\s)"']+\/releases\/download\/v(?<version>\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?)/g },
    { label: "GitHub release tag URL", regex: /github\.com\/[^\s)"']+\/releases\/tag\/v(?<version>\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?)/g },
    { label: "GitHub install ref", regex: /github:[^\s)"']+#v(?<version>\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?)/g },
    { label: "release tarball filename", regex: /officegen-v(?<version>\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?)\.tgz/g },
    { label: "README cliVersion example", regex: /"cliVersion":\s*"(?<version>\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?)"/g },
    { label: "benchmark result version filename", regex: /benchmark-results\/v(?<version>\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?)\.json/g },
    { label: "version bump literal", regex: /version:bump -- (?<version>\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?)/g }
  ];

  try {
    const text = await readText(readmePath);
    const findings = [];
    const seen = new Set();
    for (const { label, regex } of patterns) {
      for (const match of text.matchAll(regex)) {
        const version = match.groups?.version;
        if (version && version !== expectedVersion) {
          const location = textLocation(text, match.index ?? 0);
          const issue = `${readmePath}:${location.line}:${location.column}: ${label} uses ${version}, expected ${expectedVersion}`;
          if (!seen.has(issue)) {
            findings.push(issue);
            seen.add(issue);
          }
        }
      }
    }
    if (findings.length > 0) {
      foundIssues.push(...findings);
    } else if (!text.includes(expectedVersion)) {
      foundIssues.push(`${readmePath}: README release references do not include ${expectedVersion}`);
    }
  } catch (error) {
    if (error?.code !== "ENOENT") throw error;
  }
}

async function assertCargoManifestVersion(expectedVersion, foundIssues) {
  let text;
  try {
    text = await readText(cargoManifestPath);
  } catch (error) {
    if (error?.code === "ENOENT") return;
    throw error;
  }
  const version = readCargoPackageVersion(text);
  if (version !== expectedVersion) {
    foundIssues.push(`${cargoManifestPath}: package.version expected ${expectedVersion}, found ${version ?? "<missing>"}`);
  }
}

async function assertCargoLockVersion(expectedVersion, foundIssues) {
  let text;
  try {
    text = await readText(cargoLockPath);
  } catch (error) {
    if (error?.code === "ENOENT") return;
    throw error;
  }
  const version = readCargoLockPackageVersion(text);
  if (version !== expectedVersion) {
    foundIssues.push(`${cargoLockPath}: officegen package.version expected ${expectedVersion}, found ${version ?? "<missing>"}`);
  }
}

function readCargoPackageVersion(text) {
  return /^\[package\][\s\S]*?^\s*version\s*=\s*"(?<version>[^"]+)"\s*$/m.exec(text)?.groups?.version;
}

function replaceCargoPackageVersion(text, version) {
  let replaced = false;
  const after = text.replace(/(^\[package\][\s\S]*?^\s*version\s*=\s*")[^"]+(".*$)/m, (_section, before, afterVersion) => {
    replaced = true;
    return `${before}${version}${afterVersion}`;
  });
  if (!replaced) {
    throw new Error(`${cargoManifestPath}: could not find [package] version`);
  }
  return after;
}

function readCargoLockPackageVersion(text) {
  const section = /\[\[package\]\]\s*\nname = "officegen"\s*\nversion = "(?<version>[^"]+)"/m.exec(text);
  return section?.groups?.version;
}

function replaceCargoLockPackageVersion(text, version) {
  let replaced = false;
  const after = text.replace(/(\[\[package\]\]\s*\nname = "officegen"\s*\nversion = ")[^"]+(")/m, (_match, before, afterVersion) => {
    replaced = true;
    return `${before}${version}${afterVersion}`;
  });
  if (!replaced) {
    throw new Error(`${cargoLockPath}: could not find officegen package version`);
  }
  return after;
}

function textLocation(text, index) {
  const before = text.slice(0, index);
  const lines = before.split(/\r?\n/);
  return { line: lines.length, column: lines.at(-1).length + 1 };
}

async function writeJson(filePath, value) {
  await writeFile(path.resolve(filePath), `${JSON.stringify(value, null, 2)}\n`, "utf8");
}

function usage() {
  console.error("Usage: npm run version:bump -- <patch|minor|major|x.y.z>");
  process.exit(2);
}
