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
  await assertVersionText(readmePath, expectedVersion, foundIssues, "README release URLs");
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

async function updateCoreVersionConstant(currentVersion, version) {
  await replaceVersion(coreTypesPath, currentVersion, version);
  await replaceVersionIfExists(coreDistTypesPath, currentVersion, version);
  await replaceVersionIfExists(coreDistDtsPath, currentVersion, version);
}

async function updateReadme(currentVersion, version) {
  await replaceVersion(readmePath, currentVersion, version);
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

async function writeJson(filePath, value) {
  await writeFile(path.resolve(filePath), `${JSON.stringify(value, null, 2)}\n`, "utf8");
}

function usage() {
  console.error("Usage: npm run version:bump -- <patch|minor|major|x.y.z>");
  process.exit(2);
}
