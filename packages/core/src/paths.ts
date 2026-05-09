import { constants } from "node:fs";
import { access, lstat, mkdir, realpath, stat } from "node:fs/promises";
import path from "node:path";
import { expandHome } from "./config.js";
import { OfficegenError } from "./errors.js";
import type { OfficegenConfig, PathValidationOptions, ValidatedPath } from "./types.js";

function normalizeCase(filePath: string): string {
  return process.platform === "win32" ? filePath.toLowerCase() : filePath;
}

function isInside(child: string, parent: string): boolean {
  const relative = path.relative(parent, child);
  return relative === "" || (!relative.startsWith("..") && !path.isAbsolute(relative));
}

async function exists(filePath: string): Promise<boolean> {
  try {
    await access(filePath, constants.F_OK);
    return true;
  } catch {
    return false;
  }
}

export function resolveOfficegenPath(config: OfficegenConfig, inputPath: string): string {
  const expanded = expandHome(inputPath);
  return path.resolve(config.paths.projectRoot, expanded);
}

export async function canonicalizePath(config: OfficegenConfig, inputPath: string): Promise<ValidatedPath> {
  const absolutePath = resolveOfficegenPath(config, inputPath);
  const existed = await exists(absolutePath);
  const realPath = existed ? await realpath(absolutePath) : await realpathForMissingPath(absolutePath);
  return {
    inputPath,
    absolutePath,
    realPath,
    existed,
    warnings: []
  };
}

async function realpathForMissingPath(absolutePath: string): Promise<string> {
  const missingParts: string[] = [];
  let current = absolutePath;
  while (!(await exists(current))) {
    const parent = path.dirname(current);
    if (parent === current) return path.resolve(absolutePath);
    missingParts.unshift(path.basename(current));
    current = parent;
  }
  return path.join(await realpath(current), ...missingParts);
}

async function allowedRoots(config: OfficegenConfig, extraRoots: string[] = []): Promise<string[]> {
  const rawRoots = [...config.security.trustedRoots, ...extraRoots];
  const resolved = rawRoots.map((root) => resolveOfficegenPath(config, root));
  const roots: string[] = [];
  for (const root of resolved) {
    const lexicalRoot = path.resolve(root);
    roots.push(lexicalRoot);
    roots.push((await exists(root)) ? await realpath(root) : lexicalRoot);
  }
  return [...new Set(roots.map(normalizeCase))];
}

async function assertNoSymlinkComponents(targetPath: string, basePath: string): Promise<void> {
  const absoluteBase = path.resolve(basePath);
  const relative = path.relative(absoluteBase, targetPath);
  const startsInsideBase = relative === "" || (!relative.startsWith("..") && !path.isAbsolute(relative));
  const parsed = path.parse(targetPath);
  const relativeParts = (startsInsideBase ? relative : path.relative(parsed.root, targetPath)).split(path.sep).filter(Boolean);
  let current = startsInsideBase ? absoluteBase : parsed.root;
  for (const part of relativeParts) {
    current = path.join(current, part);
    try {
      const stats = await lstat(current);
      if (stats.isSymbolicLink()) {
        throw new OfficegenError("SECURITY_SYMLINK_DENIED", `Symlink or reparse point denied: ${current}`);
      }
    } catch (error) {
      const nodeError = error as NodeJS.ErrnoException;
      if (nodeError.code === "ENOENT") return;
      throw error;
    }
  }
}

export async function validatePath(config: OfficegenConfig, options: PathValidationOptions): Promise<ValidatedPath> {
  if (options.kind === "output" && path.isAbsolute(options.path) && !(options.allowAbsoluteOut ?? config.security.allowAbsoluteOutputPaths)) {
    throw new OfficegenError("SECURITY_ABSOLUTE_OUT_DENIED", `Absolute output path denied: ${options.path}`);
  }
  if (options.kind === "input" && path.isAbsolute(options.path) && !config.security.allowAbsoluteInputPaths) {
    throw new OfficegenError("SECURITY_PATH_OUTSIDE_ROOT", `Absolute input path denied: ${options.path}`);
  }

  const candidate = await canonicalizePath(config, options.path);
  const roots = await allowedRoots(config, options.allowRoots);
  const realCandidate = normalizeCase(candidate.realPath);
  const lexicalCandidate = normalizeCase(candidate.absolutePath);
  const realInside = roots.some((root) => isInside(realCandidate, root));
  const lexicalInside = roots.some((root) => isInside(lexicalCandidate, root));
  if (!realInside || !lexicalInside) {
    if (options.kind === "output" && config.security.outOfProjectPolicy === "warn") {
      candidate.warnings.push("outOfProjectOutput");
    } else if (options.kind === "output" && config.security.outOfProjectPolicy === "allow") {
      candidate.warnings.push("outOfProjectOutputAllowed");
    } else {
      throw new OfficegenError("SECURITY_PATH_OUTSIDE_ROOT", `Path resolves outside allowed roots: ${options.path}`);
    }
  }

  if (!config.security.followSymlinks) {
    await assertNoSymlinkComponents(candidate.absolutePath, config.paths.projectRoot);
    if (candidate.existed && normalizeCase(candidate.absolutePath) !== normalizeCase(candidate.realPath)) {
      throw new OfficegenError("SECURITY_SYMLINK_DENIED", `Symlink or reparse point denied: ${options.path}`);
    }
  }

  if (options.kind === "output") {
    if (candidate.existed && !(options.overwrite ?? config.security.allowOverwrite)) {
      throw new OfficegenError("EDIT_TRANSACTION_FAILED", `Output already exists and overwrite is not enabled: ${options.path}`);
    }
    if (candidate.existed && !config.security.allowHardlinks) {
      const stats = await stat(candidate.absolutePath);
      if (stats.nlink > 1) {
        throw new OfficegenError("SECURITY_HARDLINK_DENIED", `Hardlinked output denied: ${options.path}`);
      }
    }
  }

  return candidate;
}

export async function ensureDirectory(filePath: string): Promise<void> {
  await mkdir(filePath, { recursive: true });
}
