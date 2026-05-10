import path from "node:path";

export const BENCHMARK_MANIFEST_PATH_DENIED = "BENCHMARK_MANIFEST_PATH_DENIED";
export const DEFAULT_BENCHMARK_MANIFEST_PATH = "benchmarks/office-corpus/manifest.json";
export const DEFAULT_BENCHMARK_STORAGE_ROOT = ".officegen/benchmark-corpus";

export function resolveBenchmarkStorageRoot(cwd, storageRoot = DEFAULT_BENCHMARK_STORAGE_ROOT) {
  return resolveBenchmarkRelativePath(cwd, String(storageRoot), "storageRoot");
}

export function resolveBenchmarkManifestPath(cwd, manifestPath = DEFAULT_BENCHMARK_MANIFEST_PATH) {
  return resolveBenchmarkRelativePath(cwd, String(manifestPath), "--manifest");
}

export function resolveBenchmarkDocumentPath(storageRoot, relativePath, field = "documents[].path") {
  return resolveBenchmarkRelativePath(storageRoot, String(relativePath), field);
}

export function assertBenchmarkRelativePath(value, field = "path") {
  const requestedPath = String(value ?? "");
  const reason = benchmarkPathDenyReason(requestedPath);
  if (reason) throw benchmarkPathDenied(requestedPath, field, reason);
}

function resolveBenchmarkRelativePath(root, value, field) {
  const requestedPath = String(value ?? "");
  const reason = benchmarkPathDenyReason(requestedPath);
  if (reason) throw benchmarkPathDenied(requestedPath, field, reason);
  const absoluteRoot = path.resolve(root);
  const absolutePath = path.resolve(absoluteRoot, requestedPath);
  if (!isWithinRoot(absoluteRoot, absolutePath)) {
    throw benchmarkPathDenied(requestedPath, field, "storageRoot escape");
  }
  return absolutePath;
}

function benchmarkPathDenyReason(value) {
  if (!value || value.includes("\0")) return "empty or invalid path";
  if (isAbsoluteOrDrivePath(value)) return "absolute path";
  if (value.split(/[\\/]+/).includes("..")) return "../../ traversal";
  return undefined;
}

function isAbsoluteOrDrivePath(value) {
  return path.isAbsolute(value) || path.win32.isAbsolute(value) || path.posix.isAbsolute(value) || /^[A-Za-z]:/.test(value);
}

function isWithinRoot(root, candidate) {
  const relative = path.relative(root, candidate);
  return relative === "" || (!relative.startsWith("..") && !path.isAbsolute(relative));
}

function benchmarkPathDenied(requestedPath, field, reason) {
  const error = new Error(`Benchmark manifest path denied: ${field} contains ${reason}.`);
  error.code = BENCHMARK_MANIFEST_PATH_DENIED;
  error.details = { field, requestedPath, reason };
  return error;
}
