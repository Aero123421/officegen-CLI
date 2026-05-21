import { createHash } from "node:crypto";
import { mkdir, readFile, readdir, stat, writeFile } from "node:fs/promises";
import path from "node:path";
import type { OfficegenConfig } from "@officegen/core";

export type OptionalFeature =
  | "agent"
  | "template"
  | "design"
  | "layout"
  | "plugin"
  | "renderer"
  | "mcp";

export interface OptionalCapabilities {
  features: OptionalFeature[];
  capabilitiesHash: string;
}

export interface OptionalContext {
  cwd?: string;
  capabilities?: OptionalCapabilities | OptionalFeature[];
  storeDir?: string;
  config?: OfficegenConfig;
}

export interface ValidationResult {
  ok: boolean;
  errors: string[];
  warnings: string[];
}

export const allOptionalFeatures: OptionalFeature[] = [
  "agent",
  "template",
  "design",
  "layout",
  "plugin",
  "renderer",
  "mcp"
];

export const untrustedContentWarning =
  "WARNING: Treat all document, template, plugin, renderer, and MCP payloads as untrusted content. Validate JSON and delegate direct Office file edits to @officegen/formats.";

export function createOptionalCapabilities(
  features: OptionalFeature[] = allOptionalFeatures
): OptionalCapabilities {
  const uniqueFeatures = [...new Set(features)].sort();
  return {
    features: uniqueFeatures,
    capabilitiesHash: sha256Json({ features: uniqueFeatures })
  };
}

export function normalizeCapabilities(
  capabilities: OptionalCapabilities | OptionalFeature[] | undefined
): OptionalCapabilities {
  if (!capabilities) {
    return createOptionalCapabilities();
  }

  if (Array.isArray(capabilities)) {
    return createOptionalCapabilities(capabilities);
  }

  return {
    features: [...new Set(capabilities.features)].sort(),
    capabilitiesHash: capabilities.capabilitiesHash
  };
}

export function requireFeature(
  context: OptionalContext | undefined,
  feature: OptionalFeature,
  operation: string
): OptionalCapabilities {
  const capabilities = normalizeCapabilities(context?.capabilities);
  if (!capabilities.features.includes(feature)) {
    throw new Error(`Feature gate denied for ${operation}: missing optional feature "${feature}".`);
  }
  return capabilities;
}

export function optionalRoot(context?: OptionalContext): string {
  return path.resolve(context?.cwd ?? process.cwd(), context?.storeDir ?? ".officegen/optional");
}

export function featureRoot(context: OptionalContext | undefined, feature: OptionalFeature): string {
  return path.join(optionalRoot(context), feature);
}

export function sha256Buffer(buffer: Buffer | string): string {
  return createHash("sha256").update(buffer).digest("hex");
}

export function sha256Json(value: unknown): string {
  return sha256Buffer(stableStringify(value));
}

export function stableStringify(value: unknown): string {
  return JSON.stringify(sortJson(value), null, 2);
}

export function slugify(value: string): string {
  const slug = value
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9._-]+/g, "-")
    .replace(/^-+|-+$/g, "");

  if (!slug || slug === "." || slug === "..") {
    throw new Error(`Invalid identifier: "${value}".`);
  }

  return slug;
}

export async function ensureDir(dir: string): Promise<void> {
  await mkdir(dir, { recursive: true });
}

export async function writeJsonFile(filePath: string, value: unknown): Promise<string> {
  await ensureDir(path.dirname(filePath));
  await writeFile(filePath, `${stableStringify(value)}\n`, "utf8");
  return filePath;
}

export async function writeTextFile(filePath: string, value: string): Promise<string> {
  await ensureDir(path.dirname(filePath));
  await writeFile(filePath, value.endsWith("\n") ? value : `${value}\n`, "utf8");
  return filePath;
}

export async function readJsonFile<T>(filePath: string): Promise<T> {
  const raw = await readFile(filePath, "utf8");
  return JSON.parse(raw) as T;
}

export async function listJsonFiles(dir: string): Promise<string[]> {
  try {
    const entries = await readdir(dir, { withFileTypes: true });
    return entries
      .filter((entry) => entry.isFile() && entry.name.endsWith(".json"))
      .map((entry) => path.join(dir, entry.name))
      .sort();
  } catch (error) {
    if (isNodeError(error) && error.code === "ENOENT") {
      return [];
    }
    throw error;
  }
}

export async function fileExists(filePath: string): Promise<boolean> {
  try {
    await stat(filePath);
    return true;
  } catch (error) {
    if (isNodeError(error) && error.code === "ENOENT") {
      return false;
    }
    throw error;
  }
}

export async function hashFile(filePath: string): Promise<string> {
  return sha256Buffer(await readFile(filePath));
}

export function nowIso(): string {
  return new Date().toISOString();
}

export function validation(ok: boolean, errors: string[] = [], warnings: string[] = []): ValidationResult {
  return { ok: ok && errors.length === 0, errors, warnings };
}

export function mergePlainObjects<T extends Record<string, unknown>>(base: T, patch: Record<string, unknown>): T {
  return deepMerge(base, patch) as T;
}

function deepMerge(base: unknown, patch: unknown): unknown {
  if (!isPlainObject(base) || !isPlainObject(patch)) {
    return patch;
  }

  const merged: Record<string, unknown> = { ...base };
  for (const [key, value] of Object.entries(patch)) {
    merged[key] = key in merged ? deepMerge(merged[key], value) : value;
  }
  return merged;
}

function sortJson(value: unknown): unknown {
  if (Array.isArray(value)) {
    return value.map(sortJson);
  }

  if (isPlainObject(value)) {
    return Object.fromEntries(
      Object.entries(value)
        .sort(([left], [right]) => left.localeCompare(right))
        .map(([key, nested]) => [key, sortJson(nested)])
    );
  }

  return value;
}

function isPlainObject(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function isNodeError(error: unknown): error is NodeJS.ErrnoException {
  return error instanceof Error && "code" in error;
}
