import path from "node:path";

import {
  OptionalContext,
  ValidationResult,
  featureRoot,
  hashFile,
  listJsonFiles,
  mergePlainObjects,
  nowIso,
  readJsonFile,
  requireFeature,
  sha256Json,
  slugify,
  validation,
  writeJsonFile
} from "./common.js";

export interface DesignProfile {
  id: string;
  name: string;
  version?: string;
  tokens: Record<string, unknown>;
  assets?: Record<string, unknown>;
  sourceCapture?: Record<string, unknown>;
  createdAt?: string;
  updatedAt?: string;
  hash?: string;
}

export interface DesignInitOptions extends OptionalContext {
  id: string;
  name?: string;
}

export interface DesignInspectOptions extends OptionalContext {
  id: string;
}

export interface DesignUpdateOptions extends DesignInspectOptions {
  patch: Record<string, unknown>;
}

export interface DesignCaptureOptions extends DesignInspectOptions {
  sourcePath: string;
  label?: string;
}

export interface DesignApplyOptions extends DesignInspectOptions {
  targetPath?: string;
  outputPath?: string;
}

export async function initDesign(options: DesignInitOptions): Promise<DesignProfile> {
  requireFeature(options, "design", "design init");
  const now = nowIso();
  const profile: DesignProfile = withDesignHash({
    id: slugify(options.id),
    name: options.name ?? options.id,
    version: "1.0.0",
    tokens: {
      color: {},
      typography: {},
      spacing: {},
      layout: {}
    },
    assets: {},
    createdAt: now,
    updatedAt: now
  });

  await writeJsonFile(designPath(options, profile.id), profile);
  return profile;
}

export async function listDesigns(options: OptionalContext = {}): Promise<DesignProfile[]> {
  requireFeature(options, "design", "design list");
  const files = await listJsonFiles(featureRoot(options, "design"));
  const designs = await Promise.all(files.map((file) => readJsonFile<DesignProfile>(file)));
  return designs.sort((left, right) => left.id.localeCompare(right.id));
}

export async function inspectDesign(options: DesignInspectOptions): Promise<DesignProfile> {
  requireFeature(options, "design", "design inspect");
  return readJsonFile<DesignProfile>(designPath(options, options.id));
}

export async function updateDesign(options: DesignUpdateOptions): Promise<DesignProfile> {
  requireFeature(options, "design", "design update");
  const current = await inspectDesign(options);
  const updated = withDesignHash({
    ...mergePlainObjects(current as unknown as Record<string, unknown>, options.patch),
    id: current.id,
    updatedAt: nowIso()
  } as unknown as DesignProfile);
  const result = validateDesignProfile(updated);
  if (!result.ok) {
    throw new Error(`Invalid design update: ${result.errors.join("; ")}`);
  }
  await writeJsonFile(designPath(options, current.id), updated);
  return updated;
}

export async function captureDesign(options: DesignCaptureOptions): Promise<DesignProfile> {
  requireFeature(options, "design", "design capture");
  const current = await inspectDesign(options);
  const capture = {
    label: options.label ?? path.basename(options.sourcePath),
    sourcePath: path.resolve(options.cwd ?? process.cwd(), options.sourcePath),
    sha256: await hashFile(path.resolve(options.cwd ?? process.cwd(), options.sourcePath)),
    capturedAt: nowIso()
  };

  return updateDesign({
    ...options,
    patch: {
      sourceCapture: capture
    }
  });
}

export async function applyDesign(options: DesignApplyOptions): Promise<Record<string, unknown>> {
  requireFeature(options, "design", "design apply");
  const design = await inspectDesign(options);
  const plan = {
    kind: "officegen.design.apply",
    generatedAt: nowIso(),
    designId: design.id,
    designHash: design.hash,
    targetPath: options.targetPath ? path.resolve(options.cwd ?? process.cwd(), options.targetPath) : undefined,
    tokens: design.tokens,
    note: "This is a design application plan for @officegen/formats."
  };
  const outputPath = options.outputPath ?? path.join(featureRoot(options, "design"), "runs", `${slugify(design.id)}.apply.json`);
  await writeJsonFile(outputPath, plan);
  return plan;
}

export async function validateDesign(options: DesignInspectOptions): Promise<ValidationResult> {
  requireFeature(options, "design", "design validate");
  return validateDesignProfile(await inspectDesign(options));
}

export function validateDesignProfile(design: DesignProfile): ValidationResult {
  const errors: string[] = [];
  const warnings: string[] = [];

  if (!design.id?.trim()) errors.push("id is required");
  if (!design.name?.trim()) errors.push("name is required");
  if (!design.tokens || typeof design.tokens !== "object") errors.push("tokens object is required");
  if (Object.keys(design.tokens ?? {}).length === 0) warnings.push("design has no tokens");

  return validation(errors.length === 0, errors, warnings);
}

function withDesignHash(design: DesignProfile): DesignProfile {
  return {
    ...design,
    hash: sha256Json({ ...design, hash: undefined })
  };
}

function designPath(context: OptionalContext, id: string): string {
  return path.join(featureRoot(context, "design"), `${slugify(id)}.json`);
}
