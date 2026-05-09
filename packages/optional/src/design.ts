import { readFile } from "node:fs/promises";
import path from "node:path";
import { loadZip } from "@officegen/formats";

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
  sourceCapture?: DesignSourceCapture | Record<string, unknown>;
  createdAt?: string;
  updatedAt?: string;
  hash?: string;
}

export interface DesignColorCandidate {
  value: string;
  count: number;
  sources: string[];
}

export interface DesignTextSizeBucket {
  minPt: number;
  maxPt: number;
  count: number;
}

export interface DesignPreviewCandidate {
  stableObjectId: string;
  slide: number;
  title?: string;
  textSnippet?: string;
  shapeCount: number;
  pictureCount: number;
}

export interface DesignContextCandidate {
  key: string;
  value: unknown;
  confidence: number;
  source: string;
}

export interface DesignMapCandidate {
  field: string;
  stableObjectId: string;
  slide: number;
  text: string;
  confidence: number;
}

export interface PptxDesignSignals {
  metadata: {
    format: "pptx";
    slides: number;
    textObjects: number;
    assets: number;
    macros: number;
    byteLength: number;
    title?: string;
    creator?: string;
    created?: string;
    modified?: string;
  };
  colors: DesignColorCandidate[];
  textSizeDistribution: DesignTextSizeBucket[];
  previewCandidates: DesignPreviewCandidate[];
  contextCandidates: DesignContextCandidate[];
  mapCandidates: DesignMapCandidate[];
}

export interface DesignSourceCapture {
  label: string;
  sourcePath: string;
  sha256: string;
  capturedAt: string;
  metadata?: PptxDesignSignals["metadata"];
  colors?: DesignColorCandidate[];
  textSizeDistribution?: DesignTextSizeBucket[];
  previewCandidates?: DesignPreviewCandidate[];
  contextCandidates?: DesignContextCandidate[];
  mapCandidates?: DesignMapCandidate[];
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
  const sourcePath = path.resolve(options.cwd ?? process.cwd(), options.sourcePath);
  const pptxSignals = await capturePptxDesignSignals(sourcePath);
  const capture = {
    label: options.label ?? path.basename(options.sourcePath),
    sourcePath,
    sha256: await hashFile(sourcePath),
    capturedAt: nowIso(),
    ...(pptxSignals
      ? {
          metadata: pptxSignals.metadata,
          colors: pptxSignals.colors,
          textSizeDistribution: pptxSignals.textSizeDistribution,
          previewCandidates: pptxSignals.previewCandidates,
          contextCandidates: pptxSignals.contextCandidates,
          mapCandidates: pptxSignals.mapCandidates
        }
      : {})
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

export async function capturePptxDesignSignals(sourcePath: string): Promise<PptxDesignSignals | undefined> {
  if (path.extname(sourcePath).toLowerCase() !== ".pptx") return undefined;
  const bytes = await readFile(sourcePath);
  const input = { bytes, path: sourcePath, format: "pptx" as const, trusted: false };

  const zip = await loadZip(input);
  const paths = sortedZipFiles(zip);
  const slidePaths = paths.filter((zipPath) => /^ppt\/slides\/slide\d+\.xml$/i.test(zipPath)).sort(naturalSort);
  const mediaPaths = paths.filter((zipPath) => /^ppt\/media\//i.test(zipPath));
  const themePaths = paths.filter((zipPath) => /^ppt\/theme\/theme\d+\.xml$/i.test(zipPath));
  const macros = paths.filter((zipPath) => /vbaProject\.bin$/i.test(zipPath));
  const coreXml = (await readZipText(zip, "docProps/core.xml")) ?? "";
  const colorCounts = new Map<string, { count: number; sources: Set<string> }>();
  const textSizes: number[] = [];
  const previewCandidates: DesignPreviewCandidate[] = [];
  const contextCandidates: DesignContextCandidate[] = [];
  const mapCandidates: DesignMapCandidate[] = [];
  let textObjects = 0;

  for (const themePath of themePaths) {
    collectColors((await readZipText(zip, themePath)) ?? "", themePath, colorCounts);
  }

  for (const [slideIndex, slidePath] of slidePaths.entries()) {
    const slideNo = slideIndex + 1;
    const xml = (await readZipText(zip, slidePath)) ?? "";
    const texts = extractXmlTexts(xml, "t");
    const shapeCount = (xml.match(/<p:sp[\s>]/g) ?? []).length;
    const pictureCount = (xml.match(/<p:pic[\s>]/g) ?? []).length;
    collectColors(xml, slidePath, colorCounts);
    collectTextSizes(xml, textSizes);
    textObjects += texts.length;

    const title = texts.find((text) => text.trim().length > 0);
    previewCandidates.push({
      stableObjectId: stableObjectId("slide", slideNo, 1),
      slide: slideNo,
      title,
      textSnippet: texts.join(" ").slice(0, 240),
      shapeCount,
      pictureCount
    });

    if (title) {
      contextCandidates.push({
        key: slideNo === 1 ? "deckTitle" : `slide${slideNo}Title`,
        value: title,
        confidence: slideNo === 1 ? 0.8 : 0.6,
        source: slidePath
      });
    }

    for (const [textIndex, text] of texts.entries()) {
      if (mapCandidates.length >= 24) break;
      const field = textToFieldName(text);
      if (!field) continue;
      mapCandidates.push({
        field,
        stableObjectId: stableObjectId("text", slideNo, textIndex + 1),
        slide: slideNo,
        text: text.slice(0, 160),
        confidence: fieldConfidence(text),
      });
    }
  }

  return {
    metadata: {
      format: "pptx",
      slides: slidePaths.length,
      textObjects,
      assets: mediaPaths.length,
      macros: macros.length,
      byteLength: input.bytes.byteLength,
      title: firstXmlText(coreXml, "title"),
      creator: firstXmlText(coreXml, "creator"),
      created: firstXmlText(coreXml, "created"),
      modified: firstXmlText(coreXml, "modified")
    },
    colors: [...colorCounts.entries()]
      .map(([value, entry]) => ({ value, count: entry.count, sources: [...entry.sources].slice(0, 6) }))
      .sort((left, right) => right.count - left.count || left.value.localeCompare(right.value))
      .slice(0, 16),
    textSizeDistribution: bucketTextSizes(textSizes),
    previewCandidates: previewCandidates.slice(0, 8),
    contextCandidates: contextCandidates.slice(0, 16),
    mapCandidates
  };
}

function withDesignHash(design: DesignProfile): DesignProfile {
  return {
    ...design,
    hash: sha256Json({ ...design, hash: undefined })
  };
}

function collectColors(xml: string, source: string, colors: Map<string, { count: number; sources: Set<string> }>): void {
  for (const match of xml.matchAll(/<a:srgbClr\b[^>]*\bval="([0-9a-fA-F]{6})"/g)) {
    const value = `#${(match[1] ?? "").toUpperCase()}`;
    const entry = colors.get(value) ?? { count: 0, sources: new Set<string>() };
    entry.count += 1;
    entry.sources.add(source);
    colors.set(value, entry);
  }
}

function sortedZipFiles(zip: Awaited<ReturnType<typeof loadZip>>): string[] {
  return Object.keys(zip.files)
    .filter((name) => !zip.files[name]?.dir)
    .sort((left, right) => left.localeCompare(right));
}

async function readZipText(zip: Awaited<ReturnType<typeof loadZip>>, zipPath: string): Promise<string | undefined> {
  const file = zip.file(zipPath);
  return file ? file.async("string") : undefined;
}

function extractXmlTexts(xml: string, localName: string): string[] {
  const pattern = new RegExp(`<[^>]*:?${localName}(?:\\s[^>]*)?>([\\s\\S]*?)<\\/[^>]*:?${localName}>`, "g");
  return [...xml.matchAll(pattern)]
    .map((match) => decodeXmlEntities(match[1] ?? "").trim())
    .filter(Boolean);
}

function decodeXmlEntities(value: string): string {
  return value
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

function collectTextSizes(xml: string, sizes: number[]): void {
  for (const match of xml.matchAll(/\bsz="(\d+)"/g)) {
    const raw = Number(match[1]);
    if (Number.isFinite(raw) && raw > 0) {
      sizes.push(Math.round(raw / 100));
    }
  }
}

function bucketTextSizes(sizes: number[]): DesignTextSizeBucket[] {
  const buckets = new Map<string, DesignTextSizeBucket>();
  for (const size of sizes) {
    const minPt = Math.floor(size / 4) * 4;
    const maxPt = minPt + 3;
    const key = `${minPt}-${maxPt}`;
    const bucket = buckets.get(key) ?? { minPt, maxPt, count: 0 };
    bucket.count += 1;
    buckets.set(key, bucket);
  }
  return [...buckets.values()].sort((left, right) => left.minPt - right.minPt);
}

function textToFieldName(text: string): string | undefined {
  const normalized = text.trim().replace(/\s+/g, " ");
  if (normalized.length < 2 || normalized.length > 80) return undefined;
  const label = normalized.replace(/[:：]$/, "");
  const field = label
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "");
  return field && /^[a-z]/.test(field) ? field : undefined;
}

function fieldConfidence(text: string): number {
  if (/[:：]\s*$/.test(text)) return 0.75;
  if (/\{\{[^}]+\}\}/.test(text)) return 0.9;
  if (text.length <= 32) return 0.6;
  return 0.35;
}

function firstXmlText(xml: string, localName: string): string | undefined {
  return extractXmlTexts(xml, localName)[0];
}

function stableObjectId(kind: string, slide: number, ordinal: number): string {
  return `pptx:s${String(slide).padStart(3, "0")}:${kind}:${String(ordinal).padStart(4, "0")}`;
}

function naturalSort(left: string, right: string): number {
  return left.localeCompare(right, undefined, { numeric: true, sensitivity: "base" });
}

function designPath(context: OptionalContext, id: string): string {
  return path.join(featureRoot(context, "design"), `${slugify(id)}.json`);
}
