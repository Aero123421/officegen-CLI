import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { loadZip, stableHashId } from "@officegen/formats";

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
  sha256Buffer,
  sha256Json,
  slugify,
  untrustedContentWarning,
  validation,
  writeJsonFile,
  writeTextFile
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

export interface DesignColorRoleCandidate {
  role: "background" | "text" | "accent";
  value: string;
  confidence: number;
  sources: string[];
  reason: string;
}

export interface DesignTextSizeBucket {
  minPt: number;
  maxPt: number;
  count: number;
}

export interface DesignBounds {
  x: number;
  y: number;
  width: number;
  height: number;
  unit: "ratio";
}

export interface DesignBBoxPattern {
  kind: "title" | "body" | "image";
  count: number;
  slides: number[];
  average: DesignBounds;
}

export interface DesignPreviewCandidate {
  stableObjectId: string;
  slide: number;
  title?: string;
  textSnippet?: string;
  shapeCount: number;
  pictureCount: number;
  slideType?: DesignSlideType;
  densityScore?: number;
  previewPath?: string;
  evidencePath?: string;
}

export interface DesignContextCandidate {
  key: string;
  value: unknown;
  confidence: number;
  source: string;
  untrusted?: true;
}

export interface DesignMapCandidate {
  field: string;
  stableObjectId: string;
  slide: number;
  text: string;
  confidence: number;
  untrusted?: true;
}

export interface TemplatePlaceholderCandidate {
  field: string;
  stableObjectId: string;
  slide: number;
  text?: string;
  name?: string;
  placeholderType?: string;
  bounds?: DesignBounds;
  confidence: number;
  source: string;
  untrusted?: true;
}

export interface NamedShapeCandidate {
  name: string;
  stableObjectId: string;
  slide: number;
  kind: "shape" | "picture" | "chart" | "diagram";
  text?: string;
  bounds?: DesignBounds;
  confidence: number;
  source: string;
  untrusted?: true;
}

export interface TemplateSchemaCandidate {
  name: string;
  type: "string" | "number" | "boolean" | "date" | "json";
  required: boolean;
  confidence: number;
  reason: string;
}

export interface TemplateMapSuggestion {
  schema: "officegen.template.map@1.2";
  mapping: Record<string, string>;
  confidence: number;
  candidateCount: number;
}

export type DesignSlideType =
  | "title"
  | "title-body"
  | "section"
  | "image"
  | "chart"
  | "diagram"
  | "mixed"
  | "blank";

export interface DesignSlideSignal {
  stableObjectId: string;
  slide: number;
  title?: string;
  slideType: DesignSlideType;
  shapeCount: number;
  pictureCount: number;
  chartCount: number;
  diagramCount: number;
  textObjectCount: number;
  densityScore: number;
  titleBounds?: DesignBounds;
  bodyBounds?: DesignBounds;
  imageBounds: DesignBounds[];
  previewPath?: string;
  evidencePath?: string;
}

export interface PptxDesignArtifactPaths {
  contextPath?: string;
  evidencePath?: string;
  templateMapSuggestedPath?: string;
  schemaCandidatesPath?: string;
  previewPaths: string[];
}

export interface PptxDesignTrustEnvelope {
  trusted: {
    schema: "officegen.design.signals.trusted@1.2";
    format: "pptx";
    sourcePath: string;
    sha256: string;
    byteLength: number;
    generatedAt: string;
    summary: Record<string, unknown>;
  };
  untrusted: {
    schema: "officegen.design.signals.untrusted@1.2";
    slideTitles: string[];
    textSamples: string[];
    shapeNames: string[];
  };
  agentInstruction: string;
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
  colorRoleCandidates: DesignColorRoleCandidate[];
  textSizeDistribution: DesignTextSizeBucket[];
  bboxPatterns: DesignBBoxPattern[];
  slideSignals: DesignSlideSignal[];
  densityScore: number;
  chartPresence: { count: number; slides: number[] };
  diagramPresence: { count: number; slides: number[] };
  previewCandidates: DesignPreviewCandidate[];
  contextCandidates: DesignContextCandidate[];
  mapCandidates: DesignMapCandidate[];
  placeholderCandidates: TemplatePlaceholderCandidate[];
  namedShapeCandidates: NamedShapeCandidate[];
  schemaCandidates: TemplateSchemaCandidate[];
  templateMapSuggested: TemplateMapSuggestion;
  artifactPaths?: PptxDesignArtifactPaths;
  trust: PptxDesignTrustEnvelope;
}

export interface DesignSourceCapture {
  label: string;
  sourcePath: string;
  sha256: string;
  capturedAt: string;
  metadata?: PptxDesignSignals["metadata"];
  colors?: DesignColorCandidate[];
  colorRoleCandidates?: DesignColorRoleCandidate[];
  textSizeDistribution?: DesignTextSizeBucket[];
  bboxPatterns?: DesignBBoxPattern[];
  slideSignals?: DesignSlideSignal[];
  densityScore?: number;
  chartPresence?: PptxDesignSignals["chartPresence"];
  diagramPresence?: PptxDesignSignals["diagramPresence"];
  previewCandidates?: DesignPreviewCandidate[];
  contextCandidates?: DesignContextCandidate[];
  mapCandidates?: DesignMapCandidate[];
  placeholderCandidates?: TemplatePlaceholderCandidate[];
  namedShapeCandidates?: NamedShapeCandidate[];
  schemaCandidates?: TemplateSchemaCandidate[];
  templateMapSuggested?: TemplateMapSuggestion;
  artifactPaths?: PptxDesignArtifactPaths;
  trust?: PptxDesignTrustEnvelope;
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

export interface PptxDesignSignalOptions {
  cwd?: string;
  artifactsDir?: string;
}

interface RawBounds {
  x: number;
  y: number;
  width: number;
  height: number;
}

interface SlideObjectSignal {
  stableObjectId: string;
  slide: number;
  kind: "shape" | "picture" | "chart" | "diagram";
  ordinal: number;
  name?: string;
  placeholderType?: string;
  text?: string;
  bounds?: DesignBounds;
  source: string;
  colors: string[];
  fontSizes: number[];
}

interface SlideSize {
  width: number;
  height: number;
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
  const pptxSignals = await capturePptxDesignSignals(sourcePath, {
    cwd: options.cwd,
    artifactsDir: path.join(featureRoot(options, "design"), "captures", slugify(current.id), slugify(path.basename(sourcePath, path.extname(sourcePath))))
  });
  const capture = {
    label: options.label ?? path.basename(options.sourcePath),
    sourcePath,
    sha256: await hashFile(sourcePath),
    capturedAt: nowIso(),
    ...(pptxSignals
      ? {
          metadata: pptxSignals.metadata,
          colors: pptxSignals.colors,
          colorRoleCandidates: pptxSignals.colorRoleCandidates,
          textSizeDistribution: pptxSignals.textSizeDistribution,
          bboxPatterns: pptxSignals.bboxPatterns,
          slideSignals: pptxSignals.slideSignals,
          densityScore: pptxSignals.densityScore,
          chartPresence: pptxSignals.chartPresence,
          diagramPresence: pptxSignals.diagramPresence,
          previewCandidates: pptxSignals.previewCandidates,
          contextCandidates: pptxSignals.contextCandidates,
          mapCandidates: pptxSignals.mapCandidates,
          placeholderCandidates: pptxSignals.placeholderCandidates,
          namedShapeCandidates: pptxSignals.namedShapeCandidates,
          schemaCandidates: pptxSignals.schemaCandidates,
          templateMapSuggested: pptxSignals.templateMapSuggested,
          artifactPaths: pptxSignals.artifactPaths,
          trust: pptxSignals.trust
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
  if (options.targetPath && options.outputPath && path.extname(options.targetPath).toLowerCase() === ".pptx" && path.extname(options.outputPath).toLowerCase() === ".pptx") {
    const applied = await applyPptxDesign(design, path.resolve(options.cwd ?? process.cwd(), options.targetPath), options.outputPath);
    return {
      kind: "officegen.design.apply",
      planOnly: false,
      mutatesOffice: true,
      generatedAt: nowIso(),
      designId: design.id,
      designHash: design.hash,
      targetPath: options.targetPath,
      out: options.outputPath,
      ...applied
    };
  }
  const plan = {
    kind: "officegen.design.apply",
    planOnly: true,
    mutatesOffice: false,
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

async function applyPptxDesign(design: DesignProfile, targetPath: string, outputPath: string): Promise<Record<string, unknown>> {
  const bytes = await readFile(targetPath);
  const zip = await loadZip({ bytes, path: targetPath, format: "pptx", trusted: false });
  const accent = designAccentColor(design);
  const fontFace = designFontFace(design);
  const changedParts: string[] = [];
  for (const file of Object.values(zip.files)) {
    if (file.dir || !/^ppt\/theme\/theme\d+\.xml$/i.test(file.name)) continue;
    const xml = await file.async("string");
    let next = xml;
    if (accent) {
      next = replaceThemeColor(next, "accent1", accent);
      next = replaceThemeColor(next, "accent2", accent);
    }
    if (fontFace) {
      next = next.replace(/(<a:latin\b[^>]*typeface=")[^"]*(")/g, `$1${escapeXmlAttr(fontFace)}$2`);
      next = next.replace(/(<a:ea\b[^>]*typeface=")[^"]*(")/g, `$1${escapeXmlAttr(fontFace)}$2`);
    }
    if (next !== xml) {
      zip.file(file.name, next);
      changedParts.push(file.name);
    }
  }
  await writeFile(outputPath, await zip.generateAsync({ type: "uint8array", compression: "DEFLATE", compressionOptions: { level: 6 } }));
  return {
    changedParts,
    tokensApplied: { accent, fontFace },
    mastersAndLayoutsPreserved: true,
    note: changedParts.length ? "Updated PPTX theme tokens while preserving slide masters, layouts, placeholders, relationships, and content." : "No theme parts were changed."
  };
}

function designAccentColor(design: DesignProfile): string | undefined {
  const tokens = design.tokens as Record<string, unknown>;
  const color = tokens.color as Record<string, unknown> | undefined;
  const token = color?.accent ?? color?.primary ?? color?.brand;
  if (typeof token === "string") return normalizeHexColor(token);
  const capture = design.sourceCapture as DesignSourceCapture | undefined;
  return normalizeHexColor(capture?.colorRoleCandidates?.find((candidate) => candidate.role === "accent")?.value);
}

function designFontFace(design: DesignProfile): string | undefined {
  const tokens = design.tokens as Record<string, unknown>;
  const typography = tokens.typography as Record<string, unknown> | undefined;
  const family = typography?.fontFamily ?? typography?.bodyFont ?? typography?.headingFont;
  return typeof family === "string" ? family : undefined;
}

function replaceThemeColor(xml: string, role: string, hex: string): string {
  const re = new RegExp(`(<a:${role}>[\\s\\S]*?<a:srgbClr\\b[^>]*val=")[^"]*("[\\s\\S]*?</a:${role}>)`, "i");
  return re.test(xml) ? xml.replace(re, `$1${hex}$2`) : xml;
}

function normalizeHexColor(value: string | undefined): string | undefined {
  const match = /^#?([0-9a-f]{6})$/i.exec(value ?? "");
  return match?.[1]?.toUpperCase();
}

function escapeXmlAttr(value: string): string {
  return value.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
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

export async function capturePptxDesignSignals(
  sourcePath: string,
  options: PptxDesignSignalOptions = {}
): Promise<PptxDesignSignals | undefined> {
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
  const presentationXml = (await readZipText(zip, "ppt/presentation.xml")) ?? "";
  const slideSize = readSlideSize(presentationXml);
  const colorCounts = new Map<string, { count: number; sources: Set<string> }>();
  const backgroundColorCounts = new Map<string, { count: number; sources: Set<string> }>();
  const textColorCounts = new Map<string, { count: number; sources: Set<string> }>();
  const textSizes: number[] = [];
  const previewCandidates: DesignPreviewCandidate[] = [];
  const contextCandidates: DesignContextCandidate[] = [];
  const mapCandidates: DesignMapCandidate[] = [];
  const placeholderCandidates: TemplatePlaceholderCandidate[] = [];
  const namedShapeCandidates: NamedShapeCandidate[] = [];
  const slideSignals: DesignSlideSignal[] = [];
  const textSamples: string[] = [];
  const shapeNames: string[] = [];
  const chartSlides = new Set<number>();
  const diagramSlides = new Set<number>();
  let chartCount = 0;
  let diagramCount = 0;
  let textObjects = 0;

  for (const themePath of themePaths) {
    collectColors((await readZipText(zip, themePath)) ?? "", themePath, colorCounts);
  }

  for (const [slideIndex, slidePath] of slidePaths.entries()) {
    const slideNo = slideIndex + 1;
    const xml = (await readZipText(zip, slidePath)) ?? "";
    collectColors(xml, slidePath, colorCounts);
    collectBackgroundColors(xml, slidePath, backgroundColorCounts);
    collectTextSizes(xml, textSizes);

    const objects = extractSlideObjects(xml, slideNo, slidePath, slideSize);
    for (const object of objects) {
      if (object.text) {
        textObjects += 1;
        textSamples.push(object.text);
        for (const color of object.colors) incrementColor(textColorCounts, color, slidePath);
        for (const size of object.fontSizes) textSizes.push(size);
      }
      if (object.name) shapeNames.push(object.name);
      if (object.kind === "chart") {
        chartCount += 1;
        chartSlides.add(slideNo);
      }
      if (object.kind === "diagram") {
        diagramCount += 1;
        diagramSlides.add(slideNo);
      }
    }

    const shapeObjects = objects.filter((object) => object.kind === "shape");
    const pictureObjects = objects.filter((object) => object.kind === "picture");
    const textObjectsOnSlide = objects.filter((object) => object.text?.trim());
    const titleObject = pickTitleObject(textObjectsOnSlide);
    const bodyObject = pickBodyObject(textObjectsOnSlide, titleObject);
    const title = titleObject?.text?.trim() || textObjectsOnSlide[0]?.text?.trim();
    const slideChartCount = objects.filter((object) => object.kind === "chart").length;
    const slideDiagramCount = objects.filter((object) => object.kind === "diagram").length;
    const densityScore = scoreSlideDensity(objects);
    const slideType = classifySlide({
      title,
      textObjectCount: textObjectsOnSlide.length,
      pictureCount: pictureObjects.length,
      chartCount: slideChartCount,
      diagramCount: slideDiagramCount,
      densityScore
    });

    previewCandidates.push({
      stableObjectId: stableSlideObjectId(slidePath),
      slide: slideNo,
      title,
      textSnippet: textObjectsOnSlide.map((object) => object.text).filter(Boolean).join(" ").slice(0, 240),
      shapeCount: shapeObjects.length,
      pictureCount: pictureObjects.length,
      slideType,
      densityScore
    });

    slideSignals.push({
      stableObjectId: stableSlideObjectId(slidePath),
      slide: slideNo,
      title,
      slideType,
      shapeCount: shapeObjects.length,
      pictureCount: pictureObjects.length,
      chartCount: slideChartCount,
      diagramCount: slideDiagramCount,
      textObjectCount: textObjectsOnSlide.length,
      densityScore,
      titleBounds: titleObject?.bounds,
      bodyBounds: bodyObject?.bounds,
      imageBounds: pictureObjects.map((object) => object.bounds).filter((bounds): bounds is DesignBounds => Boolean(bounds))
    });

    if (title) {
      contextCandidates.push({
        key: slideNo === 1 ? "deckTitle" : `slide${slideNo}Title`,
        value: title,
        confidence: slideNo === 1 ? 0.8 : 0.6,
        source: slidePath,
        untrusted: true
      });
    }

    for (const object of textObjectsOnSlide) {
      if (mapCandidates.length >= 32) break;
      const field = textToFieldName(object.text ?? "");
      if (!field) continue;
      mapCandidates.push({
        field,
        stableObjectId: object.stableObjectId,
        slide: slideNo,
        text: (object.text ?? "").slice(0, 160),
        confidence: fieldConfidence(object.text ?? "", object.placeholderType),
        untrusted: true
      });
    }

    for (const object of objects) {
      const placeholder = object.placeholderType || placeholderFieldFromText(object.text);
      const field = textToFieldName(placeholder ?? object.text ?? object.name ?? "");
      if (field && placeholderCandidates.length < 32) {
        placeholderCandidates.push({
          field,
          stableObjectId: object.stableObjectId,
          slide: slideNo,
          text: object.text?.slice(0, 160),
          name: object.name,
          placeholderType: object.placeholderType,
          bounds: object.bounds,
          confidence: placeholder ? 0.85 : fieldConfidence(object.text ?? object.name ?? "", object.placeholderType),
          source: slidePath,
          untrusted: true
        });
      }

      if (object.name && !isGenericShapeName(object.name) && namedShapeCandidates.length < 32) {
        namedShapeCandidates.push({
          name: object.name,
          stableObjectId: object.stableObjectId,
          slide: slideNo,
          kind: object.kind,
          text: object.text?.slice(0, 160),
          bounds: object.bounds,
          confidence: object.placeholderType ? 0.8 : 0.65,
          source: slidePath,
          untrusted: true
        });
      }
    }
  }

  const schemaCandidates = buildSchemaCandidates(placeholderCandidates, mapCandidates);
  const templateMapSuggested = buildTemplateMapSuggestion(schemaCandidates, placeholderCandidates, mapCandidates);
  const bboxPatterns = buildBBoxPatterns(slideSignals);
  const colorRoleCandidates = buildColorRoleCandidates(colorCounts, backgroundColorCounts, textColorCounts);
  const averageDensity = round(slideSignals.reduce((sum, slide) => sum + slide.densityScore, 0) / Math.max(1, slideSignals.length), 3);
  const generatedAt = nowIso();
  const trust: PptxDesignTrustEnvelope = {
    trusted: {
      schema: "officegen.design.signals.trusted@1.2",
      format: "pptx",
      sourcePath,
      sha256: sha256Buffer(bytes),
      byteLength: input.bytes.byteLength,
      generatedAt,
      summary: {
        slides: slidePaths.length,
        textObjects,
        assets: mediaPaths.length,
        macros: macros.length,
        densityScore: averageDensity,
        chartCount,
        diagramCount
      }
    },
    untrusted: {
      schema: "officegen.design.signals.untrusted@1.2",
      slideTitles: slideSignals.map((slide) => slide.title).filter((title): title is string => Boolean(title)),
      textSamples: textSamples.slice(0, 24),
      shapeNames: [...new Set(shapeNames)].slice(0, 24)
    },
    agentInstruction: untrustedContentWarning
  };

  const signals: PptxDesignSignals = {
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
    colorRoleCandidates,
    textSizeDistribution: bucketTextSizes(textSizes),
    bboxPatterns,
    slideSignals,
    densityScore: averageDensity,
    chartPresence: { count: chartCount, slides: [...chartSlides].sort((left, right) => left - right) },
    diagramPresence: { count: diagramCount, slides: [...diagramSlides].sort((left, right) => left - right) },
    previewCandidates: previewCandidates.slice(0, 8),
    contextCandidates: contextCandidates.slice(0, 16),
    mapCandidates,
    placeholderCandidates,
    namedShapeCandidates,
    schemaCandidates,
    templateMapSuggested,
    trust
  };

  if (options.artifactsDir) {
    signals.artifactPaths = await writeSignalArtifacts(signals, sourcePath, options);
  }

  return signals;
}

function withDesignHash(design: DesignProfile): DesignProfile {
  return {
    ...design,
    hash: sha256Json({ ...design, hash: undefined })
  };
}

function collectColors(xml: string, source: string, colors: Map<string, { count: number; sources: Set<string> }>): void {
  for (const match of xml.matchAll(/<a:srgbClr\b[^>]*\bval="([0-9a-fA-F]{6})"/g)) {
    incrementColor(colors, `#${(match[1] ?? "").toUpperCase()}`, source);
  }
}

function collectBackgroundColors(xml: string, source: string, colors: Map<string, { count: number; sources: Set<string> }>): void {
  for (const match of xml.matchAll(/<p:bg[\s\S]*?<\/p:bg>/g)) {
    collectColors(match[0], source, colors);
  }
}

function incrementColor(colors: Map<string, { count: number; sources: Set<string> }>, value: string, source: string): void {
  const normalized = value.startsWith("#") ? value.toUpperCase() : `#${value.toUpperCase()}`;
  const entry = colors.get(normalized) ?? { count: 0, sources: new Set<string>() };
  entry.count += 1;
  entry.sources.add(source);
  colors.set(normalized, entry);
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
  const tag = localName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const pattern = new RegExp(
    `<(?:[A-Za-z_][\\w.-]*:)?${tag}\\b[^>]*>([\\s\\S]*?)<\\/(?:[A-Za-z_][\\w.-]*:)?${tag}>`,
    "g"
  );
  return [...xml.matchAll(pattern)]
    .map((match) => decodeXmlEntities(stripXmlTags(match[1] ?? "")).trim())
    .filter(Boolean);
}

function stripXmlTags(value: string): string {
  return value.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ");
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

function extractSlideObjects(xml: string, slide: number, source: string, slideSize: SlideSize): SlideObjectSignal[] {
  const objects: SlideObjectSignal[] = [];
  let shapeOrdinal = 0;
  for (const block of extractBlocks(xml, "p:sp")) {
    shapeOrdinal += 1;
    const shapeId = readShapeId(block);
    const text = extractXmlTexts(block, "t").join(" ").replace(/\s+/g, " ").trim();
    objects.push({
      stableObjectId: shapeId
        ? stableHashId("pptx", slideScope(source), "shape", `${source}#${shapeId}`)
        : stableObjectId("shape", source, shapeOrdinal),
      slide,
      kind: "shape",
      ordinal: shapeOrdinal,
      name: readAttr(block, "name"),
      placeholderType: readPlaceholderType(block),
      text: text || undefined,
      bounds: normalizeBounds(readRawBounds(block), slideSize),
      source,
      colors: extractColors(block),
      fontSizes: extractFontSizes(block)
    });
  }

  let pictureOrdinal = 0;
  for (const block of extractBlocks(xml, "p:pic")) {
    pictureOrdinal += 1;
    objects.push({
      stableObjectId: stableObjectId("picture", source, pictureOrdinal),
      slide,
      kind: "picture",
      ordinal: pictureOrdinal,
      name: readAttr(block, "name"),
      bounds: normalizeBounds(readRawBounds(block), slideSize),
      source,
      colors: extractColors(block),
      fontSizes: []
    });
  }

  let frameOrdinal = 0;
  for (const block of extractBlocks(xml, "p:graphicFrame")) {
    frameOrdinal += 1;
    const kind = /<c:chart\b/i.test(block) ? "chart" : /<dgm:/i.test(block) ? "diagram" : "diagram";
    objects.push({
      stableObjectId: stableObjectId(kind, source, frameOrdinal),
      slide,
      kind,
      ordinal: frameOrdinal,
      name: readAttr(block, "name"),
      bounds: normalizeBounds(readRawBounds(block), slideSize),
      source,
      colors: extractColors(block),
      fontSizes: []
    });
  }

  if (/<c:chart\b/i.test(xml) && !objects.some((object) => object.kind === "chart")) {
    objects.push({ stableObjectId: stableObjectId("chart", source, 1), slide, kind: "chart", ordinal: 1, source, colors: [], fontSizes: [] });
  }
  if (/<dgm:/i.test(xml) && !objects.some((object) => object.kind === "diagram")) {
    objects.push({ stableObjectId: stableObjectId("diagram", source, 1), slide, kind: "diagram", ordinal: 1, source, colors: [], fontSizes: [] });
  }
  return objects;
}

function extractBlocks(xml: string, tagName: string): string[] {
  const escaped = tagName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  return [...xml.matchAll(new RegExp(`<${escaped}\\b[\\s\\S]*?<\\/${escaped}>`, "g"))].map((match) => match[0]);
}

function readAttr(xml: string, attr: string): string | undefined {
  const match = xml.match(new RegExp(`\\b${attr}="([^"]+)"`));
  return match ? decodeXmlEntities(match[1] ?? "") : undefined;
}

function readShapeId(xml: string): string | undefined {
  const cNvPr = /<p:cNvPr\b([^>]*)\/>/.exec(xml)?.[1] ?? "";
  return readAttr(cNvPr, "id");
}

function readPlaceholderType(xml: string): string | undefined {
  const ph = xml.match(/<p:ph\b[^>]*>/);
  if (!ph) return undefined;
  return readAttr(ph[0], "type") ?? "body";
}

function readRawBounds(xml: string): RawBounds | undefined {
  const off = xml.match(/<a:off\b[^>]*\bx="(-?\d+)"[^>]*\by="(-?\d+)"/);
  const ext = xml.match(/<a:ext\b[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"/);
  if (!off || !ext) return undefined;
  return {
    x: Number(off[1]),
    y: Number(off[2]),
    width: Number(ext[1]),
    height: Number(ext[2])
  };
}

function normalizeBounds(bounds: RawBounds | undefined, slideSize: SlideSize): DesignBounds | undefined {
  if (!bounds || slideSize.width <= 0 || slideSize.height <= 0) return undefined;
  return {
    x: round(bounds.x / slideSize.width),
    y: round(bounds.y / slideSize.height),
    width: round(bounds.width / slideSize.width),
    height: round(bounds.height / slideSize.height),
    unit: "ratio"
  };
}

function extractColors(xml: string): string[] {
  return [...xml.matchAll(/<a:srgbClr\b[^>]*\bval="([0-9a-fA-F]{6})"/g)].map((match) => `#${(match[1] ?? "").toUpperCase()}`);
}

function extractFontSizes(xml: string): number[] {
  const sizes: number[] = [];
  collectTextSizes(xml, sizes);
  return sizes;
}

function readSlideSize(presentationXml: string): SlideSize {
  const match = presentationXml.match(/<p:sldSz\b[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"/);
  return {
    width: Number(match?.[1] ?? 9144000),
    height: Number(match?.[2] ?? 6858000)
  };
}

function pickTitleObject(objects: SlideObjectSignal[]): SlideObjectSignal | undefined {
  return (
    objects.find((object) => object.placeholderType === "title" || object.placeholderType === "ctrTitle") ??
    objects
      .filter((object) => object.bounds)
      .sort((left, right) => (left.bounds?.y ?? 1) - (right.bounds?.y ?? 1) || (right.fontSizes[0] ?? 0) - (left.fontSizes[0] ?? 0))[0] ??
    objects[0]
  );
}

function pickBodyObject(objects: SlideObjectSignal[], titleObject: SlideObjectSignal | undefined): SlideObjectSignal | undefined {
  return objects
    .filter((object) => object !== titleObject)
    .sort((left, right) => (right.bounds?.height ?? 0) * (right.bounds?.width ?? 0) - (left.bounds?.height ?? 0) * (left.bounds?.width ?? 0))[0];
}

function classifySlide(input: {
  title?: string;
  textObjectCount: number;
  pictureCount: number;
  chartCount: number;
  diagramCount: number;
  densityScore: number;
}): DesignSlideType {
  if (input.chartCount > 0) return "chart";
  if (input.diagramCount > 0) return "diagram";
  if (input.pictureCount > 0 && input.textObjectCount <= 1) return "image";
  if (input.textObjectCount === 0 && input.pictureCount === 0) return "blank";
  if (input.title && input.textObjectCount === 1) return "title";
  if (input.title && input.textObjectCount >= 2) return "title-body";
  if (input.densityScore < 0.2) return "section";
  return "mixed";
}

function scoreSlideDensity(objects: SlideObjectSignal[]): number {
  const area = objects.reduce((sum, object) => {
    const bounds = object.bounds;
    return sum + (bounds ? Math.max(0, bounds.width) * Math.max(0, bounds.height) : 0.03);
  }, 0);
  const textChars = objects.reduce((sum, object) => sum + (object.text?.length ?? 0), 0);
  return round(Math.min(1, area * 0.7 + Math.min(1, textChars / 1200) * 0.3), 3);
}

function buildColorRoleCandidates(
  colors: Map<string, { count: number; sources: Set<string> }>,
  backgroundColors: Map<string, { count: number; sources: Set<string> }>,
  textColors: Map<string, { count: number; sources: Set<string> }>
): DesignColorRoleCandidate[] {
  const roles: DesignColorRoleCandidate[] = [];
  const topBackground = topColor(backgroundColors) ?? topColor(colors);
  const topText = topColor(textColors);
  if (topBackground) {
    roles.push({
      role: "background",
      value: topBackground.value,
      confidence: backgroundColors.size > 0 ? 0.82 : 0.45,
      sources: topBackground.sources,
      reason: backgroundColors.size > 0 ? "explicit slide background color" : "most frequent deck color"
    });
  }
  if (topText) {
    roles.push({
      role: "text",
      value: topText.value,
      confidence: 0.72,
      sources: topText.sources,
      reason: "most frequent color inside text-bearing shapes"
    });
  }
  const excluded = new Set(roles.map((role) => role.value));
  for (const accent of [...colors.entries()]
    .map(([value, entry]) => ({ value, count: entry.count, sources: [...entry.sources].slice(0, 6) }))
    .filter((entry) => !excluded.has(entry.value))
    .sort((left, right) => right.count - left.count)
    .slice(0, 3)) {
    roles.push({
      role: "accent",
      value: accent.value,
      confidence: 0.55,
      sources: accent.sources,
      reason: "frequent non-background/non-text color"
    });
  }
  return roles;
}

function topColor(colors: Map<string, { count: number; sources: Set<string> }>): { value: string; sources: string[] } | undefined {
  const first = [...colors.entries()].sort((left, right) => right[1].count - left[1].count || left[0].localeCompare(right[0]))[0];
  return first ? { value: first[0], sources: [...first[1].sources].slice(0, 6) } : undefined;
}

function buildBBoxPatterns(slides: DesignSlideSignal[]): DesignBBoxPattern[] {
  const patterns: DesignBBoxPattern[] = [];
  const title = averageBounds(slides.map((slide) => ({ slide: slide.slide, bounds: slide.titleBounds })).filter(hasBounds));
  if (title) patterns.push({ kind: "title", ...title });
  const body = averageBounds(slides.map((slide) => ({ slide: slide.slide, bounds: slide.bodyBounds })).filter(hasBounds));
  if (body) patterns.push({ kind: "body", ...body });
  const images = averageBounds(slides.flatMap((slide) => slide.imageBounds.map((bounds) => ({ slide: slide.slide, bounds }))).filter(hasBounds));
  if (images) patterns.push({ kind: "image", ...images });
  return patterns;
}

function hasBounds(value: { slide: number; bounds?: DesignBounds }): value is { slide: number; bounds: DesignBounds } {
  return Boolean(value.bounds);
}

function averageBounds(items: { slide: number; bounds: DesignBounds }[]): Omit<DesignBBoxPattern, "kind"> | undefined {
  if (items.length === 0) return undefined;
  return {
    count: items.length,
    slides: [...new Set(items.map((item) => item.slide))],
    average: {
      x: round(items.reduce((sum, item) => sum + item.bounds.x, 0) / items.length),
      y: round(items.reduce((sum, item) => sum + item.bounds.y, 0) / items.length),
      width: round(items.reduce((sum, item) => sum + item.bounds.width, 0) / items.length),
      height: round(items.reduce((sum, item) => sum + item.bounds.height, 0) / items.length),
      unit: "ratio"
    }
  };
}

function buildSchemaCandidates(
  placeholders: TemplatePlaceholderCandidate[],
  mapCandidates: DesignMapCandidate[]
): TemplateSchemaCandidate[] {
  const byName = new Map<string, TemplateSchemaCandidate>();
  for (const candidate of [
    ...placeholders.map((placeholder) => ({
      name: placeholder.field,
      confidence: placeholder.confidence,
      reason: placeholder.placeholderType ? `placeholder:${placeholder.placeholderType}` : "shape/text candidate"
    })),
    ...mapCandidates.map((candidate) => ({
      name: candidate.field,
      confidence: candidate.confidence,
      reason: "text label candidate"
    }))
  ]) {
    const existing = byName.get(candidate.name);
    if (!existing || candidate.confidence > existing.confidence) {
      byName.set(candidate.name, {
        name: candidate.name,
        type: inferFieldType(candidate.name),
        required: candidate.confidence >= 0.7,
        confidence: round(candidate.confidence, 2),
        reason: candidate.reason
      });
    }
  }
  return [...byName.values()].sort((left, right) => right.confidence - left.confidence || left.name.localeCompare(right.name)).slice(0, 24);
}

function buildTemplateMapSuggestion(
  schemaCandidates: TemplateSchemaCandidate[],
  placeholders: TemplatePlaceholderCandidate[],
  mapCandidates: DesignMapCandidate[]
): TemplateMapSuggestion {
  const mapping: Record<string, string> = {};
  for (const field of schemaCandidates) {
    const placeholder = placeholders.find((candidate) => candidate.field === field.name);
    const map = mapCandidates.find((candidate) => candidate.field === field.name);
    const stableObjectId = placeholder?.stableObjectId ?? map?.stableObjectId;
    if (stableObjectId) mapping[field.name] = stableObjectId;
  }
  return {
    schema: "officegen.template.map@1.2",
    mapping,
    confidence: round(schemaCandidates.reduce((sum, field) => sum + field.confidence, 0) / Math.max(1, schemaCandidates.length), 2),
    candidateCount: Object.keys(mapping).length
  };
}

function inferFieldType(name: string): TemplateSchemaCandidate["type"] {
  const tokens = name
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "")
    .split("_")
    .filter(Boolean);
  const tokenSet = new Set(tokens);

  if (/^(is|has|can|should|enabled|active)(_|$)/i.test(name)) return "boolean";
  const numberTokens = new Set(["amount", "price", "revenue", "sales", "count", "number", "score", "rate", "ratio", "percent", "total", "qty", "quantity"]);
  const dateTokens = new Set(["date", "day", "month", "year", "deadline"]);

  if (tokens.some((token) => numberTokens.has(token))) return "number";
  if (tokens.some((token) => dateTokens.has(token))) return "date";
  if ((tokenSet.has("created") || tokenSet.has("modified") || tokenSet.has("updated")) && tokenSet.has("at")) return "date";
  return "string";
}

function textToFieldName(text: string): string | undefined {
  const placeholder = placeholderFieldFromText(text);
  if (placeholder) return placeholder;
  const normalized = text.trim().replace(/\s+/g, " ");
  if (normalized.length < 2 || normalized.length > 80) return undefined;
  const label = normalized.replace(/[:：]$/, "");
  const field = label
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "");
  return field && /^[a-z]/.test(field) ? field : undefined;
}

function placeholderFieldFromText(text: string | undefined): string | undefined {
  const match = text?.match(/\{\{\s*([A-Za-z][\w.-]*)\s*\}\}/);
  return match?.[1]?.toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_+|_+$/g, "");
}

function fieldConfidence(text: string, placeholderType?: string): number {
  if (placeholderType) return 0.8;
  if (/\{\{[^}]+\}\}/.test(text)) return 0.9;
  if (/[:：]\s*$/.test(text)) return 0.75;
  if (text.length <= 32) return 0.6;
  return 0.35;
}

function isGenericShapeName(name: string): boolean {
  return /^(Title|Subtitle|TextBox|Text Box|Rectangle|Oval|Picture|Content Placeholder|Chart|Diagram)\s*\d*$/i.test(name.trim());
}

function firstXmlText(xml: string, localName: string): string | undefined {
  return extractXmlTexts(xml, localName)[0];
}

function stableSlideObjectId(sourcePath: string): string {
  return stableHashId("pptx", "deck", "slide", sourcePath);
}

function stableObjectId(kind: string, sourcePath: string, ordinal: number): string {
  return `pptx:${slideScope(sourcePath)}:${kind}:${String(ordinal).padStart(4, "0")}`;
}

function slideScope(sourcePath: string): string {
  return `slide-${stablePathToken(sourcePath)}`;
}

function stablePathToken(sourcePath: string): string {
  let hash = 2166136261;
  for (const char of sourcePath.toLowerCase()) {
    hash ^= char.charCodeAt(0);
    hash = Math.imul(hash, 16777619);
  }
  return (hash >>> 0).toString(16).padStart(8, "0").slice(0, 8);
}

function naturalSort(left: string, right: string): number {
  return left.localeCompare(right, undefined, { numeric: true, sensitivity: "base" });
}

async function writeSignalArtifacts(
  signals: PptxDesignSignals,
  sourcePath: string,
  options: PptxDesignSignalOptions
): Promise<PptxDesignArtifactPaths> {
  const baseDir = path.resolve(options.cwd ?? process.cwd(), options.artifactsDir ?? "");
  const previewPaths: string[] = [];
  for (const slide of signals.slideSignals.slice(0, 8)) {
    const previewPath = path.join(baseDir, `preview-slide-${String(slide.slide).padStart(3, "0")}.svg`);
    await writeTextFile(previewPath, makeSlidePreviewSvg(slide));
    previewPaths.push(previewPath);
    const preview = signals.previewCandidates.find((candidate) => candidate.slide === slide.slide);
    if (preview) preview.previewPath = previewPath;
    slide.previewPath = previewPath;
  }

  const evidencePath = await writeJsonFile(path.join(baseDir, "evidence.json"), {
    schema: "officegen.design.evidence@1.2",
    sourcePath,
    metadata: signals.metadata,
    colorRoleCandidates: signals.colorRoleCandidates,
    textSizeDistribution: signals.textSizeDistribution,
    bboxPatterns: signals.bboxPatterns,
    slideSignals: signals.slideSignals,
    placeholderCandidates: signals.placeholderCandidates,
    namedShapeCandidates: signals.namedShapeCandidates,
    schemaCandidates: signals.schemaCandidates,
    templateMapSuggested: signals.templateMapSuggested,
    trust: signals.trust
  });
  for (const slide of signals.slideSignals) slide.evidencePath = evidencePath;
  for (const preview of signals.previewCandidates) preview.evidencePath = evidencePath;

  const templateMapSuggestedPath = await writeJsonFile(path.join(baseDir, "template-map.suggested.json"), signals.templateMapSuggested);
  const schemaCandidatesPath = await writeJsonFile(path.join(baseDir, "schema-candidates.json"), {
    schema: "officegen.template.schema-candidates@1.2",
    fields: signals.schemaCandidates
  });
  const contextPath = await writeTextFile(path.join(baseDir, "context.md"), makeContextMarkdown(signals, sourcePath));
  return { contextPath, evidencePath, templateMapSuggestedPath, schemaCandidatesPath, previewPaths };
}

function makeContextMarkdown(signals: PptxDesignSignals, sourcePath: string): string {
  const lines = [
    "# PPTX design context",
    "",
    "## Trusted summary",
    "",
    `- Source: ${sourcePath}`,
    `- Slides: ${signals.metadata.slides}`,
    `- Assets: ${signals.metadata.assets}`,
    `- Macros: ${signals.metadata.macros}`,
    `- Density score: ${signals.densityScore}`,
    "",
    "## Candidate roles",
    "",
    ...signals.colorRoleCandidates.map((candidate) => `- ${candidate.role}: ${candidate.value} (${candidate.reason})`),
    "",
    "## Slide types",
    "",
    ...signals.slideSignals.map((slide) => `- Slide ${slide.slide}: ${slide.slideType}, density ${slide.densityScore}`),
    "",
    "## Template map suggestion",
    "",
    ...Object.entries(signals.templateMapSuggested.mapping).map(([field, stableObjectId]) => `- ${field}: ${stableObjectId}`),
    "",
    "## Untrusted deck excerpts",
    "",
    untrustedContentWarning,
    "",
    ...signals.trust.untrusted.textSamples.map((sample, index) => `> [${index + 1}] ${sample.replace(/\s+/g, " ").slice(0, 240)}`)
  ];
  return lines.join("\n");
}

function makeSlidePreviewSvg(slide: DesignSlideSignal): string {
  const rects = [
    slide.titleBounds ? previewRect(slide.titleBounds, "#2563EB", "title") : "",
    slide.bodyBounds ? previewRect(slide.bodyBounds, "#16A34A", "body") : "",
    ...slide.imageBounds.map((bounds) => previewRect(bounds, "#DC2626", "image"))
  ].filter(Boolean);
  return [
    '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 960 540" width="960" height="540">',
    '<rect width="960" height="540" fill="#FFFFFF" stroke="#CBD5E1"/>',
    `<text x="24" y="34" font-family="Arial, sans-serif" font-size="20" fill="#0F172A">Slide ${slide.slide}: ${escapeXml(slide.slideType)}</text>`,
    ...rects,
    "</svg>"
  ].join("");
}

function previewRect(bounds: DesignBounds, color: string, label: string): string {
  const x = bounds.x * 960;
  const y = bounds.y * 540;
  const width = bounds.width * 960;
  const height = bounds.height * 540;
  return `<g><rect x="${x}" y="${y}" width="${width}" height="${height}" fill="${color}" fill-opacity="0.12" stroke="${color}" stroke-width="2"/><text x="${x + 6}" y="${y + 20}" font-family="Arial, sans-serif" font-size="16" fill="${color}">${escapeXml(label)}</text></g>`;
}

function escapeXml(value: unknown): string {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function round(value: number, digits = 4): number {
  const factor = 10 ** digits;
  return Math.round(value * factor) / factor;
}

function designPath(context: OptionalContext, id: string): string {
  return path.join(featureRoot(context, "design"), `${slugify(id)}.json`);
}
