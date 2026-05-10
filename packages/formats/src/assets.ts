import { mkdir, writeFile } from "node:fs/promises";
import { join } from "node:path";
import {
  type InputLike,
  type OfficegenConfig,
  getLoadedZipSafetyReport,
  loadZip,
  normalizeInput,
  readZipBytes,
  sha256,
  stableHashId,
  sortedZipFiles,
  writeOutput,
  zipSafetyCaveats,
  zipPathBasename,
  zipToBytes
} from "./shared.js";
import { OfficegenError } from "@officegen/core";
import { parseRelationships, relationshipTarget, type Relationship } from "./ooxml/relationships.js";

export interface AssetInfo {
  schema: "officegen.asset.info@1.2";
  source?: string;
  mediaType: string;
  byteLength: number;
  sha256: string;
  width?: number;
  height?: number;
  trusted: false;
}

export interface ExtractAssetsOptions {
  outDir?: string;
  images?: boolean;
  config?: OfficegenConfig;
}

export interface ExtractAssetsResult {
  schema: "officegen.asset.extract.result@1.2";
  assets: Array<AssetInfo & { path: string; outPath?: string }>;
  caveats: string[];
}

export interface EmbeddedAssetUsage {
  kind: "picture" | "chartEmbeddedWorkbook" | "worksheetDrawingImage" | "embeddedObject" | "relationship";
  partPath: string;
  relationshipId: string;
  relationshipType?: string;
  targetMode?: string;
  slide?: number;
  sheet?: number;
  story?: string;
}

export interface EmbeddedAssetInfo extends Omit<AssetInfo, "schema"> {
  schema: "officegen.asset.embedded.info@2.5";
  stableAssetId: string;
  zipPath: string;
  path: string;
  fileName: string;
  usageCount: number;
  usages: EmbeddedAssetUsage[];
  orphaned: boolean;
  replaceCommand: string;
  extractCommand: string;
  supportedActions: string[];
  limitation?: string;
  untrusted: true;
}

export interface InspectEmbeddedAssetsResult {
  schema: "officegen.asset.embedded.result@2.5";
  mode: "embedded";
  format: "pptx" | "docx" | "xlsx";
  trusted: {
    schema: "officegen.asset.embedded.trusted@2.5";
    format: "pptx" | "docx" | "xlsx";
    source?: string;
    summary: {
      assets: number;
      mediaAssets: number;
      embeddedObjects: number;
      usages: number;
      orphanedAssets: number;
      zipEntries: number;
    };
  };
  untrusted: {
    schema: "officegen.asset.embedded.untrusted@2.5";
    assets: EmbeddedAssetInfo[];
  };
  assets: EmbeddedAssetInfo[];
  caveats: string[];
}

export interface ReplaceAssetOptions {
  out?: string;
  assetPath: string;
  replacement: Uint8Array | Buffer;
  replacementPath?: string;
  allowMediaTypeChange?: boolean;
  config?: OfficegenConfig;
}

export async function inspectAsset(input: InputLike): Promise<AssetInfo> {
  const normalized = await normalizeInput(input, "unknown");
  const mediaType = detectMediaType(normalized.bytes, normalized.path);
  const dimensions = detectDimensions(normalized.bytes, mediaType);
  return {
    schema: "officegen.asset.info@1.2",
    source: normalized.path,
    mediaType,
    byteLength: normalized.bytes.byteLength,
    sha256: sha256(normalized.bytes),
    ...dimensions,
    trusted: false
  };
}

export async function inspectEmbeddedAssets(input: InputLike, options: ExtractAssetsOptions = {}): Promise<InspectEmbeddedAssetsResult> {
  const normalized = await normalizeInput(input, "unknown");
  const format = normalized.format;
  const mediaPrefix =
    format === "pptx" ? "ppt/media/" : format === "docx" ? "word/media/" : format === "xlsx" ? "xl/media/" : "";
  const embeddingPrefix =
    format === "pptx" ? "ppt/embeddings/" : format === "docx" ? "word/embeddings/" : format === "xlsx" ? "xl/embeddings/" : "";
  if (!mediaPrefix) {
    throw new OfficegenError("UNSUPPORTED_FORMAT", `Embedded asset inspection is not supported for ${format}.`, {
      format,
      supported: ["pptx", "docx", "xlsx"]
    });
  }
  const officeFormat = format as "pptx" | "docx" | "xlsx";
  const zip = await loadZip(normalized, { zipSafety: { config: options.config } });
  const files = sortedZipFiles(zip);
  const assetPaths = files.filter((filePath) => filePath.startsWith(mediaPrefix) || filePath.startsWith(embeddingPrefix));
  const usageMap = await collectEmbeddedAssetUsages(zip, files, assetPaths);
  const assets: EmbeddedAssetInfo[] = [];
  for (const zipPath of assetPaths) {
    const bytes = (await readZipBytes(zip, zipPath)) ?? new Uint8Array();
    const mediaType = detectMediaType(bytes, zipPath);
    const usages = usageMap.get(zipPath) ?? [];
    const stableAssetId = stableHashId(officeFormat, "package", zipPath.startsWith(embeddingPrefix) ? "embeddedObject" : "asset", zipPath);
    assets.push({
      schema: "officegen.asset.embedded.info@2.5",
      stableAssetId,
      source: normalized.path,
      zipPath,
      path: zipPath,
      fileName: zipPathBasename(zipPath),
      mediaType,
      byteLength: bytes.byteLength,
      sha256: sha256(bytes),
      ...detectDimensions(bytes, mediaType),
      usageCount: usages.length,
      usages,
      orphaned: usages.length === 0,
      replaceCommand: zipPath.startsWith(mediaPrefix) ? `officegen asset replace <office-file> --asset ${zipPath} <replacement> --out <output-file> --agent --json` : "",
      extractCommand: zipPath.startsWith(mediaPrefix) ? `officegen asset extract <office-file> --images --out .officegen/assets --agent --json` : "",
      supportedActions: zipPath.startsWith(mediaPrefix) ? ["extract", "replace"] : ["inspect-only"],
      ...(zipPath.startsWith(embeddingPrefix) ? { limitation: "Embedded OLE/package objects are inspect-only and blocked for mutation by default." } : {}),
      trusted: false,
      untrusted: true
    });
  }
  const mediaAssets = assets.filter((asset) => asset.zipPath.startsWith(mediaPrefix)).length;
  const embeddedObjects = assets.filter((asset) => asset.zipPath.startsWith(embeddingPrefix)).length;
  const usages = assets.reduce((sum, asset) => sum + asset.usageCount, 0);
  return {
    schema: "officegen.asset.embedded.result@2.5",
    mode: "embedded",
    format: officeFormat,
    trusted: {
      schema: "officegen.asset.embedded.trusted@2.5",
      format: officeFormat,
      source: normalized.path,
      summary: {
        assets: assets.length,
        mediaAssets,
        embeddedObjects,
        usages,
        orphanedAssets: assets.filter((asset) => asset.orphaned).length,
        zipEntries: files.length
      }
    },
    untrusted: {
      schema: "officegen.asset.embedded.untrusted@2.5",
      assets
    },
    assets,
    caveats: ["Embedded asset paths and relationship metadata are untrusted document content.", ...zipSafetyCaveats(getLoadedZipSafetyReport(zip))]
  };
}

export async function extractAssets(input: InputLike, options: ExtractAssetsOptions = {}): Promise<ExtractAssetsResult> {
  const normalized = await normalizeInput(input, "unknown");
  const mediaPrefix =
    normalized.format === "pptx" ? "ppt/media/" : normalized.format === "docx" ? "word/media/" : normalized.format === "xlsx" ? "xl/media/" : "";
  if (!mediaPrefix) {
    throw new OfficegenError("UNSUPPORTED_FORMAT", `Asset extraction is not supported for ${normalized.format}.`, {
      format: normalized.format,
      supported: ["pptx", "docx", "xlsx"]
    });
  }
  const zip = await loadZip(normalized, { zipSafety: { config: options.config } });
  const mediaPaths = sortedZipFiles(zip).filter((path) => path.startsWith(mediaPrefix));
  const assets = [];
  for (const path of mediaPaths) {
    const bytes = (await readZipBytes(zip, path)) ?? new Uint8Array();
    const outPath = options.outDir ? join(options.outDir, zipPathBasename(path)) : undefined;
    if (outPath) {
      await mkdir(options.outDir as string, { recursive: true });
      await writeFile(outPath, bytes);
    }
    const mediaType = detectMediaType(bytes, path);
    assets.push({
      schema: "officegen.asset.info@1.2" as const,
      path,
      outPath,
      mediaType,
      byteLength: bytes.byteLength,
      sha256: sha256(bytes),
      ...detectDimensions(bytes, mediaType),
      trusted: false as const
    });
  }
  return {
    schema: "officegen.asset.extract.result@1.2",
    assets,
    caveats: ["Extracted assets are untrusted document content.", ...zipSafetyCaveats(getLoadedZipSafetyReport(zip))]
  };
}

export async function replaceAsset(input: InputLike, options: ReplaceAssetOptions): Promise<{ schema: "officegen.asset.replace.result@1.2"; changed: boolean; out?: string; bytes?: Uint8Array; media: Record<string, unknown>; caveats: string[] }> {
  const normalized = await normalizeInput(input, "unknown");
  const zip = await loadZip(normalized, { zipSafety: { config: options.config } });
  const target = zip.file(options.assetPath);
  if (!target) throw new Error(`Asset path not found: ${options.assetPath}`);
  const currentBytes = (await readZipBytes(zip, options.assetPath)) ?? new Uint8Array();
  const currentMediaType = detectMediaType(currentBytes, options.assetPath);
  const replacementMediaType = detectMediaType(options.replacement, options.replacementPath ?? options.assetPath);
  const expectedMediaType = mediaTypeFromExtension(options.assetPath);
  const existingExtensionMismatch = Boolean(expectedMediaType && currentMediaType !== "application/octet-stream" && currentMediaType !== expectedMediaType);
  const replacementExt = extensionFromMediaType(replacementMediaType);
  const targetAssetPath = replacementExt && expectedMediaType !== replacementMediaType && (existingExtensionMismatch || options.allowMediaTypeChange)
    ? withExtension(options.assetPath, replacementExt)
    : options.assetPath;
  const caveats = [];
  if (existingExtensionMismatch) {
    caveats.push(`Existing asset extension does not match content: ${options.assetPath} is ${currentMediaType}, expected ${expectedMediaType}.`);
  }
  if (!replacementExt) {
    throw new OfficegenError("ASSET_UNSUPPORTED_FORMAT", `ASSET_UNSUPPORTED_FORMAT: unsupported replacement media type ${replacementMediaType}.`, {
      replacementMediaType,
      supported: ["image/png", "image/jpeg", "image/svg+xml", "image/gif"]
    });
  }
  if (
    currentMediaType !== "application/octet-stream" &&
    replacementMediaType !== currentMediaType &&
    expectedMediaType !== replacementMediaType &&
    targetAssetPath === options.assetPath &&
    !options.allowMediaTypeChange
  ) {
    throw new Error(`ASSET_UNSUPPORTED_FORMAT: replacement media type ${replacementMediaType} does not match existing asset type ${currentMediaType}.`);
  }
  if (targetAssetPath !== options.assetPath) {
    await rewriteRelationshipTargets(zip, options.assetPath, targetAssetPath);
    zip.remove(options.assetPath);
    caveats.push(`Updated media relationship target from ${options.assetPath} to ${targetAssetPath} to match ${replacementMediaType}.`);
  }
  await ensureMediaContentType(zip, targetAssetPath, replacementMediaType);
  zip.file(targetAssetPath, options.replacement);
  const bytes = await zipToBytes(zip);
  await writeOutput(options.out, bytes);
  return {
    schema: "officegen.asset.replace.result@1.2",
    changed: true,
    out: options.out,
    bytes: options.out ? undefined : bytes,
    media: {
      assetPath: options.assetPath,
      targetAssetPath,
      existingMediaType: currentMediaType,
      replacementMediaType,
      expectedMediaType,
      replacementPath: options.replacementPath
    },
    caveats: [
      ...caveats,
      "Replaced asset bytes after media type validation; relationship targets and content types are repaired when the replacement extension changes.",
      ...zipSafetyCaveats(getLoadedZipSafetyReport(zip))
    ]
  };
}

export const assetInspect = inspectAsset;
export const assetInspectEmbedded = inspectEmbeddedAssets;
export const assetExtract = extractAssets;
export const assetReplace = replaceAsset;

async function collectEmbeddedAssetUsages(zip: Awaited<ReturnType<typeof loadZip>>, files: string[], assetPaths: string[]): Promise<Map<string, EmbeddedAssetUsage[]>> {
  const assetSet = new Set(assetPaths);
  const usageMap = new Map<string, EmbeddedAssetUsage[]>();
  const ownerContext = await collectRelationshipOwnerContext(zip, files);
  for (const relsPath of files.filter((filePath) => filePath.endsWith(".rels"))) {
    const relsXml = await zip.file(relsPath)?.async("string");
    const ownerPart = relationshipOwnerPart(relsPath);
    const baseDir = ownerPart.includes("/") ? ownerPart.slice(0, ownerPart.lastIndexOf("/")) : "";
    for (const rel of parseRelationships(relsXml)) {
      if (/^https?:|^file:/i.test(rel.target) || rel.targetMode === "External") continue;
      const target = relationshipTarget(baseDir, rel.target);
      if (!assetSet.has(target)) continue;
      const usage = embeddedAssetUsage(ownerPart, rel, target, ownerContext.get(ownerPart));
      usageMap.set(target, [...(usageMap.get(target) ?? []), usage]);
    }
  }
  return usageMap;
}

async function collectRelationshipOwnerContext(zip: Awaited<ReturnType<typeof loadZip>>, files: string[]): Promise<Map<string, Partial<EmbeddedAssetUsage>>> {
  const context = new Map<string, Partial<EmbeddedAssetUsage>>();
  for (const relsPath of files.filter((filePath) => filePath.endsWith(".rels"))) {
    const relsXml = await zip.file(relsPath)?.async("string");
    const ownerPart = relationshipOwnerPart(relsPath);
    const baseDir = ownerPart.includes("/") ? ownerPart.slice(0, ownerPart.lastIndexOf("/")) : "";
    const sheet = Number(ownerPart.match(/^xl\/worksheets\/sheet(\d+)\.xml$/)?.[1]);
    for (const rel of parseRelationships(relsXml)) {
      if (/^https?:|^file:/i.test(rel.target) || rel.targetMode === "External") continue;
      const target = relationshipTarget(baseDir, rel.target);
      if (Number.isFinite(sheet) && sheet > 0 && target.startsWith("xl/drawings/")) {
        context.set(target, { sheet });
      }
    }
  }
  return context;
}

function embeddedAssetUsage(partPath: string, rel: Relationship, target: string, inherited?: Partial<EmbeddedAssetUsage>): EmbeddedAssetUsage {
  const slide = Number(partPath.match(/^ppt\/slides\/slide(\d+)\.xml$/)?.[1]);
  const sheet = Number(partPath.match(/^xl\/worksheets\/sheet(\d+)\.xml$/)?.[1]);
  const story = partPath.startsWith("word/")
    ? partPath.replace(/^word\//, "").replace(/\.xml$/i, "")
    : undefined;
  const isEmbedded = target.includes("/embeddings/") || /oleObject|package/i.test(rel.type);
  const isChartWorkbook = partPath.startsWith("ppt/charts/") && target.includes("/embeddings/");
  const kind: EmbeddedAssetUsage["kind"] = isChartWorkbook
    ? "chartEmbeddedWorkbook"
    : isEmbedded
      ? "embeddedObject"
      : partPath.startsWith("xl/drawings/")
        ? "worksheetDrawingImage"
        : partPath.startsWith("ppt/slides/")
          ? "picture"
          : "relationship";
  return {
    kind,
    partPath,
    relationshipId: rel.id,
    relationshipType: rel.type,
    targetMode: rel.targetMode,
    ...(Number.isFinite(slide) && slide > 0 ? { slide } : {}),
    ...(Number.isFinite(sheet) && sheet > 0 ? { sheet } : inherited?.sheet ? { sheet: inherited.sheet } : {}),
    ...(story ? { story } : {})
  };
}

function relationshipOwnerPart(relsPath: string): string {
  if (!relsPath.includes("_rels/")) return relsPath.replace(/\.rels$/i, "");
  return relsPath.replace(/\/_rels\/([^/]+)\.rels$/i, "/$1");
}

function detectMediaType(bytes: Uint8Array, path?: string): string {
  const ext = path?.split(".").pop()?.toLowerCase();
  if (bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47) return "image/png";
  if (bytes[0] === 0xff && bytes[1] === 0xd8) return "image/jpeg";
  if (String.fromCharCode(...bytes.slice(0, 120)).includes("<svg")) return "image/svg+xml";
  if (bytes.length >= 6 && (Buffer.from(bytes.slice(0, 6)).toString("ascii") === "GIF87a" || Buffer.from(bytes.slice(0, 6)).toString("ascii") === "GIF89a")) return "image/gif";
  if (ext === "emf" || ext === "wmf" || ext === "gif") return "application/octet-stream";
  return "application/octet-stream";
}

function detectDimensions(bytes: Uint8Array, mediaType: string): { width?: number; height?: number } {
  if (mediaType === "image/png" && bytes.length >= 24) {
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    return { width: view.getUint32(16), height: view.getUint32(20) };
  }
  if (mediaType === "image/jpeg") {
    let offset = 2;
    while (offset + 9 < bytes.length) {
      if (bytes[offset] !== 0xff) break;
      const marker = bytes[offset + 1];
      const length = (bytes[offset + 2] << 8) + bytes[offset + 3];
      if (marker >= 0xc0 && marker <= 0xc3) {
        return { height: (bytes[offset + 5] << 8) + bytes[offset + 6], width: (bytes[offset + 7] << 8) + bytes[offset + 8] };
      }
      offset += 2 + length;
    }
  }
  return {};
}

function mediaTypeFromExtension(path: string): string | undefined {
  const ext = path.split(".").pop()?.toLowerCase();
  if (ext === "png") return "image/png";
  if (ext === "jpg" || ext === "jpeg") return "image/jpeg";
  if (ext === "svg") return "image/svg+xml";
  if (ext === "gif") return "image/gif";
  if (ext === "emf") return "image/x-emf";
  if (ext === "wmf") return "image/x-wmf";
  return undefined;
}

function extensionFromMediaType(mediaType: string): string | undefined {
  if (mediaType === "image/png") return "png";
  if (mediaType === "image/jpeg") return "jpg";
  if (mediaType === "image/svg+xml") return "svg";
  if (mediaType === "image/gif") return "gif";
  return undefined;
}

function withExtension(path: string, extension: string): string {
  return path.replace(/\.[^/.]+$/, `.${extension}`);
}

async function rewriteRelationshipTargets(zip: Awaited<ReturnType<typeof loadZip>>, fromPath: string, toPath: string): Promise<void> {
  await Promise.all(Object.entries(zip.files).map(async ([relsPath, file]) => {
    if (file.dir || !relsPath.endsWith(".rels")) return;
    const xml = await file.async("string");
    const base = relationshipBase(relsPath);
    const next = xml.replace(/\bTarget="([^"]+)"/g, (match, target: string) => {
      const resolved = normalizeZipTarget(base, target);
      if (resolved !== fromPath) return match;
      return `Target="${target.replace(/[^/\\]+$/, toPath.split("/").pop() ?? "")}"`;
    });
    if (next !== xml) zip.file(relsPath, next);
  }));
}

async function ensureMediaContentType(zip: Awaited<ReturnType<typeof loadZip>>, assetPath: string, mediaType: string): Promise<void> {
  const extension = assetPath.split(".").pop()?.toLowerCase();
  if (!extension) return;
  const xml = (await zip.file("[Content_Types].xml")?.async("string")) ?? '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>';
  if (new RegExp(`<Default\\b[^>]*\\bExtension="${escapeRegExp(extension)}"[^>]*\\bContentType="${escapeRegExp(mediaType)}"`).test(xml)) return;
  if (new RegExp(`<Default\\b[^>]*\\bExtension="${escapeRegExp(extension)}"`).test(xml)) {
    zip.file("[Content_Types].xml", xml.replace(new RegExp(`<Default\\b([^>]*)\\bExtension="${escapeRegExp(extension)}"([^>]*)/>`), `<Default Extension="${extension}" ContentType="${mediaType}"/>`));
    return;
  }
  zip.file("[Content_Types].xml", xml.replace(/<\/Types>\s*$/, `<Default Extension="${extension}" ContentType="${mediaType}"/></Types>`));
}

function relationshipBase(relsPath: string): string {
  if (relsPath === "_rels/.rels") return "";
  return relsPath.replace(/\/_rels\/[^/]+\.rels$/, "");
}

function normalizeZipTarget(base: string, target: string): string {
  const normalizedTarget = target.replace(/\\/g, "/");
  const packageAbsolute = normalizedTarget.startsWith("/");
  const parts = `${packageAbsolute || !base ? "" : `${base}/`}${packageAbsolute ? normalizedTarget.slice(1) : normalizedTarget}`.split("/");
  const normalized: string[] = [];
  for (const part of parts) {
    if (!part || part === ".") continue;
    if (part === "..") normalized.pop();
    else normalized.push(part);
  }
  return normalized.join("/");
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}
