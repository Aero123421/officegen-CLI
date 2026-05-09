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
  sortedZipFiles,
  writeOutput,
  zipSafetyCaveats,
  zipPathBasename,
  zipToBytes
} from "./shared.js";
import { OfficegenError } from "@officegen/core";

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
export const assetExtract = extractAssets;
export const assetReplace = replaceAsset;

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
