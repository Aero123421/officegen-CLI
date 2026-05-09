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
  const zip = await loadZip(normalized, { zipSafety: { config: options.config } });
  const mediaPrefix =
    normalized.format === "pptx" ? "ppt/media/" : normalized.format === "docx" ? "word/media/" : normalized.format === "xlsx" ? "xl/media/" : "";
  if (!mediaPrefix) throw new Error(`Unsupported asset extraction format: ${normalized.format}`);
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
  const caveats = [];
  if (expectedMediaType && currentMediaType !== "application/octet-stream" && currentMediaType !== expectedMediaType) {
    caveats.push(`Existing asset extension does not match content: ${options.assetPath} is ${currentMediaType}, expected ${expectedMediaType}.`);
  }
  if (expectedMediaType && replacementMediaType !== expectedMediaType && !options.allowMediaTypeChange) {
    throw new Error(`ASSET_UNSUPPORTED_FORMAT: replacement media type ${replacementMediaType} does not match ${options.assetPath} (${expectedMediaType}).`);
  }
  if (currentMediaType !== "application/octet-stream" && replacementMediaType !== currentMediaType && !options.allowMediaTypeChange) {
    throw new Error(`ASSET_UNSUPPORTED_FORMAT: replacement media type ${replacementMediaType} does not match existing asset type ${currentMediaType}.`);
  }
  zip.file(options.assetPath, options.replacement);
  const bytes = await zipToBytes(zip);
  await writeOutput(options.out, bytes);
  return {
    schema: "officegen.asset.replace.result@1.2",
    changed: true,
    out: options.out,
    bytes: options.out ? undefined : bytes,
    media: {
      assetPath: options.assetPath,
      existingMediaType: currentMediaType,
      replacementMediaType,
      expectedMediaType,
      replacementPath: options.replacementPath
    },
    caveats: [
      ...caveats,
      "Replaced asset bytes after media type validation; relationships and shape crop are preserved.",
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
