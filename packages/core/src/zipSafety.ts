import JSZip from "jszip";
import path from "node:path";
import type { OfficegenConfig, ZipSafetyReport, ZipSafetyWarning } from "./types.js";

interface ZipPrivateData {
  compressedSize?: number;
  uncompressedSize?: number;
}

function privateSizes(file: JSZip.JSZipObject): { compressedSize: number; uncompressedSize: number } {
  const data = (file as unknown as { _data?: ZipPrivateData })._data;
  return {
    compressedSize: data?.compressedSize ?? 0,
    uncompressedSize: data?.uncompressedSize ?? 0
  };
}

function isTraversal(entryName: string): boolean {
  const normalized = path.posix.normalize(entryName.replace(/\\/g, "/"));
  return normalized === ".." || normalized.startsWith("../") || path.posix.isAbsolute(normalized);
}

function isZipLike(entryName: string): boolean {
  return /\.(zip|docx|pptx|xlsx|xlsm|pptm|docm)$/i.test(entryName);
}

function isXmlLike(entryName: string): boolean {
  return /\.(xml|rels)$/i.test(entryName);
}

function isMacro(entryName: string): boolean {
  return /(^|\/)vbaProject\.bin$/i.test(entryName) || /\.(xlsm|docm|pptm)$/i.test(entryName);
}

function warning(code: ZipSafetyWarning["code"], message: string, entry?: string, severity: ZipSafetyWarning["severity"] = "warning"): ZipSafetyWarning {
  return { code, severity, message, entry };
}

export async function inspectZipSafety(
  input: Buffer | Uint8Array | ArrayBuffer,
  config: OfficegenConfig,
  options: { depth?: number; compressionRatioLimit?: number } = {}
): Promise<ZipSafetyReport> {
  const depth = options.depth ?? 0;
  const compressionRatioLimit = options.compressionRatioLimit ?? 100;
  const limits = config.security.untrustedInput;
  const zip = await JSZip.loadAsync(input);
  const files = Object.values(zip.files).filter((file) => !file.dir);
  const warnings: ZipSafetyWarning[] = [];
  let expandedBytes = 0;
  let compressedBytes = 0;
  let relationshipCount = 0;
  let hasMacros = false;
  const externalRelationships: string[] = [];

  if (files.length > limits.maxZipEntries) {
    warnings.push(warning("ZIP_ENTRY_LIMIT_EXCEEDED", `Archive has ${files.length} entries.`, undefined, "error"));
  }

  for (const file of files) {
    const entryName = file.unsafeOriginalName ?? file.name;
    const sizes = privateSizes(file);
    expandedBytes += sizes.uncompressedSize;
    compressedBytes += sizes.compressedSize;

    if (isTraversal(entryName)) {
      warnings.push(warning("ZIP_PATH_TRAVERSAL", `Zip entry escapes extraction root: ${entryName}`, entryName, "error"));
    }
    if (sizes.uncompressedSize > 0 && sizes.compressedSize > 0 && sizes.uncompressedSize / Math.max(1, sizes.compressedSize) > compressionRatioLimit) {
      warnings.push(warning("ZIP_COMPRESSION_RATIO_EXCEEDED", `Zip entry compression ratio is too high: ${entryName}`, entryName, "error"));
    }
    if (isZipLike(entryName) && depth >= limits.maxNestedZipDepth) {
      warnings.push(warning("ZIP_NESTED_ZIP_DETECTED", `Nested archive exceeds max depth: ${entryName}`, entryName, "error"));
    }
    if (isMacro(entryName)) {
      hasMacros = true;
      warnings.push(warning("ZIP_MACRO_DETECTED", `Macro part detected: ${entryName}`, entryName, "warning"));
    }
    if (/embeddings\//i.test(entryName)) {
      warnings.push(warning("ZIP_EMBEDDED_OBJECT", `Embedded object detected: ${entryName}`, entryName, "warning"));
    }
    if (isXmlLike(entryName)) {
      if (sizes.uncompressedSize > limits.maxSingleXmlPartBytes) {
        warnings.push(warning("ZIP_XML_PART_TOO_LARGE", `XML part exceeds safe size: ${entryName}`, entryName, "error"));
      }
      if (/\.rels$/i.test(entryName)) {
        relationshipCount += 1;
      }
      const xml = await file.async("string");
      if (limits.xmlExternalEntities === "deny" && /<!DOCTYPE|<!ENTITY/i.test(xml)) {
        warnings.push(warning("ZIP_XML_ENTITY_DENIED", `XML entity declaration denied: ${entryName}`, entryName, "error"));
      }
      if (/TargetMode\s*=\s*["']External["']/i.test(xml)) {
        externalRelationships.push(entryName);
        warnings.push(warning("ZIP_EXTERNAL_RELATIONSHIP", `External relationship detected: ${entryName}`, entryName, "warning"));
      }
      relationshipCount += (xml.match(/<Relationship\b/g) ?? []).length;
    }
  }

  if (expandedBytes > limits.maxZipExpandedBytes) {
    warnings.push(warning("ZIP_EXPANDED_BYTES_EXCEEDED", `Archive expands to ${expandedBytes} bytes.`, undefined, "error"));
  }
  if (relationshipCount > limits.maxRelationships) {
    warnings.push(warning("ZIP_RELATIONSHIP_LIMIT_EXCEEDED", `Archive has ${relationshipCount} relationships.`, undefined, "error"));
  }

  return {
    ok: !warnings.some((item) => item.severity === "error" || item.severity === "critical"),
    entryCount: files.length,
    expandedBytes,
    compressedBytes,
    hasMacros,
    externalRelationships,
    warnings
  };
}
