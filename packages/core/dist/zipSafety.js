import JSZip from "jszip";
import path from "node:path";
const LOCAL_FILE_HEADER_SIGNATURE = 0x04034b50;
const CENTRAL_DIRECTORY_SIGNATURE = 0x02014b50;
const END_OF_CENTRAL_DIRECTORY_SIGNATURE = 0x06054b50;
const ZIP64_END_OF_CENTRAL_DIRECTORY_LOCATOR_SIGNATURE = 0x07064b50;
const ZIP64_EXTRA_FIELD_ID = 0x0001;
function isTraversal(entryName) {
    const slashed = entryName.replace(/\\/g, "/");
    const normalized = path.posix.normalize(slashed);
    return normalized === ".." || normalized.startsWith("../") || path.posix.isAbsolute(normalized) || /^[A-Za-z]:\//.test(slashed);
}
function isZipLike(entryName) {
    return /\.(zip|docx|pptx|xlsx|xlsm|pptm|docm)$/i.test(entryName);
}
function isXmlLike(entryName) {
    return /\.(xml|rels)$/i.test(entryName);
}
function isMacro(entryName) {
    return /(^|\/)vbaProject\.bin$/i.test(entryName) || /\.(xlsm|docm|pptm)$/i.test(entryName);
}
function warning(code, message, entry, severity = "warning") {
    return { code, severity, message, entry };
}
function asUint8Array(input) {
    if (input instanceof ArrayBuffer)
        return new Uint8Array(input);
    return new Uint8Array(input.buffer, input.byteOffset, input.byteLength);
}
function canRead(view, offset, length) {
    return Number.isInteger(offset) && offset >= 0 && length >= 0 && offset + length <= view.byteLength;
}
function readUInt16(view, offset) {
    return view.getUint16(offset, true);
}
function readUInt32(view, offset) {
    return view.getUint32(offset, true);
}
function findEndOfCentralDirectory(view) {
    const minOffset = Math.max(0, view.byteLength - 22 - 0xffff);
    for (let offset = view.byteLength - 22; offset >= minOffset; offset -= 1) {
        if (readUInt32(view, offset) === END_OF_CENTRAL_DIRECTORY_SIGNATURE) {
            const commentLength = readUInt16(view, offset + 20);
            if (offset + 22 + commentLength === view.byteLength)
                return offset;
        }
    }
    return -1;
}
function decodeEntryName(bytes, utf8) {
    return new TextDecoder(utf8 ? "utf-8" : "latin1", { fatal: false }).decode(bytes);
}
function hasExtraField(bytes, fieldId) {
    let offset = 0;
    while (offset + 4 <= bytes.byteLength) {
        const headerId = bytes[offset] | ((bytes[offset + 1] ?? 0) << 8);
        const dataSize = (bytes[offset + 2] ?? 0) | ((bytes[offset + 3] ?? 0) << 8);
        offset += 4;
        if (offset + dataSize > bytes.byteLength)
            return false;
        if (headerId === fieldId)
            return true;
        offset += dataSize;
    }
    return false;
}
function hasZip64Locator(view, eocdOffset) {
    const locatorOffset = eocdOffset - 20;
    return canRead(view, locatorOffset, 4) && readUInt32(view, locatorOffset) === ZIP64_END_OF_CENTRAL_DIRECTORY_LOCATOR_SIGNATURE;
}
function hasBlockingError(report) {
    return report.warnings.some((item) => item.severity === "error" || item.severity === "critical");
}
function makeReport(scan) {
    return {
        ok: !scan.warnings.some((item) => item.severity === "error" || item.severity === "critical"),
        entryCount: scan.entryCount,
        expandedBytes: scan.expandedBytes,
        compressedBytes: scan.compressedBytes,
        hasMacros: scan.hasMacros,
        externalRelationships: scan.externalRelationships ?? [],
        warnings: scan.warnings
    };
}
function invalidCentralDirectory(message) {
    const warnings = [warning("ZIP_CENTRAL_DIRECTORY_INVALID", message, undefined, "error")];
    return {
        entries: [],
        report: makeReport({
            warnings,
            entryCount: 0,
            expandedBytes: 0,
            compressedBytes: 0,
            hasMacros: false
        })
    };
}
function scanZipPreload(input, config, options = {}) {
    const bytes = asUint8Array(input);
    if (bytes.byteLength < 22) {
        return invalidCentralDirectory("Archive is too small to contain a zip central directory.");
    }
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const eocdOffset = findEndOfCentralDirectory(view);
    if (eocdOffset < 0) {
        return invalidCentralDirectory("Archive is missing an end of central directory record.");
    }
    const limits = config.security.untrustedInput;
    const depth = options.depth ?? 0;
    const compressionRatioLimit = options.compressionRatioLimit ?? 100;
    const warnings = [];
    const entries = [];
    const names = new Set();
    let expandedBytes = 0;
    let compressedBytes = 0;
    let hasMacros = false;
    const diskNumber = readUInt16(view, eocdOffset + 4);
    const centralDirectoryDisk = readUInt16(view, eocdOffset + 6);
    const entriesOnDisk = readUInt16(view, eocdOffset + 8);
    const totalEntries = readUInt16(view, eocdOffset + 10);
    const centralDirectorySize = readUInt32(view, eocdOffset + 12);
    const centralDirectoryOffset = readUInt32(view, eocdOffset + 16);
    if (hasZip64Locator(view, eocdOffset) ||
        diskNumber === 0xffff ||
        centralDirectoryDisk === 0xffff ||
        entriesOnDisk === 0xffff ||
        totalEntries === 0xffff ||
        centralDirectorySize === 0xffffffff ||
        centralDirectoryOffset === 0xffffffff) {
        warnings.push(warning("ZIP64_UNSUPPORTED", "Zip64 archives are not accepted for untrusted Office inputs.", undefined, "error"));
    }
    if (diskNumber !== 0 || centralDirectoryDisk !== 0 || entriesOnDisk !== totalEntries) {
        warnings.push(warning("ZIP_CENTRAL_DIRECTORY_INVALID", "Split or inconsistent zip central directory records are not accepted.", undefined, "error"));
    }
    if (!canRead(view, centralDirectoryOffset, centralDirectorySize) || centralDirectoryOffset + centralDirectorySize > eocdOffset) {
        warnings.push(warning("ZIP_CENTRAL_DIRECTORY_INVALID", "Zip central directory points outside the archive.", undefined, "error"));
    }
    if (totalEntries > limits.maxZipEntries) {
        warnings.push(warning("ZIP_ENTRY_LIMIT_EXCEEDED", `Archive has ${totalEntries} entries.`, undefined, "error"));
    }
    let offset = centralDirectoryOffset;
    for (let index = 0; index < totalEntries; index += 1) {
        if (!canRead(view, offset, 46) || readUInt32(view, offset) !== CENTRAL_DIRECTORY_SIGNATURE) {
            warnings.push(warning("ZIP_CENTRAL_DIRECTORY_INVALID", `Invalid central directory entry at index ${index}.`, undefined, "error"));
            break;
        }
        const flags = readUInt16(view, offset + 8);
        const compressedSize = readUInt32(view, offset + 20);
        const uncompressedSize = readUInt32(view, offset + 24);
        const fileNameLength = readUInt16(view, offset + 28);
        const extraLength = readUInt16(view, offset + 30);
        const commentLength = readUInt16(view, offset + 32);
        const diskStart = readUInt16(view, offset + 34);
        const localHeaderOffset = readUInt32(view, offset + 42);
        const variableOffset = offset + 46;
        const nextOffset = variableOffset + fileNameLength + extraLength + commentLength;
        if (!canRead(view, variableOffset, fileNameLength + extraLength + commentLength)) {
            warnings.push(warning("ZIP_CENTRAL_DIRECTORY_INVALID", `Central directory entry at index ${index} is truncated.`, undefined, "error"));
            break;
        }
        const nameBytes = bytes.subarray(variableOffset, variableOffset + fileNameLength);
        const extraBytes = bytes.subarray(variableOffset + fileNameLength, variableOffset + fileNameLength + extraLength);
        const entryName = decodeEntryName(nameBytes, (flags & 0x0800) !== 0);
        const identityName = entryName.replace(/\\/g, "/");
        const hasZip64 = compressedSize === 0xffffffff ||
            uncompressedSize === 0xffffffff ||
            localHeaderOffset === 0xffffffff ||
            diskStart === 0xffff ||
            hasExtraField(extraBytes, ZIP64_EXTRA_FIELD_ID);
        if ((flags & 0x0001) !== 0) {
            warnings.push(warning("ZIP_ENCRYPTED_ENTRY", `Encrypted zip entry is not accepted: ${entryName}`, entryName, "error"));
        }
        if (hasZip64) {
            warnings.push(warning("ZIP64_UNSUPPORTED", `Zip64 entry metadata is not accepted: ${entryName}`, entryName, "error"));
        }
        if (names.has(identityName)) {
            warnings.push(warning("ZIP_DUPLICATE_ENTRY", `Duplicate zip entry name: ${entryName}`, entryName, "error"));
        }
        else {
            names.add(identityName);
        }
        if (isTraversal(entryName)) {
            warnings.push(warning("ZIP_PATH_TRAVERSAL", `Zip entry escapes extraction root: ${entryName}`, entryName, "error"));
        }
        const safeCompressedSize = compressedSize === 0xffffffff ? 0 : compressedSize;
        const safeUncompressedSize = uncompressedSize === 0xffffffff ? 0 : uncompressedSize;
        compressedBytes += safeCompressedSize;
        expandedBytes += safeUncompressedSize;
        if (safeUncompressedSize > 0 && safeCompressedSize > 0 && safeUncompressedSize / Math.max(1, safeCompressedSize) > compressionRatioLimit) {
            warnings.push(warning("ZIP_COMPRESSION_RATIO_EXCEEDED", `Zip entry compression ratio is too high: ${entryName}`, entryName, "error"));
        }
        if (isZipLike(entryName) && depth >= limits.maxNestedZipDepth) {
            warnings.push(warning("ZIP_NESTED_ZIP_DETECTED", `Nested archive exceeds max depth: ${entryName}`, entryName, "error"));
        }
        if (isMacro(entryName)) {
            hasMacros = true;
            warnings.push(warning("ZIP_MACRO_DETECTED", `Macro part detected: ${entryName}`, entryName, "warning"));
        }
        if (isXmlLike(entryName) && safeUncompressedSize > limits.maxSingleXmlPartBytes) {
            warnings.push(warning("ZIP_XML_PART_TOO_LARGE", `XML part exceeds safe size: ${entryName}`, entryName, "error"));
        }
        if (canRead(view, localHeaderOffset, 30) && readUInt32(view, localHeaderOffset) === LOCAL_FILE_HEADER_SIGNATURE) {
            const localFlags = readUInt16(view, localHeaderOffset + 6);
            const localCompressedSize = readUInt32(view, localHeaderOffset + 18);
            const localUncompressedSize = readUInt32(view, localHeaderOffset + 22);
            const localNameLength = readUInt16(view, localHeaderOffset + 26);
            const localExtraLength = readUInt16(view, localHeaderOffset + 28);
            const localExtraOffset = localHeaderOffset + 30 + localNameLength;
            const localExtraBytes = canRead(view, localExtraOffset, localExtraLength) ? bytes.subarray(localExtraOffset, localExtraOffset + localExtraLength) : new Uint8Array();
            if ((localFlags & 0x0001) !== 0) {
                warnings.push(warning("ZIP_ENCRYPTED_ENTRY", `Encrypted local zip entry is not accepted: ${entryName}`, entryName, "error"));
            }
            if (localCompressedSize === 0xffffffff || localUncompressedSize === 0xffffffff || hasExtraField(localExtraBytes, ZIP64_EXTRA_FIELD_ID)) {
                warnings.push(warning("ZIP64_UNSUPPORTED", `Zip64 local entry metadata is not accepted: ${entryName}`, entryName, "error"));
            }
        }
        else {
            warnings.push(warning("ZIP_CENTRAL_DIRECTORY_INVALID", `Central directory entry has an invalid local header offset: ${entryName}`, entryName, "error"));
        }
        entries.push({
            name: entryName,
            uncompressedSize: safeUncompressedSize
        });
        offset = nextOffset;
    }
    if (offset !== centralDirectoryOffset + centralDirectorySize && canRead(view, centralDirectoryOffset, centralDirectorySize)) {
        warnings.push(warning("ZIP_CENTRAL_DIRECTORY_INVALID", "Zip central directory size does not match parsed entries.", undefined, "error"));
    }
    if (expandedBytes > limits.maxZipExpandedBytes) {
        warnings.push(warning("ZIP_EXPANDED_BYTES_EXCEEDED", `Archive expands to ${expandedBytes} bytes.`, undefined, "error"));
    }
    return {
        entries,
        report: makeReport({
            warnings,
            entryCount: totalEntries,
            expandedBytes,
            compressedBytes,
            hasMacros
        })
    };
}
export function scanZipSafetyMetadata(input, config, options = {}) {
    return scanZipPreload(input, config, options).report;
}
export async function inspectZipSafety(input, config, options = {}) {
    const depth = options.depth ?? 0;
    const compressionRatioLimit = options.compressionRatioLimit ?? 100;
    const limits = config.security.untrustedInput;
    const preload = scanZipPreload(input, config, { depth, compressionRatioLimit });
    if (hasBlockingError(preload.report))
        return preload.report;
    const zip = options.preloadedZip ?? (await JSZip.loadAsync(input, { checkCRC32: false }));
    const warnings = [...preload.report.warnings];
    const expandedBytes = preload.report.expandedBytes;
    const compressedBytes = preload.report.compressedBytes;
    let relationshipCount = 0;
    const hasMacros = preload.report.hasMacros;
    const externalRelationships = [];
    const embeddedEntries = new Set(preload.entries
        .filter((entry) => /(^|\/)embeddings\//i.test(entry.name) && !entry.name.replace(/\\/g, "/").endsWith("/"))
        .map((entry) => entry.name.replace(/\\/g, "/")));
    const referencedEmbeddedEntries = new Set();
    const chartWorkbookEntries = new Set();
    for (const entry of preload.entries) {
        if (isXmlLike(entry.name)) {
            if (entry.uncompressedSize > limits.maxSingleXmlPartBytes) {
                if (/\.rels$/i.test(entry.name))
                    relationshipCount += 1;
                continue;
            }
            if (/\.rels$/i.test(entry.name)) {
                relationshipCount += 1;
            }
            const file = zip.file(entry.name);
            if (!file)
                continue;
            const xml = await file.async("string");
            if (limits.xmlExternalEntities === "deny" && /<!DOCTYPE|<!ENTITY/i.test(xml)) {
                warnings.push(warning("ZIP_XML_ENTITY_DENIED", `XML entity declaration denied: ${entry.name}`, entry.name, "error"));
            }
            if (/TargetMode\s*=\s*["']External["']/i.test(xml)) {
                externalRelationships.push(entry.name);
                warnings.push(warning("ZIP_EXTERNAL_RELATIONSHIP", `External relationship detected: ${entry.name}`, entry.name, "warning"));
            }
            if (/\.rels$/i.test(entry.name)) {
                for (const rel of relationshipRecords(xml)) {
                    if (/^https?:|^file:/i.test(rel.target) || rel.targetMode === "External")
                        continue;
                    const target = resolveRelationshipTarget(entry.name, rel.target);
                    if (!/(^|\/)embeddings\//i.test(target))
                        continue;
                    referencedEmbeddedEntries.add(target);
                    if (isPptxChartWorkbookRelationship(entry.name, rel, target)) {
                        chartWorkbookEntries.add(target);
                        warnings.push(warning("ZIP_CHART_EMBEDDED_WORKBOOK", `Editable PPTX chart workbook detected: ${target}`, target, "info"));
                    }
                    else if (!chartWorkbookEntries.has(target)) {
                        warnings.push(warning("ZIP_EMBEDDED_OBJECT", `Embedded object detected: ${target}`, target, "warning"));
                    }
                }
            }
            relationshipCount += (xml.match(/<Relationship\b/g) ?? []).length;
        }
    }
    for (const entryName of embeddedEntries) {
        if (referencedEmbeddedEntries.has(entryName))
            continue;
        if (chartWorkbookEntries.has(entryName))
            continue;
        warnings.push(warning("ZIP_EMBEDDED_OBJECT", `Embedded object detected: ${entryName}`, entryName, "warning"));
    }
    if (relationshipCount > limits.maxRelationships) {
        warnings.push(warning("ZIP_RELATIONSHIP_LIMIT_EXCEEDED", `Archive has ${relationshipCount} relationships.`, undefined, "error"));
    }
    return makeReport({
        warnings,
        entryCount: preload.report.entryCount,
        expandedBytes,
        compressedBytes,
        hasMacros,
        externalRelationships
    });
}
function relationshipRecords(xml) {
    return [...xml.matchAll(/<Relationship\b([^>]*)\/?>/g)].flatMap((match) => {
        const attrs = match[1] ?? "";
        const target = attrValue(attrs, "Target");
        if (!target)
            return [];
        return [{
                type: attrValue(attrs, "Type") ?? "",
                target,
                targetMode: attrValue(attrs, "TargetMode")
            }];
    });
}
function attrValue(attrs, name) {
    const match = new RegExp(`\\b${name}=["']([^"']*)["']`, "i").exec(attrs);
    return match?.[1];
}
function isPptxChartWorkbookRelationship(relsPath, rel, target) {
    return /^ppt\/charts\/_rels\/chart\d+\.xml\.rels$/i.test(relsPath) &&
        /\/officeDocument\/2006\/relationships\/package$/i.test(rel.type) &&
        /^ppt\/embeddings\/.+\.xlsx$/i.test(target);
}
function resolveRelationshipTarget(relsPath, target) {
    const normalizedTarget = target.replace(/\\/g, "/");
    if (normalizedTarget.startsWith("/"))
        return normalizeZipPath(normalizedTarget.slice(1));
    const ownerPart = relationshipOwnerPart(relsPath);
    const base = ownerPart.includes("/") ? ownerPart.slice(0, ownerPart.lastIndexOf("/")) : "";
    return normalizeZipPath(`${base ? `${base}/` : ""}${normalizedTarget}`);
}
function relationshipOwnerPart(relsPath) {
    if (relsPath === "_rels/.rels")
        return "";
    return relsPath.replace(/\/_rels\/([^/]+)\.rels$/i, "/$1");
}
function normalizeZipPath(pathValue) {
    const normalized = [];
    for (const part of pathValue.split("/")) {
        if (!part || part === ".")
            continue;
        if (part === "..")
            normalized.pop();
        else
            normalized.push(part);
    }
    return normalized.join("/");
}
//# sourceMappingURL=zipSafety.js.map