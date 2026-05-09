import { mkdir, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { getLoadedZipSafetyReport, loadZip, normalizeInput, readZipBytes, sha256, sortedZipFiles, writeOutput, zipSafetyCaveats, zipPathBasename, zipToBytes } from "./shared.js";
export async function inspectAsset(input) {
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
export async function extractAssets(input, options = {}) {
    const normalized = await normalizeInput(input, "unknown");
    const zip = await loadZip(normalized);
    const mediaPrefix = normalized.format === "pptx" ? "ppt/media/" : normalized.format === "docx" ? "word/media/" : normalized.format === "xlsx" ? "xl/media/" : "";
    if (!mediaPrefix)
        throw new Error(`Unsupported asset extraction format: ${normalized.format}`);
    const mediaPaths = sortedZipFiles(zip).filter((path) => path.startsWith(mediaPrefix));
    const assets = [];
    for (const path of mediaPaths) {
        const bytes = (await readZipBytes(zip, path)) ?? new Uint8Array();
        const outPath = options.outDir ? join(options.outDir, zipPathBasename(path)) : undefined;
        if (outPath) {
            await mkdir(options.outDir, { recursive: true });
            await writeFile(outPath, bytes);
        }
        const mediaType = detectMediaType(bytes, path);
        assets.push({
            schema: "officegen.asset.info@1.2",
            path,
            outPath,
            mediaType,
            byteLength: bytes.byteLength,
            sha256: sha256(bytes),
            ...detectDimensions(bytes, mediaType),
            trusted: false
        });
    }
    return {
        schema: "officegen.asset.extract.result@1.2",
        assets,
        caveats: ["Extracted assets are untrusted document content.", ...zipSafetyCaveats(getLoadedZipSafetyReport(zip))]
    };
}
export async function replaceAsset(input, options) {
    const normalized = await normalizeInput(input, "unknown");
    const zip = await loadZip(normalized);
    const target = zip.file(options.assetPath);
    if (!target)
        throw new Error(`Asset path not found: ${options.assetPath}`);
    zip.file(options.assetPath, options.replacement);
    const bytes = await zipToBytes(zip);
    await writeOutput(options.out, bytes);
    return {
        schema: "officegen.asset.replace.result@1.2",
        changed: true,
        out: options.out,
        bytes: options.out ? undefined : bytes,
        caveats: ["Replaced asset bytes are written verbatim.", ...zipSafetyCaveats(getLoadedZipSafetyReport(zip))]
    };
}
export const assetInspect = inspectAsset;
export const assetExtract = extractAssets;
export const assetReplace = replaceAsset;
function detectMediaType(bytes, path) {
    const ext = path?.split(".").pop()?.toLowerCase();
    if (bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47)
        return "image/png";
    if (bytes[0] === 0xff && bytes[1] === 0xd8)
        return "image/jpeg";
    if (String.fromCharCode(...bytes.slice(0, 120)).includes("<svg"))
        return "image/svg+xml";
    if (ext === "gif")
        return "image/gif";
    if (ext === "emf")
        return "image/x-emf";
    if (ext === "wmf")
        return "image/x-wmf";
    return "application/octet-stream";
}
function detectDimensions(bytes, mediaType) {
    if (mediaType === "image/png" && bytes.length >= 24) {
        const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
        return { width: view.getUint32(16), height: view.getUint32(20) };
    }
    if (mediaType === "image/jpeg") {
        let offset = 2;
        while (offset + 9 < bytes.length) {
            if (bytes[offset] !== 0xff)
                break;
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
//# sourceMappingURL=assets.js.map