import { inspect } from "./inspect.js";
import { view } from "./view.js";
import { exportDocument } from "./export.js";
import { loadZip, normalizeInput, sortedZipFiles } from "./shared.js";
import { PDFDocument } from "pdf-lib";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { createHash } from "node:crypto";
import { comparePngPixels } from "./visualDiff.js";
export async function diffDocuments(before, after, options = {}) {
    const beforeInspect = await inspect(before, { depth: "shallow", config: options.config });
    const afterInspect = await inspect(after, { depth: "shallow", config: options.config });
    const semantic = semanticDiff(beforeInspect, afterInspect);
    semantic.partChanges = await semanticPartDiff(before, after, beforeInspect.trusted.format, afterInspect.trusted.format);
    const visual = options.visual ? await visualDiff(before, after, beforeInspect, afterInspect, options) : undefined;
    const visualRegressionScore = visual?.pageScores.length
        ? Number((visual.pageScores.reduce((sum, page) => sum + page.score, 0) / visual.pageScores.length).toFixed(4))
        : undefined;
    const beforePages = pageLikeCount(beforeInspect);
    const afterPages = pageLikeCount(afterInspect);
    const pageCountChanged = beforePages !== afterPages;
    const changed = semantic.added.length > 0 || semantic.removed.length > 0 || semantic.changedText.length > 0 || semantic.changedGeometry.length > 0 || semantic.changedSemantic.length > 0 || (semantic.partChanges?.length ?? 0) > 0 || pageCountChanged || (visualRegressionScore ?? 0) > 0;
    return {
        schema: "officegen.diff.result@1.2",
        formatBefore: beforeInspect.trusted.format,
        formatAfter: afterInspect.trusted.format,
        changed,
        summary: {
            addedObjects: semantic.added.length,
            removedObjects: semantic.removed.length,
            changedTextObjects: semantic.changedText.length,
            changedGeometryObjects: semantic.changedGeometry.length,
            changedSemanticObjects: semantic.changedSemantic.length,
            beforePages,
            afterPages,
            pageCountChanged,
            changedParts: semantic.partChanges?.length ?? 0,
            visualRegressionScore
        },
        semantic,
        visual,
        caveats: [
            visualCaveat(visual),
            "StableObjectId matching is best-effort across generated files and preserves strongest value for edits within the same document lineage."
        ]
    };
}
function visualCaveat(visual) {
    if (visual?.status === "blocked")
        return `Visual diff was blocked: ${visual.message ?? "renderer unavailable"}.`;
    if (visual?.kind === "raster-pixel")
        return "Visual diff compares rasterized page pixels with deterministic thresholding and changed-pixel bounding boxes.";
    if (visual?.fidelity === "native")
        return "Native visual regression compares trusted renderer PDF outputs; fidelity depends on installed renderer filters and fonts.";
    if (visual?.renderer === "pdf-bytes")
        return "PDF visual diff without native renderer compares page-aware PDF byte windows; it avoids zero-content false negatives but is not raster fidelity.";
    return "Visual diff is based on officegen's approximate SVG/HTML view, not a native Office rasterization.";
}
function semanticDiff(before, after) {
    const beforeMap = new Map(before.objectMap.map((entry) => [entry.stableObjectId, entry]));
    const afterMap = new Map(after.objectMap.map((entry) => [entry.stableObjectId, entry]));
    const added = [...afterMap.values()].filter((entry) => !beforeMap.has(entry.stableObjectId));
    const removed = [...beforeMap.values()].filter((entry) => !afterMap.has(entry.stableObjectId));
    const changedText = [...beforeMap.entries()]
        .map(([stableObjectId, beforeEntry]) => {
        const afterEntry = afterMap.get(stableObjectId);
        if (!afterEntry || (beforeEntry.text ?? beforeEntry.textPreview) === (afterEntry.text ?? afterEntry.textPreview))
            return undefined;
        return {
            stableObjectId,
            kind: beforeEntry.kind,
            before: beforeEntry.text ?? beforeEntry.textPreview,
            after: afterEntry.text ?? afterEntry.textPreview
        };
    })
        .filter((entry) => Boolean(entry));
    const changedGeometry = [...beforeMap.entries()]
        .map(([stableObjectId, beforeEntry]) => {
        const afterEntry = afterMap.get(stableObjectId);
        const beforeBbox = normalizedBbox(beforeEntry);
        const afterBbox = afterEntry ? normalizedBbox(afterEntry) : undefined;
        if (!afterEntry)
            return undefined;
        if (!beforeBbox && !afterBbox)
            return undefined;
        if (beforeBbox && afterBbox && bboxEqual(beforeBbox, afterBbox))
            return undefined;
        return {
            stableObjectId,
            kind: beforeEntry.kind,
            beforeBbox,
            afterBbox,
            delta: geometryDelta(beforeBbox, afterBbox)
        };
    })
        .filter((entry) => Boolean(entry));
    for (const afterEntry of added) {
        const afterBbox = normalizedBbox(afterEntry);
        if (!afterBbox)
            continue;
        changedGeometry.push({
            stableObjectId: afterEntry.stableObjectId,
            kind: afterEntry.kind,
            beforeBbox: undefined,
            afterBbox,
            delta: geometryDelta(undefined, afterBbox)
        });
    }
    const changedSemantic = [...beforeMap.entries()]
        .map(([stableObjectId, beforeEntry]) => {
        const afterEntry = afterMap.get(stableObjectId);
        if (!afterEntry)
            return undefined;
        const beforeSemantic = comparableTextSemantic(beforeEntry);
        const afterSemantic = comparableTextSemantic(afterEntry);
        if (!beforeSemantic && !afterSemantic)
            return undefined;
        if (stableStringify(beforeSemantic) === stableStringify(afterSemantic))
            return undefined;
        return {
            stableObjectId,
            kind: beforeEntry.kind,
            changes: classifyTextSemanticChanges(beforeSemantic, afterSemantic),
            before: beforeSemantic,
            after: afterSemantic
        };
    })
        .filter((entry) => Boolean(entry));
    return { added, removed, changedText, changedGeometry, changedSemantic };
}
function comparableTextSemantic(entry) {
    const semantic = entry.semantic;
    if (semantic?.kind !== "pptxText" || !Array.isArray(semantic.paragraphs))
        return undefined;
    return {
        text: {
            paragraphSeparated: typeof semantic.text?.paragraphSeparated === "string" ? semantic.text.paragraphSeparated : entry.text,
            hasExplicitLineBreaks: Boolean(semantic.text?.hasExplicitLineBreaks),
            explicitLineBreakCount: Number(semantic.text?.explicitLineBreakCount ?? 0)
        },
        paragraphs: semantic.paragraphs.map((paragraph) => {
            const item = paragraph;
            return {
                index: item.index,
                text: item.text,
                level: item.level,
                bullet: item.bullet,
                numbering: item.numbering,
                runs: Array.isArray(item.runs)
                    ? item.runs.map((run) => {
                        const runItem = run;
                        return {
                            index: runItem.index,
                            text: runItem.text,
                            bold: runItem.bold,
                            italic: runItem.italic,
                            fontSizePt: runItem.fontSizePt,
                            fontFamilyLatin: runItem.fontFamilyLatin,
                            fontFamilyEastAsia: runItem.fontFamilyEastAsia,
                            fontFamilyComplexScript: runItem.fontFamilyComplexScript,
                            lang: runItem.lang,
                            noProof: runItem.noProof
                        };
                    })
                    : []
            };
        })
    };
}
function classifyTextSemanticChanges(before, after) {
    const changes = new Set();
    const beforeParagraphs = Array.isArray(before?.paragraphs) ? before.paragraphs : [];
    const afterParagraphs = Array.isArray(after?.paragraphs) ? after.paragraphs : [];
    if (beforeParagraphs.length !== afterParagraphs.length || stableStringify(before?.text) !== stableStringify(after?.text))
        changes.add("paragraph");
    const count = Math.max(beforeParagraphs.length, afterParagraphs.length);
    for (let index = 0; index < count; index += 1) {
        const beforeParagraph = beforeParagraphs[index];
        const afterParagraph = afterParagraphs[index];
        if (stableStringify(beforeParagraph?.bullet) !== stableStringify(afterParagraph?.bullet))
            changes.add("bullet");
        if (stableStringify(beforeParagraph?.numbering) !== stableStringify(afterParagraph?.numbering))
            changes.add("numbering");
        if (stableStringify(runFormats(beforeParagraph)) !== stableStringify(runFormats(afterParagraph)))
            changes.add("run-format");
    }
    return [...changes];
}
function runFormats(paragraph) {
    const runs = paragraph?.runs;
    if (!Array.isArray(runs))
        return [];
    return runs.map((run) => {
        const item = run;
        return {
            index: item.index,
            bold: item.bold,
            italic: item.italic,
            fontSizePt: item.fontSizePt,
            fontFamilyLatin: item.fontFamilyLatin,
            fontFamilyEastAsia: item.fontFamilyEastAsia,
            fontFamilyComplexScript: item.fontFamilyComplexScript,
            lang: item.lang,
            noProof: item.noProof
        };
    });
}
function stableStringify(value) {
    return JSON.stringify(value ?? null);
}
async function semanticPartDiff(before, after, formatBefore, formatAfter) {
    if (formatBefore !== formatAfter || !["pptx", "docx", "xlsx"].includes(formatBefore))
        return [];
    const beforeNormalized = await normalizeInput(before);
    const afterNormalized = await normalizeInput(after);
    const beforeZip = await loadZip(beforeNormalized);
    const afterZip = await loadZip(afterNormalized);
    const beforeParts = await packagePartHashes(beforeZip);
    const afterParts = await packagePartHashes(afterZip);
    const keys = new Set([...beforeParts.keys(), ...afterParts.keys()]);
    const changes = [];
    for (const key of [...keys].sort()) {
        const beforeHash = beforeParts.get(key);
        const afterHash = afterParts.get(key);
        if (beforeHash === afterHash)
            continue;
        changes.push({
            path: key,
            kind: classifyPackagePart(key),
            beforeHash,
            afterHash,
            status: beforeHash ? afterHash ? "changed" : "removed" : "added"
        });
    }
    return changes;
}
function normalizedBbox(entry) {
    if (entry.bbox && entry.bbox.length === 4)
        return entry.bbox;
    if (!entry.bounds)
        return undefined;
    return [entry.bounds.x, entry.bounds.y, entry.bounds.width, entry.bounds.height];
}
function bboxEqual(left, right) {
    return left.every((value, index) => Math.abs(value - (right[index] ?? 0)) < 0.01);
}
function geometryDelta(beforeBbox, afterBbox) {
    return {
        x: Number(((afterBbox?.[0] ?? 0) - (beforeBbox?.[0] ?? 0)).toFixed(2)),
        y: Number(((afterBbox?.[1] ?? 0) - (beforeBbox?.[1] ?? 0)).toFixed(2)),
        width: Number(((afterBbox?.[2] ?? 0) - (beforeBbox?.[2] ?? 0)).toFixed(2)),
        height: Number(((afterBbox?.[3] ?? 0) - (beforeBbox?.[3] ?? 0)).toFixed(2))
    };
}
function pageLikeCount(inspected) {
    const summary = inspected.trusted.summary;
    return Number(summary.pages ?? summary.slides ?? summary.sheets ?? (inspected.trusted.format === "docx" ? 1 : 0));
}
async function packagePartHashes(zip) {
    const map = new Map();
    const interesting = sortedZipFiles(zip).filter((file) => /\/charts\/chart\d+\.xml$/i.test(file) ||
        /\/embeddings\//i.test(file) ||
        /\/theme\/theme\d+\.xml$/i.test(file) ||
        /\/tables\/table\d+\.xml$/i.test(file) ||
        /\/media\//i.test(file) ||
        /\/comments.*\.xml$/i.test(file) ||
        /\/styles\.xml$/i.test(file));
    for (const file of interesting) {
        const entry = zip.file(file);
        if (!entry)
            continue;
        const bytes = await entry.async("uint8array");
        map.set(file, bytesHash(bytes));
    }
    return map;
}
function classifyPackagePart(file) {
    if (/\/charts\//i.test(file))
        return "chartXml";
    if (/\/embeddings\//i.test(file))
        return "embeddedWorkbook";
    if (/\/theme\//i.test(file))
        return "theme";
    if (/\/tables\//i.test(file))
        return "table";
    if (/\/media\//i.test(file))
        return "imageOrMedia";
    if (/\/comments/i.test(file))
        return "comments";
    if (/\/styles\.xml$/i.test(file))
        return "styles";
    return "packagePart";
}
function bytesHash(bytes) {
    return `sha256:${createHash("sha256").update(bytes).digest("hex")}`;
}
async function visualDiff(beforeInput, afterInput, before, after, options) {
    if (options.native)
        return nativeVisualDiff(beforeInput, afterInput, options);
    if (before.trusted.format === "pdf" && after.trusted.format === "pdf")
        return pdfByteVisualDiff(beforeInput, afterInput, options);
    const beforeView = await view(before, { format: "svg", maxPages: options.maxPages, config: options.config });
    const afterView = await view(after, { format: "svg", maxPages: options.maxPages, config: options.config });
    const pagesCompared = Math.min(beforeView.pages.length, afterView.pages.length);
    const pageScores = [];
    for (let index = 0; index < pagesCompared; index += 1) {
        const beforeHash = textHash(beforeView.pages[index]?.content ?? "");
        const afterHash = textHash(afterView.pages[index]?.content ?? "");
        pageScores.push({
            page: index + 1,
            beforeHash,
            afterHash,
            score: beforeHash === afterHash ? 0 : normalizedStringDistance(beforeView.pages[index]?.content ?? "", afterView.pages[index]?.content ?? "")
        });
    }
    return {
        status: "compared",
        kind: "approximate-string",
        fidelity: "approximate",
        pagesCompared,
        beforePages: beforeView.pages.length,
        afterPages: afterView.pages.length,
        pageCountChanged: beforeView.pages.length !== afterView.pages.length,
        pageScores
    };
}
async function pdfByteVisualDiff(beforeInput, afterInput, options) {
    const beforeNormalized = await normalizeInput(beforeInput);
    const afterNormalized = await normalizeInput(afterInput);
    const beforeDoc = await loadPdfForReporting(beforeNormalized.bytes);
    const afterDoc = await loadPdfForReporting(afterNormalized.bytes);
    const raster = await rasterPdfVisualDiff(beforeNormalized.bytes, afterNormalized.bytes, beforeDoc.getPageCount(), afterDoc.getPageCount(), "approximate", "pdfjs-canvas", options);
    if (raster.status === "compared")
        return raster;
    const pagesCompared = Math.min(beforeDoc.getPageCount(), afterDoc.getPageCount(), options.maxPages ?? Number.MAX_SAFE_INTEGER);
    const beforeWindows = byteWindows(beforeNormalized.bytes, pagesCompared);
    const afterWindows = byteWindows(afterNormalized.bytes, pagesCompared);
    const pageScores = [];
    for (let index = 0; index < pagesCompared; index += 1) {
        pageScores.push({
            page: index + 1,
            beforeHash: textHash(beforeWindows[index] ?? ""),
            afterHash: textHash(afterWindows[index] ?? ""),
            score: normalizedStringDistance(beforeWindows[index] ?? "", afterWindows[index] ?? "")
        });
    }
    return {
        status: "compared",
        kind: "pdf-byte-window",
        fidelity: "approximate",
        pagesCompared,
        beforePages: beforeDoc.getPageCount(),
        afterPages: afterDoc.getPageCount(),
        pageCountChanged: beforeDoc.getPageCount() !== afterDoc.getPageCount(),
        pageScores,
        renderer: "pdf-bytes",
        fallback: true,
        message: raster.message ? `Raster pixel diff unavailable; used deterministic byte-window fallback. ${raster.message}` : "Raster pixel diff unavailable; used deterministic byte-window fallback."
    };
}
async function nativeVisualDiff(beforeInput, afterInput, options) {
    const dir = await mkdtemp(path.join(os.tmpdir(), "officegen-diff-native-"));
    try {
        const beforePdf = path.join(dir, "before.pdf");
        const afterPdf = path.join(dir, "after.pdf");
        const beforeExport = await exportDocument(beforeInput, { to: "pdf", mode: "native", out: beforePdf, config: options.config });
        await exportDocument(afterInput, { to: "pdf", mode: "native", out: afterPdf, config: options.config });
        const beforeBytes = await readFile(beforePdf);
        const afterBytes = await readFile(afterPdf);
        const beforeDoc = await loadPdfForReporting(beforeBytes);
        const afterDoc = await loadPdfForReporting(afterBytes);
        const raster = await rasterPdfVisualDiff(beforeBytes, afterBytes, beforeDoc.getPageCount(), afterDoc.getPageCount(), "native", `${beforeExport.renderer?.id ?? "native"}+pdfjs-canvas`, options);
        if (raster.status === "compared")
            return raster;
        const pagesCompared = Math.min(beforeDoc.getPageCount(), afterDoc.getPageCount(), options.maxPages ?? Number.MAX_SAFE_INTEGER);
        const beforeText = byteWindows(beforeBytes, pagesCompared);
        const afterText = byteWindows(afterBytes, pagesCompared);
        const pageScores = [];
        for (let index = 0; index < pagesCompared; index += 1) {
            pageScores.push({
                page: index + 1,
                beforeHash: textHash(beforeText[index] ?? ""),
                afterHash: textHash(afterText[index] ?? ""),
                score: normalizedStringDistance(beforeText[index] ?? "", afterText[index] ?? "")
            });
        }
        return {
            status: "compared",
            kind: "pdf-byte-window",
            fidelity: "native",
            pagesCompared,
            beforePages: beforeDoc.getPageCount(),
            afterPages: afterDoc.getPageCount(),
            pageCountChanged: beforeDoc.getPageCount() !== afterDoc.getPageCount(),
            pageScores,
            renderer: `${beforeExport.renderer?.id ?? "native"}+pdf-bytes`,
            fallback: true,
            message: raster.message ? `Raster pixel diff unavailable; used deterministic byte-window fallback. ${raster.message}` : "Raster pixel diff unavailable; used deterministic byte-window fallback."
        };
    }
    catch (error) {
        return blockedVisualDiff("native", error instanceof Error ? error.message : String(error));
    }
    finally {
        await rm(dir, { recursive: true, force: true });
    }
}
async function rasterPdfVisualDiff(beforePdf, afterPdf, beforePages, afterPages, fidelity, renderer, options) {
    try {
        const [beforeView, afterView] = await Promise.all([
            view({ data: beforePdf, format: "pdf" }, { format: "png", maxPages: options.maxPages, config: options.config }),
            view({ data: afterPdf, format: "pdf" }, { format: "png", maxPages: options.maxPages, config: options.config })
        ]);
        const pagesCompared = Math.min(beforeView.pages.length, afterView.pages.length);
        const pageScores = [];
        for (let index = 0; index < pagesCompared; index += 1) {
            const beforeBytes = beforeView.pages[index]?.bytes ?? new Uint8Array();
            const afterBytes = afterView.pages[index]?.bytes ?? new Uint8Array();
            const pixelDiff = await comparePngPixels(beforeBytes, afterBytes, { threshold: options.pixelThreshold });
            if (pixelDiff.status !== "compared") {
                return blockedVisualDiff(fidelity, pixelDiff.message ?? "Raster page decode failed.", `${renderer}+blocked`);
            }
            pageScores.push({
                page: index + 1,
                beforeHash: pixelDiff.beforeHash,
                afterHash: pixelDiff.afterHash,
                score: pixelDiff.changedRatio,
                pixelDiff
            });
        }
        return {
            status: "compared",
            kind: "raster-pixel",
            fidelity,
            pagesCompared,
            beforePages,
            afterPages,
            pageCountChanged: beforePages !== afterPages,
            pageScores,
            renderer
        };
    }
    catch (error) {
        return blockedVisualDiff(fidelity, error instanceof Error ? error.message : String(error), renderer);
    }
}
function blockedVisualDiff(fidelity, message, renderer = "native") {
    return {
        status: "blocked",
        kind: "raster-pixel",
        fidelity,
        pagesCompared: 0,
        beforePages: 0,
        afterPages: 0,
        pageCountChanged: false,
        pageScores: [],
        renderer,
        message
    };
}
function loadPdfForReporting(bytes) {
    return PDFDocument.load(bytes, { ignoreEncryption: true });
}
function byteWindows(bytes, windows) {
    if (windows <= 0)
        return [];
    const chunkSize = Math.max(1, Math.ceil(bytes.length / windows));
    const chunks = [];
    for (let index = 0; index < windows; index += 1) {
        const chunk = bytes.subarray(index * chunkSize, Math.min(bytes.length, (index + 1) * chunkSize));
        chunks.push(Buffer.from(chunk).toString("latin1"));
    }
    return chunks;
}
function normalizedStringDistance(before, after) {
    if (!before && !after)
        return 0;
    const max = Math.max(before.length, after.length, 1);
    let changed = Math.abs(before.length - after.length);
    const limit = Math.min(before.length, after.length);
    for (let index = 0; index < limit; index += 1) {
        if (before.charCodeAt(index) !== after.charCodeAt(index))
            changed += 1;
    }
    return Number(Math.min(1, changed / max).toFixed(4));
}
function textHash(value) {
    return `sha256:${createHash("sha256").update(value).digest("hex")}`;
}
export const diff = diffDocuments;
//# sourceMappingURL=diff.js.map