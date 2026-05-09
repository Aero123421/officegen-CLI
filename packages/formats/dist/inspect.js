import { AGENT_UNTRUSTED_INSTRUCTION, getLoadedZipSafetyReport, loadZip, makeStableObjectId, normalizeInput, readZipText, sortedZipFiles, trustedMeta, zipSafetyCaveats } from "./shared.js";
import { inspectParagraphs } from "./ooxml/docx.js";
import { inspectSlides } from "./ooxml/pptx.js";
import { inspectSheets } from "./ooxml/xlsx.js";
import { PDFDocument } from "pdf-lib";
export async function inspect(input, options = {}) {
    const normalized = await normalizeInput(input, options.format ?? "unknown");
    if (normalized.format === "pptx")
        return inspectPptx(normalized, options);
    if (normalized.format === "docx")
        return inspectDocx(normalized, options);
    if (normalized.format === "xlsx")
        return inspectXlsx(normalized, options);
    if (normalized.format === "pdf")
        return inspectPdf(normalized, options);
    throw new Error(`Unsupported inspect format: ${normalized.format}`);
}
export const inspectDocument = inspect;
export const inspectOfficeFile = inspect;
async function inspectPptx(input, options) {
    const zip = await loadZip(input, { zipSafety: { config: options.config } });
    const paths = sortedZipFiles(zip);
    const mediaPaths = paths.filter((path) => /^ppt\/media\//i.test(path));
    const { slides, objectMap } = await inspectSlides(zip);
    const summaryDepth = options.depth === "summary";
    const slidePayload = summaryDepth
        ? slides.map((slide) => ({
            stableObjectId: slide.stableObjectId,
            index: slide.index,
            sourcePath: slide.sourcePath,
            textPreview: slide.text.slice(0, 300),
            textObjectCount: slide.textObjects.length,
            shapeCount: slide.shapeCount,
            pictureCount: slide.pictureCount,
            untrusted: true
        }))
        : slides;
    const macros = paths.filter((path) => /vbaProject\.bin$/i.test(path));
    return {
        schema: "officegen.inspect.result@1.2",
        trusted: trustedMeta("officegen.inspect.result@1.2", input, {
            slides: slides.length,
            textObjects: objectMap.length,
            assets: mediaPaths.length,
            macros: macros.length,
            zipEntries: paths.length
        }, ["PPTX inspect is zip/XML based; animation and theme resolution are summarized only.", ...zipSafetyCaveats(getLoadedZipSafetyReport(zip))]),
        untrusted: {
            slides: slidePayload,
            assets: mediaPaths.map((path, index) => ({
                stableObjectId: makeStableObjectId("pptx", "deck", "asset", index + 1),
                path,
                fileName: path.split("/").pop(),
                untrusted: true
            })),
            ...(options.depth === "full" || options.include?.includes("rawPaths") ? { rawPaths: paths } : {})
        },
        objectMap: summaryDepth ? compactObjectMap(objectMap, 25) : objectMap,
        agentInstruction: AGENT_UNTRUSTED_INSTRUCTION
    };
}
async function inspectDocx(input, options) {
    const zip = await loadZip(input, { zipSafety: { config: options.config } });
    const paths = sortedZipFiles(zip);
    const { paragraphs, objectMap } = await inspectParagraphs(zip);
    const mediaPaths = paths.filter((path) => /^word\/media\//i.test(path));
    const macros = paths.filter((path) => /vbaProject\.bin$/i.test(path));
    const summaryDepth = options.depth === "summary";
    return {
        schema: "officegen.inspect.result@1.2",
        trusted: trustedMeta("officegen.inspect.result@1.2", input, {
            paragraphs: paragraphs.length,
            textObjects: objectMap.length,
            assets: mediaPaths.length,
            macros: macros.length,
            zipEntries: paths.length
        }, ["DOCX inspect reads main document XML; headers, footers, fields, and styles are summarized.", ...zipSafetyCaveats(getLoadedZipSafetyReport(zip))]),
        untrusted: {
            paragraphs: summaryDepth ? paragraphs.slice(0, 50).map((paragraph) => ({ ...paragraph, text: paragraph.text.slice(0, 300) })) : paragraphs,
            assets: mediaPaths.map((path, index) => ({
                stableObjectId: makeStableObjectId("docx", "body", "asset", index + 1),
                path,
                fileName: path.split("/").pop(),
                untrusted: true
            })),
            ...(options.depth === "full" || options.include?.includes("rawPaths") ? { rawPaths: paths } : {})
        },
        objectMap: summaryDepth ? compactObjectMap(objectMap, 25) : objectMap,
        agentInstruction: AGENT_UNTRUSTED_INSTRUCTION
    };
}
async function inspectXlsx(input, options) {
    const zip = await loadZip(input, { zipSafety: { config: options.config } });
    const paths = sortedZipFiles(zip);
    const { sheets, objectMap, sharedStrings } = await inspectSheets(zip);
    const macros = paths.filter((path) => /vbaProject\.bin$/i.test(path));
    const summaryDepth = options.depth === "summary";
    const worksheetXml = await Promise.all(paths
        .filter((path) => /^xl\/worksheets\/sheet\d+\.xml$/i.test(path))
        .map(async (path) => (await readZipText(zip, path)) ?? ""));
    const formulaCount = worksheetXml.reduce((count, xml) => count + (xml.match(/<f\b/g) ?? []).length, 0);
    const sheetSummaries = summaryDepth
        ? sheets.map((sheet) => ({
            stableObjectId: sheet.stableObjectId,
            index: sheet.index,
            sourcePath: sheet.sourcePath,
            cellCount: sheet.cells.length,
            usedRange: usedRangeFromCells(sheet.cells.map((cell) => cell.ref)),
            previewCells: sheet.cells.slice(0, 20).map((cell) => ({
                stableObjectId: cell.stableObjectId,
                ref: cell.ref,
                valuePreview: cell.value.slice(0, 120),
                sourcePath: cell.sourcePath,
                untrusted: true
            }))
        }))
        : sheets;
    return {
        schema: "officegen.inspect.result@1.2",
        trusted: trustedMeta("officegen.inspect.result@1.2", input, {
            sheets: sheets.length,
            cells: objectMap.length,
            sharedStrings: sharedStrings.length,
            formulas: formulaCount,
            tables: paths.filter((path) => /^xl\/tables\//i.test(path)).length,
            charts: paths.filter((path) => /^xl\/charts\//i.test(path)).length,
            macros: macros.length,
            zipEntries: paths.length
        }, ["XLSX inspect reads cached cell values; formulas, styles, and charts are summarized.", ...zipSafetyCaveats(getLoadedZipSafetyReport(zip))]),
        untrusted: {
            sheets: sheetSummaries,
            ...(options.depth === "full" || options.include?.includes("rawPaths") ? { rawPaths: paths } : {})
        },
        objectMap: summaryDepth ? compactObjectMap(objectMap, 50) : objectMap,
        agentInstruction: AGENT_UNTRUSTED_INSTRUCTION
    };
}
async function inspectPdf(input, options) {
    const pdf = await PDFDocument.load(input.bytes, { ignoreEncryption: true });
    const extractedText = extractPdfTextPreview(input.bytes);
    const summaryDepth = options.depth === "summary";
    const pages = pdf.getPages().map((page, index) => {
        const size = page.getSize();
        const text = extractedText.pages[index] ?? "";
        return {
            stableObjectId: makeStableObjectId("pdf", "document", "page", index + 1),
            index: index + 1,
            width: size.width,
            height: size.height,
            textPreview: text.slice(0, summaryDepth ? 300 : 1000),
            untrusted: true
        };
    });
    return {
        schema: "officegen.inspect.result@1.2",
        trusted: trustedMeta("officegen.inspect.result@1.2", input, { pages: pages.length, textBlocks: extractedText.pages.filter(Boolean).length, images: extractedText.imageRefs }, [
            "PDF text extraction is best-effort and does not replace OCR or native renderer output.",
            ...(extractedText.caveats.length ? extractedText.caveats : [])
        ]),
        untrusted: {
            pages,
            ...(summaryDepth ? {} : { text: extractedText.pages })
        },
        objectMap: pages
            .map((page, index) => extractedText.pages[index]
            ? {
                stableObjectId: makeStableObjectId("pdf", "document", "text", index + 1),
                kind: "pdfText",
                text: summaryDepth ? undefined : extractedText.pages[index],
                textPreview: extractedText.pages[index]?.slice(0, 240),
                selectorHints: { page: page.index },
                trust: { level: "untrusted", reason: "document-content" },
                untrusted: true
            }
            : undefined)
            .filter((entry) => Boolean(entry)),
        agentInstruction: AGENT_UNTRUSTED_INSTRUCTION
    };
}
function extractPdfTextPreview(bytes) {
    const raw = Buffer.from(bytes).toString("latin1");
    const pageChunks = raw.split(/\/Type\s*\/Page\b/g).slice(1);
    const chunks = pageChunks.length ? pageChunks : [raw];
    const pages = chunks.map((chunk) => extractPdfTextFromChunk(chunk).slice(0, 8000));
    const imageRefs = (raw.match(/\/Subtype\s*\/Image\b/g) ?? []).length;
    const caveats = [];
    if (/\/Filter\s*\/(?:FlateDecode|DCTDecode|JPXDecode|LZWDecode)/.test(raw)) {
        caveats.push("Some PDF streams are compressed or image-based; text preview may be incomplete.");
    }
    if (!pages.some(Boolean)) {
        caveats.push("No plain text operators were found; use OCR or native PDF tooling for scanned/compressed PDFs.");
    }
    return { pages, imageRefs, caveats };
}
function extractPdfTextFromChunk(chunk) {
    const strings = [];
    for (const match of chunk.matchAll(/\((?:\\.|[^\\)])*\)\s*Tj/g)) {
        strings.push(decodePdfLiteral(match[0].replace(/\)\s*Tj$/, "").slice(1)));
    }
    for (const match of chunk.matchAll(/\[((?:\s*\((?:\\.|[^\\)])*\)\s*-?\d*)+)\]\s*TJ/g)) {
        const array = match[1] ?? "";
        for (const item of array.matchAll(/\((?:\\.|[^\\)])*\)/g))
            strings.push(decodePdfLiteral(item[0].slice(1, -1)));
    }
    return strings.join(" ").replace(/\s+/g, " ").trim();
}
function decodePdfLiteral(value) {
    return value
        .replace(/\\([nrtbf()\\])/g, (_match, code) => ({ n: "\n", r: "\r", t: "\t", b: "\b", f: "\f", "(": "(", ")": ")", "\\": "\\" })[code] ?? code)
        .replace(/\\([0-7]{1,3})/g, (_match, octal) => String.fromCharCode(Number.parseInt(octal, 8)))
        .replace(/\)$/g, "");
}
function compactObjectMap(objectMap, limit) {
    return objectMap.slice(0, limit).map((entry) => ({
        stableObjectId: entry.stableObjectId,
        kind: entry.kind,
        label: entry.label,
        textPreview: entry.textPreview,
        selectorHints: entry.selectorHints,
        editableOps: entry.editableOps,
        media: entry.media,
        trust: entry.trust,
        untrusted: true
    }));
}
function usedRangeFromCells(refs) {
    const cells = refs
        .map((ref) => /^([A-Z]+)(\d+)$/i.exec(ref))
        .filter((match) => Boolean(match))
        .map((match) => ({ col: columnIndex(match[1] ?? "A"), row: Number(match[2]) }));
    if (!cells.length)
        return undefined;
    const minCol = Math.min(...cells.map((cell) => cell.col));
    const maxCol = Math.max(...cells.map((cell) => cell.col));
    const minRow = Math.min(...cells.map((cell) => cell.row));
    const maxRow = Math.max(...cells.map((cell) => cell.row));
    return `${columnName(minCol)}${minRow}:${columnName(maxCol)}${maxRow}`;
}
function columnIndex(name) {
    let value = 0;
    for (const char of name.toUpperCase())
        value = value * 26 + (char.charCodeAt(0) - 64);
    return value || 1;
}
function columnName(index) {
    let value = index;
    let name = "";
    while (value > 0) {
        value -= 1;
        name = String.fromCharCode(65 + (value % 26)) + name;
        value = Math.floor(value / 26);
    }
    return name || "A";
}
//# sourceMappingURL=inspect.js.map