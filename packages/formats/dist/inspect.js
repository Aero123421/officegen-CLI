import { AGENT_UNTRUSTED_INSTRUCTION, getLoadedZipSafetyReport, loadZip, makeStableObjectId, normalizeInput, readZipText, sortedZipFiles, trustedMeta, zipSafetyCaveats } from "./shared.js";
import { inspectParagraphs } from "./ooxml/docx.js";
import { inspectSlides } from "./ooxml/pptx.js";
import { inspectSheets } from "./ooxml/xlsx.js";
import { PDFDocument } from "pdf-lib";
import { spawn } from "node:child_process";
import { mkdtemp, readdir, rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
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
    const themePaths = paths.filter((path) => /^ppt\/theme\/theme\d+\.xml$/i.test(path));
    const masterPaths = paths.filter((path) => /^ppt\/slideMasters\/slideMaster\d+\.xml$/i.test(path));
    const layoutPaths = paths.filter((path) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/i.test(path));
    const chartPaths = paths.filter((path) => /^ppt\/charts\/chart\d+\.xml$/i.test(path));
    const slidePayload = summaryDepth
        ? slides.map((slide) => ({
            stableObjectId: slide.stableObjectId,
            index: slide.index,
            sourcePath: slide.sourcePath,
            textPreview: slide.text.slice(0, 300),
            textObjectCount: slide.textObjects.length,
            shapeCount: slide.shapeCount,
            pictureCount: slide.pictureCount,
            chartCount: slide.chartCount,
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
            charts: chartPaths.length,
            masters: masterPaths.length,
            layouts: layoutPaths.length,
            themes: themePaths.length,
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
            designInventory: {
                themes: themePaths,
                masters: masterPaths,
                layouts: layoutPaths,
                charts: chartPaths,
                placeholders: objectMap
                    .filter((entry) => entry.selectorHints?.placeholder)
                    .slice(0, summaryDepth ? 40 : undefined)
                    .map((entry) => ({
                    stableObjectId: entry.stableObjectId,
                    slide: entry.selectorHints?.slide,
                    placeholder: entry.selectorHints?.placeholder,
                    label: entry.label,
                    untrusted: true
                }))
            },
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
    const headerPaths = paths.filter((path) => /^word\/header\d+\.xml$/i.test(path));
    const footerPaths = paths.filter((path) => /^word\/footer\d+\.xml$/i.test(path));
    const commentPaths = paths.filter((path) => /^word\/comments\.xml$/i.test(path));
    const stylePaths = paths.filter((path) => /^word\/styles\.xml$/i.test(path));
    const macros = paths.filter((path) => /vbaProject\.bin$/i.test(path));
    const summaryDepth = options.depth === "summary";
    return {
        schema: "officegen.inspect.result@1.2",
        trusted: trustedMeta("officegen.inspect.result@1.2", input, {
            paragraphs: paragraphs.length,
            textObjects: objectMap.length,
            assets: mediaPaths.length,
            headers: headerPaths.length,
            footers: footerPaths.length,
            comments: commentPaths.length,
            styles: stylePaths.length,
            macros: macros.length,
            zipEntries: paths.length
        }, ["DOCX inspect reads main document XML; headers, footers, fields, and styles are summarized.", ...zipSafetyCaveats(getLoadedZipSafetyReport(zip))]),
        untrusted: {
            paragraphs: summaryDepth ? paragraphs.slice(0, 50).map((paragraph) => ({ ...paragraph, text: paragraph.text.slice(0, 300) })) : paragraphs,
            documentParts: {
                headers: headerPaths,
                footers: footerPaths,
                comments: commentPaths,
                styles: stylePaths
            },
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
    const tablePaths = paths.filter((path) => /^xl\/tables\//i.test(path));
    const chartPaths = paths.filter((path) => /^xl\/charts\//i.test(path));
    const pivotPaths = paths.filter((path) => /^xl\/pivotTables\//i.test(path));
    const slicerPaths = paths.filter((path) => /^xl\/slicers\//i.test(path) || /^xl\/slicerCaches\//i.test(path));
    const definedNames = await readDefinedNames(zip);
    const cellCount = sheets.reduce((count, sheet) => count + sheet.cells.length, 0);
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
            cells: cellCount,
            sharedStrings: sharedStrings.length,
            formulas: formulaCount,
            tables: tablePaths.length,
            charts: chartPaths.length,
            pivotTables: pivotPaths.length,
            slicers: slicerPaths.length,
            definedNames: definedNames.length,
            macros: macros.length,
            zipEntries: paths.length
        }, ["XLSX inspect reads cached cell values; formulas, styles, and charts are summarized.", ...zipSafetyCaveats(getLoadedZipSafetyReport(zip))]),
        untrusted: {
            sheets: sheetSummaries,
            workbookObjects: {
                tables: tablePaths,
                charts: chartPaths,
                pivotTables: pivotPaths,
                slicers: slicerPaths,
                definedNames
            },
            ...(options.depth === "full" || options.include?.includes("rawPaths") ? { rawPaths: paths } : {})
        },
        objectMap: summaryDepth ? compactObjectMap(objectMap, 50) : objectMap,
        agentInstruction: AGENT_UNTRUSTED_INSTRUCTION
    };
}
async function inspectPdf(input, options) {
    const pdf = await PDFDocument.load(input.bytes, { ignoreEncryption: true });
    const extractedText = extractPdfTextPreview(input.bytes);
    const ocr = options.ocr && !extractedText.pages.some(Boolean)
        ? options.config?.security.externalProcess === "allow" && options.config.security.renderers === "enabled"
            ? await tryOcrPdf(input.path, pdf.getPageCount())
            : { ok: false, engine: "tesseract", message: "OCR requires security.externalProcess=allow and security.renderers=enabled.", pages: [] }
        : undefined;
    const textPages = ocr?.pages?.length ? ocr.pages : extractedText.pages;
    const summaryDepth = options.depth === "summary";
    const pages = pdf.getPages().map((page, index) => {
        const size = page.getSize();
        const text = textPages[index] ?? "";
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
            ...(ocr ? [`OCR ${ocr.ok ? "completed" : "not available"}: ${ocr.message}`] : []),
            ...(extractedText.caveats.length ? extractedText.caveats : [])
        ]),
        untrusted: {
            pages,
            ...(summaryDepth ? {} : { text: textPages }),
            ...(ocr ? { ocr } : {})
        },
        objectMap: pages
            .map((page, index) => textPages[index]
            ? {
                stableObjectId: makeStableObjectId("pdf", "document", "text", index + 1),
                kind: "pdfText",
                text: summaryDepth ? undefined : textPages[index],
                textPreview: textPages[index]?.slice(0, 240),
                selectorHints: { page: page.index },
                trust: { level: "untrusted", reason: "document-content" },
                untrusted: true
            }
            : undefined)
            .filter((entry) => Boolean(entry)),
        agentInstruction: AGENT_UNTRUSTED_INSTRUCTION
    };
}
async function tryOcrPdf(inputPath, pageCount) {
    if (!inputPath)
        return { ok: false, engine: "tesseract", message: "OCR requires a PDF input path.", pages: [] };
    const executable = await firstRunnable(["tesseract"]);
    if (!executable)
        return { ok: false, engine: "tesseract", message: "tesseract executable was not found.", pages: [] };
    const result = await runCapture(executable, [inputPath, "stdout", "-l", process.env.OFFICEGEN_OCR_LANG ?? "eng+jpn"], 30000);
    if (result.ok && result.stdout.trim()) {
        return { ok: true, engine: "tesseract", message: "OCR text extracted by Tesseract.", pages: result.stdout.split(/\f/g).map((page) => page.trim()) };
    }
    const rasterized = await tryRasterizedOcr(inputPath, pageCount, executable);
    if (rasterized.ok)
        return rasterized;
    return {
        ok: false,
        engine: "tesseract",
        message: [result.stderr || `direct OCR exit ${result.code}`, rasterized.message].filter(Boolean).join(" / "),
        pages: []
    };
}
async function tryRasterizedOcr(inputPath, pageCount, tesseract) {
    const pdftoppm = await firstRunnable(["pdftoppm"]);
    if (!pdftoppm)
        return { ok: false, engine: "tesseract+pdftoppm", message: "pdftoppm executable was not found for PDF raster OCR fallback.", pages: [] };
    const maxPages = Math.max(1, Number.parseInt(process.env.OFFICEGEN_OCR_MAX_PAGES ?? "20", 10) || 20);
    const lastPage = Math.min(Math.max(pageCount, 1), maxPages);
    const dir = await mkdtemp(path.join(os.tmpdir(), "officegen-ocr-"));
    try {
        const prefix = path.join(dir, "page");
        const raster = await runCapture(pdftoppm, ["-png", "-r", "200", "-f", "1", "-l", String(lastPage), inputPath, prefix], 60000);
        if (!raster.ok)
            return { ok: false, engine: "tesseract+pdftoppm", message: raster.stderr || `pdftoppm exit ${raster.code}`, pages: [] };
        const files = (await readdir(dir))
            .filter((file) => file.toLowerCase().endsWith(".png"))
            .sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
        const pages = [];
        for (const file of files) {
            const imagePath = path.join(dir, file);
            const ocr = await runCapture(tesseract, [imagePath, "stdout", "-l", process.env.OFFICEGEN_OCR_LANG ?? "eng+jpn"], 30000);
            pages.push(ocr.ok ? ocr.stdout.trim() : "");
        }
        const extracted = pages.some(Boolean);
        return {
            ok: extracted,
            engine: "tesseract+pdftoppm",
            message: extracted
                ? `OCR text extracted from ${pages.length} rasterized PDF page(s).${pageCount > lastPage ? ` Limited to ${lastPage}/${pageCount} pages by OFFICEGEN_OCR_MAX_PAGES.` : ""}`
                : "Raster OCR completed but produced no text.",
            pages
        };
    }
    finally {
        await rm(dir, { recursive: true, force: true });
    }
}
async function firstRunnable(commands) {
    for (const command of commands) {
        const result = await runCapture(command, ["--version"], 5000);
        if (result.ok)
            return command;
    }
    return undefined;
}
function runCapture(command, args, timeoutMs) {
    return new Promise((resolve) => {
        const child = spawn(command, args, { windowsHide: true });
        let stdout = "";
        let stderr = "";
        const timer = setTimeout(() => {
            child.kill();
            resolve({ ok: false, stdout, stderr: `${stderr}\nTimed out after ${timeoutMs}ms.`.trim() });
        }, timeoutMs);
        child.stdout.on("data", (chunk) => { stdout += String(chunk); });
        child.stderr.on("data", (chunk) => { stderr += String(chunk); });
        child.on("error", (error) => {
            clearTimeout(timer);
            resolve({ ok: false, stdout, stderr: error.message });
        });
        child.on("exit", (code) => {
            clearTimeout(timer);
            resolve({ ok: code === 0, code, stdout, stderr });
        });
    });
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
async function readDefinedNames(zip) {
    const workbookXml = (await readZipText(zip, "xl/workbook.xml")) ?? "";
    return [...workbookXml.matchAll(/<definedName\b([^>]*)>([\s\S]*?)<\/definedName>/g)]
        .map((match) => {
        const attrs = match[1] ?? "";
        const name = /\bname="([^"]+)"/.exec(attrs)?.[1] ?? "";
        const ref = (match[2] ?? "").replace(/&apos;/g, "'").replace(/&amp;/g, "&");
        return name ? { name, ref, untrusted: true } : undefined;
    })
        .filter((item) => Boolean(item));
}
//# sourceMappingURL=inspect.js.map