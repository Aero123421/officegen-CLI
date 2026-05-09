import { inspect } from "./inspect.js";
import { render } from "./render.js";
import { inspectInputZipSafety, normalizeInput, writeOutput, zipSafetyCaveats } from "./shared.js";
import { embedPdfFonts, ensurePdfTextEncodable } from "./pdfFonts.js";
import { spawn } from "node:child_process";
import { mkdtemp, readFile, readdir, rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { OfficegenError } from "../../core/dist/index.js";
import { PDFDocument, rgb } from "pdf-lib";
export async function exportDocument(input, options) {
    assertOutputExtensionMatches(options.to, options.out);
    if (typeof input === "object" && !("data" in input) && !("path" in input) && !isByteInput(input)) {
        const rendered = await render(input, { target: options.to, out: options.out, config: options.config });
        return {
            schema: "officegen.export.result@1.2",
            from: "ir",
            to: options.to,
            mode: options.mode ?? "fast",
            out: rendered.out,
            bytes: rendered.bytes instanceof Uint8Array ? rendered.bytes : rendered.bytes ? new Uint8Array(rendered.bytes) : undefined,
            fidelity: "internal",
            caveats: rendered.caveats
        };
    }
    const normalized = await normalizeInput(input, "unknown");
    if (normalized.format === "pdf" && options.to === "pdf") {
        const pdf = await PDFDocument.load(normalized.bytes, { ignoreEncryption: true });
        const bytes = await pdf.save({ useObjectStreams: false });
        await writeOutput(options.out, bytes);
        return result(normalized.format, options, bytes, ["PDF was normalized through pdf-lib."]);
    }
    if (options.mode === "native" && options.to === "pdf") {
        assertNativeExportAllowed(options.config);
        return exportOfficeToPdfWithLibreOffice(normalized, options);
    }
    if (options.to === "pdf") {
        const inspected = await inspect({ data: normalized.bytes, format: normalized.format }, { config: options.config });
        const pdf = await PDFDocument.create();
        const pages = normalized.format === "pptx"
            ? (inspected.untrusted.slides ?? [])
            : normalized.format === "docx"
                ? [{ title: "Document", text: (inspected.untrusted.paragraphs ?? []).map((p) => p.text).join("\n") }]
                : normalized.format === "xlsx"
                    ? (inspected.untrusted.sheets ?? []).map((sheet) => ({ title: `Sheet ${sheet.index}`, text: (sheet.cells ?? []).map((cell) => `${cell.ref}: ${cell.value}`).join("\n") }))
                    : [];
        const fontSet = await embedPdfFonts(pdf, pages.flatMap((pageInfo, index) => [String(pageInfo.title ?? `Page ${index + 1}`), String(pageInfo.text ?? "")]));
        for (const [index, pageInfo] of pages.entries()) {
            const page = pdf.addPage([612, 792]);
            page.drawText(ensurePdfTextEncodable(String(pageInfo.title ?? `Page ${index + 1}`), fontSet.bold, "export.pdf.title"), { x: 54, y: 735, size: 18, font: fontSet.bold, color: rgb(0.07, 0.07, 0.07) });
            const text = String(pageInfo.text ?? "");
            let y = 700;
            for (const line of text.split(/\r?\n/).slice(0, 36)) {
                page.drawText(ensurePdfTextEncodable(pdfSafeLine(line).slice(0, 95), fontSet.font, "export.pdf.body"), { x: 54, y, size: 10, font: fontSet.font, color: rgb(0.2, 0.2, 0.2) });
                y -= 16;
            }
        }
        const bytes = await pdf.save({ useObjectStreams: false });
        await writeOutput(options.out, bytes);
        return result(normalized.format, options, bytes, [
            "Fast Office-to-PDF export is approximate and text-summary based.",
            ...fontSet.caveats,
            ...inspected.trusted.caveats
        ]);
    }
    throw new OfficegenError("EXPORT_UNSUPPORTED", `Unsupported export: ${normalized.format} to ${options.to}`, {
        from: normalized.format,
        to: options.to
    });
}
function pdfSafeLine(value) {
    return value.replace(/\t/g, "    ").replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, " ");
}
export const exportFile = exportDocument;
export async function mergePdfs(inputs, options = {}) {
    assertOutputExtensionMatches("pdf", options.out);
    const output = await PDFDocument.create();
    for (const input of inputs) {
        const normalized = await normalizeInput(input, "pdf");
        const source = await PDFDocument.load(normalized.bytes, { ignoreEncryption: true });
        const pages = await output.copyPages(source, source.getPageIndices());
        for (const page of pages)
            output.addPage(page);
    }
    const bytes = await output.save({ useObjectStreams: false });
    await writeOutput(options.out, bytes);
    return result("pdf", { to: "pdf", out: options.out }, bytes, ["Merged PDFs with pdf-lib; outlines and advanced annotations may not be preserved."]);
}
export async function splitPdf(input, ranges, options = {}) {
    assertOutputExtensionMatches("pdf", options.out);
    const normalized = await normalizeInput(input, "pdf");
    const source = await PDFDocument.load(normalized.bytes, { ignoreEncryption: true });
    const results = [];
    for (const [rangeIndex, range] of ranges.entries()) {
        const output = await PDFDocument.create();
        const indices = range.map((page) => page - 1).filter((page) => page >= 0 && page < source.getPageCount());
        const pages = await output.copyPages(source, indices);
        for (const page of pages)
            output.addPage(page);
        const bytes = await output.save({ useObjectStreams: false });
        const out = options.out ? options.out.replace(/(\.pdf)?$/i, `.${rangeIndex + 1}.pdf`) : undefined;
        await writeOutput(out, bytes);
        results.push(result("pdf", { to: "pdf", out }, bytes, ["Split PDF with pdf-lib."]));
    }
    return results;
}
export async function reorderPdf(input, order, options = {}) {
    assertOutputExtensionMatches("pdf", options.out);
    const normalized = await normalizeInput(input, "pdf");
    const source = await PDFDocument.load(normalized.bytes, { ignoreEncryption: true });
    const output = await PDFDocument.create();
    const indices = order.map((page) => page - 1).filter((page) => page >= 0 && page < source.getPageCount());
    const pages = await output.copyPages(source, indices);
    for (const page of pages)
        output.addPage(page);
    const bytes = await output.save({ useObjectStreams: false });
    await writeOutput(options.out, bytes);
    return result("pdf", { to: "pdf", out: options.out }, bytes, ["Reordered PDF pages with pdf-lib."]);
}
function result(from, options, bytes, caveats) {
    return {
        schema: "officegen.export.result@1.2",
        from,
        to: options.to,
        mode: options.mode ?? "fast",
        out: options.out,
        bytes: options.out ? undefined : bytes,
        fidelity: "approximate",
        caveats
    };
}
async function exportOfficeToPdfWithLibreOffice(input, options) {
    if (!input.path) {
        throw new OfficegenError("EXPORT_UNSUPPORTED", "Native Office-to-PDF export requires an input file path.");
    }
    if (!["pptx", "docx", "xlsx"].includes(input.format)) {
        throw new OfficegenError("EXPORT_UNSUPPORTED", `Native PDF export is not supported for ${input.format}.`);
    }
    const zipSafety = await inspectInputZipSafety(input, { config: options.config });
    const executable = await findLibreOfficeExecutable();
    if (!executable) {
        throw new OfficegenError("EXPORT_UNSUPPORTED", "LibreOffice/soffice was not found. Install LibreOffice or use --mode fast for approximate PDF export.");
    }
    const outDir = options.out ? path.dirname(options.out) : await mkdtemp(path.join(os.tmpdir(), "officegen-pdf-"));
    const cleanup = options.out ? undefined : outDir;
    try {
        await runProcess(executable, [
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            outDir,
            input.path
        ]);
        const generated = await findConvertedPdf(outDir, input.path);
        const pdf = await PDFDocument.load(await readFile(generated), { ignoreEncryption: true });
        const bytes = await pdf.save({ useObjectStreams: false });
        await writeOutput(options.out, bytes);
        return {
            schema: "officegen.export.result@1.2",
            from: input.format,
            to: options.to,
            mode: "native",
            out: options.out,
            bytes: options.out ? undefined : bytes,
            fidelity: "native",
            caveats: [
                "Converted with LibreOffice in headless mode; fidelity depends on installed fonts and LibreOffice filters.",
                ...zipSafetyCaveats(zipSafety)
            ],
            renderer: { id: "libreoffice", executable, status: "used" }
        };
    }
    finally {
        if (cleanup)
            await rm(cleanup, { recursive: true, force: true });
    }
}
export async function findLibreOfficeExecutable() {
    const candidates = process.platform === "win32"
        ? [
            "soffice.exe",
            "libreoffice.exe",
            "C:\\Program Files\\LibreOffice\\program\\soffice.exe",
            "C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe"
        ]
        : ["soffice", "libreoffice"];
    for (const candidate of candidates) {
        if (await canRun(candidate))
            return candidate;
    }
    return undefined;
}
async function canRun(command) {
    return new Promise((resolve) => {
        const child = spawn(command, ["--version"], { stdio: "ignore", windowsHide: true });
        child.on("error", () => resolve(false));
        child.on("exit", (code) => resolve(code === 0));
    });
}
async function runProcess(command, args) {
    return new Promise((resolve, reject) => {
        const child = spawn(command, args, { windowsHide: true });
        let stderr = "";
        child.stderr.on("data", (chunk) => {
            stderr += String(chunk);
        });
        child.on("error", reject);
        child.on("exit", (code) => {
            if (code === 0)
                resolve();
            else
                reject(new OfficegenError("EXPORT_UNSUPPORTED", `LibreOffice conversion failed with exit code ${code}. ${stderr.trim()}`));
        });
    });
}
async function findConvertedPdf(outDir, inputPath) {
    const expected = path.join(outDir, `${path.basename(inputPath, path.extname(inputPath))}.pdf`);
    const files = await readdir(outDir);
    const match = files.find((file) => file.toLowerCase() === path.basename(expected).toLowerCase())
        ?? files.find((file) => file.toLowerCase().endsWith(".pdf"));
    if (!match)
        throw new OfficegenError("EXPORT_UNSUPPORTED", "LibreOffice did not produce a PDF.");
    return path.join(outDir, match);
}
function isByteInput(value) {
    return value instanceof Uint8Array || value instanceof ArrayBuffer;
}
function assertNativeExportAllowed(config) {
    if (config?.security.externalProcess === "allow" && config.security.renderers === "enabled")
        return;
    throw new OfficegenError("SECURITY_EXTERNAL_PROCESS_DENIED", "Native LibreOffice export is disabled by the active configuration. Set security.externalProcess to allow and security.renderers to enabled to use native export.", {
        externalProcess: config?.security.externalProcess ?? "deny",
        renderers: config?.security.renderers ?? "disabled"
    }, { feature: "renderer" });
}
function assertOutputExtensionMatches(target, out) {
    if (!out)
        return;
    const ext = path.extname(out).slice(1).toLowerCase();
    if (!ext || ext === target)
        return;
    throw new OfficegenError("TARGET_EXTENSION_MISMATCH", `Export target ${target} does not match output extension .${ext} for ${out}.`, { target, outputExtension: ext, out });
}
//# sourceMappingURL=export.js.map