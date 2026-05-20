import { inspect } from "./inspect.js";
import { render } from "./render.js";
import { inspectInputZipSafety, normalizeInput, writeOutput, zipSafetyCaveats } from "./shared.js";
import { embedPdfFonts, ensurePdfTextEncodable } from "./pdfFonts.js";
import { spawn } from "node:child_process";
import { mkdtemp, readFile, readdir, rm, writeFile } from "node:fs/promises";
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
        assertPdfMutationAllowed(normalized.bytes, "PDF export/normalization");
        const pdf = await PDFDocument.load(normalized.bytes);
        const bytes = await pdf.save({ useObjectStreams: false });
        await writeOutput(options.out, bytes);
        return result(normalized.format, options, bytes, ["PDF was normalized through pdf-lib."]);
    }
    if ((options.mode === "native" || options.mode === "proof") && options.to === "pdf") {
        assertNativeExportAllowed(options.config);
        return exportOfficeToPdfNative(normalized, options);
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
        assertPdfMutationAllowed(normalized.bytes, "PDF merge");
        const source = await PDFDocument.load(normalized.bytes);
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
    assertPdfMutationAllowed(normalized.bytes, "PDF split");
    const source = await PDFDocument.load(normalized.bytes);
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
    assertPdfMutationAllowed(normalized.bytes, "PDF reorder");
    const source = await PDFDocument.load(normalized.bytes);
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
export const MIN_NATIVE_RENDERER_TIMEOUT_MS = 1000;
export const DEFAULT_NATIVE_RENDERER_TIMEOUT_MS = 120000;
export function resolveNativeRendererTimeoutMs(timeoutMs) {
    if (timeoutMs === undefined)
        return DEFAULT_NATIVE_RENDERER_TIMEOUT_MS;
    if (!Number.isFinite(timeoutMs) || timeoutMs <= 0)
        return DEFAULT_NATIVE_RENDERER_TIMEOUT_MS;
    return Math.max(MIN_NATIVE_RENDERER_TIMEOUT_MS, Math.trunc(timeoutMs));
}
async function exportOfficeToPdfNative(input, options) {
    if (!input.path) {
        throw new OfficegenError("EXPORT_UNSUPPORTED", "Native Office-to-PDF export requires an input file path.");
    }
    if (!["pptx", "docx", "xlsx"].includes(input.format)) {
        throw new OfficegenError("EXPORT_UNSUPPORTED", `Native PDF export is not supported for ${input.format}.`);
    }
    const zipSafety = await inspectInputZipSafety(input, { config: options.config });
    const officeCom = await findOfficeComRenderer(input.format);
    if (officeCom) {
        return exportOfficeToPdfWithCom(input, options, officeCom, zipSafetyCaveats(zipSafety));
    }
    const executable = await findLibreOfficeExecutable();
    if (!executable) {
        throw new OfficegenError("EXPORT_UNSUPPORTED", `No Office COM renderer was available for ${input.format}, and LibreOffice/soffice was not found. Install Microsoft Office with COM automation, install LibreOffice, or use --mode fast for approximate PDF export.`);
    }
    const outDir = options.out ? path.dirname(options.out) : await mkdtemp(path.join(os.tmpdir(), "officegen-pdf-"));
    const cleanup = options.out ? undefined : outDir;
    const timeoutMs = resolveNativeRendererTimeoutMs(options.timeoutMs);
    try {
        await runProcess(executable, [
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            outDir,
            input.path
        ], timeoutMs);
        const generated = await findConvertedPdf(outDir, input.path);
        const pdf = await PDFDocument.load(await readFile(generated));
        const bytes = await pdf.save({ useObjectStreams: false });
        await writeOutput(options.out, bytes);
        return {
            schema: "officegen.export.result@1.2",
            from: input.format,
            to: options.to,
            mode: options.mode === "proof" ? "proof" : "native",
            out: options.out,
            bytes: options.out ? undefined : bytes,
            fidelity: "native",
            caveats: [
                "Converted with LibreOffice in headless mode; fidelity depends on installed fonts and LibreOffice filters.",
                ...zipSafetyCaveats(zipSafety)
            ],
            renderer: { id: "libreoffice", executable, status: "used", backend: "libreoffice", repairDialogExpected: false },
            nativeProof: {
                status: "passed",
                renderer: "libreoffice",
                artifact: options.out,
                reason: "Native proof rendered through LibreOffice headless conversion."
            }
        };
    }
    finally {
        if (cleanup)
            await rm(cleanup, { recursive: true, force: true });
    }
}
async function exportOfficeToPdfWithCom(input, options, renderer, caveats) {
    if (!input.path)
        throw new OfficegenError("EXPORT_UNSUPPORTED", "Office COM export requires an input file path.");
    const scriptDir = await mkdtemp(path.join(os.tmpdir(), "officegen-com-script-"));
    const reportDir = await mkdtemp(path.join(os.tmpdir(), "officegen-com-report-"));
    const outputDir = options.out ? undefined : await mkdtemp(path.join(os.tmpdir(), "officegen-com-pdf-"));
    const outPath = options.out ?? path.join(outputDir, `${path.basename(input.path, path.extname(input.path))}.pdf`);
    const reportPath = path.join(reportDir, "native-report.json");
    const scriptPath = path.join(scriptDir, "convert.ps1");
    try {
        await writeFile(scriptPath, officeComExportScript(input.format), "utf8");
        const result = await runProcessCapture("powershell.exe", [
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            scriptPath,
            "-InputPath",
            input.path,
            "-OutputPath",
            outPath,
            "-ReportPath",
            reportPath
        ], resolveNativeRendererTimeoutMs(options.timeoutMs));
        if (!result.ok) {
            throw new OfficegenError("EXPORT_UNSUPPORTED", `Office COM native export failed. ${result.stderr || result.stdout}`);
        }
        const bytes = await readFile(outPath);
        const pdf = await PDFDocument.load(bytes);
        const saved = await pdf.save({ useObjectStreams: false });
        await writeOutput(options.out, saved);
        const report = await readJsonLoose(reportPath);
        return {
            schema: "officegen.export.result@1.2",
            from: input.format,
            to: options.to,
            mode: options.mode === "proof" ? "proof" : "native",
            out: options.out,
            bytes: options.out ? undefined : saved,
            fidelity: "native",
            caveats: [
                `Converted with ${renderer.id} through Windows Office COM.`,
                "Repair-dialog status is inferred from COM open/export errors and repair-mode flags; a visible user dialog is not shown.",
                ...caveats
            ],
            renderer: {
                id: renderer.id,
                executable: renderer.executable,
                status: "used",
                backend: "office-com",
                repairDialogExpected: Boolean(report.repairDialogExpected)
            },
            nativeProof: {
                status: "passed",
                renderer: renderer.id === "powerpoint-com" ? "powerpoint" : "office-com",
                artifact: options.out,
                reason: `Native proof rendered through ${renderer.id}.`
            }
        };
    }
    finally {
        await rm(scriptDir, { recursive: true, force: true });
        await rm(reportDir, { recursive: true, force: true });
        if (outputDir)
            await rm(outputDir, { recursive: true, force: true });
    }
}
export async function nativeRendererDoctor(config) {
    const libreOffice = await findLibreOfficeExecutable();
    const office = await Promise.all(["pptx", "docx", "xlsx"].map((format) => findOfficeComRenderer(format)));
    const officeAvailable = office.filter((item) => Boolean(item));
    return {
        schema: "officegen.renderer.doctor@2.2",
        platform: process.platform,
        policy: {
            externalProcess: config?.security.externalProcess,
            renderers: config?.security.renderers
        },
        renderers: [
            {
                id: "libreoffice",
                backend: "libreoffice",
                available: Boolean(libreOffice),
                executable: libreOffice,
                formats: ["pptx", "docx", "xlsx"],
                message: libreOffice ? "LibreOffice headless conversion is available." : "LibreOffice/soffice was not found."
            },
            ...officeAvailable.map((renderer) => ({
                id: renderer.id,
                backend: "office-com",
                available: true,
                executable: renderer.executable,
                formats: [renderer.format],
                message: `${renderer.id} COM automation is available.`
            }))
        ]
    };
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
        const child = spawn(command, ["--version"], { stdio: "ignore", windowsHide: true, env: safeExternalEnv() });
        child.on("error", () => resolve(false));
        child.on("exit", (code) => resolve(code === 0));
    });
}
async function runProcess(command, args, timeoutMs = DEFAULT_NATIVE_RENDERER_TIMEOUT_MS) {
    return new Promise((resolve, reject) => {
        const child = spawn(command, args, { windowsHide: true, env: safeExternalEnv() });
        let stderr = "";
        let settled = false;
        let timer;
        const finish = (callback) => {
            if (settled)
                return;
            settled = true;
            clearTimeout(timer);
            callback();
        };
        timer = setTimeout(() => {
            child.kill();
            finish(() => reject(new OfficegenError("EXPORT_UNSUPPORTED", `LibreOffice conversion timed out after ${timeoutMs}ms.`)));
        }, timeoutMs);
        child.stderr.on("data", (chunk) => {
            stderr += String(chunk);
        });
        child.on("error", (error) => finish(() => reject(error)));
        child.on("exit", (code) => {
            finish(() => {
                if (code === 0)
                    resolve();
                else
                    reject(new OfficegenError("EXPORT_UNSUPPORTED", `LibreOffice conversion failed with exit code ${code}. ${stderr.trim()}`));
            });
        });
    });
}
function runProcessCapture(command, args, timeoutMs) {
    return new Promise((resolve) => {
        const child = spawn(command, args, { windowsHide: true, env: safeExternalEnv() });
        let stdout = "";
        let stderr = "";
        let settled = false;
        let timer;
        const finish = (value) => {
            if (settled)
                return;
            settled = true;
            clearTimeout(timer);
            resolve(value);
        };
        timer = setTimeout(() => {
            child.kill();
            finish({ ok: false, stdout, stderr: `${stderr}\nTimed out after ${timeoutMs}ms.`.trim() });
        }, timeoutMs);
        child.stdout.on("data", (chunk) => { stdout += String(chunk); });
        child.stderr.on("data", (chunk) => { stderr += String(chunk); });
        child.on("error", (error) => {
            finish({ ok: false, stdout, stderr: error.message });
        });
        child.on("exit", (code) => {
            finish({ ok: code === 0, stdout, stderr, code });
        });
    });
}
async function findOfficeComRenderer(format) {
    if (process.platform !== "win32")
        return undefined;
    const candidates = {
        pptx: { id: "powerpoint-com", progId: "PowerPoint.Application", executable: "POWERPNT.EXE" },
        docx: { id: "word-com", progId: "Word.Application", executable: "WINWORD.EXE" },
        xlsx: { id: "excel-com", progId: "Excel.Application", executable: "EXCEL.EXE" }
    };
    const candidate = candidates[format];
    if (!candidate)
        return undefined;
    const probe = await runProcessCapture("powershell.exe", [
        "-NoProfile",
        "-Command",
        `$ErrorActionPreference='Stop'; $app=New-Object -ComObject '${candidate.progId}'; $app.Quit(); 'ok'`
    ], 15000);
    return probe.ok ? { ...candidate, format } : undefined;
}
function officeComExportScript(format) {
    if (format === "pptx")
        return `
param([string]$InputPath,[string]$OutputPath,[string]$ReportPath)
$ErrorActionPreference='Stop'
$app=New-Object -ComObject PowerPoint.Application
$repair=$false
try {
  $presentation=$app.Presentations.Open($InputPath, $true, $false, $false)
  $presentation.SaveAs($OutputPath, 32)
  $presentation.Close()
} catch {
  $repair=$true
  throw
} finally {
  $app.Quit()
  @{ ok=$true; backend='office-com'; app='PowerPoint'; repairDialogExpected=$repair } | ConvertTo-Json | Set-Content -Encoding UTF8 $ReportPath
}`.trim();
    if (format === "docx")
        return `
param([string]$InputPath,[string]$OutputPath,[string]$ReportPath)
$ErrorActionPreference='Stop'
$app=New-Object -ComObject Word.Application
$app.Visible=$false
$repair=$false
try {
  $doc=$app.Documents.Open($InputPath, $false, $true, $false)
  $doc.ExportAsFixedFormat($OutputPath, 17)
  $doc.Close($false)
} catch {
  $repair=$true
  throw
} finally {
  $app.Quit()
  @{ ok=$true; backend='office-com'; app='Word'; repairDialogExpected=$repair } | ConvertTo-Json | Set-Content -Encoding UTF8 $ReportPath
}`.trim();
    return `
param([string]$InputPath,[string]$OutputPath,[string]$ReportPath)
$ErrorActionPreference='Stop'
$app=New-Object -ComObject Excel.Application
$app.Visible=$false
$app.DisplayAlerts=$false
$repair=$false
try {
  $wb=$app.Workbooks.Open($InputPath, 3, $false)
  $wb.RefreshAll()
  $app.CalculateFullRebuild()
  $wb.ExportAsFixedFormat(0, $OutputPath)
  $wb.Close($false)
} catch {
  $repair=$true
  throw
} finally {
  $app.Quit()
  @{ ok=$true; backend='office-com'; app='Excel'; repairDialogExpected=$repair } | ConvertTo-Json | Set-Content -Encoding UTF8 $ReportPath
}`.trim();
}
async function readJsonLoose(file) {
    try {
        return JSON.parse(await readFile(file, "utf8"));
    }
    catch {
        return {};
    }
}
function safeExternalEnv() {
    const allowed = [
        "PATHEXT",
        "ComSpec",
        "SystemRoot",
        "WINDIR",
        "TEMP",
        "TMP",
        "TMPDIR",
        "HOME",
        "USERPROFILE",
        "LANG",
        "LC_ALL",
        "LC_CTYPE"
    ];
    const env = {};
    const pathValue = process.platform === "win32" ? (process.env.Path ?? process.env.PATH) : (process.env.PATH ?? process.env.Path);
    if (pathValue !== undefined)
        env[process.platform === "win32" ? "Path" : "PATH"] = pathValue;
    for (const key of allowed) {
        const value = process.env[key];
        if (value !== undefined)
            env[key] = value;
    }
    return env;
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
    throw new OfficegenError("SECURITY_EXTERNAL_PROCESS_DENIED", "Native renderer export is disabled by the active configuration. This blocks PowerPoint/Word/Excel COM and LibreOffice renderers; set security.externalProcess to allow and security.renderers to enabled to use native export.", {
        externalProcess: config?.security.externalProcess ?? "deny",
        renderers: config?.security.renderers ?? "disabled"
    }, { feature: "renderer" });
}
function assertPdfMutationAllowed(bytes, operation) {
    if (!hasPdfEncryptEntry(bytes))
        return;
    throw new OfficegenError("EXPORT_UNSUPPORTED", `PDF_ENCRYPTED_BLOCKED: ${operation} is blocked for encrypted PDFs. Inspect can report PDF_ENCRYPTED risk flags, but mutation/export does not use ignoreEncryption by default.`, { operation }, { feature: "security" });
}
function hasPdfEncryptEntry(bytes) {
    return /\/Encrypt\b/.test(Buffer.from(bytes).toString("latin1"));
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