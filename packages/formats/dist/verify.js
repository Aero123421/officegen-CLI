import { mkdtemp, rm, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { diagnose } from "./diagnose.js";
import { exportDocument } from "./export.js";
import { inspect } from "./inspect.js";
import { normalizeInput } from "./shared.js";
import { view } from "./view.js";
import { PDFDocument } from "pdf-lib";
export async function verify(input, options = {}) {
    const normalized = await normalizeInput(input, "unknown");
    const artifacts = {};
    let openable = true;
    const warnings = [];
    const blockingIssues = [];
    const inspected = await inspect({ data: normalized.bytes, path: normalized.path, format: normalized.format }, { depth: "summary", config: options.config }).catch((error) => {
        openable = false;
        blockingIssues.push(error instanceof Error ? error.message : String(error));
        return undefined;
    });
    const diagnosed = inspected ? await diagnose({ data: normalized.bytes, path: normalized.path, format: normalized.format }, { config: options.config }) : undefined;
    for (const issue of diagnosed?.issues ?? []) {
        if (issue.severity === "error")
            blockingIssues.push(`${issue.code}: ${issue.message}`);
        if (issue.severity === "warning")
            warnings.push(`${issue.code}: ${issue.message}`);
    }
    const noRepairDialogExpected = ![...(diagnosed?.issues ?? [])].some((issue) => issue.code.startsWith("OFFICE_REPAIR_RISK"));
    const visual = options.visual && inspected
        ? await verifyVisual({ data: normalized.bytes, format: normalized.format }, options.config)
        : undefined;
    if (visual?.blankPages)
        warnings.push(`VISUAL_BLANK_PAGE: ${visual.blankPages} blank preview pages detected.`);
    const nativeRenderer = options.native
        ? await verifyNative(normalized, options, artifacts)
        : undefined;
    if (nativeRenderer && !nativeRenderer.ok)
        warnings.push(nativeRenderer.message ?? "Native renderer verification did not complete.");
    if (!openable)
        blockingIssues.push("INPUT_NOT_OPENABLE");
    const readiness = blockingIssues.length ? "blocked" : warnings.length ? "warning" : "pass";
    const score = Number(Math.max(0, 1 - blockingIssues.length * 0.35 - warnings.length * 0.08).toFixed(2));
    const result = {
        schema: "officegen.verify.result@1.2",
        readiness,
        score,
        format: normalized.format,
        openable,
        noRepairDialogExpected,
        nativeRenderer,
        visual,
        blockingIssues,
        warnings,
        artifacts
    };
    if (options.out)
        await writeFile(options.out, `${JSON.stringify(result, null, 2)}\n`, "utf8");
    return result;
}
async function verifyVisual(input, config) {
    const preview = await view(input, { format: "svg", maxPages: 10, config });
    const blankPages = preview.pages.filter((page) => !/<text\b|<rect\b|data-kind=/.test(page.content)).length;
    return { fidelity: "approximate", pagesChecked: preview.pages.length, blankPages };
}
async function verifyNative(input, options, artifacts) {
    if (!["pptx", "docx", "xlsx"].includes(input.format))
        return { attempted: false, ok: false, message: "Native renderer verification is only available for Office inputs." };
    if (!input.path)
        return { attempted: false, ok: false, message: "Native renderer verification requires an input file path." };
    const dir = await mkdtemp(path.join(os.tmpdir(), "officegen-verify-"));
    const pdfPath = path.join(dir, "native.pdf");
    try {
        const exported = await exportDocument(input.path, { to: "pdf", mode: "native", out: pdfPath, config: options.config });
        const pdf = await PDFDocument.load(await import("node:fs/promises").then((fs) => fs.readFile(pdfPath)), { ignoreEncryption: true });
        artifacts.nativePdf = pdfPath;
        return { attempted: true, ok: true, artifact: pdfPath, message: `Native renderer produced ${pdf.getPageCount()} PDF page(s) with ${exported.renderer?.id ?? "renderer"}.` };
    }
    catch (error) {
        await rm(dir, { recursive: true, force: true });
        return { attempted: true, ok: false, message: error instanceof Error ? error.message : String(error) };
    }
}
export const verifyDocument = verify;
//# sourceMappingURL=verify.js.map