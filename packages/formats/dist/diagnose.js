import { inspect } from "./inspect.js";
import { loadZip, normalizeInput, readZipText, sortedZipFiles } from "./shared.js";
export async function diagnose(input, options = {}) {
    const inspected = isInspectResult(input) ? input : await inspect(input, { depth: "shallow", config: options.config });
    const issues = [];
    const maxTextLength = options.maxTextLength ?? 220;
    for (const entry of inspected.objectMap) {
        const overflow = overflowRiskForEntry(entry, maxTextLength, inspected.trusted.format);
        if (overflow) {
            issues.push({
                code: "TEXT_OVERFLOW_RISK",
                severity: "warning",
                message: overflow.message,
                stableObjectId: entry.stableObjectId,
                location: objectLocation(entry),
                metrics: overflow.metrics,
                suggestedOps: [{ type: "setText", selector: { stableObjectId: entry.stableObjectId }, text: `${String(entry.text ?? entry.textPreview ?? "").slice(0, maxTextLength - 1)}…` }]
            });
        }
    }
    if (Number(inspected.trusted.summary.macros ?? 0) > 0) {
        issues.push({
            code: "MACRO_PRESENT",
            severity: "warning",
            message: "Document contains a VBA project. Treat the file as untrusted and avoid executing embedded code."
        });
    }
    if (!inspected.objectMap.length && inspected.trusted.format !== "pdf") {
        issues.push({
            code: "NO_TEXT_OBJECTS",
            severity: "info",
            message: "No editable text objects were detected by the XML inspector."
        });
    }
    if (!isInspectResult(input) && ["pptx", "docx", "xlsx"].includes(inspected.trusted.format)) {
        issues.push(...await officeRepairRiskIssues(input, inspected.trusted.format, options.config));
    }
    return {
        schema: "officegen.diagnose.result@1.2",
        issues,
        caveats: ["Diagnosis is based on approximate inspect/view data and does not execute external renderers."]
    };
}
export const diagnoseDocument = diagnose;
function overflowRiskForEntry(entry, maxTextLength, format) {
    const text = String(entry.text ?? entry.textPreview ?? "");
    if (!text.trim())
        return undefined;
    const semantic = asPlainRecord(entry.semantic);
    const paragraphs = Array.isArray(semantic.paragraphs) ? semantic.paragraphs.map(asPlainRecord) : [];
    const runs = paragraphs.flatMap((paragraph) => Array.isArray(paragraph.runs) ? paragraph.runs.map(asPlainRecord) : []);
    const fontSizes = runs.map((run) => Number(run.fontSizePt)).filter((size) => Number.isFinite(size) && size > 0);
    const fontSizePt = median(fontSizes) ?? (entry.kind === "shape" ? 18 : 11);
    const estimatedHeight = entry.bounds ? estimatedTextHeight(text, entry.bounds.width, fontSizePt, Math.max(1, paragraphs.length)) : undefined;
    if (format !== "pdf" && entry.bounds && estimatedHeight !== undefined && estimatedHeight > entry.bounds.height * 1.15) {
        return {
            message: "Estimated wrapped text height exceeds the object bounds in approximate layout analysis.",
            metrics: {
                textLength: text.length,
                paragraphCount: paragraphs.length,
                fontSizePt,
                bounds: entry.bounds,
                estimatedTextHeight: Number(estimatedHeight.toFixed(1))
            }
        };
    }
    if (text.length > maxTextLength) {
        return {
            message: "Text object is long enough to risk overflow in approximate rendering.",
            metrics: { textLength: text.length, maxTextLength, paragraphCount: paragraphs.length, fontSizePt }
        };
    }
    return undefined;
}
function objectLocation(entry) {
    return {
        slide: typeof entry.selectorHints?.slide === "number" ? entry.selectorHints.slide : undefined,
        page: typeof entry.selectorHints?.page === "number" ? entry.selectorHints.page : undefined,
        stableObjectId: entry.stableObjectId
    };
}
function asPlainRecord(value) {
    return value && typeof value === "object" && !Array.isArray(value) ? value : {};
}
function median(values) {
    const sorted = values.filter((value) => Number.isFinite(value)).sort((left, right) => left - right);
    if (!sorted.length)
        return undefined;
    return sorted[Math.floor(sorted.length / 2)];
}
function estimatedTextHeight(text, widthPx, fontSizePt, paragraphCount) {
    const fontPx = Math.max(8, fontSizePt * 1.333);
    const charsPerLine = Math.max(8, Math.floor(Math.max(48, widthPx) / Math.max(4.5, fontPx * 0.52)));
    const explicitLines = text.split(/\r?\n/).reduce((sum, line) => sum + Math.max(1, Math.ceil(line.length / charsPerLine)), 0);
    const lines = Math.max(explicitLines, paragraphCount);
    return lines * fontPx * 1.22 + Math.max(0, paragraphCount - 1) * fontPx * 0.28 + 8;
}
function isInspectResult(value) {
    return Boolean(value && typeof value === "object" && value.schema === "officegen.inspect.result@1.2");
}
async function officeRepairRiskIssues(input, format, config) {
    const normalized = await normalizeInput(input, format);
    const zip = await loadZip(normalized, { zipSafety: { config } });
    const paths = new Set(sortedZipFiles(zip));
    const issues = [];
    const required = format === "pptx"
        ? ["[Content_Types].xml", "_rels/.rels", "ppt/presentation.xml"]
        : format === "docx"
            ? ["[Content_Types].xml", "_rels/.rels", "word/document.xml"]
            : ["[Content_Types].xml", "_rels/.rels", "xl/workbook.xml"];
    for (const requiredPath of required) {
        if (!paths.has(requiredPath)) {
            issues.push({
                code: "OFFICE_REPAIR_RISK_MISSING_PART",
                severity: "error",
                message: `Required OOXML part is missing: ${requiredPath}. Office may show a repair dialog.`
            });
        }
    }
    for (const relsPath of [...paths].filter((path) => path.endsWith(".rels"))) {
        const relsXml = (await readZipText(zip, relsPath)) ?? "";
        const base = relationshipBase(relsPath);
        for (const match of relsXml.matchAll(/<Relationship\b([^>]*)\/>/g)) {
            const attrs = match[1] ?? "";
            const target = /\bTarget="([^"]+)"/.exec(attrs)?.[1];
            const mode = /\bTargetMode="([^"]+)"/.exec(attrs)?.[1];
            if (!target || mode === "External")
                continue;
            const resolved = normalizeZipTarget(base, target);
            if (!paths.has(resolved)) {
                issues.push({
                    code: "OFFICE_REPAIR_RISK_BROKEN_RELATIONSHIP",
                    severity: "warning",
                    message: `Relationship target is missing: ${relsPath} -> ${target}. Office may repair or drop the relationship.`
                });
            }
        }
    }
    return issues;
}
function relationshipBase(relsPath) {
    if (relsPath === "_rels/.rels")
        return "";
    return relsPath.replace(/\/_rels\/[^/]+\.rels$/, "");
}
function normalizeZipTarget(base, target) {
    const normalizedTarget = target.replace(/\\/g, "/");
    const packageAbsolute = normalizedTarget.startsWith("/");
    const parts = `${packageAbsolute || !base ? "" : `${base}/`}${packageAbsolute ? normalizedTarget.slice(1) : normalizedTarget}`.split("/");
    const normalized = [];
    for (const part of parts) {
        if (!part || part === ".")
            continue;
        if (part === "..")
            normalized.pop();
        else
            normalized.push(part);
    }
    return normalized.join("/");
}
//# sourceMappingURL=diagnose.js.map