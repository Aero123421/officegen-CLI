import { inspect } from "./inspect.js";
import { loadZip, normalizeInput, readZipText, sortedZipFiles } from "./shared.js";
export async function diagnose(input, options = {}) {
    const inspected = isInspectResult(input) ? input : await inspect(input, { depth: "shallow", config: options.config });
    const issues = [];
    const maxTextLength = options.maxTextLength ?? 220;
    for (const entry of inspected.objectMap) {
        if ((entry.text?.length ?? 0) > maxTextLength) {
            issues.push({
                code: "TEXT_OVERFLOW_RISK",
                severity: "warning",
                message: "Text object is long enough to risk overflow in approximate rendering.",
                stableObjectId: entry.stableObjectId,
                suggestedOps: [{ type: "setText", selector: { stableObjectId: entry.stableObjectId }, text: `${entry.text?.slice(0, maxTextLength - 1)}…` }]
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