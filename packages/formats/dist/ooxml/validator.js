import { XMLValidator } from "fast-xml-parser";
import JSZip from "jszip";
import { normalizeInput, readZipText, sortedZipFiles } from "../shared.js";
import { parseRelationships, relationshipTarget } from "./relationships.js";
export async function validateOoxml(input, options = {}) {
    const normalized = await normalizeInput(input, options.format ?? "unknown");
    const zip = await JSZip.loadAsync(normalized.bytes, { checkCRC32: false });
    const format = options.format ?? (isOoxmlFormat(normalized.format) ? normalized.format : "unknown");
    return validateOoxmlZip(zip, { format });
}
export async function validateOoxmlZip(zip, options = {}) {
    const paths = sortedZipFiles(zip);
    const pathSet = new Set(paths);
    const issues = [];
    const relationships = [];
    for (const requiredPath of ["[Content_Types].xml", "_rels/.rels"]) {
        if (!pathSet.has(requiredPath)) {
            issues.push({
                code: "OOXML_MISSING_REQUIRED_PART",
                severity: "error",
                path: requiredPath,
                message: `Required OPC part is missing: ${requiredPath}.`
            });
        }
    }
    for (const requiredPath of requiredFormatParts(options.format ?? "unknown")) {
        if (!pathSet.has(requiredPath)) {
            issues.push({
                code: "OOXML_MISSING_FORMAT_MAIN_PART",
                severity: "error",
                path: requiredPath,
                message: `Required ${options.format} part is missing: ${requiredPath}.`
            });
        }
    }
    const xmlPaths = paths.filter(isXmlPartPath);
    for (const path of xmlPaths) {
        const xml = await readZipText(zip, path);
        const validation = XMLValidator.validate(xml ?? "");
        if (validation !== true) {
            issues.push({
                code: "OOXML_XML_NOT_WELL_FORMED",
                severity: "error",
                path,
                line: validation.err.line,
                column: validation.err.col,
                message: `XML is not well-formed in ${path}: ${validation.err.code} ${validation.err.msg}`
            });
        }
    }
    const relsPaths = paths.filter((path) => path.endsWith(".rels"));
    for (const path of relsPaths) {
        const xml = await readZipText(zip, path);
        if (!xml || XMLValidator.validate(xml) !== true)
            continue;
        const seenIds = new Set();
        const duplicateIds = new Set();
        const baseDir = relationshipBaseDir(path);
        for (const rel of parseRelationships(xml)) {
            if (seenIds.has(rel.id) && !duplicateIds.has(rel.id)) {
                duplicateIds.add(rel.id);
                issues.push({
                    code: "OOXML_DUPLICATE_RELATIONSHIP_ID",
                    severity: "error",
                    path,
                    relationshipId: rel.id,
                    message: `Duplicate relationship id ${rel.id} in ${path}.`
                });
            }
            seenIds.add(rel.id);
            const external = isExternalRelationship(rel);
            const resolvedTarget = external ? undefined : resolveInternalRelationshipTarget(baseDir, rel.target);
            relationships.push({ ...rel, path, resolvedTarget, external });
            if (external) {
                issues.push({
                    code: "OOXML_EXTERNAL_RELATIONSHIP_TARGET",
                    severity: "warning",
                    path,
                    relationshipId: rel.id,
                    target: rel.target,
                    message: `Relationship ${rel.id} in ${path} targets an external resource.`
                });
                continue;
            }
            if (resolvedTarget && !pathSet.has(resolvedTarget)) {
                issues.push({
                    code: "OOXML_MISSING_INTERNAL_RELATIONSHIP_TARGET",
                    severity: "error",
                    path,
                    relationshipId: rel.id,
                    target: rel.target,
                    resolvedTarget,
                    message: `Relationship ${rel.id} in ${path} points to missing internal target ${resolvedTarget}.`
                });
            }
        }
    }
    const riskyParts = collectOoxmlRiskyParts(paths, relationships);
    const errors = issues.filter((issue) => issue.severity === "error").length;
    return {
        schema: "officegen.ooxml.validation@1",
        ok: errors === 0,
        format: options.format ?? "unknown",
        issues,
        riskyParts,
        relationships,
        summary: {
            entries: paths.length,
            xmlParts: xmlPaths.length,
            relationshipParts: relsPaths.length,
            riskyParts: riskyParts.length
        }
    };
}
function requiredFormatParts(format) {
    if (format === "pptx")
        return ["ppt/presentation.xml"];
    if (format === "docx")
        return ["word/document.xml"];
    if (format === "xlsx")
        return ["xl/workbook.xml"];
    return [];
}
export async function detectOoxmlRiskyParts(input, options = {}) {
    return (await validateOoxml(input, options)).riskyParts;
}
export function collectOoxmlRiskyParts(paths, relationships = []) {
    const riskyParts = [];
    for (const path of paths) {
        if (/(^|\/)vbaProject\.bin$/i.test(path)) {
            riskyParts.push({
                kind: "macro",
                path,
                message: `VBA project part detected: ${path}.`
            });
        }
        else if (/(^|\/)embeddings\/[^/]+$/i.test(path)) {
            riskyParts.push({
                kind: "embeddedObject",
                path,
                message: `Embedded object part detected: ${path}.`
            });
        }
    }
    for (const rel of relationships) {
        if (!isExternalRelationship(rel))
            continue;
        riskyParts.push({
            kind: "externalRelationship",
            path: rel.path,
            relationshipId: rel.id,
            target: rel.target,
            targetMode: rel.targetMode,
            message: `External relationship detected: ${rel.path} ${rel.id}.`
        });
    }
    return riskyParts;
}
function isOoxmlFormat(format) {
    return format === "pptx" || format === "docx" || format === "xlsx";
}
function isXmlPartPath(path) {
    return path.endsWith(".xml") || path.endsWith(".rels");
}
function relationshipBaseDir(relsPath) {
    if (relsPath === "_rels/.rels")
        return "";
    return relsPath.replace(/\/_rels\/[^/]+\.rels$/, "");
}
function resolveInternalRelationshipTarget(baseDir, target) {
    const cleanTarget = target.replace(/\\/g, "/").split("#", 1)[0]?.split("?", 1)[0] ?? "";
    if (!cleanTarget)
        return undefined;
    return relationshipTarget(baseDir, cleanTarget);
}
function isExternalRelationship(rel) {
    return /^external$/i.test(rel.targetMode ?? "") || /^[A-Za-z][A-Za-z0-9+.-]*:/.test(rel.target);
}
//# sourceMappingURL=validator.js.map