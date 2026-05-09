import path from "node:path";
import { readFile, stat } from "node:fs/promises";
import { edit, inspect } from "../../formats/dist/index.js";
import { featureRoot, hashFile, listJsonFiles, nowIso, readJsonFile, requireFeature, sha256Json, slugify, validation, writeJsonFile } from "./common.js";
import { capturePptxDesignSignals } from "./design.js";
export class TemplateFillError extends Error {
    details;
    constructor(message, details = {}) {
        super(message);
        this.name = "TemplateFillError";
        this.details = details;
    }
}
export async function createTemplate(options) {
    requireFeature(options, "template", "template create");
    const now = nowIso();
    const id = slugify(options.template.id);
    const sourcePath = options.sourcePath ?? options.template.source?.path;
    const resolvedSourcePath = sourcePath ? path.resolve(options.cwd ?? process.cwd(), sourcePath) : undefined;
    const sourceSignals = resolvedSourcePath
        ? await capturePptxDesignSignals(resolvedSourcePath, {
            cwd: options.cwd,
            artifactsDir: path.join(featureRoot(options, "template"), "captures", id, slugify(path.basename(resolvedSourcePath, path.extname(resolvedSourcePath))))
        })
        : undefined;
    const inferredFields = sourceSignals?.schemaCandidates.map((field) => ({
        name: field.name,
        type: field.type,
        fieldType: field.type,
        required: field.required,
        description: `${field.reason}; confidence ${field.confidence}`,
        confidence: field.confidence,
        reason: field.reason,
        ...fieldBindingMetadata(field.name, sourceSignals)
    }));
    const template = {
        ...options.template,
        id,
        fields: options.template.fields && options.template.fields.length > 0 ? options.template.fields : inferredFields ?? options.template.fields,
        mapping: normalizeTemplateMapping(options.template.mapping ?? sourceSignals?.templateMapSuggested),
        source: resolvedSourcePath
            ? {
                path: resolvedSourcePath,
                format: path.extname(resolvedSourcePath).replace(/^\./, "").toLowerCase() || options.template.source?.format,
                sha256: await hashFile(resolvedSourcePath)
            }
            : options.template.source,
        sourceCapture: sourceSignals
            ? {
                metadata: sourceSignals.metadata,
                artifactPaths: sourceSignals.artifactPaths,
                placeholderCandidates: sourceSignals.placeholderCandidates,
                namedShapeCandidates: sourceSignals.namedShapeCandidates,
                schemaCandidates: sourceSignals.schemaCandidates,
                templateMapSuggested: sourceSignals.templateMapSuggested,
                trust: sourceSignals.trust
            }
            : options.template.sourceCapture,
        createdAt: now,
        updatedAt: now
    };
    template.hash = sha256Json({ ...template, hash: undefined });
    assertTemplateValid(template);
    await writeJsonFile(templatePath(options, template.id), template);
    return template;
}
export async function listTemplates(options = {}) {
    requireFeature(options, "template", "template list");
    const files = await listJsonFiles(featureRoot(options, "template"));
    const templates = await Promise.all(files.map((file) => readJsonFile(file)));
    return templates.sort((left, right) => left.id.localeCompare(right.id));
}
export async function inspectTemplate(options) {
    requireFeature(options, "template", "template inspect");
    return readJsonFile(templatePath(options, options.id));
}
export async function templateCandidates(options = {}) {
    const templates = await listTemplates(options);
    const query = options.query?.trim().toLowerCase();
    const tags = new Set((options.tags ?? []).map((tag) => tag.toLowerCase()));
    const fields = new Set((options.fields ?? []).map((field) => field.toLowerCase()));
    const resolvedSourcePath = options.sourcePath ? path.resolve(options.cwd ?? process.cwd(), options.sourcePath) : undefined;
    const sourceSignals = resolvedSourcePath
        ? await capturePptxDesignSignals(resolvedSourcePath, {
            cwd: options.cwd,
            artifactsDir: path.join(featureRoot(options, "template"), "candidates", slugify(path.basename(resolvedSourcePath, path.extname(resolvedSourcePath))))
        })
        : undefined;
    const registryCandidates = templates
        .map((template) => {
        const reasons = [];
        let score = 0;
        const mapCandidates = sourceSignals ? matchTemplateMapCandidates(template, sourceSignals) : undefined;
        const searchable = [template.id, template.name, template.description, ...(template.tags ?? [])]
            .filter(Boolean)
            .join(" ")
            .toLowerCase();
        if (query && searchable.includes(query)) {
            score += 5;
            reasons.push("query");
        }
        const matchedTags = (template.tags ?? []).filter((tag) => tags.has(tag.toLowerCase()));
        if (matchedTags.length > 0) {
            score += matchedTags.length * 3;
            reasons.push(`tags:${matchedTags.join(",")}`);
        }
        const matchedFields = (template.fields ?? []).filter((field) => fields.has(field.name.toLowerCase()));
        if (matchedFields.length > 0) {
            score += matchedFields.length * 2;
            reasons.push(`fields:${matchedFields.map((field) => field.name).join(",")}`);
        }
        if (sourceSignals) {
            score += 1;
            reasons.push("source:pptx");
            if (mapCandidates && mapCandidates.length > 0) {
                score += mapCandidates.length;
                reasons.push(`map:${mapCandidates.map((candidate) => candidate.field).join(",")}`);
            }
        }
        if (!query && tags.size === 0 && fields.size === 0 && !sourceSignals) {
            score = 1;
            reasons.push("available");
        }
        return {
            template,
            score,
            reasons,
            ...(sourceSignals
                ? {
                    sourceMetadata: sourceSignals.metadata,
                    previewCandidates: sourceSignals.previewCandidates,
                    contextCandidates: sourceSignals.contextCandidates,
                    mapCandidates,
                    placeholderCandidates: sourceSignals.placeholderCandidates,
                    namedShapeCandidates: sourceSignals.namedShapeCandidates,
                    schemaCandidates: sourceSignals.schemaCandidates,
                    templateMapSuggested: sourceSignals.templateMapSuggested,
                    artifactPaths: sourceSignals.artifactPaths,
                    trust: sourceSignals.trust
                }
                : {})
        };
    })
        .filter((candidate) => candidate.score > 0);
    const sourceDerivedCandidate = sourceSignals && resolvedSourcePath ? makeSourceDerivedTemplateCandidate(resolvedSourcePath, sourceSignals) : undefined;
    return [...(sourceDerivedCandidate ? [sourceDerivedCandidate] : []), ...registryCandidates]
        .sort((left, right) => right.score - left.score || left.template.id.localeCompare(right.template.id));
}
export async function applyTemplateMap(options) {
    requireFeature(options, "template", "template apply-map");
    const template = await inspectTemplate(options);
    const updated = {
        ...template,
        mapping: normalizeTemplateMapping(options.mapping),
        updatedAt: nowIso()
    };
    updated.hash = sha256Json({ ...updated, hash: undefined });
    await writeJsonFile(templatePath(options, updated.id), updated);
    const plan = {
        kind: "officegen.template.apply-map",
        planOnly: true,
        generatedAt: nowIso(),
        templateId: updated.id,
        templateHash: updated.hash,
        mapping: updated.mapping,
        persisted: true,
        note: "Template mapping was persisted; template fill can mutate Office files when an Office --out path is supplied."
    };
    await writeJsonFile(options.outputPath ?? path.join(featureRoot(options, "template"), "runs", `${slugify(template.id)}.map.json`), plan);
    return plan;
}
export async function fillTemplate(options) {
    requireFeature(options, "template", "template fill");
    const template = await inspectTemplate(options);
    const bindings = templateBindingDiagnostics(template);
    const result = validateTemplateValues(template, options.values);
    if (!result.ok) {
        throw new TemplateFillError(`Template fill failed validation: ${result.errors.join("; ")}`, {
            errors: result.errors,
            bindings
        });
    }
    const sourcePath = template.source?.path ? path.resolve(options.cwd ?? process.cwd(), template.source.path) : undefined;
    const outputPath = options.outputPath;
    const outputFormat = outputPath ? path.extname(outputPath).replace(/^\./, "").toLowerCase() : "";
    const canMutateOffice = Boolean(sourcePath && outputPath && ["pptx", "docx", "xlsx"].includes(outputFormat));
    const ops = canMutateOffice ? await templateFillOperations(template, options.values, options.cwd) : [];
    const feasibility = await templateFillFeasibility(template, { sourcePath, outputPath, outputFormat, canMutateOffice, ops, cwd: options.cwd });
    if (options.validateOnly) {
        const resolver = canMutateOffice && sourcePath && ops.length
            ? await edit(sourcePath, ops, {
                format: outputFormat,
                dryRun: true,
                resolveSelectors: true,
                validateFirst: true,
                atomic: true
            }).catch((error) => ({ changed: false, errors: [error instanceof Error ? error.message : String(error)] }))
            : undefined;
        const resolverRecord = resolver;
        const resolverErrors = Array.isArray(resolverRecord?.errors) ? resolverRecord.errors : [];
        const resolverResolved = Array.isArray(resolverRecord?.resolvedSelectors) ? resolverRecord.resolvedSelectors : [];
        if (canMutateOffice && (!ops.length || resolverErrors.length > 0)) {
            throw new TemplateFillError("Template validate-only failed: one or more bindings cannot be resolved by the real edit resolver.", {
                validateOnly: true,
                validationCode: "TEMPLATE_VALIDATE_FAILED",
                errors: resolverErrors.length ? resolverErrors : ["no edit operations were produced"],
                resolver,
                bindings: enrichBindingsWithResolver(bindings, resolverResolved, resolverErrors),
                operations: ops.map(operationSummary),
                recommendedFix: "Run template candidates/create against the source Office file, inspect the objectMap, then correct template.mapping before retrying."
            });
        }
        return {
            kind: "officegen.template.fill",
            schema: "officegen.template.fill-validation@2.2",
            planOnly: true,
            validateOnly: true,
            mutatesOffice: false,
            generatedAt: nowIso(),
            templateId: template.id,
            templateHash: template.hash,
            supported: feasibility.supported,
            noopReason: feasibility.noopReason,
            formatCapabilities: feasibility.formatCapabilities,
            bindings: enrichBindingsWithResolver(bindings, resolverResolved, resolverErrors),
            operations: ops.map(operationSummary),
            validation: result,
            resolver
        };
    }
    if (canMutateOffice && !ops.length) {
        throw new TemplateFillError("Template fill produced no Office edit operations. Add template.mapping bindings for supplied values or run template apply-map before requesting an Office output.", {
            bindings,
            supported: false,
            noopReason: "no-edit-operations",
            formatCapabilities: feasibility.formatCapabilities,
            artifacts: expectedArtifact(outputPath, "output", outputFormat, false, "No edit operations were produced.")
        });
    }
    if (canMutateOffice && sourcePath) {
        const editResult = await edit(sourcePath, ops, {
            out: outputPath,
            format: outputFormat,
            resolveSelectors: true,
            validateFirst: true,
            atomic: true
        });
        const artifact = await outputArtifact(outputPath, "office-artifact", outputFormat, "template fill");
        const editRecord = editResult;
        const errors = Array.isArray(editRecord.errors) ? editRecord.errors : [];
        const changed = editRecord.changed !== false && errors.length === 0 && artifact.exists !== false;
        if (!changed) {
            throw new TemplateFillError("Template fill did not create a valid Office artifact.", {
                editResult,
                bindings,
                operations: ops.map(operationSummary),
                artifacts: [artifact],
                recommendedFix: "Run template fill --validate-only, inspect the template source with objectMap, then correct mappings before retrying."
            });
        }
        return {
            kind: "officegen.template.fill",
            planOnly: false,
            mutatesOffice: true,
            generatedAt: nowIso(),
            templateId: template.id,
            templateHash: template.hash,
            sourcePath,
            out: outputPath,
            supported: true,
            formatCapabilities: feasibility.formatCapabilities,
            bindings,
            operations: ops.map(operationSummary),
            editResult,
            artifacts: [artifact],
            values: options.values
        };
    }
    if (outputPath && ["pptx", "docx", "xlsx"].includes(outputFormat)) {
        throw new TemplateFillError("Template fill cannot write an Office artifact without a source Office template and resolvable mappings.", {
            supported: false,
            noopReason: sourcePath ? "missing-editable-bindings" : "missing-template-source",
            formatCapabilities: feasibility.formatCapabilities,
            requiredForOfficeMutation: {
                sourcePath: sourcePath ?? "template.source.path",
                outputPath,
                mapping: template.mapping ? "present" : "template.mapping"
            },
            bindings,
            artifacts: expectedArtifact(outputPath, "office-artifact", outputFormat, false, "Template source or mappings are missing.")
        });
    }
    const filled = {
        kind: "officegen.template.fill",
        planOnly: true,
        mutatesOffice: false,
        generatedAt: nowIso(),
        templateId: template.id,
        templateHash: template.hash,
        values: options.values,
        supported: false,
        noopReason: "plan-only-no-office-output",
        formatCapabilities: feasibility.formatCapabilities,
        bindings,
        requiredForOfficeMutation: {
            sourcePath: sourcePath ?? "template.source.path",
            outputPath: outputPath ?? "--out <file.pptx|file.docx|file.xlsx>",
            mapping: template.mapping ? "present" : "template.mapping"
        },
        note: "This JSON is a fill plan because no Office output path/source was available."
    };
    await writeJsonFile(options.outputPath ?? path.join(featureRoot(options, "template"), "runs", `${slugify(template.id)}.fill.json`), filled);
    return filled;
}
async function templateFillOperations(template, values, cwd) {
    const mapping = normalizeTemplateMapping(template.mapping) ?? {};
    const ops = [];
    for (const field of template.fields ?? []) {
        const value = values[field.name] ?? field.defaultValue;
        if (value === undefined)
            continue;
        const binding = normalizeTemplateBinding(mapping[field.name], field.type);
        if (!binding)
            continue;
        if (binding.kind === "chartData" && isRecord(value)) {
            ops.push({
                op: "pptx.updateChartData",
                selector: binding.selector,
                seriesName: typeof value.seriesName === "string" ? value.seriesName : field.name,
                categories: Array.isArray(value.categories) ? value.categories.map(String) : [],
                values: Array.isArray(value.values) ? value.values.map((item) => Number(item)) : []
            });
            continue;
        }
        if (binding.kind === "image") {
            const imagePath = typeof value === "string" ? path.resolve(cwd ?? process.cwd(), value) : undefined;
            if (imagePath) {
                ops.push({
                    op: "pptx.replaceImageByShape",
                    selector: binding.selector,
                    replacementBase64: (await readFile(imagePath)).toString("base64"),
                    replacementPath: imagePath,
                    fit: binding.fit
                });
            }
            continue;
        }
        if (binding.kind === "table" && Array.isArray(value)) {
            ops.push({
                op: "xlsx.writeTable",
                selector: binding.selector,
                startCell: binding.startCell ?? "A1",
                tableName: binding.tableName ?? field.name,
                rows: value
            });
            continue;
        }
        ops.push({
            op: "setText",
            selector: binding.selector,
            text: stringifyTemplateValue(value)
        });
    }
    return ops;
}
function normalizeTemplateBinding(value, fieldType) {
    if (typeof value === "string")
        return { selector: { stableObjectId: value }, kind: bindingKindFromFieldType(fieldType) };
    if (!isRecord(value))
        return undefined;
    const selector = isRecord(value.selector)
        ? value.selector
        : typeof value.stableObjectId === "string"
            ? { stableObjectId: value.stableObjectId }
            : undefined;
    if (!selector)
        return undefined;
    return {
        selector,
        kind: typeof value.kind === "string" ? value.kind : typeof value.type === "string" ? value.type : bindingKindFromFieldType(fieldType),
        fit: value.fit === "contain" || value.fit === "cover" || value.fit === "stretch" ? value.fit : undefined,
        startCell: typeof value.startCell === "string" ? value.startCell : undefined,
        tableName: typeof value.tableName === "string" ? value.tableName : undefined
    };
}
function bindingKindFromFieldType(fieldType) {
    if (fieldType === "image" || fieldType === "chartData" || fieldType === "table")
        return fieldType;
    if (fieldType === "list")
        return "table";
    return undefined;
}
function stringifyTemplateValue(value) {
    if (value === null || value === undefined)
        return "";
    if (typeof value === "string")
        return value;
    if (typeof value === "number" || typeof value === "boolean")
        return String(value);
    return JSON.stringify(value);
}
function isRecord(value) {
    return Boolean(value && typeof value === "object" && !Array.isArray(value));
}
export async function validateTemplate(options) {
    requireFeature(options, "template", "template validate");
    return validateTemplateDefinition(await inspectTemplate(options));
}
export function validateTemplateDefinition(template) {
    const errors = [];
    const warnings = [];
    if (!template.id?.trim())
        errors.push("id is required");
    if (!template.name?.trim())
        errors.push("name is required");
    for (const field of template.fields ?? []) {
        if (!field.name?.trim())
            errors.push("field.name is required");
        if (field.type && !supportedTemplateFieldTypes().includes(field.type)) {
            errors.push(`unsupported field type: ${field.type}`);
        }
    }
    if (!template.fields || template.fields.length === 0) {
        warnings.push("template has no fields");
    }
    return validation(errors.length === 0, errors, warnings);
}
function validateTemplateValues(template, values) {
    const errors = [];
    for (const field of template.fields ?? []) {
        const value = values[field.name] ?? field.defaultValue;
        if (field.required && value === undefined) {
            errors.push(`missing required field: ${field.name}`);
            continue;
        }
        if (value !== undefined && field.type && field.type !== "json" && !matchesFieldType(value, field.type)) {
            errors.push(`field ${field.name} expected ${field.type}`);
        }
    }
    return validation(errors.length === 0, errors);
}
function matchesFieldType(value, type) {
    if (type === "date")
        return typeof value === "string" && !Number.isNaN(Date.parse(value));
    if (type === "image")
        return typeof value === "string";
    if (type === "chartData")
        return isRecord(value) && Array.isArray(value.values);
    if (type === "table" || type === "list")
        return Array.isArray(value);
    return typeof value === type;
}
function supportedTemplateFieldTypes() {
    return ["string", "number", "boolean", "date", "json", "image", "chartData", "table", "list"];
}
function assertTemplateValid(template) {
    const result = validateTemplateDefinition(template);
    if (!result.ok) {
        throw new Error(`Invalid template: ${result.errors.join("; ")}`);
    }
}
function matchTemplateMapCandidates(template, sourceSignals) {
    const templateFields = new Set((template.fields ?? []).map((field) => normalizeField(field.name)).filter(Boolean));
    if (templateFields.size === 0)
        return sourceSignals.mapCandidates.slice(0, 8);
    const matched = [];
    for (const candidate of sourceSignals.mapCandidates) {
        const normalizedCandidate = normalizeField(candidate.field);
        if (!normalizedCandidate)
            continue;
        if (templateFields.has(normalizedCandidate) ||
            [...templateFields].some((field) => field.includes(normalizedCandidate) || normalizedCandidate.includes(field))) {
            matched.push(candidate);
        }
        if (matched.length >= 8)
            break;
    }
    return matched;
}
function makeSourceDerivedTemplateCandidate(sourcePath, sourceSignals) {
    const name = path.basename(sourcePath, path.extname(sourcePath));
    const fields = sourceSignals.schemaCandidates.map((field) => ({
        name: field.name,
        type: field.type,
        fieldType: field.type,
        required: field.required,
        description: `${field.reason}; confidence ${field.confidence}`,
        confidence: field.confidence,
        reason: field.reason,
        ...fieldBindingMetadata(field.name, sourceSignals)
    }));
    const template = {
        id: slugify(`suggested-${name}`),
        name: `Suggested template from ${name}`,
        version: "0.1.0",
        description: "Generated from sourcePath PPTX analysis. Review untrusted text before use.",
        tags: ["pptx", "source-derived"],
        fields,
        source: {
            path: sourcePath,
            format: "pptx",
            sha256: sourceSignals.trust.trusted.sha256
        },
        mapping: normalizeTemplateMapping(sourceSignals.templateMapSuggested),
        sourceCapture: {
            metadata: sourceSignals.metadata,
            artifactPaths: sourceSignals.artifactPaths,
            placeholderCandidates: sourceSignals.placeholderCandidates,
            namedShapeCandidates: sourceSignals.namedShapeCandidates,
            schemaCandidates: sourceSignals.schemaCandidates,
            templateMapSuggested: sourceSignals.templateMapSuggested,
            trust: sourceSignals.trust
        },
        createdAt: sourceSignals.trust.trusted.generatedAt,
        updatedAt: sourceSignals.trust.trusted.generatedAt
    };
    template.hash = sha256Json({ ...template, hash: undefined });
    return {
        template,
        score: 100 + sourceSignals.schemaCandidates.length + sourceSignals.placeholderCandidates.length,
        reasons: ["source:pptx-analysis", "template-map:suggested", "schema:candidates"],
        sourceMetadata: sourceSignals.metadata,
        previewCandidates: sourceSignals.previewCandidates,
        contextCandidates: sourceSignals.contextCandidates,
        mapCandidates: sourceSignals.mapCandidates,
        placeholderCandidates: sourceSignals.placeholderCandidates,
        namedShapeCandidates: sourceSignals.namedShapeCandidates,
        schemaCandidates: sourceSignals.schemaCandidates,
        templateMapSuggested: sourceSignals.templateMapSuggested,
        artifactPaths: sourceSignals.artifactPaths,
        trust: sourceSignals.trust,
        generatedFromSource: true
    };
}
function normalizeTemplateMapping(value) {
    if (!isRecord(value))
        return undefined;
    const candidate = isRecord(value.mapping) ? value.mapping : value;
    const normalized = {};
    for (const [field, mappingValue] of Object.entries(candidate)) {
        if (mappingValue === undefined)
            continue;
        if (isRecord(mappingValue) && isRecord(mappingValue.mapping)) {
            Object.assign(normalized, normalizeTemplateMapping(mappingValue));
            continue;
        }
        normalized[field] = normalizeMappingValue(mappingValue);
    }
    return normalized;
}
function normalizeMappingValue(value) {
    if (typeof value === "string")
        return value;
    if (!isRecord(value))
        return value;
    const selector = isRecord(value.selector)
        ? value.selector
        : typeof value.stableObjectId === "string"
            ? { stableObjectId: value.stableObjectId }
            : undefined;
    return selector ? { ...value, selector } : value;
}
function fieldBindingMetadata(fieldName, signals) {
    const stableObjectId = normalizeTemplateMapping(signals?.templateMapSuggested)?.[fieldName];
    const binding = normalizeTemplateBinding(stableObjectId, signals?.schemaCandidates.find((field) => field.name === fieldName)?.type);
    const placeholder = signals?.placeholderCandidates.find((candidate) => candidate.field === fieldName);
    const shape = binding?.selector?.stableObjectId
        ? signals?.namedShapeCandidates.find((candidate) => candidate.stableObjectId === binding.selector.stableObjectId)
        : undefined;
    const editableOps = editableOpsForBinding(binding?.kind, shape?.kind ?? placeholder?.placeholderType);
    return {
        selector: binding?.selector,
        editable: Boolean(binding?.selector),
        editableOps,
        confidence: placeholder?.confidence,
        reason: placeholder?.source ?? shape?.source
    };
}
function templateBindingDiagnostics(template) {
    const mapping = normalizeTemplateMapping(template.mapping);
    return (template.fields ?? []).map((field) => {
        const binding = normalizeTemplateBinding(mapping?.[field.name], field.type);
        const editableOps = editableOpsForBinding(binding?.kind);
        return {
            field: field.name,
            fieldType: field.type ?? field.fieldType ?? "string",
            kind: binding?.kind ?? "text",
            selector: binding?.selector,
            editable: Boolean(binding?.selector),
            selectorStatus: binding?.selector ? "unverified" : "missing",
            matchCount: binding?.selector ? undefined : 0,
            sourcePart: "slide",
            unsupportedReason: binding?.selector ? undefined : "missing-template-mapping",
            editableOps,
            confidence: field.confidence,
            reason: field.reason ?? field.description
        };
    });
}
function enrichBindingsWithResolver(bindings, resolvedSelectors, errors) {
    const errorText = errors.map((error) => typeof error === "string" ? error : JSON.stringify(error)).join("\n");
    return bindings.map((binding) => {
        const selector = isRecord(binding.selector) ? binding.selector : {};
        const stableObjectId = typeof selector.stableObjectId === "string" ? selector.stableObjectId : undefined;
        const matched = stableObjectId && resolvedSelectors.some((item) => JSON.stringify(item).includes(stableObjectId));
        const failed = stableObjectId && errorText.includes(stableObjectId);
        const selectorStatus = !stableObjectId ? "missing" : failed ? "not-found" : matched ? "resolved" : errors.length ? "unverified" : "resolved";
        return {
            ...binding,
            selectorStatus,
            matchCount: selectorStatus === "resolved" ? 1 : 0,
            editable: selectorStatus === "resolved",
            sourcePart: "slide",
            unsupportedReason: selectorStatus === "resolved" ? undefined : selectorStatus
        };
    });
}
function editableOpsForBinding(kind, objectKind) {
    if (kind === "image" || objectKind === "picture" || objectKind === "pic")
        return ["pptx.replaceImageByShape"];
    if (kind === "chartData" || objectKind === "chart")
        return ["pptx.updateChartData"];
    if (kind === "table")
        return ["xlsx.writeTable", "pptx.table.replaceData"];
    return ["setText"];
}
async function templateFillFeasibility(template, input) {
    let selectorCoverage;
    if (input.sourcePath) {
        const inspected = await inspect(input.sourcePath, { depth: "shallow", format: input.outputFormat }).catch(() => undefined);
        const objectIds = new Set(inspected?.objectMap.map((entry) => entry.stableObjectId) ?? []);
        const bindings = templateBindingDiagnostics(template);
        selectorCoverage = {
            checked: Boolean(inspected),
            missing: bindings.filter((binding) => {
                const selector = isRecord(binding.selector) ? binding.selector : {};
                return typeof selector.stableObjectId === "string" && !objectIds.has(String(selector.stableObjectId));
            }).map((binding) => binding.field)
        };
    }
    return {
        supported: input.canMutateOffice && input.ops.length > 0,
        noopReason: input.canMutateOffice ? (input.ops.length ? undefined : "no-edit-operations") : "requires-source-and-office-out",
        formatCapabilities: {
            pptx: { text: true, image: true, chartData: true, table: "limited" },
            docx: { text: true, comments: true, image: "limited", table: "limited" },
            xlsx: { text: true, table: true, chartData: "limited", pivot: "inspect-only" }
        },
        selectorCoverage
    };
}
function operationSummary(op) {
    const record = op;
    return {
        op: String(record.op ?? record.type ?? "unknown"),
        selector: record.selector,
        riskLevel: ["pptx.replaceImageByShape", "pptx.updateChartData", "xlsx.writeTable"].includes(String(record.op)) ? "medium" : "low"
    };
}
async function outputArtifact(filePath, kind, format, sourceCommand) {
    try {
        const stats = await stat(filePath);
        return { path: filePath, exists: true, bytes: stats.size, kind, format, sourceCommand };
    }
    catch {
        return { path: filePath, exists: false, kind, format, sourceCommand, reason: "expected output artifact was not created" };
    }
}
function expectedArtifact(filePath, kind, format, exists, reason) {
    return filePath ? [{ path: filePath, exists, kind, format, sourceCommand: "template fill", reason }] : [];
}
function normalizeField(value) {
    return value.toLowerCase().replace(/[^a-z0-9]+/g, "");
}
function templatePath(context, id) {
    return path.join(featureRoot(context, "template"), `${slugify(id)}.json`);
}
//# sourceMappingURL=template.js.map