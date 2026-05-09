import { Ajv } from "ajv/dist/ajv.js";
const commonSchemaField = { const: "" };
function schemaField(id) {
    return { const: id };
}
const envelopeSchema = {
    $id: "officegen.envelope@1.2",
    type: "object",
    required: ["schema", "ok", "cliVersion", "pathsRedacted", "warnings", "diagnostics", "artifacts", "nextSuggestedCommands"],
    allOf: [
        {
            if: { properties: { ok: { const: true } }, required: ["ok"] },
            then: { required: ["result"] }
        },
        {
            if: { properties: { ok: { const: false } }, required: ["ok"] },
            then: { required: ["error", "availableCommands"] }
        }
    ],
    properties: {
        schema: schemaField("officegen.envelope@1.2"),
        ok: { type: "boolean" },
        command: { type: "string" },
        runId: { type: "string" },
        cliVersion: { type: "string" },
        capabilitiesHash: { type: "string", pattern: "^sha256:" },
        pathsRedacted: { type: "boolean" },
        truncated: { type: "boolean" },
        result: {},
        error: {
            type: "object",
            required: ["code", "category", "severity", "message"],
            additionalProperties: true,
            properties: {
                code: { type: "string" },
                category: { type: "string" },
                severity: { enum: ["info", "warning", "error", "critical"] },
                message: { type: "string" }
            }
        },
        availableCommands: { type: "array", items: { type: "string" } },
        warnings: { type: "array" },
        diagnostics: { type: "array" },
        artifacts: { type: "array" },
        nextSuggestedCommands: { type: "array", items: { type: "string" } }
    },
    additionalProperties: false
};
const selectorSchema = {
    type: "object",
    minProperties: 1,
    additionalProperties: false,
    properties: {
        stableObjectId: { type: "string" },
        contains: { type: "string" },
        placeholderKey: { type: "string" },
        shapeName: { type: "string" },
        contentControlTag: { type: "string" },
        namedRange: { type: "string" },
        textMatch: {
            type: "object",
            required: ["text"],
            additionalProperties: false,
            properties: {
                text: { type: "string" },
                exact: { type: "boolean" }
            }
        }
    }
};
function editOperationSchemas() {
    const fields = {
        selector: selectorSchema,
        from: { type: "string" },
        to: { type: "string" },
        text: { type: "string" },
        slide: { type: "integer", minimum: 1 },
        after: { type: "integer", minimum: 0 },
        order: { type: "array", minItems: 1, items: { type: "integer", minimum: 1 } },
        items: { type: "array", minItems: 1, items: { type: "string" } },
        sheet: { type: "integer", minimum: 1 },
        rowIndex: { type: "integer", minimum: 1 },
        rows: { type: "array", minItems: 1, items: { type: "array" } },
        cell: { type: "string", pattern: "^[A-Za-z]+[1-9][0-9]*$" },
        value: { type: ["string", "number", "boolean", "null"] },
        formula: { type: "string", minLength: 1 },
        ref: { type: "string", pattern: "^[A-Za-z]+[1-9][0-9]*:[A-Za-z]+[1-9][0-9]*$" },
        startCell: { type: "string", pattern: "^[A-Za-z]+[1-9][0-9]*$" },
        page: { type: "integer", minimum: 1 },
        x: { type: "number" },
        y: { type: "number" },
        size: { type: "number", minimum: 1 },
        color: { type: "string" },
        width: { type: "number", minimum: 0 },
        height: { type: "number", minimum: 0 },
        replacementPath: { type: "string" },
        replacementBase64: { type: "string" },
        fit: { enum: ["contain", "cover", "stretch"] },
        crop: {
            type: "object",
            additionalProperties: false,
            properties: {
                left: { type: "number", minimum: 0, maximum: 1 },
                right: { type: "number", minimum: 0, maximum: 1 },
                top: { type: "number", minimum: 0, maximum: 1 },
                bottom: { type: "number", minimum: 0, maximum: 1 }
            }
        },
        bounds: {
            type: "object",
            required: ["x", "y", "width", "height"],
            additionalProperties: false,
            properties: {
                x: { type: "number" },
                y: { type: "number" },
                width: { type: "number", minimum: 0 },
                height: { type: "number", minimum: 0 }
            }
        },
        categories: { type: "array", minItems: 1, items: { type: "string" } },
        values: { type: "array", minItems: 1, items: { type: "number" } },
        seriesName: { type: "string" },
        author: { type: "string" },
        tableName: { type: "string" },
        styleId: { type: "string" },
        font: { type: "string" },
        bold: { type: "boolean" }
    };
    const pickFields = (allowed) => Object.fromEntries(allowed.map((field) => [field, fields[field]]));
    const op = (name, required, allowed, extra = {}, constraints = {}) => ({
        type: "object",
        required: ["op", ...required],
        additionalProperties: false,
        ...constraints,
        properties: {
            ...pickFields(allowed),
            ...extra,
            op: { const: name }
        }
    });
    return [
        op("replaceText", ["from", "to"], ["selector", "from", "to"]),
        op("setText", ["selector", "text"], ["selector", "text"]),
        op("pptx.duplicateSlide", [], ["slide", "after", "selector"], {}, { anyOf: [{ required: ["slide"] }, { required: ["selector"] }] }),
        op("pptx.reorderSlides", ["order"], ["order"]),
        op("pptx.insertBulletItems", ["selector", "items"], ["selector", "items"]),
        op("pptx.replaceBulletItems", ["selector", "items"], ["selector", "items"]),
        op("pptx.replaceImageByShape", ["selector"], ["selector", "replacementPath", "replacementBase64", "fit", "crop"], {}, { anyOf: [{ required: ["replacementPath"] }, { required: ["replacementBase64"] }] }),
        op("pptx.updateChartData", ["selector", "categories", "values"], ["selector", "categories", "values", "seriesName"]),
        op("pptx.setBounds", ["selector", "bounds"], ["selector", "bounds"]),
        op("docx.insertParagraphAfter", ["selector", "text"], ["selector", "text"]),
        op("docx.setHeader", ["text"], ["text"]),
        op("docx.setFooter", ["text"], ["text"]),
        op("docx.setStyle", ["styleId"], ["styleId", "font", "size", "bold"]),
        op("docx.addComment", ["selector", "text"], ["selector", "text", "author"]),
        op("docx.addRedline", ["selector", "text"], ["selector", "text", "author"]),
        op("xlsx.insertRows", ["rowIndex", "rows"], ["sheet", "rowIndex", "rows"]),
        op("xlsx.appendRows", ["rows"], ["sheet", "rows"]),
        op("xlsx.setCell", ["cell", "value"], ["sheet", "cell", "value"]),
        op("xlsx.setFormula", ["cell", "formula"], ["sheet", "cell", "formula"]),
        op("xlsx.updateTable", ["startCell", "rows"], ["sheet", "startCell", "rows"]),
        op("xlsx.writeTable", ["startCell", "rows"], ["sheet", "startCell", "rows", "tableName"]),
        op("xlsx.table.resize", ["selector", "ref"], ["selector", "ref"]),
        op("xlsx.chart.setData", ["selector", "categories", "values"], ["selector", "categories", "values", "seriesName"]),
        op("xlsx.pivot.refreshDefinition", ["selector"], ["selector"]),
        op("pdf.textOverlay", ["page", "text", "x", "y"], ["page", "text", "x", "y", "size", "color"]),
        op("pdf.annotation", ["page", "text", "x", "y"], ["page", "text", "x", "y", "width", "height"])
    ];
}
const editOpsSchema = {
    $id: "officegen.edit.ops@1.2",
    type: "object",
    required: ["schema", "target", "ops"],
    additionalProperties: false,
    allOf: [
        targetOpsCompatibility("pptx", "^(replaceText|setText|pptx\\.)"),
        targetOpsCompatibility("docx", "^(replaceText|setText|docx\\.)"),
        targetOpsCompatibility("xlsx", "^(replaceText|setText|xlsx\\.)"),
        targetOpsCompatibility("pdf", "^pdf\\.")
    ],
    properties: {
        schema: schemaField("officegen.edit.ops@1.2"),
        target: { enum: ["pptx", "docx", "xlsx", "pdf"] },
        options: {
            type: "object",
            additionalProperties: false,
            properties: {
                atomic: { type: "boolean", default: true },
                continueOnError: { type: "boolean", default: false },
                validateFirst: { type: "boolean", default: true },
                idempotencyKey: { type: "string" },
                preserveUnknownParts: { type: "boolean", default: true },
                preserveAnimations: { type: "boolean" }
            }
        },
        ops: {
            type: "array",
            minItems: 1,
            items: { oneOf: editOperationSchemas() }
        }
    }
};
function targetOpsCompatibility(target, pattern) {
    return {
        if: {
            properties: { target: { const: target } },
            required: ["target"]
        },
        then: {
            properties: {
                ops: {
                    type: "array",
                    items: {
                        type: "object",
                        required: ["op"],
                        properties: {
                            op: { type: "string", pattern }
                        }
                    }
                }
            }
        }
    };
}
const documentIrSchema = {
    $id: "officegen.ir.document@1.2",
    type: "object",
    required: ["schema", "targets", "sections"],
    additionalProperties: false,
    properties: {
        schema: schemaField("officegen.ir.document@1.2"),
        title: { type: "string" },
        header: { type: "string" },
        footer: { type: "string" },
        metadata: {
            type: "object",
            additionalProperties: true,
            properties: {
                title: { type: "string" },
                author: { type: "string" }
            }
        },
        targets: { type: "array", minItems: 1, items: { enum: ["pptx", "docx", "xlsx", "pdf"] } },
        design: { type: "object", additionalProperties: true },
        assets: { type: "array", items: { type: "object", additionalProperties: true } },
        slides: { type: "array", items: { type: "object", additionalProperties: true } },
        sheets: { type: "array", items: { type: "object", additionalProperties: true } },
        sections: {
            type: "array",
            minItems: 1,
            items: {
                type: "object",
                required: ["blocks"],
                additionalProperties: false,
                properties: {
                    id: { type: "string" },
                    title: { type: "string" },
                    body: { anyOf: [{ type: "string" }, { type: "array", items: { type: "string" } }] },
                    rows: { type: "array", items: { anyOf: [{ type: "array" }, { type: "object", additionalProperties: true }] } },
                    items: { type: "array", items: { type: "string" } },
                    blocks: {
                        type: "array",
                        items: {
                            type: "object",
                            required: ["type"],
                            additionalProperties: true,
                            properties: {
                                type: { type: "string" },
                                text: { type: "string" },
                                title: { type: "string" },
                                role: { type: "string" },
                                specRef: { type: "string" },
                                items: { type: "array", items: { type: "string" } },
                                rows: { type: "array", items: { anyOf: [{ type: "array" }, { type: "object", additionalProperties: true }] } },
                                path: { type: "string" }
                            }
                        }
                    }
                }
            }
        }
    }
};
const assetSpecSchema = {
    $id: "officegen.asset.spec@1.2",
    type: "object",
    required: ["schema", "assets"],
    additionalProperties: false,
    properties: {
        schema: schemaField("officegen.asset.spec@1.2"),
        assets: {
            type: "array",
            items: {
                type: "object",
                required: ["id", "type", "path"],
                additionalProperties: false,
                properties: {
                    id: { type: "string" },
                    type: { enum: ["image", "video", "audio", "data", "font"] },
                    path: { type: "string" },
                    role: { type: "string" }
                }
            }
        }
    }
};
const designPackSchema = {
    $id: "officegen.design.pack@1.2",
    type: "object",
    required: ["schema", "name"],
    additionalProperties: true,
    properties: {
        schema: schemaField("officegen.design.pack@1.2"),
        name: { type: "string" },
        colors: { type: "object", additionalProperties: { type: "string" } },
        typography: { type: "object", additionalProperties: true },
        spacing: { type: "object", additionalProperties: true }
    }
};
const templateMapSchema = {
    $id: "officegen.template.map@1.2",
    type: "object",
    required: ["schema", "bindings"],
    additionalProperties: false,
    properties: {
        schema: schemaField("officegen.template.map@1.2"),
        bindings: {
            type: "array",
            items: {
                type: "object",
                required: ["key", "selector"],
                additionalProperties: false,
                properties: {
                    key: { type: "string" },
                    selector: selectorSchema,
                    locked: { type: "boolean" }
                }
            }
        }
    }
};
const viewObjectMapSchema = {
    $id: "officegen.view.objectMap@1.2",
    type: "object",
    required: ["schema", "page", "coordinateSystem", "fidelity", "objects"],
    additionalProperties: false,
    properties: {
        schema: schemaField("officegen.view.objectMap@1.2"),
        page: { type: "integer", minimum: 1 },
        coordinateSystem: { enum: ["px", "pt", "emu"] },
        fidelity: { enum: ["exact", "high", "approximate", "low"] },
        objects: {
            type: "array",
            items: {
                type: "object",
                required: ["stableObjectId", "type", "bbox", "editable", "untrusted"],
                additionalProperties: false,
                properties: {
                    stableObjectId: { type: "string" },
                    type: { type: "string" },
                    name: { type: "string" },
                    bbox: { type: "array", minItems: 4, maxItems: 4, items: { type: "number" } },
                    textPreview: { type: "string" },
                    selectorHints: { type: "object", additionalProperties: true },
                    trust: {
                        type: "object",
                        required: ["level", "reason"],
                        additionalProperties: false,
                        properties: {
                            level: { enum: ["untrusted"] },
                            reason: { type: "string" }
                        }
                    },
                    editable: { type: "boolean" },
                    untrusted: { type: "boolean" }
                }
            }
        }
    }
};
const diagnosticsSchema = {
    $id: "officegen.diagnostics@1.2",
    type: "object",
    required: ["schema", "diagnostics"],
    additionalProperties: false,
    properties: {
        schema: schemaField("officegen.diagnostics@1.2"),
        diagnostics: {
            type: "array",
            items: {
                type: "object",
                required: ["code", "severity", "message"],
                additionalProperties: true,
                properties: {
                    code: { type: "string" },
                    severity: { enum: ["info", "warning", "error", "critical"] },
                    message: { type: "string" },
                    details: {}
                }
            }
        }
    }
};
const entries = [
    entry("officegen.envelope@1.2", envelopeSchema, undefined),
    entry("officegen.edit.ops@1.2", editOpsSchema, "edit"),
    entry("officegen.ir.document@1.2", documentIrSchema, "render"),
    entry("officegen.asset.spec@1.2", assetSpecSchema, "asset"),
    entry("officegen.design.pack@1.2", designPackSchema, "design"),
    entry("officegen.template.map@1.2", templateMapSchema, "template"),
    entry("officegen.view.objectMap@1.2", viewObjectMapSchema, "view"),
    entry("officegen.diagnostics@1.2", diagnosticsSchema, "diagnose")
];
function entry(id, schema, feature) {
    return {
        id,
        schema: {
            ...schema,
            "x-officegen-schema-id": id,
            "x-officegen-stability": "stable",
            "x-officegen-introduced-in": "1.2.0",
            "x-officegen-deprecated": false
        },
        stability: "stable",
        introducedIn: "1.2.0",
        deprecated: false,
        feature,
        visibleToAgents: true
    };
}
export class SchemaRegistry {
    ajv;
    entriesById = new Map();
    validators = new Map();
    constructor(registryEntries = entries) {
        this.ajv = new Ajv({ allErrors: true, strict: false });
        for (const registryEntry of registryEntries) {
            this.entriesById.set(registryEntry.id, registryEntry);
            this.validators.set(registryEntry.id, this.ajv.compile(registryEntry.schema));
        }
    }
    list(options = {}) {
        return [...this.entriesById.values()].filter((registryEntry) => {
            if (!options.agent)
                return true;
            if (!registryEntry.visibleToAgents)
                return false;
            if (!registryEntry.feature || !options.config)
                return true;
            const feature = options.config.features[registryEntry.feature];
            return feature.enabled && feature.visibleToAgents;
        });
    }
    get(id) {
        return this.entriesById.get(id);
    }
    validate(id, value) {
        const validate = this.validators.get(id);
        if (!validate) {
            return {
                ok: false,
                errors: [
                    {
                        keyword: "schema",
                        instancePath: "",
                        schemaPath: "",
                        params: { id },
                        message: `Unknown schema: ${id}`
                    }
                ]
            };
        }
        const ok = validate(value);
        return ok ? { ok: true } : { ok: false, errors: cloneErrors(validate.errors ?? []) };
    }
}
function cloneErrors(errors) {
    return errors.map((error) => ({
        ...error,
        params: { ...error.params }
    }));
}
export const defaultSchemaRegistry = new SchemaRegistry();
export function listSchemas(options = {}) {
    return defaultSchemaRegistry.list(options);
}
export function getSchema(id) {
    return defaultSchemaRegistry.get(id);
}
export function validateSchema(id, value) {
    return defaultSchemaRegistry.validate(id, value);
}
void commonSchemaField;
//# sourceMappingURL=schemas.js.map