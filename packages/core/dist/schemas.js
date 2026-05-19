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
        runtimeEnvelope: schemaField("officegen.envelope@2"),
        ok: { type: "boolean" },
        command: { type: "string" },
        runId: { type: "string" },
        cliVersion: { type: "string" },
        capabilitiesHash: { type: "string", pattern: "^sha256:" },
        pathsRedacted: { type: "boolean" },
        truncated: { type: "boolean" },
        executionOk: { type: "boolean" },
        objectiveOk: { type: "boolean" },
        mutationStatus: { type: "string" },
        artifactStatus: { type: "string" },
        readiness: { type: "string" },
        partial: { type: "boolean" },
        failureClass: { enum: ["none", "unsupported", "partial", "blocked", "runtime", "input", "schema", "security", "usage"] },
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
        nextSuggestedCommands: { type: "array", items: { type: "string" } },
        nextActions: { type: "array", items: { type: "string" } }
    },
    additionalProperties: false
};
const envelopeV2Schema = {
    ...envelopeSchema,
    $id: "officegen.envelope@2",
    required: [
        "schema",
        "runtimeEnvelope",
        "ok",
        "command",
        "runId",
        "cliVersion",
        "capabilitiesHash",
        "pathsRedacted",
        "executionOk",
        "objectiveOk",
        "readiness",
        "mutationStatus",
        "artifactStatus",
        "partial",
        "failureClass",
        "warnings",
        "diagnostics",
        "artifacts",
        "nextSuggestedCommands",
        "nextActions"
    ],
    properties: {
        ...envelopeSchema.properties,
        schema: { enum: ["officegen.envelope@1.2", "officegen.envelope@2"] },
        runtimeEnvelope: schemaField("officegen.envelope@2"),
        mutationStatus: { enum: ["changed", "noop", "plan_only", "failed", "not_applicable"] },
        artifactStatus: { enum: ["complete", "missing", "not_expected"] },
        readiness: { enum: ["pass", "pass_with_environment_gap", "warning", "partial", "blocked"] }
    }
};
const capabilityContractSchema = {
    type: "object",
    required: ["area", "formats", "support", "summary", "limitations"],
    additionalProperties: false,
    properties: {
        area: { type: "string" },
        formats: { type: "array", minItems: 1, items: { enum: ["pptx", "docx", "xlsx", "pdf", "json", "svg", "html", "image"] } },
        support: { enum: ["supported", "limited", "unsupported", "optional-gated", "overlay-only"] },
        summary: { type: "string" },
        limitations: { type: "array", items: { type: "string" } }
    }
};
const runtimeProfileCapabilitySchema = {
    type: "object",
    required: ["id", "area", "support", "summary", "evidence", "gaps"],
    additionalProperties: false,
    properties: {
        id: { type: "string" },
        area: { type: "string" },
        support: { enum: ["supported", "limited", "unsupported", "target-only"] },
        summary: { type: "string" },
        evidence: { type: "array", items: { type: "string" } },
        gaps: { type: "array", items: { type: "string" } }
    }
};
const runtimeProfileSchema = {
    type: "object",
    required: ["id", "role", "runtime", "summary", "capabilities"],
    additionalProperties: false,
    properties: {
        id: { enum: ["current-limited-v3.1", "perfect-runtime-target"] },
        role: { enum: ["current", "target"] },
        runtime: { type: "string" },
        summary: { type: "string" },
        capabilities: { type: "array", minItems: 1, items: runtimeProfileCapabilitySchema }
    }
};
const capabilitiesSchema = {
    $id: "officegen.capabilities@1.2",
    type: "object",
    required: [
        "schema",
        "ok",
        "profile",
        "capabilitiesHash",
        "visibleCommands",
        "hiddenFromAgents",
        "disabled",
        "agentInstructionsPath",
        "jsonBudgetBytes",
        "featureContracts",
        "formatCapabilities",
        "runtimeProfiles",
        "specProfile",
        "knownLimitations",
        "unsupportedNow",
        "nextSuggestedCommands"
    ],
    additionalProperties: true,
    properties: {
        schema: schemaField("officegen.capabilities@1.2"),
        ok: { const: true },
        profile: { enum: ["substrate", "authoring", "enterprise"] },
        capabilitiesHash: { type: "string", pattern: "^sha256:" },
        visibleCommands: { type: "array", items: { type: "string" } },
        hiddenFromAgents: { type: "array", items: { type: "string" } },
        disabled: { type: "array", items: { type: "string" } },
        agentInstructionsPath: { type: "string" },
        jsonBudgetBytes: { type: "integer", minimum: 0 },
        featureContracts: { type: "array", minItems: 1, items: capabilityContractSchema },
        formatCapabilities: {
            type: "object",
            required: ["pptx", "docx", "xlsx", "pdf"],
            additionalProperties: true,
            properties: {
                pptx: { type: "object", additionalProperties: true },
                docx: { type: "object", additionalProperties: true },
                xlsx: { type: "object", additionalProperties: true },
                pdf: { type: "object", additionalProperties: true }
            }
        },
        runtimeProfiles: {
            type: "object",
            required: ["current-limited-v3.1", "perfect-runtime-target"],
            additionalProperties: false,
            properties: {
                "current-limited-v3.1": runtimeProfileSchema,
                "perfect-runtime-target": runtimeProfileSchema
            }
        },
        specProfile: {
            type: "object",
            required: ["currentProfileId", "targetProfileId", "runtimeProjection", "truthfulnessPolicy", "currentEvidence", "targetGapIds"],
            additionalProperties: false,
            properties: {
                currentProfileId: { const: "current-limited-v3.1" },
                targetProfileId: { const: "perfect-runtime-target" },
                runtimeProjection: { const: "runtime-v2" },
                truthfulnessPolicy: { type: "string" },
                currentEvidence: { type: "array", minItems: 1, items: { type: "string" } },
                targetGapIds: { type: "array", minItems: 1, items: { type: "string" } }
            }
        },
        knownLimitations: { type: "array", items: { type: "string" } },
        unsupportedNow: { type: "array", items: { type: "string" } },
        nextSuggestedCommands: { type: "array", items: { type: "string" } }
    }
};
const selectorSchema = {
    type: "object",
    minProperties: 1,
    additionalProperties: false,
    properties: {
        stableObjectId: { type: "string" },
        slide: { type: "integer", minimum: 1 },
        shapeId: { type: "string" },
        contains: { type: "string" },
        placeholderKey: { type: "string" },
        placeholder: { type: "string" },
        shapeName: { type: "string" },
        sourcePath: { type: "string" },
        xmlPath: { type: "string" },
        page: { type: "integer", minimum: 1 },
        story: { type: "string" },
        paragraph: { type: "integer", minimum: 1 },
        table: { type: "integer", minimum: 1 },
        row: { type: "integer", minimum: 1 },
        column: { type: "integer", minimum: 1 },
        range: { type: "string" },
        relationshipId: { type: "string" },
        assetPath: { type: "string" },
        commentId: { type: "string" },
        revisionId: { type: "string" },
        contentControlTag: { type: "string" },
        namedRange: { type: "string" },
        sheetName: { type: "string" },
        cell: { type: "string", pattern: "^[A-Za-z]+[1-9][0-9]*$" },
        tableName: { type: "string" },
        chartPath: { type: "string" },
        textHash: { type: "string" },
        positionHash: { type: "string" },
        nearestTo: {
            type: "object",
            required: ["x", "y"],
            additionalProperties: false,
            properties: {
                slide: { type: "integer", minimum: 1 },
                x: { type: "number" },
                y: { type: "number" }
            }
        },
        rightOf: {
            oneOf: [
                { type: "string" },
                {
                    type: "object",
                    required: ["text"],
                    additionalProperties: false,
                    properties: {
                        text: { type: "string" },
                        slide: { type: "integer", minimum: 1 }
                    }
                }
            ]
        },
        largestTextOnSlide: { oneOf: [{ type: "integer", minimum: 1 }, { type: "boolean" }] },
        nthBodyShape: {
            type: "object",
            required: ["slide", "n"],
            additionalProperties: false,
            properties: {
                slide: { type: "integer", minimum: 1 },
                n: { type: "integer", minimum: 1 }
            }
        },
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
        selectors: { type: "array", minItems: 1, items: selectorSchema },
        from: { type: "string" },
        to: { type: "string" },
        text: { type: "string" },
        slide: { type: "integer", minimum: 1 },
        after: { type: "integer", minimum: 0 },
        order: { type: "array", minItems: 1, items: { type: "integer", minimum: 1 } },
        items: { type: "array", minItems: 1, items: { type: "string" } },
        sheet: { oneOf: [{ type: "integer", minimum: 1 }, { type: "string", minLength: 1 }] },
        sheetName: { type: "string", minLength: 1 },
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
        name: { type: "string" },
        title: { type: "string" },
        description: { type: "string" },
        decorative: { type: "boolean" },
        layout: { oneOf: [{ type: "string" }, { type: "integer", minimum: 1 }] },
        mode: { enum: ["replace", "append", "left", "right", "center", "top", "bottom", "middle"] },
        axis: { enum: ["x", "y"] },
        minFontSize: { type: "number", minimum: 1 },
        values2d: { type: "array", minItems: 1, items: { type: "array" } },
        fontSize: { type: "number", minimum: 1 },
        bold: { type: "boolean" },
        level: { type: "integer", minimum: 0, maximum: 8 },
        startAt: { type: "integer", minimum: 1 },
        lineSpacing: { type: "number", minimum: 0 },
        spaceBefore: { type: "number", minimum: 0 },
        textCase: { enum: ["upper", "lower", "title", "sentence"] },
        kind: { enum: ["header", "footer"] },
        selected: { type: "array", minItems: 1, items: { type: "string" } }
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
        op("pptx.addSlide", [], ["after"]),
        op("pptx.addSlideFromLayout", [], ["after", "layout"]),
        op("pptx.reorderSlides", ["order"], ["order"]),
        op("pptx.addTextbox", ["slide", "text", "bounds"], ["slide", "text", "bounds", "name", "fontSize", "bold"]),
        op("pptx.formatTitle", ["selector"], ["selector", "fontSize", "bold", "textCase"]),
        op("pptx.formatAllTitles", [], ["fontSize", "bold", "textCase"]),
        op("pptx.replaceBodyBullets", ["slide", "items"], ["slide", "items"], {
            items: {
                type: "array",
                minItems: 1,
                items: {
                    oneOf: [
                        { type: "string" },
                        {
                            type: "object",
                            required: ["text"],
                            additionalProperties: false,
                            properties: {
                                text: { type: "string" },
                                level: { type: "integer", minimum: 0, maximum: 8 },
                                bold: { type: "boolean" },
                                numbering: { type: "boolean" }
                            }
                        }
                    ]
                }
            },
            spaceBeforeForLevel1ExceptFirst: { type: "number", minimum: 0 }
        }),
        op("pptx.fitContentToPlaceholder", ["selector"], ["selector", "minFontSize"]),
        op("pptx.alignObjects", ["selectors", "mode"], ["selectors", "mode"]),
        op("pptx.distributeObjects", ["selectors", "axis"], ["selectors", "axis"]),
        op("pptx.setAltText", ["selector"], ["selector", "title", "description", "decorative"], {
            title: { type: "string" }
        }, { anyOf: [{ required: ["title"] }, { required: ["description"] }, { required: ["decorative"] }] }),
        op("pptx.setSpeakerNotes", ["slide", "text"], ["slide", "text", "mode"], {
            mode: { enum: ["replace", "append"] }
        }),
        op("pptx.replaceWithBulletList", ["selector", "items"], ["selector", "items"], {
            items: {
                type: "array",
                minItems: 1,
                items: {
                    oneOf: [
                        { type: "string" },
                        {
                            type: "object",
                            required: ["text"],
                            additionalProperties: false,
                            properties: {
                                text: { type: "string" },
                                level: { type: "integer", minimum: 0, maximum: 8 },
                                bold: { type: "boolean" },
                                numbering: { type: "boolean" }
                            }
                        }
                    ]
                }
            },
            spaceBeforeForLevel1ExceptFirst: { type: "number", minimum: 0 }
        }),
        op("pptx.insertBulletItems", ["selector", "items"], ["selector", "items"]),
        op("pptx.replaceBulletItems", ["selector", "items"], ["selector", "items"]),
        op("pptx.setFontSize", ["selector", "fontSize"], ["selector", "fontSize"]),
        op("pptx.setBold", ["selector", "bold"], ["selector", "bold"]),
        op("pptx.setBulletLevel", ["selector", "level"], ["selector", "level"]),
        op("pptx.setNumbering", ["selector"], ["selector", "level", "startAt"]),
        op("pptx.setLineSpacing", ["selector", "lineSpacing"], ["selector", "lineSpacing"]),
        op("pptx.setSpaceBefore", ["selector", "spaceBefore"], ["selector", "spaceBefore"]),
        op("pptx.setTextCase", ["selector", "textCase"], ["selector", "textCase"]),
        op("pptx.setTableCellText", ["selector", "text"], ["selector", "text"]),
        op("pptx.replaceImageByShape", ["selector"], ["selector", "replacementPath", "replacementBase64", "fit", "crop"], {}, { anyOf: [{ required: ["replacementPath"] }, { required: ["replacementBase64"] }] }),
        op("pptx.updateChartData", ["selector", "categories", "values"], ["selector", "categories", "values", "seriesName"], {}, {
            description: "Updates one existing PPTX chart series by replacing categories and values. Multi-series, secondary-axis, and combo-chart editing are unsupported.",
            "x-officegen-support": {
                support: "limited",
                series: "single-series-only",
                unsupported: ["multi-series", "secondary-axis", "combo-chart", "per-series-chart-type"]
            }
        }),
        op("pptx.setBounds", ["selector", "bounds"], ["selector", "bounds"]),
        op("docx.insertParagraphAfter", ["selector", "text"], ["selector", "text"]),
        op("docx.replaceTextSmart", ["from", "to"], ["selector", "from", "to"]),
        op("docx.setTableCellText", ["selector", "text"], ["selector", "text"]),
        op("docx.setHeader", ["text"], ["text"]),
        op("docx.setFooter", ["text"], ["text"]),
        op("docx.setStyle", ["styleId"], ["styleId", "font", "size", "bold"]),
        op("docx.addComment", ["selector", "text"], ["selector", "text", "author"]),
        op("docx.addRedline", ["selector", "text"], ["selector", "text", "author"]),
        op("docx.redline.insert", ["selector", "text"], ["selector", "text", "author"]),
        op("docx.redline.delete", ["selector"], ["selector", "author"]),
        op("docx.redline.replace", ["selector", "text"], ["selector", "text", "author"]),
        op("docx.applyStyle", ["selector", "styleId"], ["selector", "styleId"]),
        op("docx.headerFooter.setText", ["kind", "text"], ["kind", "text"]),
        op("xlsx.insertRows", ["rowIndex", "rows"], ["sheet", "sheetName", "rowIndex", "rows"]),
        op("xlsx.appendRows", ["rows"], ["sheet", "sheetName", "rows"]),
        op("xlsx.setCell", ["cell", "value"], ["sheet", "sheetName", "cell", "value"]),
        op("xlsx.setFormula", ["cell", "formula"], ["sheet", "sheetName", "cell", "formula"]),
        op("xlsx.definedName.set", ["name", "ref"], ["name", "ref"]),
        op("xlsx.definedName.delete", ["name"], ["name"]),
        op("xlsx.setRange", ["startCell", "values"], ["sheet", "sheetName", "startCell"], {
            values: { type: "array", minItems: 1, items: { type: "array" } }
        }),
        op("xlsx.updateTable", ["startCell", "rows"], ["sheet", "sheetName", "startCell", "rows"]),
        op("xlsx.writeTable", ["startCell", "rows"], ["sheet", "sheetName", "startCell", "rows", "tableName"]),
        op("xlsx.table.resize", ["selector", "ref"], ["selector", "ref"]),
        op("xlsx.chart.setData", ["selector", "categories", "values"], ["selector", "categories", "values", "seriesName"], {}, {
            description: "Updates one existing XLSX chart series by replacing categories and values. Multi-series, secondary-axis, and combo-chart editing are unsupported.",
            "x-officegen-support": {
                support: "limited",
                series: "single-series-only",
                unsupported: ["multi-series", "secondary-axis", "combo-chart", "per-series-chart-type"]
            }
        }),
        op("xlsx.pivot.refreshDefinition", ["selector"], ["selector"], {}, {
            description: "Marks an existing pivot definition for refresh. Pivot layout, field, and value editing are unsupported.",
            "x-officegen-support": { support: "limited", mode: "refresh-flag-only" }
        }),
        op("xlsx.pivot.refreshAll", [], [], {}, {
            description: "Marks workbook pivot definitions/caches for refresh. Pivot layout, field, and value editing are unsupported.",
            "x-officegen-support": { support: "limited", mode: "refresh-flag-only" }
        }),
        op("xlsx.slicer.setSelection", ["selector", "selected"], ["selector", "selected"], {}, {
            description: "Updates selected slicer items only. Slicer creation, caption editing, and styling are unsupported.",
            "x-officegen-support": { support: "limited", mode: "selection-only" }
        }),
        op("pdf.textOverlay", ["page", "text", "x", "y"], ["page", "text", "x", "y", "size", "color"], {}, {
            description: "Draws overlay text on a PDF page. This is not redaction and does not remove or rewrite underlying PDF content.",
            "x-officegen-support": { support: "overlay-only", redaction: "unsupported", removesUnderlyingContent: false }
        }),
        op("pdf.annotation", ["page", "text", "x", "y"], ["page", "text", "x", "y", "width", "height"], {}, {
            description: "Adds a PDF annotation. This is not redaction and does not remove or rewrite underlying PDF content.",
            "x-officegen-support": { support: "overlay-only", redaction: "unsupported", removesUnderlyingContent: false }
        })
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
                allowPartial: { type: "boolean", default: false },
                validateFirst: { type: "boolean", default: true },
                idempotencyKey: { type: "string" },
                expectedInputSha256: { type: "string", pattern: "^sha256:" },
                expectedObjectMapHash: { type: "string", pattern: "^sha256:" },
                expectedObjectGraphHash: { type: "string", pattern: "^sha256:" },
                selectionLock: {
                    type: "object",
                    required: ["objectGraphHash"],
                    additionalProperties: false,
                    properties: {
                        objectGraphHash: { type: "string", pattern: "^sha256:" },
                        nodeId: { type: "string" },
                        sourceFingerprint: { type: "string", pattern: "^sha256:" }
                    }
                },
                minSelectorConfidence: { type: "number", minimum: 0, maximum: 1 },
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
const objectGraphRiskFlagSchema = {
    type: "object",
    required: ["code", "severity", "message"],
    additionalProperties: true,
    properties: {
        code: { type: "string" },
        severity: { enum: ["info", "warning", "error"] },
        message: { type: "string" },
        source: { type: "string" }
    }
};
const objectGraphEvidenceSchema = {
    type: "object",
    required: ["kind", "confidence", "message"],
    additionalProperties: true,
    properties: {
        kind: { enum: ["object-map", "selector-hint", "geometry", "derived"] },
        confidence: { type: "number", minimum: 0, maximum: 1 },
        message: { type: "string" },
        sourceField: { type: "string" }
    }
};
const objectGraphSourceSchema = {
    type: "object",
    additionalProperties: true,
    properties: {
        format: { type: "string" },
        inputPath: { type: "string" },
        inputSha256: { type: "string" },
        sourcePath: { type: "string" },
        xmlPath: { type: "string" },
        slide: { type: "integer", minimum: 1 },
        page: { type: "integer", minimum: 1 },
        sheet: { type: "integer", minimum: 1 },
        sheetName: { type: "string" },
        story: { type: "string" }
    }
};
const objectGraphNodeSchema = {
    type: "object",
    required: ["schema", "version", "graphVersion", "index", "nodeId", "stableId", "type", "source", "provenance", "confidence", "riskFlags", "evidence"],
    additionalProperties: true,
    properties: {
        schema: schemaField("officegen.objectGraph@2"),
        version: { const: 2 },
        graphVersion: schemaField("officegen.objectGraph@2"),
        index: { type: "integer", minimum: 0 },
        nodeId: { type: "string" },
        stableId: { type: "string" },
        type: { type: "string" },
        label: { type: "string" },
        bbox: { type: "array", minItems: 4, maxItems: 4, items: { type: "number" } },
        text: { type: "object", additionalProperties: true },
        style: { type: "object", additionalProperties: true },
        source: objectGraphSourceSchema,
        provenance: {
            type: "object",
            required: ["schema", "source", "objectMapIndex", "stableObjectId"],
            additionalProperties: true,
            properties: {
                schema: schemaField("officegen.objectGraph@2"),
                source: { const: "inspect.objectMap" },
                objectMapIndex: { type: "integer", minimum: 0 },
                stableObjectId: { type: "string" }
            }
        },
        confidence: { type: "number", minimum: 0, maximum: 1 },
        riskFlags: { type: "array", items: objectGraphRiskFlagSchema },
        evidence: { type: "array", items: objectGraphEvidenceSchema }
    }
};
const objectGraphEdgeSchema = {
    type: "object",
    required: ["schema", "version", "graphVersion", "index", "edgeId", "from", "to", "relation", "confidence", "riskFlags", "evidence"],
    additionalProperties: true,
    properties: {
        schema: schemaField("officegen.objectGraph@2"),
        version: { const: 2 },
        graphVersion: schemaField("officegen.objectGraph@2"),
        index: { type: "integer", minimum: 0 },
        edgeId: { type: "string" },
        from: { type: "string" },
        to: { type: "string" },
        relation: { enum: ["contains", "rightOf", "below"] },
        confidence: { type: "number", minimum: 0, maximum: 1 },
        riskFlags: { type: "array", items: objectGraphRiskFlagSchema },
        evidence: { type: "array", items: objectGraphEvidenceSchema }
    }
};
const objectGraphSchema = {
    $id: "officegen.objectGraph@2",
    type: "object",
    required: ["schema", "version", "graphVersion", "source", "provenance", "confidence", "riskFlags", "pagination", "index", "nodes", "edges"],
    additionalProperties: true,
    properties: {
        schema: schemaField("officegen.objectGraph@2"),
        version: { const: 2 },
        graphVersion: schemaField("officegen.objectGraph@2"),
        source: {
            type: "object",
            required: ["objectMapCount", "builder"],
            additionalProperties: true,
            properties: {
                format: { type: "string" },
                inputPath: { type: "string" },
                inputSha256: { type: "string" },
                objectMapCount: { type: "integer", minimum: 0 },
                builder: { const: "inspect.objectMap" }
            }
        },
        provenance: {
            type: "object",
            required: ["generatedFrom", "sourceField"],
            additionalProperties: true,
            properties: {
                generatedFrom: schemaField("officegen.inspect.result@1.2"),
                sourceField: { const: "objectMap" }
            }
        },
        confidence: { type: "number", minimum: 0, maximum: 1 },
        riskFlags: { type: "array", items: objectGraphRiskFlagSchema },
        pagination: {
            type: "object",
            required: ["nodeOffset", "nodeLimit", "nodeCount", "totalNodes", "edgeOffset", "edgeLimit", "edgeCount", "totalEdges", "truncated"],
            additionalProperties: false,
            properties: {
                nodeOffset: { type: "integer", minimum: 0 },
                nodeLimit: { type: "integer", minimum: 0 },
                nodeCount: { type: "integer", minimum: 0 },
                totalNodes: { type: "integer", minimum: 0 },
                edgeOffset: { type: "integer", minimum: 0 },
                edgeLimit: { type: "integer", minimum: 0 },
                edgeCount: { type: "integer", minimum: 0 },
                totalEdges: { type: "integer", minimum: 0 },
                truncated: { type: "boolean" },
                nextNodeOffset: { type: "integer", minimum: 0 },
                nextEdgeOffset: { type: "integer", minimum: 0 }
            }
        },
        index: {
            type: "object",
            required: ["nodesByStableId", "nodesByType", "edgesByRelation"],
            additionalProperties: true,
            properties: {
                nodesByStableId: { type: "object", additionalProperties: { type: "string" } },
                nodesByType: { type: "object", additionalProperties: { type: "array", items: { type: "string" } } },
                edgesByRelation: { type: "object", additionalProperties: { type: "array", items: { type: "string" } } }
            }
        },
        nodes: { type: "array", items: objectGraphNodeSchema },
        edges: { type: "array", items: objectGraphEdgeSchema }
    }
};
const selectorResolutionV2Schema = {
    $id: "officegen.selectorResolution@2",
    type: "object",
    required: ["schema", "status", "candidates", "evidence", "nextActions", "selectionLock"],
    additionalProperties: false,
    properties: {
        schema: schemaField("officegen.selectorResolution@2"),
        status: { enum: ["matched", "not_found", "ambiguous", "low_confidence", "stale", "unsupported"] },
        confidence: { type: "number", minimum: 0, maximum: 1 },
        candidates: {
            type: "array",
            items: {
                type: "object",
                required: ["stableObjectId", "type"],
                additionalProperties: true,
                properties: {
                    nodeId: { type: "string" },
                    stableObjectId: { type: "string" },
                    type: { type: "string" },
                    label: { type: "string" },
                    text: { type: "string" },
                    confidence: { type: "number", minimum: 0, maximum: 1 },
                    source: { type: "object", additionalProperties: true }
                }
            }
        },
        evidence: {
            type: "array",
            items: {
                type: "object",
                required: ["kind", "confidence", "message"],
                additionalProperties: true,
                properties: {
                    kind: { enum: ["object-map", "selector-hint", "geometry", "derived"] },
                    confidence: { type: "number", minimum: 0, maximum: 1 },
                    message: { type: "string" },
                    sourceField: { type: "string" }
                }
            }
        },
        ambiguityReason: { type: "string" },
        nextActions: { type: "array", items: { type: "string" } },
        selectionLock: {
            type: "object",
            required: ["objectGraphHash"],
            additionalProperties: false,
            properties: {
                objectGraphHash: { type: "string", pattern: "^sha256:" },
                nodeId: { type: "string" },
                sourceFingerprint: { type: "string", pattern: "^sha256:" }
            }
        }
    }
};
const sourceFingerprintSchema = {
    type: "object",
    required: ["algorithm", "hash", "byteLength"],
    additionalProperties: false,
    properties: {
        algorithm: { const: "sha256" },
        hash: { type: "string", pattern: "^[a-f0-9]{64}$" },
        byteLength: { type: "integer", minimum: 0 },
        path: { type: "string" }
    }
};
const editPlanV2Schema = {
    $id: "officegen.editPlan@2",
    type: "object",
    required: ["schema", "target", "operations", "wouldWrite"],
    additionalProperties: true,
    properties: {
        schema: schemaField("officegen.editPlan@2"),
        input: { type: "string" },
        target: { enum: ["pptx", "docx", "xlsx", "pdf", "unknown"] },
        inputSha256: { type: "string", pattern: "^sha256:" },
        objectMapHash: { type: "string", pattern: "^sha256:" },
        objectGraphHash: { type: "string", pattern: "^sha256:" },
        operations: { type: "array", items: { type: "object", additionalProperties: true } },
        selectorResolution: {
            type: "object",
            required: ["schema", "status", "candidates", "evidence", "nextActions", "selectionLock"],
            additionalProperties: true,
            properties: {
                schema: schemaField("officegen.selectorResolution@2"),
                status: { enum: ["matched", "not_found", "ambiguous", "low_confidence", "stale", "unsupported"] },
                candidates: { type: "array" },
                evidence: { type: "array" },
                nextActions: { type: "array", items: { type: "string" } },
                selectionLock: { type: "object", additionalProperties: true }
            }
        },
        selectorResolutions: {
            type: "object",
            required: ["schema", "resolutions"],
            additionalProperties: true,
            properties: {
                schema: schemaField("officegen.edit.selectors@1.2"),
                resolutions: { type: "array", items: { type: "object", additionalProperties: true } }
            }
        },
        patchPlan: { type: "object", additionalProperties: true },
        wouldWrite: { type: "boolean" }
    }
};
const patchPlanV2Schema = {
    $id: "officegen.patchPlan@2",
    type: "object",
    required: ["schema", "format", "wouldWrite", "inputSha256", "sourceFingerprint", "operations", "touchedParts", "expectedChangedParts", "sourceFingerprints", "blocked"],
    additionalProperties: false,
    properties: {
        schema: schemaField("officegen.patchPlan@2"),
        format: { enum: ["pptx", "docx", "xlsx", "pdf", "unknown"] },
        wouldWrite: { const: false },
        inputSha256: { type: "string", pattern: "^sha256:" },
        objectMapHash: { type: "string", pattern: "^sha256:" },
        objectGraphHash: { type: "string", pattern: "^sha256:" },
        sourceFingerprint: sourceFingerprintSchema,
        operations: {
            type: "array",
            items: {
                type: "object",
                required: ["operationIndex", "op", "wouldApply"],
                additionalProperties: false,
                properties: {
                    operationIndex: { type: "integer", minimum: 0 },
                    op: { type: "string" },
                    wouldApply: { type: "boolean" },
                    reason: { enum: ["not-found", "ambiguous", "low-confidence", "unsupported", "validation-failed", "idempotency-replay", "skipped-after-error", "stale-plan"] },
                    message: { type: "string" },
                    selector: selectorSchema
                }
            }
        },
        touchedParts: {
            type: "array",
            items: {
                type: "object",
                required: ["path", "change"],
                additionalProperties: false,
                properties: {
                    path: { type: "string" },
                    change: { enum: ["modified", "created", "deleted"] },
                    beforeSha256: { type: "string", pattern: "^sha256:" },
                    afterSha256: { type: "string", pattern: "^sha256:" },
                    sourceFingerprint: sourceFingerprintSchema
                }
            }
        },
        expectedChangedParts: { type: "array", items: { type: "string" } },
        sourceFingerprints: { type: "array", items: sourceFingerprintSchema },
        blocked: { type: "array", items: { type: "object", additionalProperties: true } }
    }
};
const repairPlanV2Schema = {
    $id: "officegen.repairPlan@2",
    type: "object",
    required: ["schema", "version", "target", "inputSha256", "wouldWrite", "operations", "failureTaxonomy", "steps", "verify"],
    additionalProperties: true,
    properties: {
        schema: schemaField("officegen.repairPlan@2"),
        version: { const: 2 },
        target: { enum: ["pptx", "docx", "xlsx", "pdf", "svg", "html", "unknown"] },
        input: { type: "string" },
        inputSha256: { type: "string", pattern: "^sha256:" },
        wouldWrite: { const: false },
        operations: { type: "array", items: { type: "object", additionalProperties: true } },
        failureTaxonomy: {
            type: "array",
            items: {
                type: "object",
                required: ["code", "category", "severity", "autoRepairable", "evidence", "nextCommand"],
                additionalProperties: true,
                properties: {
                    code: { type: "string" },
                    category: { enum: ["quality", "compatibility", "security", "environment"] },
                    severity: { enum: ["info", "warning", "error", "critical"] },
                    autoRepairable: { type: "boolean" },
                    evidence: {
                        type: "array",
                        items: {
                            type: "object",
                            required: ["kind", "message"],
                            additionalProperties: true,
                            properties: {
                                kind: { type: "string" },
                                message: { type: "string" },
                                issueCode: { type: "string" },
                                stableObjectId: { type: "string" }
                            }
                        }
                    },
                    nextCommand: { type: "string" }
                }
            }
        },
        steps: {
            type: "array",
            items: {
                type: "object",
                required: ["id", "command", "dryRun", "reason"],
                additionalProperties: true,
                properties: {
                    id: { type: "string" },
                    command: { type: "string" },
                    dryRun: { type: "boolean" },
                    reason: { type: "string" }
                }
            }
        },
        verify: {
            type: "object",
            required: ["status", "requiredAfterRepair", "command", "readinessNote"],
            additionalProperties: true,
            properties: {
                status: { enum: ["not_run"] },
                requiredAfterRepair: { type: "boolean" },
                command: { type: "string" },
                readinessNote: { type: "string" }
            }
        }
    }
};
const officeAgentPhaseSchema = {
    type: "object",
    required: ["id", "standardName", "manifestRole", "commandTemplate", "mutatesOffice", "status"],
    additionalProperties: true,
    properties: {
        id: { type: "string", pattern: "^phase-\\d{2}-" },
        standardName: { type: "string" },
        manifestRole: { type: "string" },
        commandTemplate: { type: "string" },
        mutatesOffice: { type: "boolean" },
        status: { enum: ["skeleton", "manual-ready"] },
        execution: { enum: ["skeleton", "manual-gated"] }
    }
};
const officeAgentManifestV31Schema = {
    $id: "officegen.office-agent.manifest@3.1",
    type: "object",
    required: ["schema", "release", "runtimeProjection", "mode", "status", "phaseCount", "phases", "limitations", "requiredPhaseNames"],
    additionalProperties: true,
    properties: {
        schema: schemaField("officegen.office-agent.manifest@3.1"),
        release: { const: "3.1.0" },
        runtimeProjection: { const: "runtime-v2" },
        mode: { const: "skeleton-evidence" },
        status: { const: "skeleton" },
        phaseCount: { const: 13 },
        phases: { type: "array", minItems: 13, maxItems: 13, items: officeAgentPhaseSchema },
        limitations: { type: "array", minItems: 1, items: { type: "string" } },
        requiredPhaseNames: {
            type: "array",
            contains: { const: "verify" },
            items: { type: "string" }
        }
    }
};
const officeAgentResultV31Schema = {
    $id: "officegen.office-agent.result@3.1",
    type: "object",
    required: ["schema", "release", "runtimeProjection", "mode", "readiness", "phaseCount", "phases", "artifacts", "caveats"],
    additionalProperties: true,
    properties: {
        schema: schemaField("officegen.office-agent.result@3.1"),
        release: { const: "3.1.0" },
        runtimeProjection: { const: "runtime-v2" },
        mode: { const: "skeleton-evidence" },
        readiness: { const: "warning" },
        phaseCount: { const: 13 },
        phases: { type: "array", minItems: 13, maxItems: 13, items: officeAgentPhaseSchema },
        artifacts: { type: "array", minItems: 1, items: { type: "object", additionalProperties: true } },
        caveats: { type: "array", minItems: 1, items: { type: "string" } }
    }
};
const officeAgentWorkflowV31Schema = {
    $id: "officegen.office-agent.workflow@3.1",
    type: "object",
    required: ["schema", "release", "runtimeProjection", "phaseCount", "skeletonOnly", "steps"],
    additionalProperties: true,
    properties: {
        schema: schemaField("officegen.office-agent.workflow@3.1"),
        release: { const: "3.1.0" },
        runtimeProjection: { const: "runtime-v2" },
        phaseCount: { const: 13 },
        skeletonOnly: { const: true },
        steps: { type: "array", minItems: 13, maxItems: 13, items: officeAgentPhaseSchema }
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
const verifyGatesSchema = {
    $id: "officegen.verify.gates@1.2",
    type: "object",
    additionalProperties: false,
    properties: {
        expectedSlides: { type: "integer", minimum: 0 },
        expectedPages: { type: "integer", minimum: 0 },
        requiredText: { type: "array", items: { type: "string" } },
        forbiddenText: { type: "array", items: { type: "string" } },
        maxWarnings: { type: "integer", minimum: 0 },
        requireNoRepairDialog: { type: "boolean" },
        maxBlankPages: { type: "integer", minimum: 0 }
    }
};
const verificationReportV2Schema = {
    $id: "officegen.verify@2",
    type: "object",
    required: ["schema", "version", "format", "readiness", "score", "partial", "gates", "issues", "artifacts", "recommendedRepairs"],
    additionalProperties: true,
    properties: {
        schema: schemaField("officegen.verify@2"),
        version: { const: 2 },
        format: { type: "string" },
        readiness: { enum: ["pass", "pass_with_environment_gap", "warning", "blocked"] },
        score: { type: "number", minimum: 0, maximum: 1 },
        partial: { type: "boolean" },
        gates: {
            type: "object",
            required: ["schema", "package", "semantic", "visual", "native", "security", "accessibility", "goal"],
            additionalProperties: false,
            properties: Object.fromEntries(["schema", "package", "semantic", "visual", "native", "security", "accessibility", "goal"].map((name) => [name, {
                    type: "object",
                    required: ["status", "issues"],
                    additionalProperties: true,
                    properties: {
                        status: { enum: ["pass", "warning", "fail", "skipped"] },
                        score: { type: "number", minimum: 0, maximum: 1 },
                        summary: { type: "object", additionalProperties: true },
                        issues: { type: "array", items: { type: "string" } }
                    }
                }]))
        },
        issues: {
            type: "array",
            items: {
                type: "object",
                required: ["code", "severity", "category", "message"],
                additionalProperties: true,
                properties: {
                    code: { type: "string" },
                    severity: { enum: ["info", "warning", "error"] },
                    category: { enum: ["quality", "compatibility", "security", "environment"] },
                    message: { type: "string" },
                    gate: { enum: ["schema", "package", "semantic", "visual", "native", "security", "accessibility", "goal"] }
                }
            }
        },
        artifacts: {
            type: "array",
            items: {
                type: "object",
                required: ["artifactId", "role", "managed"],
                additionalProperties: true,
                properties: {
                    artifactId: { type: "string" },
                    role: { type: "string" },
                    path: { type: "string" },
                    format: { type: "string" },
                    managed: { type: "boolean" },
                    exists: { type: "boolean" },
                    sourceCommand: { type: "string" }
                }
            }
        },
        recommendedRepairs: { type: "array" }
    }
};
const xlsxFormulaDependencySchema = {
    type: "object",
    required: ["kind", "sourceText", "untrusted"],
    additionalProperties: false,
    properties: {
        kind: { enum: ["cell", "range", "threeD", "namedRange", "tableStructuredRef"] },
        ref: { type: "string" },
        sheet: { type: "string" },
        workbook: { type: "string" },
        name: { type: "string" },
        tableName: { type: "string" },
        sourceText: { type: "string" },
        untrusted: { const: true }
    }
};
const xlsxFormulaRelatedObjectSchema = {
    type: "object",
    required: ["kind", "path", "reason", "untrusted"],
    additionalProperties: false,
    properties: {
        kind: { enum: ["table", "chart", "pivotTable", "slicer"] },
        name: { type: "string" },
        path: { type: "string" },
        ref: { type: "string" },
        reason: { type: "string" },
        untrusted: { const: true }
    }
};
const xlsxFormulaCellSchema = {
    type: "object",
    required: ["stableObjectId", "sheetIndex", "ref", "formula", "dependencies", "unsafeFlags", "relatedObjects", "sourcePath", "untrusted"],
    additionalProperties: false,
    properties: {
        stableObjectId: { type: "string" },
        sheetIndex: { type: "integer", minimum: 1 },
        sheetName: { type: "string" },
        ref: { type: "string", pattern: "^[A-Za-z]+[1-9][0-9]*$" },
        formula: { type: "string" },
        formulaType: { type: "string" },
        sharedIndex: { type: "string" },
        sharedRef: { type: "string" },
        dependencies: { type: "array", items: xlsxFormulaDependencySchema },
        unsafeFlags: { type: "array", items: { enum: ["external", "volatile", "indirect", "unsupported"] } },
        volatileFunctions: { type: "array", items: { type: "string" } },
        relatedObjects: { type: "array", items: xlsxFormulaRelatedObjectSchema },
        sourcePath: { type: "string" },
        untrusted: { const: true }
    }
};
const xlsxFormulaGraphSchema = {
    $id: "officegen.xlsx.formulaGraph@1.0",
    type: "object",
    required: ["schema", "sheetIndex", "formulaCells", "dependencies", "unsafeFlags", "relatedObjects", "untrusted"],
    additionalProperties: false,
    properties: {
        schema: schemaField("officegen.xlsx.formulaGraph@1.0"),
        sheetIndex: { type: "integer", minimum: 1 },
        sheetName: { type: "string" },
        formulaCells: { type: "array", items: xlsxFormulaCellSchema },
        dependencies: { type: "array", items: xlsxFormulaDependencySchema },
        unsafeFlags: { type: "array", items: { enum: ["external", "volatile", "indirect", "unsupported"] } },
        relatedObjects: { type: "array", items: xlsxFormulaRelatedObjectSchema },
        untrusted: { const: true }
    }
};
const transactionSchema = {
    $id: "officegen.transaction@1.2",
    type: "object",
    required: ["schema", "inputPath", "outputPath", "backupPath", "inputSha256", "createdAt"],
    additionalProperties: true,
    properties: {
        schema: schemaField("officegen.transaction@1.2"),
        inputPath: { type: "string" },
        outputPath: { type: "string" },
        backupPath: { type: "string" },
        inputSha256: { type: "string" },
        outputSha256: { type: "string" },
        objectMapHash: { type: "string", pattern: "^sha256:" },
        objectGraphHash: { type: "string", pattern: "^sha256:" },
        sourceFingerprint: sourceFingerprintSchema,
        patchPlanInputSha256: { type: "string", pattern: "^sha256:" },
        createdAt: { type: "string" },
        rollbackCommand: { type: "string" },
        attribution: { type: "object", additionalProperties: true },
        lockPath: { type: "string" },
        scope: { type: "string" }
    }
};
function looseSchema(id) {
    return {
        $id: id,
        type: "object",
        required: ["schema"],
        additionalProperties: true,
        properties: {
            schema: schemaField(id)
        }
    };
}
const entries = [
    entry("officegen.envelope@1.2", envelopeSchema, undefined),
    entry("officegen.envelope@2", envelopeV2Schema, undefined),
    entry("officegen.capabilities@1.2", capabilitiesSchema, "capabilities"),
    entry("officegen.edit.ops@1.2", editOpsSchema, "edit"),
    entry("officegen.editPlan@2", editPlanV2Schema, "edit"),
    entry("officegen.patchPlan@2", patchPlanV2Schema, "edit"),
    entry("officegen.repairPlan@2", repairPlanV2Schema, "repair"),
    entry("officegen.ir.document@1.2", documentIrSchema, "render"),
    entry("officegen.manifest@1.2", looseSchema("officegen.manifest@1.2"), "run"),
    entry("officegen.asset.spec@1.2", assetSpecSchema, "asset"),
    entry("officegen.design.pack@1.2", designPackSchema, "design"),
    entry("officegen.template.map@1.2", templateMapSchema, "template"),
    entry("officegen.template.candidates.result@2.5", looseSchema("officegen.template.candidates.result@2.5"), "template"),
    entry("officegen.view.objectMap@1.2", viewObjectMapSchema, "view"),
    entry("officegen.objectGraph@0.1", looseSchema("officegen.objectGraph@0.1"), "inspect"),
    entry("officegen.objectGraph@2", objectGraphSchema, "inspect"),
    entry("officegen.docx.storyGraph@0.1", looseSchema("officegen.docx.storyGraph@0.1"), "inspect"),
    entry("officegen.docx.runGraph@0.1", looseSchema("officegen.docx.runGraph@0.1"), "inspect"),
    entry("officegen.diagnostics@1.2", diagnosticsSchema, "diagnose"),
    entry("officegen.help@1.2", looseSchema("officegen.help@1.2"), "help"),
    entry("officegen.config@1.2", looseSchema("officegen.config@1.2"), "config"),
    entry("officegen.config.result@1.2", looseSchema("officegen.config.result@1.2"), "config"),
    entry("officegen.doctor@1.2", looseSchema("officegen.doctor@1.2"), "doctor"),
    entry("officegen.schema.list@1.2", looseSchema("officegen.schema.list@1.2"), "schema"),
    entry("officegen.schema.definition@1.2", looseSchema("officegen.schema.definition@1.2"), "schema"),
    entry("officegen.validation.result@1.2", looseSchema("officegen.validation.result@1.2"), "validate"),
    entry("officegen.schema.migration.result@1.2", looseSchema("officegen.schema.migration.result@1.2"), "schema"),
    entry("officegen.errors@1.2", looseSchema("officegen.errors@1.2"), "errors"),
    entry("officegen.error@1.2", looseSchema("officegen.error@1.2"), "errors"),
    entry("officegen.progressive-disclosure@1.2", looseSchema("officegen.progressive-disclosure@1.2"), undefined),
    entry("officegen.projected-result@2.3", looseSchema("officegen.projected-result@2.3"), undefined),
    entry("officegen.inspect.result@1.2", looseSchema("officegen.inspect.result@1.2"), "inspect"),
    entry("officegen.view.result@1.2", looseSchema("officegen.view.result@1.2"), "view"),
    entry("officegen.view.manifest@1.2", looseSchema("officegen.view.manifest@1.2"), "view"),
    entry("officegen.selectorResolution@2", selectorResolutionV2Schema, "edit"),
    entry("officegen.edit.selectors@1.2", looseSchema("officegen.edit.selectors@1.2"), "edit"),
    entry("officegen.edit.result@1.2", looseSchema("officegen.edit.result@1.2"), "edit"),
    entry("officegen.render.result@1.2", looseSchema("officegen.render.result@1.2"), "render"),
    entry("officegen.export.result@1.2", looseSchema("officegen.export.result@1.2"), "export"),
    entry("officegen.diagnose.result@1.2", looseSchema("officegen.diagnose.result@1.2"), "diagnose"),
    entry("officegen.verify.result@1.2", looseSchema("officegen.verify.result@1.2"), "verify"),
    entry("officegen.verify@2", verificationReportV2Schema, "verify"),
    entry("officegen.verify.gates@1.2", verifyGatesSchema, "verify"),
    entry("officegen.xlsx.formulaGraph@1.0", xlsxFormulaGraphSchema, "inspect"),
    entry("officegen.repair.result@1.2", looseSchema("officegen.repair.result@1.2"), "repair"),
    entry("officegen.diff.result@1.2", looseSchema("officegen.diff.result@1.2"), "diff"),
    entry("officegen.diff.artifacts@1.2", looseSchema("officegen.diff.artifacts@1.2"), "diff"),
    entry("officegen.artifact.manifest@1.2", looseSchema("officegen.artifact.manifest@1.2"), "manifest"),
    entry("officegen.manifest.verify.result@1.2", looseSchema("officegen.manifest.verify.result@1.2"), "manifest"),
    entry("officegen.plan.result@1.2", looseSchema("officegen.plan.result@1.2"), "plan"),
    entry("officegen.rollback.result@1.2", looseSchema("officegen.rollback.result@1.2"), "rollback"),
    entry("officegen.lock@1.2", looseSchema("officegen.lock@1.2"), "lock"),
    entry("officegen.merge.result@1.2", looseSchema("officegen.merge.result@1.2"), "merge"),
    entry("officegen.transaction@1.2", transactionSchema, "rollback"),
    entry("officegen.office-edit.result@1.2", looseSchema("officegen.office-edit.result@1.2"), "run"),
    entry("officegen.scaffold.result@1.2", looseSchema("officegen.scaffold.result@1.2"), "scaffold"),
    entry("officegen.chart.render.result@1.2", looseSchema("officegen.chart.render.result@1.2"), "chart"),
    entry("officegen.diagram.render.result@1.2", looseSchema("officegen.diagram.render.result@1.2"), "diagram"),
    entry("officegen.asset.info@1.2", looseSchema("officegen.asset.info@1.2"), "asset"),
    entry("officegen.asset.embedded.info@2.5", looseSchema("officegen.asset.embedded.info@2.5"), "asset"),
    entry("officegen.asset.embedded.result@2.5", looseSchema("officegen.asset.embedded.result@2.5"), "asset"),
    entry("officegen.asset.embedded.trusted@2.5", looseSchema("officegen.asset.embedded.trusted@2.5"), "asset"),
    entry("officegen.asset.embedded.untrusted@2.5", looseSchema("officegen.asset.embedded.untrusted@2.5"), "asset"),
    entry("officegen.asset.extract.result@1.2", looseSchema("officegen.asset.extract.result@1.2"), "asset"),
    entry("officegen.asset.replace.result@1.2", looseSchema("officegen.asset.replace.result@1.2"), "asset"),
    entry("officegen.asset.result@1.2", looseSchema("officegen.asset.result@1.2"), "asset"),
    entry("officegen.critique.result@2.3", looseSchema("officegen.critique.result@2.3"), "critique"),
    entry("officegen.improve.plan@2.5", looseSchema("officegen.improve.plan@2.5"), "improve"),
    entry("officegen.improve.plan@2.3", looseSchema("officegen.improve.plan@2.3"), "improve"),
    entry("officegen.benchmark.run.result@2.5", looseSchema("officegen.benchmark.run.result@2.5"), "benchmark"),
    entry("officegen.benchmark.run.result@2.3", looseSchema("officegen.benchmark.run.result@2.3"), "benchmark"),
    entry("officegen.benchmark.compare.result@2.3", looseSchema("officegen.benchmark.compare.result@2.3"), "benchmark"),
    entry("officegen.run.plan@1.2", looseSchema("officegen.run.plan@1.2"), "run"),
    entry("officegen.run.manifest@2.3", looseSchema("officegen.run.manifest@2.3"), "run"),
    entry("officegen.run.result@2.3", looseSchema("officegen.run.result@2.3"), "run"),
    entry("officegen.run.manifest@2.4", looseSchema("officegen.run.manifest@2.4"), "run"),
    entry("officegen.run.result@2.4", looseSchema("officegen.run.result@2.4"), "run"),
    entry("officegen.office-agent.manifest@3.1", officeAgentManifestV31Schema, "run"),
    entry("officegen.office-agent.result@3.1", officeAgentResultV31Schema, "run"),
    entry("officegen.office-agent.workflow@3.1", officeAgentWorkflowV31Schema, "run"),
    entry("officegen.prepare-reference.manifest@1.2", looseSchema("officegen.prepare-reference.manifest@1.2"), "run"),
    entry("officegen.prepare-reference.result@1.2", looseSchema("officegen.prepare-reference.result@1.2"), "run"),
    entry("officegen.prepare-reference.object-map@1.2", looseSchema("officegen.prepare-reference.object-map@1.2"), "run"),
    entry("officegen.run.step-error@1.2", looseSchema("officegen.run.step-error@1.2"), "run"),
    entry("officegen.ooxml.validation@1", looseSchema("officegen.ooxml.validation@1"), "validate"),
    entry("officegen.mcp.tools@1.2", looseSchema("officegen.mcp.tools@1.2"), "mcp"),
    entry("officegen.renderer.doctor@2.2", looseSchema("officegen.renderer.doctor@2.2"), "renderer"),
    entry("officegen.template.fill-validation@2.2", looseSchema("officegen.template.fill-validation@2.2"), "template"),
    entry("officegen.design.signals.trusted@1.2", looseSchema("officegen.design.signals.trusted@1.2"), "design"),
    entry("officegen.design.signals.untrusted@1.2", looseSchema("officegen.design.signals.untrusted@1.2"), "design"),
    entry("officegen.design.contact-sheet@2.2", looseSchema("officegen.design.contact-sheet@2.2"), "design"),
    entry("officegen.design.evidence@1.2", looseSchema("officegen.design.evidence@1.2"), "design"),
    entry("officegen.template.schema-candidates@1.2", looseSchema("officegen.template.schema-candidates@1.2"), "template")
];
function entry(id, schema, feature) {
    const introducedIn = schemaVersionFromId(id);
    return {
        id,
        schema: {
            ...schema,
            "x-officegen-schema-id": id,
            "x-officegen-stability": "stable",
            "x-officegen-introduced-in": introducedIn,
            "x-officegen-deprecated": false
        },
        stability: "stable",
        introducedIn,
        deprecated: false,
        feature,
        visibleToAgents: true
    };
}
function schemaVersionFromId(id) {
    const version = id.match(/@([0-9]+(?:\.[0-9]+)*)$/)?.[1];
    if (!version)
        return "1.2.0";
    const parts = version.split(".");
    while (parts.length < 3)
        parts.push("0");
    return parts.join(".");
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
    validate(id, value, options = {}) {
        const validate = this.validators.get(id);
        const entry = this.entriesById.get(id);
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
        const errors = cloneErrors(validate.errors ?? []);
        return ok
            ? { ok: true }
            : {
                ok: false,
                errors,
                ...(options.diagnostics && entry ? { diagnostics: oneOfDiagnostics(entry.schema, value, errors) } : {})
            };
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
export function validateSchema(id, value, options = {}) {
    return defaultSchemaRegistry.validate(id, value, options);
}
void commonSchemaField;
export function compactSchemaErrors(errors, diagnostics = []) {
    const diagnosticPaths = new Set(diagnostics.map((diagnostic) => diagnostic.instancePath));
    const compacted = [];
    for (const error of errors) {
        if (!diagnosticPaths.has(parentInstancePath(error.instancePath)))
            compacted.push(error);
        else if (error.keyword === "oneOf")
            compacted.push(error);
    }
    return compacted.length > 0 ? compacted : errors.slice(0, 10);
}
function oneOfDiagnostics(rootSchema, value, errors) {
    return errors
        .filter((error) => error.keyword === "oneOf")
        .map((error) => oneOfDiagnostic(rootSchema, value, error))
        .filter((diagnostic) => Boolean(diagnostic));
}
function oneOfDiagnostic(rootSchema, value, error) {
    const schemaPath = error.schemaPath.replace(/^#/, "");
    const oneOf = schemaAtPointer(rootSchema, schemaPath);
    if (!Array.isArray(oneOf))
        return undefined;
    const instance = valueAtPointer(value, error.instancePath);
    const matches = oneOf
        .map((schema, index) => scoreOneOfCandidate(schema, instance, `${error.schemaPath}/${index}`))
        .sort((left, right) => right.score - left.score);
    const best = matches[0];
    if (!best)
        return undefined;
    return {
        instancePath: error.instancePath,
        schemaPath: error.schemaPath,
        bestMatch: {
            schemaPath: best.schemaPath,
            ...(best.op ? { op: best.op } : {}),
            score: best.score
        },
        missing: best.missing,
        unexpected: best.unexpected,
        expectedTypes: best.expectedTypes
    };
}
function scoreOneOfCandidate(schema, value, schemaPath) {
    const objectSchema = asRecord(schema);
    const properties = asRecord(objectSchema.properties);
    const required = Array.isArray(objectSchema.required) ? objectSchema.required.map(String) : [];
    const input = asRecord(value);
    const allowed = new Set(Object.keys(properties));
    const missing = required.filter((field) => !(field in input));
    const unexpected = objectSchema.additionalProperties === false
        ? Object.keys(input).filter((field) => !allowed.has(field))
        : [];
    const expectedTypes = {};
    let score = 0;
    const opConst = asRecord(properties.op).const;
    if (typeof opConst === "string") {
        if (input.op === opConst)
            score += 100;
        else
            score -= 10;
    }
    for (const field of required) {
        const fieldSchema = properties[field];
        if (!(field in input)) {
            expectedTypes[field] = expectedTypeLabels(fieldSchema);
            score -= 6;
        }
        else {
            score += 4;
        }
    }
    for (const [field, fieldValue] of Object.entries(input)) {
        const fieldSchema = properties[field];
        if (!fieldSchema)
            continue;
        if (schemaAcceptsValue(fieldSchema, fieldValue)) {
            score += field === "op" ? 5 : 2;
        }
        else {
            expectedTypes[field] = expectedTypeLabels(fieldSchema);
            score -= 4;
        }
    }
    score -= unexpected.length * 3;
    return {
        schemaPath,
        op: typeof opConst === "string" ? opConst : undefined,
        score,
        missing,
        unexpected,
        expectedTypes
    };
}
function schemaAcceptsValue(schema, value) {
    const record = asRecord(schema);
    if ("const" in record)
        return value === record.const;
    if (Array.isArray(record.enum))
        return record.enum.includes(value);
    if (Array.isArray(record.oneOf))
        return record.oneOf.some((candidate) => schemaAcceptsValue(candidate, value));
    if (Array.isArray(record.anyOf))
        return record.anyOf.some((candidate) => schemaAcceptsValue(candidate, value));
    const types = Array.isArray(record.type) ? record.type.map(String) : typeof record.type === "string" ? [record.type] : [];
    if (types.length === 0)
        return true;
    return types.some((type) => jsonType(value) === type || (type === "integer" && Number.isInteger(value)));
}
function expectedTypeLabels(schema) {
    const record = asRecord(schema);
    if ("const" in record)
        return [`const:${String(record.const)}`];
    if (Array.isArray(record.enum))
        return record.enum.map((item) => `enum:${String(item)}`);
    if (Array.isArray(record.oneOf))
        return [...new Set(record.oneOf.flatMap(expectedTypeLabels))];
    if (Array.isArray(record.anyOf))
        return [...new Set(record.anyOf.flatMap(expectedTypeLabels))];
    if (Array.isArray(record.type))
        return record.type.map(String);
    if (typeof record.type === "string")
        return [record.type];
    return [];
}
function schemaAtPointer(root, pointer) {
    return pointerTokens(pointer).reduce((current, token) => {
        if (Array.isArray(current))
            return current[Number(token)];
        if (current && typeof current === "object")
            return current[token];
        return undefined;
    }, root);
}
function valueAtPointer(root, pointer) {
    return schemaAtPointer(root, pointer);
}
function pointerTokens(pointer) {
    return pointer
        .split("/")
        .slice(1)
        .map((token) => token.replace(/~1/g, "/").replace(/~0/g, "~"));
}
function asRecord(value) {
    return value && typeof value === "object" && !Array.isArray(value) ? value : {};
}
function jsonType(value) {
    if (value === null)
        return "null";
    if (Array.isArray(value))
        return "array";
    return typeof value;
}
function parentInstancePath(path) {
    const index = path.lastIndexOf("/");
    return index <= 0 ? "" : path.slice(0, index);
}
//# sourceMappingURL=schemas.js.map