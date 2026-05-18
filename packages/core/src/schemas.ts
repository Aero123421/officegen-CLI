import { Ajv, type ErrorObject, type ValidateFunction } from "ajv/dist/ajv.js";
import type { FeatureName, JsonObject, OfficegenConfig, SchemaRegistryEntry } from "./types.js";

const commonSchemaField = { const: "" };

function schemaField(id: string): JsonObject {
  return { const: id };
}

const envelopeSchema: JsonObject = {
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
    executionOk: { type: "boolean" },
    objectiveOk: { type: "boolean" },
    mutationStatus: { type: "string" },
    artifactStatus: { type: "string" },
    readiness: { type: "string" },
    partial: { type: "boolean" },
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

const selectorSchema: JsonObject = {
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

function editOperationSchemas(): JsonObject[] {
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
    name: { type: "string" },
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
  const pickFields = (allowed: string[]): Record<string, unknown> =>
    Object.fromEntries(allowed.map((field) => [field, fields[field as keyof typeof fields]]));
  const op = (
    name: string,
    required: string[],
    allowed: string[],
    extra: Record<string, unknown> = {},
    constraints: Record<string, unknown> = {}
  ): JsonObject => ({
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
    op("pptx.reorderSlides", ["order"], ["order"]),
    op("pptx.addTextbox", ["slide", "text", "bounds"], ["slide", "text", "bounds", "name", "fontSize", "bold"]),
    op("pptx.formatTitle", ["selector"], ["selector", "fontSize", "bold", "textCase"]),
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
    op("pptx.updateChartData", ["selector", "categories", "values"], ["selector", "categories", "values", "seriesName"]),
    op("pptx.setBounds", ["selector", "bounds"], ["selector", "bounds"]),
    op("docx.insertParagraphAfter", ["selector", "text"], ["selector", "text"]),
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
    op("xlsx.insertRows", ["rowIndex", "rows"], ["sheet", "rowIndex", "rows"]),
    op("xlsx.appendRows", ["rows"], ["sheet", "rows"]),
    op("xlsx.setCell", ["cell", "value"], ["sheet", "cell", "value"]),
    op("xlsx.setFormula", ["cell", "formula"], ["sheet", "cell", "formula"]),
    op("xlsx.setRange", ["startCell", "values"], ["sheet", "startCell"], {
      values: { type: "array", minItems: 1, items: { type: "array" } }
    }),
    op("xlsx.updateTable", ["startCell", "rows"], ["sheet", "startCell", "rows"]),
    op("xlsx.writeTable", ["startCell", "rows"], ["sheet", "startCell", "rows", "tableName"]),
    op("xlsx.table.resize", ["selector", "ref"], ["selector", "ref"]),
    op("xlsx.chart.setData", ["selector", "categories", "values"], ["selector", "categories", "values", "seriesName"]),
    op("xlsx.pivot.refreshDefinition", ["selector"], ["selector"]),
    op("xlsx.pivot.refreshAll", [], []),
    op("xlsx.slicer.setSelection", ["selector", "selected"], ["selector", "selected"]),
    op("pdf.textOverlay", ["page", "text", "x", "y"], ["page", "text", "x", "y", "size", "color"]),
    op("pdf.annotation", ["page", "text", "x", "y"], ["page", "text", "x", "y", "width", "height"])
  ];
}

const editOpsSchema: JsonObject = {
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
        expectedInputSha256: { type: "string", pattern: "^sha256:" },
        expectedObjectMapHash: { type: "string", pattern: "^sha256:" },
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

function targetOpsCompatibility(target: string, pattern: string): JsonObject {
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

const documentIrSchema: JsonObject = {
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

const assetSpecSchema: JsonObject = {
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

const designPackSchema: JsonObject = {
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

const templateMapSchema: JsonObject = {
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

const viewObjectMapSchema: JsonObject = {
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

const diagnosticsSchema: JsonObject = {
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

function looseSchema(id: string): JsonObject {
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

const entries: SchemaRegistryEntry[] = [
  entry("officegen.envelope@1.2", envelopeSchema, undefined),
  entry("officegen.capabilities@1.2", looseSchema("officegen.capabilities@1.2"), "capabilities"),
  entry("officegen.edit.ops@1.2", editOpsSchema, "edit"),
  entry("officegen.ir.document@1.2", documentIrSchema, "render"),
  entry("officegen.manifest@1.2", looseSchema("officegen.manifest@1.2"), "run"),
  entry("officegen.asset.spec@1.2", assetSpecSchema, "asset"),
  entry("officegen.design.pack@1.2", designPackSchema, "design"),
  entry("officegen.template.map@1.2", templateMapSchema, "template"),
  entry("officegen.template.candidates.result@2.5", looseSchema("officegen.template.candidates.result@2.5"), "template"),
  entry("officegen.view.objectMap@1.2", viewObjectMapSchema, "view"),
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
  entry("officegen.edit.selectors@1.2", looseSchema("officegen.edit.selectors@1.2"), "edit"),
  entry("officegen.edit.result@1.2", looseSchema("officegen.edit.result@1.2"), "edit"),
  entry("officegen.render.result@1.2", looseSchema("officegen.render.result@1.2"), "render"),
  entry("officegen.export.result@1.2", looseSchema("officegen.export.result@1.2"), "export"),
  entry("officegen.diagnose.result@1.2", looseSchema("officegen.diagnose.result@1.2"), "diagnose"),
  entry("officegen.verify.result@1.2", looseSchema("officegen.verify.result@1.2"), "verify"),
  entry("officegen.repair.result@1.2", looseSchema("officegen.repair.result@1.2"), "repair"),
  entry("officegen.diff.result@1.2", looseSchema("officegen.diff.result@1.2"), "diff"),
  entry("officegen.artifact.manifest@1.2", looseSchema("officegen.artifact.manifest@1.2"), "manifest"),
  entry("officegen.manifest.verify.result@1.2", looseSchema("officegen.manifest.verify.result@1.2"), "manifest"),
  entry("officegen.plan.result@1.2", looseSchema("officegen.plan.result@1.2"), "plan"),
  entry("officegen.rollback.result@1.2", looseSchema("officegen.rollback.result@1.2"), "rollback"),
  entry("officegen.lock@1.2", looseSchema("officegen.lock@1.2"), "lock"),
  entry("officegen.merge.result@1.2", looseSchema("officegen.merge.result@1.2"), "merge"),
  entry("officegen.transaction@1.2", looseSchema("officegen.transaction@1.2"), "rollback"),
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

function entry(id: string, schema: JsonObject, feature?: FeatureName): SchemaRegistryEntry {
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

function schemaVersionFromId(id: string): string {
  const version = id.match(/@([0-9]+(?:\.[0-9]+)*)$/)?.[1];
  if (!version) return "1.2.0";
  const parts = version.split(".");
  while (parts.length < 3) parts.push("0");
  return parts.join(".");
}

export class SchemaRegistry {
  readonly ajv: Ajv;
  private readonly entriesById = new Map<string, SchemaRegistryEntry>();
  private readonly validators = new Map<string, ValidateFunction>();

  constructor(registryEntries: SchemaRegistryEntry[] = entries) {
    this.ajv = new Ajv({ allErrors: true, strict: false });
    for (const registryEntry of registryEntries) {
      this.entriesById.set(registryEntry.id, registryEntry);
      this.validators.set(registryEntry.id, this.ajv.compile(registryEntry.schema));
    }
  }

  list(options: { agent?: boolean; config?: OfficegenConfig } = {}): SchemaRegistryEntry[] {
    return [...this.entriesById.values()].filter((registryEntry) => {
      if (!options.agent) return true;
      if (!registryEntry.visibleToAgents) return false;
      if (!registryEntry.feature || !options.config) return true;
      const feature = options.config.features[registryEntry.feature];
      return feature.enabled && feature.visibleToAgents;
    });
  }

  get(id: string): SchemaRegistryEntry | undefined {
    return this.entriesById.get(id);
  }

  validate(id: string, value: unknown): { ok: true } | { ok: false; errors: ErrorObject[] } {
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
          } as ErrorObject
        ]
      };
    }
    const ok = validate(value);
    return ok ? { ok: true } : { ok: false, errors: cloneErrors(validate.errors ?? []) };
  }
}

function cloneErrors(errors: ErrorObject[]): ErrorObject[] {
  return errors.map((error) => ({
    ...error,
    params: { ...error.params }
  }));
}

export const defaultSchemaRegistry = new SchemaRegistry();

export function listSchemas(options: { agent?: boolean; config?: OfficegenConfig } = {}): SchemaRegistryEntry[] {
  return defaultSchemaRegistry.list(options);
}

export function getSchema(id: string): SchemaRegistryEntry | undefined {
  return defaultSchemaRegistry.get(id);
}

export function validateSchema(id: string, value: unknown): { ok: true } | { ok: false; errors: ErrorObject[] } {
  return defaultSchemaRegistry.validate(id, value);
}

void commonSchemaField;
