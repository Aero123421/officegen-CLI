import { describe, expect, it } from "vitest";
import { getCapabilities } from "./capabilities.js";
import { getBuiltinConfig } from "./config.js";
import { listSchemas, validateSchema } from "./schemas.js";

describe("schema registry", () => {
  it("validates edit ops strictly", () => {
    const good = {
      schema: "officegen.edit.ops@1.2",
      target: "pptx",
      ops: [
        {
          op: "setText",
          selector: { stableObjectId: "pptx:s001:shape:0007" },
          text: "Hello"
        }
      ]
    };
    const bad = { ...good, unknown: true };

    expect(validateSchema("officegen.edit.ops@1.2", good).ok).toBe(true);
    expect(validateSchema("officegen.edit.ops@1.2", bad).ok).toBe(false);
    expect(validateSchema("officegen.edit.ops@1.2", {
      schema: "officegen.edit.ops@1.2",
      target: "pptx",
      ops: [{ op: "pptx.reorderSlides" }]
    }).ok).toBe(false);
    expect(validateSchema("officegen.edit.ops@1.2", {
      schema: "officegen.edit.ops@1.2",
      target: "xlsx",
      ops: [{ op: "xlsx.setCell", cell: "A1" }]
    }).ok).toBe(false);
    expect(validateSchema("officegen.edit.ops@1.2", {
      schema: "officegen.edit.ops@1.2",
      target: "xlsx",
      ops: [{ op: "xlsx.setCell", cell: "A1", value: "ok" }]
    }).ok).toBe(true);
    expect(validateSchema("officegen.edit.ops@1.2", {
      schema: "officegen.edit.ops@1.2",
      target: "pptx",
      ops: [{ op: "pptx.duplicateSlide", slide: 1 }]
    }).ok).toBe(true);
    expect(validateSchema("officegen.edit.ops@1.2", {
      schema: "officegen.edit.ops@1.2",
      target: "xlsx",
      ops: [{ op: "pptx.duplicateSlide", slide: 1 }]
    }).ok).toBe(false);
    expect(validateSchema("officegen.edit.ops@1.2", {
      schema: "officegen.edit.ops@1.2",
      target: "pdf",
      ops: [{ op: "setText", selector: { stableObjectId: "pdf:document:page:0001" }, text: "Nope" }]
    }).ok).toBe(false);
    expect(validateSchema("officegen.edit.ops@1.2", {
      schema: "officegen.edit.ops@1.2",
      target: "pptx",
      ops: [
        {
          op: "pptx.updateChartData",
          selector: { stableObjectId: "pptx:s001:chart:0001" },
          categories: ["A", "B"],
          values: [1, 2],
          seriesName: "Revenue"
        }
      ]
    }).ok).toBe(true);
    expect(validateSchema("officegen.edit.ops@1.2", {
      schema: "officegen.edit.ops@1.2",
      target: "xlsx",
      ops: [
        {
          op: "xlsx.chart.setData",
          selector: { stableObjectId: "xlsx:workbook:chart:0001" },
          categories: ["A", "B"],
          values: [1, 2],
          seriesName: "Revenue"
        }
      ]
    }).ok).toBe(true);
  });

  it("rejects structural edit ops with missing or unrelated arguments", () => {
    const invalidOps = [
      { op: "pptx.duplicateSlide" },
      { op: "pptx.duplicateSlide", after: 1 },
      { op: "pptx.reorderSlides" },
      { op: "pptx.reorderSlides", order: [1, 2], text: "ignored" },
      { op: "pptx.insertBulletItems", selector: { stableObjectId: "pptx:slide-a1b2c3d4:shape:0001" } },
      { op: "docx.insertParagraphAfter", text: "Inserted" },
      { op: "xlsx.insertRows", rowIndex: 2 },
      { op: "xlsx.updateTable", startCell: "A1" },
      { op: "pptx.updateChartData", selector: { stableObjectId: "pptx:s001:chart:0001" }, categories: ["A"], values: [1], series: [{ name: "A", values: [1] }] },
      { op: "xlsx.chart.setData", selector: { stableObjectId: "xlsx:workbook:chart:0001" }, categories: ["A"], values: [1], secondaryAxis: true },
      { op: "xlsx.pivot.refreshDefinition", selector: { stableObjectId: "xlsx:workbook:pivot:0001" }, fields: ["Region"] },
      { op: "xlsx.slicer.setSelection", selector: { stableObjectId: "xlsx:workbook:slicer:0001" }, selected: ["West"], style: "Fancy" },
      { op: "pdf.textOverlay", page: 1, text: "Overlay", x: 10 },
      { op: "pdf.annotation", page: 1, text: "Note", x: 10, y: 10, color: "#f00" },
      { op: "pdf.redact", page: 1, text: "SECRET" },
      { op: "pdf.redact", page: 1, x: 10, y: 10, width: 20, height: 20 },
      { op: "pdf.textOverlay", page: 1, text: "Overlay", x: 10, y: 10, redact: true },
      { op: "pdf.annotation", page: 1, text: "Note", x: 10, y: 10, removeUnderlyingText: true }
    ];

    for (const editOp of invalidOps) {
      expect(
        validateSchema("officegen.edit.ops@1.2", {
          schema: "officegen.edit.ops@1.2",
          target: "pptx",
          ops: [editOp]
        }).ok
      ).toBe(false);
    }

    expect(
      validateSchema("officegen.edit.ops@1.2", {
        schema: "officegen.edit.ops@1.2",
        target: "pptx",
        ops: [{ op: "pptx.duplicateSlide", selector: { stableObjectId: "pptx:slide-a1b2c3d4:shape:0001" }, after: 1 }]
      }).ok
    ).toBe(true);
  });

  it("keeps unsupported redaction and broad chart contracts out of edit ops", () => {
    const unsupportedOpsByTarget = [
      {
        target: "pdf",
        op: { op: "pdf.redact", page: 1, text: "SECRET" }
      },
      {
        target: "pdf",
        op: { op: "pdf.textOverlay", page: 1, text: "Covered", x: 10, y: 10, removeUnderlyingText: true }
      },
      {
        target: "pptx",
        op: {
          op: "pptx.updateChartData",
          selector: { stableObjectId: "pptx:s001:chart:0001" },
          categories: ["A"],
          values: [1],
          multiSeries: [{ name: "Revenue", values: [1] }]
        }
      },
      {
        target: "xlsx",
        op: {
          op: "xlsx.chart.setData",
          selector: { stableObjectId: "xlsx:workbook:chart:0001" },
          categories: ["A"],
          values: [1],
          comboChart: true,
          chartTypeBySeries: { Revenue: "line" }
        }
      }
    ];

    for (const { target, op } of unsupportedOpsByTarget) {
      expect(
        validateSchema("officegen.edit.ops@1.2", {
          schema: "officegen.edit.ops@1.2",
          target,
          ops: [op]
        }).ok
      ).toBe(false);
    }

    const editOpsSchema = listSchemas().find((entry) => entry.id === "officegen.edit.ops@1.2")?.schema;
    expect(JSON.stringify(editOpsSchema)).not.toContain("pdf.redact");
  });

  it("returns actionable validation failures for unknown schemas and schema id mismatches", () => {
    const unknown = validateSchema("officegen.unknown@1.2", {});
    const mismatch = validateSchema("officegen.edit.ops@1.2", {
      schema: "officegen.ir.document@1.2",
      target: "pptx",
      ops: [{ op: "pptx.setShapeText", selector: { stableObjectId: "pptx:s001:shape:0007" } }]
    });

    expect(unknown.ok).toBe(false);
    if (!unknown.ok) {
      expect(unknown.errors[0]).toMatchObject({ keyword: "schema", params: { id: "officegen.unknown@1.2" } });
    }
    expect(mismatch.ok).toBe(false);
    if (!mismatch.ok) {
      expect(mismatch.errors.some((error) => error.instancePath === "/schema" && error.keyword === "const")).toBe(true);
    }
  });

  it("registers required substrate schemas and filters agent-hidden feature schemas", () => {
    const config = getBuiltinConfig("substrate");
    const ids = listSchemas().map((entry) => entry.id);
    const agentIds = listSchemas({ agent: true, config }).map((entry) => entry.id);

    expect(ids).toEqual(
      expect.arrayContaining([
        "officegen.envelope@1.2",
        "officegen.edit.ops@1.2",
        "officegen.ir.document@1.2",
        "officegen.asset.spec@1.2",
        "officegen.design.pack@1.2",
        "officegen.template.map@1.2",
        "officegen.view.objectMap@1.2",
        "officegen.xlsx.formulaGraph@1.0",
        "officegen.diagnostics@1.2"
      ])
    );
    expect(agentIds).toContain("officegen.design.pack@1.2");
    expect(agentIds).toContain("officegen.template.map@1.2");
  });

  it("validates capability contracts that disclose unsupported and limited editing surfaces", () => {
    const capabilities = getCapabilities(getBuiltinConfig("substrate"), { agent: true });
    const pdfCapabilities = capabilities.formatCapabilities.pdf as Record<string, unknown>;
    const pptxCapabilities = capabilities.formatCapabilities.pptx as Record<string, unknown>;

    expect(validateSchema("officegen.capabilities@1.2", capabilities).ok).toBe(true);
    expect(pdfCapabilities.redaction).toBe("unsupported; overlays do not physically remove underlying content");
    expect(pptxCapabilities.smartArt).toBe("unsupported");
    expect(capabilities.featureContracts).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ area: "PDF editing and redaction", support: "overlay-only" }),
        expect.objectContaining({ area: "PPTX/XLSX charts", support: "limited" })
      ])
    );
    expect(capabilities.unsupportedNow.join("\n")).toContain("Multi-series");
  });

  it("validates document IR and view object maps", () => {
    expect(
      validateSchema("officegen.ir.document@1.2", {
        schema: "officegen.ir.document@1.2",
        targets: ["pptx"],
        sections: [{ title: "Intro", blocks: [{ type: "text", text: "Body" }] }]
      }).ok
    ).toBe(true);

    expect(
      validateSchema("officegen.view.objectMap@1.2", {
        schema: "officegen.view.objectMap@1.2",
        page: 1,
        coordinateSystem: "px",
        fidelity: "approximate",
        objects: [
          {
            stableObjectId: "pptx:s001:shape:0007",
            type: "text",
            bbox: [0, 0, 100, 20],
            editable: true,
            untrusted: true
          }
        ]
      }).ok
    ).toBe(true);
  });

  it("captures scaffold-compatible document IR requirements", () => {
    const scaffoldIr = {
      schema: "officegen.ir.document@1.2",
      metadata: { title: "Proposal", author: "officegen" },
      targets: ["pptx"],
      sections: [{ title: "Proposal", blocks: [{ type: "heading", text: "Proposal" }] }]
    };

    expect(validateSchema("officegen.ir.document@1.2", scaffoldIr).ok).toBe(true);
    expect(
      validateSchema("officegen.ir.document@1.2", {
        schema: "officegen.ir.document@1.2",
        kind: "pptx",
        metadata: { title: "Proposal", author: "officegen" },
        sections: []
      }).ok
    ).toBe(false);
    expect(validateSchema("officegen.ir.document@1.2", { ...scaffoldIr, targets: [] }).ok).toBe(false);
    expect(validateSchema("officegen.ir.document@1.2", { ...scaffoldIr, schema: "officegen.scaffold.result@1.2" }).ok).toBe(false);
  });
});
