import { describe, expect, it } from "vitest";
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
      { op: "pdf.textOverlay", page: 1, text: "Overlay", x: 10 },
      { op: "pdf.annotation", page: 1, text: "Note", x: 10, y: 10, color: "#f00" }
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
        "officegen.diagnostics@1.2"
      ])
    );
    expect(agentIds).not.toContain("officegen.design.pack@1.2");
    expect(agentIds).not.toContain("officegen.template.map@1.2");
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
