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
          op: "pptx.setShapeText",
          selector: { stableObjectId: "pptx:s001:shape:0007" },
          text: "Hello"
        }
      ]
    };
    const bad = { ...good, unknown: true };

    expect(validateSchema("officegen.edit.ops@1.2", good).ok).toBe(true);
    expect(validateSchema("officegen.edit.ops@1.2", bad).ok).toBe(false);
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
});
