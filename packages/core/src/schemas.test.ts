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
      target: "xlsx",
      ops: [{ op: "xlsx.setCell", sheetName: "Data", cell: "A1", value: "ok" }]
    }).ok).toBe(true);
    expect(validateSchema("officegen.edit.ops@1.2", {
      schema: "officegen.edit.ops@1.2",
      target: "xlsx",
      ops: [{ op: "xlsx.setCell", sheet: "Data", cell: "A1", value: "ok" }]
    }).ok).toBe(true);
    expect(validateSchema("officegen.edit.ops@1.2", {
      schema: "officegen.edit.ops@1.2",
      target: "pptx",
      options: {
        expectedObjectGraphHash: "sha256:abc",
        selectionLock: {
          objectGraphHash: "sha256:abc",
          nodeId: "node:0001",
          sourceFingerprint: "sha256:def"
        }
      },
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

  it("summarizes oneOf edit-op failures with the closest operation schema", () => {
    const validation = validateSchema("officegen.edit.ops@1.2", {
      schema: "officegen.edit.ops@1.2",
      target: "xlsx",
      ops: [{ op: "xlsx.setCell", sheetName: "Data", cell: "A1", values: [["wrong"]] }]
    }, { diagnostics: true });

    expect(validation.ok).toBe(false);
    if (!validation.ok) {
      expect(validation.diagnostics?.[0]).toMatchObject({
        instancePath: "/ops/0",
        bestMatch: { op: "xlsx.setCell" },
        missing: ["value"],
        unexpected: ["values"]
      });
      expect(validation.diagnostics?.[0]?.expectedTypes.value).toEqual(["string", "number", "boolean", "null"]);
    }
  });

  it("registers required substrate schemas and filters agent-hidden feature schemas", () => {
    const config = getBuiltinConfig("substrate");
    const ids = listSchemas().map((entry) => entry.id);
    const agentIds = listSchemas({ agent: true, config }).map((entry) => entry.id);

    expect(ids).toEqual(
      expect.arrayContaining([
        "officegen.envelope@1.2",
        "officegen.envelope@2",
        "officegen.edit.ops@1.2",
        "officegen.ir.document@1.2",
        "officegen.asset.spec@1.2",
        "officegen.design.pack@1.2",
        "officegen.template.map@1.2",
        "officegen.view.objectMap@1.2",
        "officegen.selectorResolution@2",
        "officegen.objectGraph@2",
        "officegen.verify@2",
        "officegen.repairPlan@2",
        "officegen.xlsx.formulaGraph@1.0",
        "officegen.diagnostics@1.2"
      ])
    );
    expect(agentIds).toContain("officegen.design.pack@1.2");
    expect(agentIds).toContain("officegen.template.map@1.2");
  });

  it("validates v2 edit and patch plan contracts", () => {
    expect(validateSchema("officegen.editPlan@2", {
      schema: "officegen.editPlan@2",
      target: "pptx",
      inputSha256: "sha256:abc",
      objectGraphHash: "sha256:def",
      wouldWrite: false,
      operations: [{ op: "replaceText", from: "A", to: "B" }],
      selectorResolution: {
        schema: "officegen.selectorResolution@2",
        status: "matched",
        candidates: [],
        evidence: [],
        nextActions: [],
        selectionLock: { objectGraphHash: "sha256:def" }
      },
      selectorResolutions: {
        schema: "officegen.edit.selectors@1.2",
        resolutions: []
      }
    }).ok).toBe(true);

    expect(validateSchema("officegen.editPlan@2", {
      schema: "officegen.editPlan@2",
      target: "pptx",
      wouldWrite: false,
      operations: [],
      selectorResolution: {
        schema: "officegen.edit.selectors@1.2",
        resolutions: []
      }
    }).ok).toBe(false);

    expect(validateSchema("officegen.patchPlan@2", {
      schema: "officegen.patchPlan@2",
      format: "pptx",
      wouldWrite: false,
      inputSha256: `sha256:${"a".repeat(64)}`,
      sourceFingerprint: { algorithm: "sha256", hash: "b".repeat(64), byteLength: 10 },
      operations: [{ operationIndex: 0, op: "replaceText", wouldApply: true }],
      touchedParts: [{
        path: "ppt/slides/slide1.xml",
        change: "modified",
        beforeSha256: `sha256:${"c".repeat(64)}`,
        afterSha256: `sha256:${"d".repeat(64)}`,
        sourceFingerprint: { algorithm: "sha256", hash: "e".repeat(64), byteLength: 20, path: "ppt/slides/slide1.xml" }
      }],
      expectedChangedParts: ["ppt/slides/slide1.xml"],
      sourceFingerprints: [{ algorithm: "sha256", hash: "e".repeat(64), byteLength: 20, path: "ppt/slides/slide1.xml" }],
      blocked: []
    }).ok).toBe(true);

    expect(validateSchema("officegen.patchPlan@2", {
      schema: "officegen.patchPlan@2",
      format: "pptx",
      wouldWrite: true
    }).ok).toBe(false);
  });

  it("validates office-agent v3.1 runtime skeleton contracts", () => {
    const phase = (index: number) => ({
      id: `phase-${String(index).padStart(2, "0")}-test`,
      standardName: index === 9 ? "verify" : `phase-${index}`,
      manifestRole: `role-${index}`,
      commandTemplate: "officegen help --json",
      mutatesOffice: false,
      status: "skeleton",
      execution: "skeleton"
    });
    const phases = Array.from({ length: 13 }, (_, index) => phase(index + 1));
    expect(validateSchema("officegen.office-agent.manifest@3.1", {
      schema: "officegen.office-agent.manifest@3.1",
      release: "3.1.0",
      runtimeProjection: "runtime-v2",
      mode: "skeleton-evidence",
      status: "skeleton",
      phaseCount: 13,
      phases,
      limitations: ["skeleton only"],
      requiredPhaseNames: ["inspect", "select", "plan", "dry-run", "edit", "verify", "diff", "repair", "report"]
    }).ok).toBe(true);

    expect(validateSchema("officegen.office-agent.workflow@3.1", {
      schema: "officegen.office-agent.workflow@3.1",
      release: "3.1.0",
      runtimeProjection: "runtime-v2",
      phaseCount: 13,
      skeletonOnly: false,
      steps: phases
    }).ok).toBe(false);
  });

  it("validates v2 repair plan taxonomy and post-repair verify notes", () => {
    const plan = {
      schema: "officegen.repairPlan@2",
      version: 2,
      target: "pptx",
      inputSha256: "sha256:abc",
      wouldWrite: false,
      operations: [{ op: "setText", selector: { stableObjectId: "pptx:s001:shape:0001" }, text: "Short" }],
      failureTaxonomy: [{
        code: "TEXT_OVERFLOW_RISK",
        category: "quality",
        severity: "warning",
        autoRepairable: true,
        evidence: [{ kind: "diagnose-issue", message: "Text object is long enough to risk overflow.", issueCode: "TEXT_OVERFLOW_RISK" }],
        nextCommand: "officegen repair input.pptx --dry-run --json"
      }],
      steps: [{ id: "post-repair-verify", command: "officegen verify repaired.pptx --visual --json", dryRun: true, reason: "Run after writing." }],
      verify: {
        status: "not_run",
        requiredAfterRepair: true,
        command: "officegen verify repaired.pptx --visual --json",
        readinessNote: "Post-repair verify has not been run."
      }
    };

    expect(validateSchema("officegen.repairPlan@2", plan).ok).toBe(true);
    expect(validateSchema("officegen.repairPlan@2", { ...plan, wouldWrite: true }).ok).toBe(false);
  });

  it("validates capability contracts that disclose unsupported and limited editing surfaces", () => {
    const capabilities = getCapabilities(getBuiltinConfig("substrate"), { agent: true });
    const pdfCapabilities = capabilities.formatCapabilities.pdf as Record<string, unknown>;
    const pptxCapabilities = capabilities.formatCapabilities.pptx as Record<string, unknown>;
    const runtimeProfiles = capabilities.runtimeProfiles as Record<string, { capabilities: Array<Record<string, unknown>> }>;
    const currentProfile = runtimeProfiles["current-limited-v3.1"];
    const targetProfile = runtimeProfiles["perfect-runtime-target"];

    expect(validateSchema("officegen.capabilities@1.2", capabilities).ok).toBe(true);
    expect(pdfCapabilities.redaction).toBe("unsupported; overlays do not physically remove underlying content");
    expect(pptxCapabilities.smartArt).toBe("unsupported");
    expect(capabilities.specProfile).toMatchObject({
      currentProfileId: "current-limited-v3.1",
      targetProfileId: "perfect-runtime-target",
      runtimeProjection: "runtime-v2"
    });
    expect(currentProfile.capabilities).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ id: "runtime-v2-projections", support: "supported" }),
        expect.objectContaining({ id: "smartart-editing", support: "unsupported" }),
        expect.objectContaining({ id: "pdf-true-redaction", support: "unsupported" })
      ])
    );
    expect(targetProfile.capabilities).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ id: "smartart-editing", support: "target-only" }),
        expect.objectContaining({ id: "pdf-true-redaction", support: "target-only" })
      ])
    );
    expect(capabilities.featureContracts).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ area: "PDF editing and redaction", support: "overlay-only" }),
        expect.objectContaining({ area: "PPTX/XLSX charts", support: "limited" })
      ])
    );
    expect(capabilities.unsupportedNow.join("\n")).toContain("Multi-series");
  });

  it("validates VerificationReportV2 projection schema", () => {
    const report = {
      schema: "officegen.verify@2",
      version: 2,
      format: "pptx",
      readiness: "pass",
      score: 1,
      partial: false,
      gates: Object.fromEntries(["schema", "package", "semantic", "visual", "native", "security", "accessibility", "goal"].map((name) => [
        name,
        { status: name === "native" ? "skipped" : "pass", issues: [] }
      ])),
      issues: [],
      artifacts: [{ artifactId: "verify-native-pdf", role: "native-render", managed: true, format: "pdf" }],
      recommendedRepairs: []
    };

    expect(validateSchema("officegen.verify@2", report).ok).toBe(true);
    expect(validateSchema("officegen.verify@2", { ...report, schema: "officegen.verify.result@1.2" }).ok).toBe(false);
  });

  it("validates document IR, view object maps, and object graphs", () => {
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

    expect(
      validateSchema("officegen.objectGraph@2", {
        schema: "officegen.objectGraph@2",
        version: 2,
        graphVersion: "officegen.objectGraph@2",
        source: { objectMapCount: 1, builder: "inspect.objectMap" },
        provenance: { generatedFrom: "officegen.inspect.result@1.2", sourceField: "objectMap" },
        confidence: 1,
        riskFlags: [],
        pagination: {
          nodeOffset: 0,
          nodeLimit: 1,
          nodeCount: 1,
          totalNodes: 1,
          edgeOffset: 0,
          edgeLimit: 0,
          edgeCount: 0,
          totalEdges: 0,
          truncated: false
        },
        index: {
          nodesByStableId: { "pptx:s001:shape:0007": "node:0001" },
          nodesByType: { shape: ["node:0001"] },
          edgesByRelation: { contains: [], rightOf: [], below: [] }
        },
        nodes: [{
          schema: "officegen.objectGraph@2",
          version: 2,
          graphVersion: "officegen.objectGraph@2",
          index: 0,
          nodeId: "node:0001",
          stableId: "pptx:s001:shape:0007",
          type: "shape",
          source: { slide: 1 },
          provenance: {
            schema: "officegen.objectGraph@2",
            source: "inspect.objectMap",
            objectMapIndex: 0,
            stableObjectId: "pptx:s001:shape:0007"
          },
          confidence: 1,
          riskFlags: [],
          evidence: [{ kind: "object-map", confidence: 1, message: "test" }]
        }],
        edges: []
      }).ok
    ).toBe(true);

    expect(validateSchema("officegen.selectorResolution@2", {
      schema: "officegen.selectorResolution@2",
      status: "matched",
      confidence: 0.91,
      candidates: [{ nodeId: "node:0001", stableObjectId: "pptx:s001:shape:0007", type: "shape", confidence: 0.91 }],
      evidence: [{ kind: "object-map", confidence: 1, message: "Matched stable object id." }],
      nextActions: ["Use selectionLock with edit options to block stale mutations."],
      selectionLock: {
        objectGraphHash: "sha256:abc",
        nodeId: "node:0001",
        sourceFingerprint: "sha256:def"
      }
    }).ok).toBe(true);
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
