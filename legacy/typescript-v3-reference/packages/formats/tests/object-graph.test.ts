import { describe, expect, it } from "vitest";
import { validateSchema } from "@officegen/core";
import {
  OBJECT_GRAPH_VERSION,
  SELECTOR_GRAPH_LOW_CONFIDENCE_THRESHOLD,
  buildObjectGraph,
  resolveGraphSelector,
  selectorResolutionV2FromGraphResolution,
  type ObjectMapEntry,
  type SelectorGraphStatus
} from "../src/index.js";

describe("ObjectGraph and SelectorGraph", () => {
  it("normalizes objectMap entries into graph nodes with provenance, evidence, and edges", () => {
    const graph = buildObjectGraph([
      objectEntry("label", "shape", "Revenue", [10, 10, 80, 20]),
      objectEntry("value", "shape", "$10M", [120, 10, 80, 20])
    ]);

    expect(graph.graphVersion).toBe(OBJECT_GRAPH_VERSION);
    expect(graph.schema).toBe("officegen.objectGraph@2");
    expect(graph.version).toBe(2);
    expect(graph.source.objectMapCount).toBe(2);
    expect(graph.pagination).toMatchObject({
      nodeOffset: 0,
      nodeCount: 2,
      totalNodes: 2,
      truncated: false
    });
    expect(graph.nodes[0]).toMatchObject({
      schema: "officegen.objectGraph@2",
      version: 2,
      graphVersion: OBJECT_GRAPH_VERSION,
      index: 0,
      nodeId: "node:0001",
      stableId: "pptx:slide-1:shape:label",
      type: "shape",
      bbox: [10, 10, 80, 20],
      text: {
        value: "Revenue",
        normalized: "revenue"
      },
      source: {
        slide: 1,
        sourcePath: "ppt/slides/slide1.xml"
      },
      provenance: {
        schema: "officegen.objectGraph@2",
        source: "inspect.objectMap",
        objectMapIndex: 0,
        stableObjectId: "pptx:slide-1:shape:label"
      },
      confidence: expect.any(Number),
      riskFlags: expect.arrayContaining([
        expect.objectContaining({ code: "UNTRUSTED_DOCUMENT_CONTENT" })
      ])
    });
    expect(validateSchema("officegen.objectGraph@2", graph).ok).toBe(true);
    expect(graph.nodes[0]?.evidence.some((item) => item.kind === "object-map" && item.confidence === 1)).toBe(true);
    expect(graph.edges).toContainEqual(expect.objectContaining({
      schema: "officegen.objectGraph@2",
      version: 2,
      from: "node:0002",
      to: "node:0001",
      relation: "rightOf",
      confidence: 0.82
    }));
  });

  it("paginates object graph nodes while preserving source totals and indexes", () => {
    const graph = buildObjectGraph([
      objectEntry("one", "shape", "One", [10, 10, 80, 20]),
      objectEntry("two", "shape", "Two", [120, 10, 80, 20]),
      objectEntry("three", "shape", "Three", [230, 10, 80, 20])
    ], { nodeOffset: 1, nodeLimit: 1, riskFlags: [{ code: "TEST_RISK", severity: "info", message: "test" }] });

    expect(graph.nodes.map((node) => node.index)).toEqual([1]);
    expect(graph.pagination).toMatchObject({
      nodeOffset: 1,
      nodeLimit: 1,
      nodeCount: 1,
      totalNodes: 3,
      truncated: true,
      nextNodeOffset: 2
    });
    expect(graph.riskFlags).toEqual([{ code: "TEST_RISK", severity: "info", message: "test" }]);
    expect(validateSchema("officegen.objectGraph@2", graph).ok).toBe(true);
  });

  it("resolves text, type, slide, bbox, relation, rightOf, and nearestTo selectors with evidence", () => {
    const graph = buildObjectGraph([
      objectEntry("label", "shape", "Revenue", [10, 10, 80, 20]),
      objectEntry("value", "shape", "$10M", [120, 10, 80, 20]),
      objectEntry("chart", "chart", "Revenue chart", [10, 80, 200, 120])
    ]);

    const textTypeSlide = resolveGraphSelector(graph, { text: "revenue chart", type: "chart", slide: 1 });
    expect(textTypeSlide.status).toBe("matched");
    expect(textTypeSlide.matches[0]?.stableId).toBe("pptx:slide-1:shape:chart");
    expect(textTypeSlide.matches[0]?.evidence.map((item) => item.sourceField)).toEqual(expect.arrayContaining(["type", "source.slide", "text"]));

    const bbox = resolveGraphSelector(graph, { bbox: { x: 8, y: 78, width: 205, height: 125, mode: "intersects" }, type: "chart" });
    expect(bbox.status).toBe("matched");
    expect(bbox.matches[0]?.stableId).toBe("pptx:slide-1:shape:chart");

    const relation = resolveGraphSelector(graph, { relation: { relation: "rightOf", anchor: { text: { text: "Revenue", exact: true } } } });
    expect(relation.status).toBe("matched");
    expect(relation.matches[0]?.stableId).toBe("pptx:slide-1:shape:value");

    const rightOf = resolveGraphSelector(graph, { slide: 1, rightOf: { text: "Revenue", type: "shape" } });
    expect(rightOf.status).toBe("matched");
    expect(rightOf.matches[0]?.text).toBe("$10M");
    expect(rightOf.confidence).toBeGreaterThanOrEqual(SELECTOR_GRAPH_LOW_CONFIDENCE_THRESHOLD);

    const nearestTo = resolveGraphSelector(graph, { nearestTo: { slide: 1, x: 160, y: 20 } });
    expect(nearestTo.status).toBe("matched");
    expect(nearestTo.matches[0]?.stableId).toBe("pptx:slide-1:shape:value");
    expect(nearestTo.matches[0]?.evidence.some((item) => item.sourceField === "nearestTo")).toBe(true);

    const runtimeResolution = selectorResolutionV2FromGraphResolution(graph, nearestTo);
    expect(runtimeResolution).toMatchObject({
      schema: "officegen.selectorResolution@2",
      status: "matched",
      selectionLock: {
        objectGraphHash: expect.stringMatching(/^sha256:/),
        nodeId: "node:0002",
        sourceFingerprint: expect.stringMatching(/^sha256:/)
      }
    });
  });

  it("fixes low-confidence selector status as a public type contract", () => {
    const graph = buildObjectGraph([
      objectEntry("only-shape", "shape", "Only shape", [10, 10, 80, 20])
    ]);

    const resolution = resolveGraphSelector(graph, { type: "shape" });
    const status: SelectorGraphStatus = resolution.status;

    expect(status).toBe("low-confidence");
    expect(resolution.lowConfidence).toBe(true);
    expect(resolution.ambiguity.ambiguous).toBe(false);
    expect(resolution.confidence).toBeLessThan(SELECTOR_GRAPH_LOW_CONFIDENCE_THRESHOLD);
  });

  it("fixes ambiguous selector status and candidate ids", () => {
    const graph = buildObjectGraph([
      objectEntry("north", "shape", "Total", [10, 10, 80, 20]),
      objectEntry("south", "shape", "Total", [10, 50, 80, 20])
    ]);

    const resolution = resolveGraphSelector(graph, { text: { text: "Total", exact: true } });

    expect(resolution.status).toBe("ambiguous");
    expect(resolution.matched).toBe(false);
    expect(resolution.ambiguity).toEqual({
      ambiguous: true,
      reason: "multiple-matches",
      candidateNodeIds: ["node:0001", "node:0002"]
    });
  });

  it("marks close nearestTo candidates as ambiguous", () => {
    const graph = buildObjectGraph([
      objectEntry("left", "shape", "Left", [10, 10, 20, 20]),
      objectEntry("right", "shape", "Right", [50, 10, 20, 20])
    ]);

    const resolution = resolveGraphSelector(graph, { nearestTo: { slide: 1, x: 40, y: 20 } });

    expect(resolution.status).toBe("ambiguous");
    expect(resolution.ambiguity.reason).toBe("nearest-distance-tie");
    expect(resolution.matches.map((match) => match.stableId)).toEqual([
      "pptx:slide-1:shape:left",
      "pptx:slide-1:shape:right"
    ]);
  });

  it("propagates ambiguous relation anchors to generic relation selectors", () => {
    const graph = buildObjectGraph([
      objectEntry("label-a", "shape", "Revenue", [10, 10, 80, 20]),
      objectEntry("label-b", "shape", "Revenue", [10, 50, 80, 20]),
      objectEntry("value", "shape", "$10M", [120, 10, 80, 20])
    ]);

    const resolution = resolveGraphSelector(graph, {
      relation: { relation: "rightOf", anchor: { text: { text: "Revenue", exact: true } } }
    });

    expect(resolution.status).toBe("ambiguous");
    expect(resolution.ambiguity).toEqual({
      ambiguous: true,
      reason: "relation-anchor-ambiguous",
      candidateNodeIds: ["node:0001", "node:0002"]
    });
  });
});

function objectEntry(suffix: string, kind: string, text: string, bbox: [number, number, number, number]): ObjectMapEntry {
  return {
    stableObjectId: `pptx:slide-1:shape:${suffix}`,
    kind,
    text,
    textPreview: text,
    sourcePath: "ppt/slides/slide1.xml",
    bbox,
    selectorHints: {
      slide: 1
    },
    untrusted: true
  };
}
