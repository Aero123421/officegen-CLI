import { describe, expect, it } from "vitest";
import {
  OBJECT_GRAPH_VERSION,
  SELECTOR_GRAPH_LOW_CONFIDENCE_THRESHOLD,
  buildObjectGraph,
  resolveGraphSelector,
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
    expect(graph.source.objectMapCount).toBe(2);
    expect(graph.nodes[0]).toMatchObject({
      graphVersion: OBJECT_GRAPH_VERSION,
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
        objectMapIndex: 0,
        stableObjectId: "pptx:slide-1:shape:label"
      }
    });
    expect(graph.nodes[0]?.evidence.some((item) => item.kind === "object-map" && item.confidence === 1)).toBe(true);
    expect(graph.edges).toContainEqual(expect.objectContaining({
      from: "node:0002",
      to: "node:0001",
      relation: "rightOf",
      confidence: 0.82
    }));
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
