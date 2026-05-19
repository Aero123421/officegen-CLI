import type { ObjectBounds, ObjectMapEntry } from "../shared.js";

export const OBJECT_GRAPH_VERSION = "officegen.objectGraph@0.1" as const;

export type ObjectGraphVersion = typeof OBJECT_GRAPH_VERSION;

export type ObjectGraphBBox = [x: number, y: number, width: number, height: number];

export type ObjectGraphRelation = "contains" | "rightOf" | "below";

export interface ObjectGraphEvidence {
  kind: "object-map" | "selector-hint" | "geometry" | "derived";
  confidence: number;
  message: string;
  sourceField?: string;
}

export interface ObjectGraphSource {
  sourcePath?: string;
  xmlPath?: string;
  slide?: number;
  page?: number;
  sheet?: number;
  sheetName?: string;
  story?: string;
}

export interface ObjectGraphText {
  value?: string;
  preview?: string;
  normalized?: string;
}

export interface ObjectGraphProvenance {
  objectMapIndex: number;
  stableObjectId: string;
  selectorHints?: Record<string, unknown>;
  editableOps?: string[];
  media?: Record<string, unknown>;
  trust?: ObjectMapEntry["trust"];
}

export interface ObjectGraphNode {
  graphVersion: ObjectGraphVersion;
  nodeId: string;
  stableId: string;
  type: string;
  label?: string;
  bbox?: ObjectGraphBBox;
  text?: ObjectGraphText;
  style?: Record<string, unknown>;
  source: ObjectGraphSource;
  provenance: ObjectGraphProvenance;
  evidence: ObjectGraphEvidence[];
}

export interface ObjectGraphEdge {
  graphVersion: ObjectGraphVersion;
  edgeId: string;
  from: string;
  to: string;
  relation: ObjectGraphRelation;
  confidence: number;
  evidence: ObjectGraphEvidence[];
}

export interface ObjectGraph {
  graphVersion: ObjectGraphVersion;
  source: {
    objectMapCount: number;
  };
  nodes: ObjectGraphNode[];
  edges: ObjectGraphEdge[];
}

export interface BuildObjectGraphOptions {
  graphVersion?: ObjectGraphVersion;
}

export function buildObjectGraph(objectMap: ObjectMapEntry[], options: BuildObjectGraphOptions = {}): ObjectGraph {
  const graphVersion = options.graphVersion ?? OBJECT_GRAPH_VERSION;
  const nodes = objectMap.map((entry, index) => objectMapEntryToNode(entry, index, graphVersion));
  const edges = buildGeometryEdges(nodes, graphVersion);
  return {
    graphVersion,
    source: {
      objectMapCount: objectMap.length
    },
    nodes,
    edges
  };
}

function objectMapEntryToNode(entry: ObjectMapEntry, index: number, graphVersion: ObjectGraphVersion): ObjectGraphNode {
  const bbox = normalizeBBox(entry);
  const source = sourceFromEntry(entry);
  const style = styleFromEntry(entry);
  const evidence: ObjectGraphEvidence[] = [{
    kind: "object-map",
    confidence: 1,
    message: "Node was normalized from inspect objectMap.",
    sourceField: "objectMap"
  }];
  if (bbox) {
    evidence.push({
      kind: "geometry",
      confidence: 0.95,
      message: "Bounding box was normalized from objectMap geometry.",
      sourceField: entry.bbox ? "bbox" : "bounds"
    });
  }
  for (const key of ["slide", "page", "sheet", "sheetName", "story", "partKind"]) {
    if (entry.selectorHints?.[key] !== undefined) {
      evidence.push({
        kind: "selector-hint",
        confidence: 0.9,
        message: `Source scope includes selector hint ${key}.`,
        sourceField: `selectorHints.${key}`
      });
    }
  }

  return {
    graphVersion,
    nodeId: `node:${String(index + 1).padStart(4, "0")}`,
    stableId: entry.stableObjectId,
    type: entry.kind,
    label: entry.label,
    bbox,
    text: textFromEntry(entry),
    style,
    source,
    provenance: {
      objectMapIndex: index,
      stableObjectId: entry.stableObjectId,
      selectorHints: entry.selectorHints,
      editableOps: entry.editableOps,
      media: entry.media,
      trust: entry.trust
    },
    evidence
  };
}

function normalizeBBox(entry: ObjectMapEntry): ObjectGraphBBox | undefined {
  if (entry.bbox) return entry.bbox;
  return boundsToBBox(entry.bounds);
}

function boundsToBBox(bounds: ObjectBounds | undefined): ObjectGraphBBox | undefined {
  if (!bounds) return undefined;
  return [bounds.x, bounds.y, bounds.width, bounds.height];
}

function textFromEntry(entry: ObjectMapEntry): ObjectGraphText | undefined {
  const value = entry.text;
  const preview = entry.textPreview;
  if (value === undefined && preview === undefined) return undefined;
  const normalized = normalizeText(value ?? preview ?? "");
  return {
    value,
    preview,
    normalized
  };
}

function sourceFromEntry(entry: ObjectMapEntry): ObjectGraphSource {
  const hints = entry.selectorHints ?? {};
  return {
    sourcePath: entry.sourcePath ?? stringHint(hints.sourcePath),
    xmlPath: entry.xmlPath ?? stringHint(hints.xmlPath),
    slide: numberHint(hints.slide),
    page: numberHint(hints.page),
    sheet: numberHint(hints.sheet),
    sheetName: stringHint(hints.sheetName),
    story: stringHint(hints.story ?? hints.partKind)
  };
}

function styleFromEntry(entry: ObjectMapEntry): Record<string, unknown> | undefined {
  const hints = entry.selectorHints ?? {};
  const style = hints.style;
  if (style && typeof style === "object" && !Array.isArray(style)) return style as Record<string, unknown>;
  const font = hints.font;
  const fontSize = hints.fontSize;
  const bold = hints.bold;
  if (font === undefined && fontSize === undefined && bold === undefined) return undefined;
  return { font, fontSize, bold };
}

function buildGeometryEdges(nodes: ObjectGraphNode[], graphVersion: ObjectGraphVersion): ObjectGraphEdge[] {
  const edges: ObjectGraphEdge[] = [];
  for (const from of nodes) {
    if (!from.bbox) continue;
    for (const to of nodes) {
      if (from.nodeId === to.nodeId || !to.bbox || !sameScope(from, to)) continue;
      if (contains(from.bbox, to.bbox)) {
        edges.push(makeGeometryEdge(graphVersion, edges.length, from.nodeId, to.nodeId, "contains", 0.86, "Source bbox contains target bbox."));
      }
      if (isRightOf(from.bbox, to.bbox)) {
        edges.push(makeGeometryEdge(graphVersion, edges.length, from.nodeId, to.nodeId, "rightOf", 0.82, "Source bbox is horizontally right of target bbox."));
      }
      if (isBelow(from.bbox, to.bbox)) {
        edges.push(makeGeometryEdge(graphVersion, edges.length, from.nodeId, to.nodeId, "below", 0.78, "Source bbox is vertically below target bbox."));
      }
    }
  }
  return edges;
}

function makeGeometryEdge(
  graphVersion: ObjectGraphVersion,
  index: number,
  from: string,
  to: string,
  relation: ObjectGraphRelation,
  confidence: number,
  message: string
): ObjectGraphEdge {
  return {
    graphVersion,
    edgeId: `edge:${String(index + 1).padStart(4, "0")}`,
    from,
    to,
    relation,
    confidence,
    evidence: [{
      kind: "geometry",
      confidence,
      message
    }]
  };
}

function sameScope(left: ObjectGraphNode, right: ObjectGraphNode): boolean {
  const leftScope = scopeKey(left);
  const rightScope = scopeKey(right);
  return leftScope !== undefined && leftScope === rightScope;
}

function scopeKey(node: ObjectGraphNode): string | undefined {
  if (node.source.slide !== undefined) return `slide:${node.source.slide}`;
  if (node.source.page !== undefined) return `page:${node.source.page}`;
  if (node.source.sheet !== undefined) return `sheet:${node.source.sheet}`;
  if (node.source.story !== undefined) return `story:${node.source.story}`;
  return undefined;
}

function contains(outer: ObjectGraphBBox, inner: ObjectGraphBBox): boolean {
  const [outerX, outerY, outerWidth, outerHeight] = outer;
  const [innerX, innerY, innerWidth, innerHeight] = inner;
  if (outerWidth * outerHeight <= innerWidth * innerHeight) return false;
  return innerX >= outerX && innerY >= outerY && innerX + innerWidth <= outerX + outerWidth && innerY + innerHeight <= outerY + outerHeight;
}

function isRightOf(source: ObjectGraphBBox, target: ObjectGraphBBox): boolean {
  const [sourceX] = source;
  const [targetX, , targetWidth] = target;
  return sourceX >= targetX + targetWidth && verticalOverlapRatio(source, target) >= 0.25;
}

function isBelow(source: ObjectGraphBBox, target: ObjectGraphBBox): boolean {
  const [, sourceY] = source;
  const [, targetY, , targetHeight] = target;
  return sourceY >= targetY + targetHeight && horizontalOverlapRatio(source, target) >= 0.25;
}

function verticalOverlapRatio(left: ObjectGraphBBox, right: ObjectGraphBBox): number {
  const top = Math.max(left[1], right[1]);
  const bottom = Math.min(left[1] + left[3], right[1] + right[3]);
  const overlap = Math.max(0, bottom - top);
  return overlap / Math.max(1, Math.min(left[3], right[3]));
}

function horizontalOverlapRatio(left: ObjectGraphBBox, right: ObjectGraphBBox): number {
  const start = Math.max(left[0], right[0]);
  const end = Math.min(left[0] + left[2], right[0] + right[2]);
  const overlap = Math.max(0, end - start);
  return overlap / Math.max(1, Math.min(left[2], right[2]));
}

export function normalizeText(value: string): string {
  return value.replace(/\s+/g, " ").trim().toLowerCase();
}

function numberHint(value: unknown): number | undefined {
  const number = Number(value);
  return Number.isFinite(number) ? number : undefined;
}

function stringHint(value: unknown): string | undefined {
  return typeof value === "string" && value.length ? value : undefined;
}
