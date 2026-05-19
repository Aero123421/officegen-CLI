import type { ObjectMapEntry } from "../shared.js";
export declare const OBJECT_GRAPH_VERSION: "officegen.objectGraph@0.1";
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
export declare function buildObjectGraph(objectMap: ObjectMapEntry[], options?: BuildObjectGraphOptions): ObjectGraph;
export declare function normalizeText(value: string): string;
