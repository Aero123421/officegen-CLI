import type { ObjectMapEntry } from "../shared.js";
export declare const OBJECT_GRAPH_VERSION: "officegen.objectGraph@2";
export type ObjectGraphVersion = typeof OBJECT_GRAPH_VERSION;
export type ObjectGraphBBox = [x: number, y: number, width: number, height: number];
export type ObjectGraphRelation = "contains" | "rightOf" | "below";
export type ObjectGraphRiskSeverity = "info" | "warning" | "error";
export interface ObjectGraphEvidence {
    kind: "object-map" | "selector-hint" | "geometry" | "derived";
    confidence: number;
    message: string;
    sourceField?: string;
}
export interface ObjectGraphRiskFlag {
    code: string;
    severity: ObjectGraphRiskSeverity;
    message: string;
    source?: string;
}
export interface ObjectGraphSource {
    format?: string;
    inputPath?: string;
    inputSha256?: string;
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
    schema: ObjectGraphVersion;
    source: "inspect.objectMap";
    objectMapIndex: number;
    stableObjectId: string;
    selectorHints?: Record<string, unknown>;
    editableOps?: string[];
    media?: Record<string, unknown>;
    trust?: ObjectMapEntry["trust"];
}
export interface ObjectGraphNode {
    schema: ObjectGraphVersion;
    version: 2;
    graphVersion: ObjectGraphVersion;
    index: number;
    nodeId: string;
    stableId: string;
    type: string;
    label?: string;
    bbox?: ObjectGraphBBox;
    text?: ObjectGraphText;
    style?: Record<string, unknown>;
    source: ObjectGraphSource;
    provenance: ObjectGraphProvenance;
    confidence: number;
    riskFlags: ObjectGraphRiskFlag[];
    evidence: ObjectGraphEvidence[];
}
export interface ObjectGraphEdge {
    schema: ObjectGraphVersion;
    version: 2;
    graphVersion: ObjectGraphVersion;
    index: number;
    edgeId: string;
    from: string;
    to: string;
    relation: ObjectGraphRelation;
    confidence: number;
    riskFlags: ObjectGraphRiskFlag[];
    evidence: ObjectGraphEvidence[];
}
export interface ObjectGraph {
    schema: ObjectGraphVersion;
    version: 2;
    graphVersion: ObjectGraphVersion;
    source: {
        format?: string;
        inputPath?: string;
        inputSha256?: string;
        objectMapCount: number;
        builder: "inspect.objectMap";
    };
    provenance: {
        generatedFrom: "officegen.inspect.result@1.2";
        sourceField: "objectMap";
    };
    confidence: number;
    riskFlags: ObjectGraphRiskFlag[];
    pagination: {
        nodeOffset: number;
        nodeLimit: number;
        nodeCount: number;
        totalNodes: number;
        edgeOffset: number;
        edgeLimit: number;
        edgeCount: number;
        totalEdges: number;
        truncated: boolean;
        nextNodeOffset?: number;
        nextEdgeOffset?: number;
    };
    index: {
        nodesByStableId: Record<string, string>;
        nodesByType: Record<string, string[]>;
        edgesByRelation: Record<ObjectGraphRelation, string[]>;
    };
    nodes: ObjectGraphNode[];
    edges: ObjectGraphEdge[];
}
export interface BuildObjectGraphOptions {
    graphVersion?: ObjectGraphVersion;
    format?: string;
    inputPath?: string;
    inputSha256?: string;
    nodeOffset?: number;
    nodeLimit?: number;
    edgeOffset?: number;
    edgeLimit?: number;
    riskFlags?: ObjectGraphRiskFlag[];
}
export declare function buildObjectGraph(objectMap: ObjectMapEntry[], options?: BuildObjectGraphOptions): ObjectGraph;
export declare function normalizeText(value: string): string;
