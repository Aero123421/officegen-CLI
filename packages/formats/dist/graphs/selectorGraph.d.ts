import { type ObjectGraph, type ObjectGraphBBox, type ObjectGraphEvidence, type ObjectGraphNode, type ObjectGraphRelation } from "./objectGraph.js";
export declare const SELECTOR_GRAPH_LOW_CONFIDENCE_THRESHOLD = 0.65;
export declare const SELECTOR_GRAPH_AMBIGUITY_DISTANCE_DELTA = 24;
export declare const SELECTOR_GRAPH_AMBIGUITY_DISTANCE_RATIO = 1.15;
export type SelectorGraphStatus = "matched" | "not-found" | "ambiguous" | "low-confidence";
export type SelectorResolutionV2Status = "matched" | "not_found" | "ambiguous" | "low_confidence" | "stale" | "unsupported";
export type SelectorGraphTextSelector = string | {
    text: string;
    exact?: boolean;
    caseSensitive?: boolean;
};
export type SelectorGraphBBoxSelector = ObjectGraphBBox | {
    x: number;
    y: number;
    width: number;
    height: number;
    tolerance?: number;
    mode?: "near" | "intersects" | "contains";
};
export interface SelectorGraphRelationSelector {
    relation: ObjectGraphRelation;
    anchor?: SelectorGraphSelector;
    nodeId?: string;
    stableId?: string;
    direction?: "outgoing" | "incoming";
}
export interface SelectorGraphSelector {
    nodeId?: string;
    stableId?: string;
    text?: SelectorGraphTextSelector;
    type?: string | string[];
    slide?: number;
    page?: number;
    sheet?: number;
    bbox?: SelectorGraphBBoxSelector;
    relation?: SelectorGraphRelationSelector;
    rightOf?: string | {
        text: string;
        slide?: number;
        type?: string | string[];
    };
    nearestTo?: {
        slide?: number;
        x: number;
        y: number;
        maxDistance?: number;
    };
}
export interface SelectorGraphMatch {
    nodeId: string;
    stableId: string;
    type: string;
    label?: string;
    text?: string;
    bbox?: ObjectGraphBBox;
    source: ObjectGraphNode["source"];
    confidence: number;
    evidence: ObjectGraphEvidence[];
}
export interface SelectorGraphAmbiguity {
    ambiguous: boolean;
    reason?: "multiple-matches" | "nearest-distance-tie" | "rightOf-anchor-ambiguous" | "relation-anchor-ambiguous";
    candidateNodeIds: string[];
}
export interface SelectorGraphResolution {
    graphVersion: ObjectGraph["graphVersion"];
    selector: SelectorGraphSelector;
    status: SelectorGraphStatus;
    matched: boolean;
    matchCount: number;
    confidence?: number;
    lowConfidence: boolean;
    ambiguity: SelectorGraphAmbiguity;
    matches: SelectorGraphMatch[];
    evidence: ObjectGraphEvidence[];
}
export interface SelectorSelectionLock {
    objectGraphHash: string;
    nodeId?: string;
    sourceFingerprint?: string;
}
export interface SelectorResolutionV2Candidate {
    nodeId?: string;
    stableObjectId: string;
    type: string;
    label?: string;
    text?: string;
    confidence?: number;
    source?: ObjectGraphNode["source"];
}
export interface SelectorResolutionV2 {
    schema: "officegen.selectorResolution@2";
    status: SelectorResolutionV2Status;
    confidence?: number;
    candidates: SelectorResolutionV2Candidate[];
    evidence: ObjectGraphEvidence[];
    ambiguityReason?: string;
    nextActions: string[];
    selectionLock: SelectorSelectionLock;
}
export declare function resolveGraphSelector(graph: ObjectGraph, selector: SelectorGraphSelector): SelectorGraphResolution;
export declare function resolveGraphSelectors(graph: ObjectGraph, selectors: SelectorGraphSelector[]): SelectorGraphResolution[];
export declare function selectorResolutionV2FromGraphResolution(graph: ObjectGraph, resolution: SelectorGraphResolution): SelectorResolutionV2;
export declare function selectorResolutionV2Status(status: SelectorGraphStatus): SelectorResolutionV2Status;
export declare function objectGraphHash(graph: ObjectGraph): string;
export declare function selectionLockForNode(graph: ObjectGraph, node: ObjectGraphNode | undefined): SelectorSelectionLock;
export declare function sourceFingerprintForNode(node: ObjectGraphNode): string;
export declare function selectorResolutionNextActions(status: SelectorResolutionV2Status): string[];
