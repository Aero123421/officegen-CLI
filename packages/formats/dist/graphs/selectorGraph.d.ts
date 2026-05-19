import { type ObjectGraph, type ObjectGraphBBox, type ObjectGraphEvidence, type ObjectGraphNode, type ObjectGraphRelation } from "./objectGraph.js";
export declare const SELECTOR_GRAPH_LOW_CONFIDENCE_THRESHOLD = 0.65;
export declare const SELECTOR_GRAPH_AMBIGUITY_DISTANCE_DELTA = 24;
export declare const SELECTOR_GRAPH_AMBIGUITY_DISTANCE_RATIO = 1.15;
export type SelectorGraphStatus = "matched" | "not-found" | "ambiguous" | "low-confidence";
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
export declare function resolveGraphSelector(graph: ObjectGraph, selector: SelectorGraphSelector): SelectorGraphResolution;
export declare function resolveGraphSelectors(graph: ObjectGraph, selectors: SelectorGraphSelector[]): SelectorGraphResolution[];
