import type JSZip from "jszip";
import type { ObjectMapEntry } from "../shared.js";
export declare const DOCX_STORY_GRAPH_VERSION: "officegen.docx.storyGraph@0.1";
export declare const DOCX_RUN_GRAPH_VERSION: "officegen.docx.runGraph@0.1";
export type DocxStoryKind = "document" | "header" | "footer" | "comments";
export type DocxRunGraphNodeType = "story" | "paragraph" | "run" | "text" | "field" | "hyperlink" | "bookmark" | "commentRange" | "commentReference" | "comment" | "contentControl" | "revision";
export interface DocxStoryNode {
    storyId: string;
    kind: DocxStoryKind;
    partKind: DocxParagraph["partKind"];
    sourcePath: string;
    index: number;
    paragraphIds: string[];
    paragraphCount: number;
    runCount: number;
    textTokenCount: number;
    markers: Record<string, number>;
    untrusted: true;
}
export interface DocxStoryGraph {
    graphVersion: typeof DOCX_STORY_GRAPH_VERSION;
    stories: DocxStoryNode[];
    edges: Array<{
        from: string;
        to: string;
        relation: "contains";
        untrusted: true;
    }>;
    summary: Record<string, number>;
    untrusted: true;
}
export interface DocxRunGraphNode {
    nodeId: string;
    stableObjectId: string;
    type: DocxRunGraphNodeType;
    storyId: string;
    storyKind: DocxStoryKind;
    paragraphId?: string;
    runId?: string;
    sourcePath: string;
    index?: number;
    text?: string;
    textPreview?: string;
    attrs?: Record<string, string | undefined>;
    untrusted: true;
}
export interface DocxRunGraph {
    graphVersion: typeof DOCX_RUN_GRAPH_VERSION;
    nodes: DocxRunGraphNode[];
    edges: Array<{
        from: string;
        to: string;
        relation: "contains" | "next";
        untrusted: true;
    }>;
    summary: Record<string, number>;
    untrusted: true;
}
export interface DocxParagraph {
    stableObjectId: string;
    index: number;
    text: string;
    sourcePath: string;
    partKind: "body" | "header" | "footer" | "comment";
    storyId?: string;
    storyKind?: DocxStoryKind;
    runCount?: number;
    textTokenCount?: number;
    markers?: Record<string, number>;
    untrusted: true;
}
export declare function inspectParagraphs(zip: JSZip): Promise<{
    paragraphs: DocxParagraph[];
    objectMap: ObjectMapEntry[];
    storyGraph: DocxStoryGraph;
    runGraph: DocxRunGraph;
}>;
export declare function setParagraphText(xml: string, ordinal: number, text: string): {
    changed: boolean;
    matched: boolean;
    xml: string;
};
export declare function insertParagraphAfter(xml: string, ordinal: number, text: string): {
    changed: boolean;
    matched: boolean;
    xml: string;
};
export declare function replaceOrCreateHeaderFooter(xml: string | undefined, kind: "header" | "footer", text: string): string;
export declare function commentXml(id: number, author: string, text: string, date?: Date): string;
export declare function insertedParagraphXml(text: string, author?: string, date?: Date, revisionId?: number): string;
