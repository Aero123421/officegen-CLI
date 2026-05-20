import type JSZip from "jszip";
import type { ObjectBounds, ObjectMapEntry } from "../shared.js";
interface PptxTextSemanticRun {
    index: number;
    text: string;
    bold?: boolean;
    italic?: boolean;
    fontSizePt?: number;
    fontFamilyLatin?: string;
    fontFamilyEastAsia?: string;
    fontFamilyComplexScript?: string;
    lang?: string;
    noProof?: boolean;
}
interface PptxTextSemanticParagraph {
    index: number;
    text: string;
    textPreview?: string;
    level?: number;
    bullet?: {
        type: "bullet";
        char?: string;
    };
    numbering?: {
        type: "numbering";
        style?: string;
        startAt?: number;
    };
    runs: PptxTextSemanticRun[];
}
interface PptxTextSemantic extends Record<string, unknown> {
    kind: "pptxText";
    text: {
        plain: string;
        paragraphSeparated: string;
        paragraphCount: number;
        runCount: number;
        hasExplicitLineBreaks: boolean;
        explicitLineBreakCount: number;
    };
    paragraphs: PptxTextSemanticParagraph[];
}
export interface PptxShape {
    stableObjectId: string;
    slideStableObjectId: string;
    slideIndex: number;
    shapeIndex: number;
    shapeId?: string;
    name?: string;
    placeholderType?: string;
    text: string;
    textPreview?: string;
    semantic?: PptxTextSemantic;
    bounds?: ObjectBounds;
    sourcePath: string;
}
export interface PptxSlide {
    stableObjectId: string;
    index: number;
    sourcePath: string;
    relationshipId?: string;
    text: string;
    textObjects: ObjectMapEntry[];
    shapeCount: number;
    pictureCount: number;
    chartCount: number;
    untrusted: true;
}
export declare function getSlidePaths(zip: JSZip): Promise<string[]>;
export declare function inspectSlides(zip: JSZip): Promise<{
    slides: PptxSlide[];
    objectMap: ObjectMapEntry[];
}>;
export declare function extractShapes(xml: string, slideNo: number, slideStableObjectId: string, sourcePath: string): PptxShape[];
export declare function replaceShapeBulletItems(xml: string, ordinal: number, items: string[], mode: "insert" | "replace"): {
    changed: boolean;
    matched: boolean;
    xml: string;
};
export declare function duplicateSlide(zip: JSZip, slideNumber: number, after?: number): Promise<void>;
export declare function addBlankSlide(zip: JSZip, after?: number): Promise<number>;
export declare function addTextBox(zip: JSZip, slideNumber: number, spec: {
    text: string;
    bounds: {
        x: number;
        y: number;
        width: number;
        height: number;
    };
    name?: string;
    fontSize?: number;
    bold?: boolean;
}): Promise<void>;
export declare function reorderSlides(zip: JSZip, order: number[]): Promise<void>;
export {};
