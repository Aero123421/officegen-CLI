import type JSZip from "jszip";
import type { ObjectBounds, ObjectMapEntry } from "../shared.js";
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
export declare function reorderSlides(zip: JSZip, order: number[]): Promise<void>;
