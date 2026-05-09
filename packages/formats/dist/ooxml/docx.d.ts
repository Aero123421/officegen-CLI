import type JSZip from "jszip";
import type { ObjectMapEntry } from "../shared.js";
export interface DocxParagraph {
    stableObjectId: string;
    index: number;
    text: string;
    sourcePath: string;
    untrusted: true;
}
export declare function inspectParagraphs(zip: JSZip): Promise<{
    paragraphs: DocxParagraph[];
    objectMap: ObjectMapEntry[];
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
