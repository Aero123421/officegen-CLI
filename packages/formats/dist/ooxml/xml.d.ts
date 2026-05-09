import { XMLBuilder, XMLParser } from "fast-xml-parser";
export declare const parser: XMLParser;
export declare const builder: XMLBuilder;
export declare function parseXml<T = unknown>(xml: string): T;
export declare function buildXml(value: unknown): string;
export declare function ensureArray<T>(value: T | T[] | undefined): T[];
export declare function xmlAttr(attrs: string, name: string): string | undefined;
export declare function localText(xml: string, localName: string): string[];
export declare function exactText(xml: string, tagName: string): string[];
export declare function replaceAllXmlText(input: string, from: string, to: string): string;
export declare function escapeXmlText(value: string): string;
export declare function paragraphXml(text: string, namespace?: "a" | "w"): string;
export declare function bulletParagraphXml(text: string): string;
export declare function stripTags(xml: string): string;
export declare function replaceNthBlock(xml: string, pattern: RegExp, ordinal: number, replacer: (block: string) => string): {
    changed: boolean;
    matched: boolean;
    xml: string;
};
export declare function replaceFirstBlock(xml: string, pattern: RegExp, predicate: (block: string, ordinal: number) => boolean, replacer: (block: string, ordinal: number) => string): {
    changed: boolean;
    matchCount: number;
    xml: string;
};
export declare function setFirstTextInBlock(block: string, tagName: string, text: string): string;
export declare function preview(text: string | undefined, limit?: number): string | undefined;
export declare function bboxFromBounds(bounds?: {
    x: number;
    y: number;
    width: number;
    height: number;
}): [number, number, number, number] | undefined;
export declare function emuToPx(value: number): number;
export declare function pxToEmu(value: number): number;
export declare function escapeRegExp(value: string): string;
export declare function xmlData(value: unknown): string;
