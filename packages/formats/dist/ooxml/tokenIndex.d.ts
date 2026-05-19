import { type SourceSpan } from "./sourceSpan.js";
export interface XmlName {
    name: string;
    prefix?: string;
    localName: string;
}
export interface XmlAttributeToken extends XmlName {
    kind: "attribute";
    elementIndex: number;
    span: SourceSpan;
    nameSpan: SourceSpan;
    rawValueSpan: SourceSpan;
    valueSpan: SourceSpan;
    quote: "\"" | "'";
    value: string;
}
export interface XmlElementToken extends XmlName {
    kind: "element";
    index: number;
    depth: number;
    parentIndex?: number;
    span: SourceSpan;
    openTagSpan: SourceSpan;
    closeTagSpan?: SourceSpan;
    contentSpan: SourceSpan;
    selfClosing: boolean;
    attributes: XmlAttributeToken[];
}
export interface XmlTextToken {
    kind: "text";
    span: SourceSpan;
    valueSpan: SourceSpan;
    parentIndex?: number;
    text: string;
    cdata: boolean;
}
export interface XmlWellFormedIssue {
    code: "XML_UNTERMINATED_MARKUP" | "XML_MISMATCHED_CLOSE_TAG" | "XML_UNCLOSED_TAG" | "XML_INVALID_TAG" | "XML_INVALID_ENTITY_REFERENCE";
    message: string;
    offset: number;
}
export declare class TokenIndex {
    readonly source: string;
    readonly elements: XmlElementToken[];
    readonly attributes: XmlAttributeToken[];
    readonly textRuns: XmlTextToken[];
    constructor(source: string, elements: XmlElementToken[], attributes: XmlAttributeToken[], textRuns: XmlTextToken[]);
    findElementsByName(name: string): XmlElementToken[];
    findElementsByLocalName(localName: string): XmlElementToken[];
    findAttributesByName(name: string): XmlAttributeToken[];
    sourceFor(span: SourceSpan): string;
}
export declare function buildTokenIndex(source: string): TokenIndex;
export declare function checkXmlWellFormed(source: string): {
    ok: boolean;
    issues: XmlWellFormedIssue[];
};
