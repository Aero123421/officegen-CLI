export interface SourceSpan {
    /** UTF-8 byte offset, inclusive. */
    start: number;
    /** UTF-8 byte offset, exclusive. */
    end: number;
    /** JavaScript string offset, inclusive. */
    charStart: number;
    /** JavaScript string offset, exclusive. */
    charEnd: number;
}
export interface SourceFingerprint {
    algorithm: "sha256";
    hash: string;
    span: SourceSpan;
    byteLength: number;
}
export declare function sourceSpanFromCharRange(source: string, charStart: number, charEnd: number): SourceSpan;
export declare function sourceSpanFromByteRange(source: string, start: number, end: number): SourceSpan;
export declare function sliceSource(source: string, span: SourceSpan): string;
export declare function createSourceFingerprint(source: string, span: SourceSpan): SourceFingerprint;
export declare function verifySourceFingerprint(source: string, fingerprint: SourceFingerprint): boolean;
export declare function charIndexToByteOffset(source: string, charIndex: number): number;
export declare function byteOffsetToCharIndex(source: string, byteOffset: number): number;
export declare function isValidByteRange(source: string, start: number, end: number): boolean;
