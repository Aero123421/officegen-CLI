export interface PdfTextBlock {
    page: number;
    index: number;
    text: string;
    x?: number;
    y?: number;
    width?: number;
    height?: number;
    source: "pdfjs" | "content-stream";
    untrusted: true;
}
export interface PdfAnnotationSummary {
    page: number;
    index: number;
    subtype?: string;
    contents?: string;
    rect?: number[];
    hasAppearance?: boolean;
    untrusted: true;
}
export interface PdfRiskFlag {
    code: string;
    severity: "info" | "warning" | "error";
    message: string;
}
export interface PdfObjectGraph {
    pageCount: number;
    textBlocks: PdfTextBlock[];
    annotations: PdfAnnotationSummary[];
    metadata: Record<string, unknown>;
    scan: {
        objects: number;
        streams: number;
        imageObjects: number;
        annotationObjects: number;
        metadataObjects: number;
        embeddedFiles: number;
        objectStreams: number;
        xrefStreams: number;
        encrypted: boolean;
        incrementalUpdates: number;
        filters: Array<{
            name: string;
            count: number;
            supported: boolean;
        }>;
        unsupportedFilters: string[];
        hasAcroForm: boolean;
        hasJavascript: boolean;
        hasRedactionAnnotations: boolean;
    };
    riskFlags: PdfRiskFlag[];
    caveats: string[];
}
export declare function inspectPdfObjectGraph(bytes: Uint8Array, pageSizes: Array<{
    width: number;
    height: number;
}>): Promise<PdfObjectGraph>;
export declare function scanPdfForForbiddenText(bytes: Uint8Array, forbidden: string[] | RegExp[]): {
    found: Array<{
        pattern: string;
        source: "raw" | "content-stream";
        sample: string;
    }>;
    checkedSources: string[];
};
