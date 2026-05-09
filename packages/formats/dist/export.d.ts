import { type DocumentIR } from "./render.js";
import { type InputLike } from "./shared.js";
import { type OfficegenConfig } from "@officegen/core";
export type ExportMode = "fast" | "internal" | "native";
export interface ExportOptions {
    to: "pdf" | "svg" | "html" | "pptx" | "docx" | "xlsx";
    out?: string;
    mode?: ExportMode;
    pages?: number[];
    config?: OfficegenConfig;
}
export interface PdfOperationOptions {
    out?: string;
    config?: OfficegenConfig;
}
export interface ExportResult {
    schema: "officegen.export.result@1.2";
    from: string;
    to: string;
    mode: ExportMode;
    out?: string;
    bytes?: Uint8Array;
    fidelity: "approximate" | "internal" | "native";
    caveats: string[];
    renderer?: {
        id: string;
        executable?: string;
        status: "used" | "unavailable";
    };
}
export declare function exportDocument(input: InputLike | DocumentIR, options: ExportOptions): Promise<ExportResult>;
export declare const exportFile: typeof exportDocument;
export declare function mergePdfs(inputs: InputLike[], options?: PdfOperationOptions): Promise<ExportResult>;
export declare function splitPdf(input: InputLike, ranges: Array<number[]>, options?: PdfOperationOptions): Promise<Array<ExportResult>>;
export declare function reorderPdf(input: InputLike, order: number[], options?: PdfOperationOptions): Promise<ExportResult>;
export declare function findLibreOfficeExecutable(): Promise<string | undefined>;
