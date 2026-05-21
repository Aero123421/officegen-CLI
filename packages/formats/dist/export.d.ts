import { type DocumentIR } from "./render.js";
import { type InputLike } from "./shared.js";
import { type OfficegenConfig } from "@officegen/core";
export type ExportMode = "fast" | "internal" | "native" | "proof";
export interface ExportOptions {
    to: "pdf" | "svg" | "html" | "pptx" | "docx" | "xlsx";
    out?: string;
    mode?: ExportMode;
    pages?: number[];
    config?: OfficegenConfig;
    timeoutMs?: number;
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
        repairDialogExpected?: boolean;
        backend?: "office-com" | "libreoffice";
    };
    nativeProof?: NativeProof;
}
export interface NativeProof {
    status: "passed" | "not_run" | "unavailable" | "failed";
    renderer?: "powerpoint" | "libreoffice" | "office-com";
    reason?: string;
    artifact?: string;
}
export declare function exportDocument(input: InputLike | DocumentIR, options: ExportOptions): Promise<ExportResult>;
export declare const exportFile: typeof exportDocument;
export declare function mergePdfs(inputs: InputLike[], options?: PdfOperationOptions): Promise<ExportResult>;
export declare function splitPdf(input: InputLike, ranges: Array<number[]>, options?: PdfOperationOptions): Promise<Array<ExportResult>>;
export declare function reorderPdf(input: InputLike, order: number[], options?: PdfOperationOptions): Promise<ExportResult>;
export declare const MIN_NATIVE_RENDERER_TIMEOUT_MS = 1000;
export declare const DEFAULT_NATIVE_RENDERER_TIMEOUT_MS = 120000;
export declare function resolveNativeRendererTimeoutMs(timeoutMs?: number): number;
export interface NativeRendererDoctorResult {
    schema: "officegen.renderer.doctor@2.2";
    platform: NodeJS.Platform;
    policy: {
        externalProcess?: string;
        renderers?: string;
    };
    nativeProof: NativeProof;
    nextActions: string[];
    renderers: Array<{
        id: string;
        backend: "office-com" | "libreoffice";
        available: boolean;
        executable?: string;
        formats: string[];
        message: string;
    }>;
}
export declare function nativeRendererDoctor(config?: OfficegenConfig): Promise<NativeRendererDoctorResult>;
export declare function findLibreOfficeExecutable(): Promise<string | undefined>;
