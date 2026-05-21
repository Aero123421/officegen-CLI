import { type InspectResult } from "./inspect.js";
import { type InputLike, type OfficegenConfig } from "./shared.js";
export type IssueSeverity = "info" | "warning" | "error";
export interface DiagnoseIssue {
    code: string;
    severity: IssueSeverity;
    message: string;
    stableObjectId?: string;
    location?: {
        slide?: number;
        page?: number;
        stableObjectId?: string;
    };
    metrics?: Record<string, unknown>;
    suggestedOps?: unknown[];
    editOps?: EditOpsDocument;
}
export interface DiagnoseOptions {
    maxTextLength?: number;
    config?: OfficegenConfig;
}
export interface DiagnoseResult {
    schema: "officegen.diagnose.result@1.2";
    issues: DiagnoseIssue[];
    suggestedOps: unknown[];
    editOps?: EditOpsDocument;
    caveats: string[];
}
interface EditOpsDocument {
    schema: "officegen.edit.ops@1.2";
    target: "pptx" | "docx" | "xlsx" | "pdf";
    ops: unknown[];
}
export declare function diagnose(input: InputLike | InspectResult, options?: DiagnoseOptions): Promise<DiagnoseResult>;
export declare const diagnoseDocument: typeof diagnose;
export {};
