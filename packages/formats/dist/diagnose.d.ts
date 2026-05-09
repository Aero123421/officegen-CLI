import { type InspectResult } from "./inspect.js";
import { type InputLike, type OfficegenConfig } from "./shared.js";
export type IssueSeverity = "info" | "warning" | "error";
export interface DiagnoseIssue {
    code: string;
    severity: IssueSeverity;
    message: string;
    stableObjectId?: string;
    suggestedOps?: unknown[];
}
export interface DiagnoseOptions {
    maxTextLength?: number;
    config?: OfficegenConfig;
}
export interface DiagnoseResult {
    schema: "officegen.diagnose.result@1.2";
    issues: DiagnoseIssue[];
    caveats: string[];
}
export declare function diagnose(input: InputLike | InspectResult, options?: DiagnoseOptions): Promise<DiagnoseResult>;
export declare const diagnoseDocument: typeof diagnose;
