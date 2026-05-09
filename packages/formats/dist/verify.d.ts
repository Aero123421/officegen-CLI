import { type InputLike, type OfficegenConfig } from "./shared.js";
export interface VerifyOptions {
    native?: boolean;
    visual?: boolean;
    out?: string;
    formulas?: boolean;
    namedRanges?: boolean;
    externalLinks?: boolean;
    protectedSheets?: boolean;
    config?: OfficegenConfig;
}
export interface VerifyResult {
    schema: "officegen.verify.result@1.2";
    readiness: "pass" | "warning" | "blocked";
    score: number;
    format: string;
    openable: boolean;
    noRepairDialogExpected: boolean;
    nativeRenderer?: {
        attempted: boolean;
        ok: boolean;
        message?: string;
        artifact?: string;
    };
    visual?: {
        fidelity: "approximate" | "native";
        pagesChecked: number;
        blankPages: number;
    };
    blockingIssues: string[];
    warnings: string[];
    warningSummary: Array<{
        code: string;
        count: number;
        severity: "warning" | "error";
        examples: string[];
    }>;
    topRisks: Array<{
        code: string;
        severity: "warning" | "error";
        count: number;
        message: string;
        repair?: string;
    }>;
    scoreBreakdown: Record<string, unknown>;
    recommendedRepairs: Array<{
        code: string;
        command?: string;
        reason: string;
    }>;
    artifacts: Record<string, unknown>;
}
export declare function verify(input: InputLike, options?: VerifyOptions): Promise<VerifyResult>;
export declare const verifyDocument: typeof verify;
