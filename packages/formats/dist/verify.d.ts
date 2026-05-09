import { type InputLike, type OfficegenConfig } from "./shared.js";
export interface VerifyOptions {
    native?: boolean;
    visual?: boolean;
    out?: string;
    formulas?: boolean;
    namedRanges?: boolean;
    externalLinks?: boolean;
    protectedSheets?: boolean;
    timeoutMs?: number;
    config?: OfficegenConfig;
}
export interface VerifyResult {
    schema: "officegen.verify.result@1.2";
    readiness: "pass" | "pass_with_environment_gap" | "warning" | "blocked";
    partial?: boolean;
    phaseTimings?: Array<{
        phase: string;
        durationMs: number;
        timeout?: boolean;
    }>;
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
        category: WarningCategory;
        examples: string[];
    }>;
    topRisks: Array<{
        code: string;
        severity: "warning" | "error";
        category: WarningCategory;
        count: number;
        message: string;
        slide?: number;
        page?: number;
        stableObjectId?: string;
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
type WarningCategory = "quality" | "compatibility" | "security" | "environment";
export declare const verifyDocument: typeof verify;
export {};
