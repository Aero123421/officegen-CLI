import { type ExportMode } from "./export.js";
import { type InputLike, type OfficegenConfig } from "./shared.js";
export interface VerifyOptions {
    native?: boolean;
    visual?: boolean;
    mode?: ExportMode;
    out?: string;
    gates?: VerifyGates;
    formulas?: boolean;
    namedRanges?: boolean;
    externalLinks?: boolean;
    protectedSheets?: boolean;
    timeoutMs?: number;
    config?: OfficegenConfig;
}
export interface VerifyGates {
    expectedSlides?: number;
    expectedPages?: number;
    requiredText?: string[];
    forbiddenText?: string[];
    maxWarnings?: number;
    requireNoRepairDialog?: boolean;
    maxBlankPages?: number;
}
export interface VerifyResult {
    schema: "officegen.verify.result@1.2";
    verificationReport: VerificationReportV2;
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
        repairDialogExpected?: boolean;
        renderer?: "powerpoint" | "libreoffice" | "office-com";
    };
    nativeProof: {
        status: "passed" | "not_run" | "unavailable" | "failed";
        renderer?: "powerpoint" | "libreoffice" | "office-com";
        reason?: string;
        artifact?: string;
    };
    visual?: {
        fidelity: "approximate" | "native";
        pagesChecked: number;
        blankPages: number;
        identicalPages: number[];
        pixelDensityWarnings: string[];
        allPagesIdentical?: boolean;
        rasterDiagnosticsUnavailable?: boolean;
    };
    visualDiff?: {
        status: "compared" | "skipped" | "blocked";
        expectedDiffOnly: boolean;
        fidelity?: "approximate" | "native";
        pagesCompared?: number;
        changedPixels?: number;
        boundingBox?: {
            x: number;
            y: number;
            width: number;
            height: number;
        };
        threshold?: number;
        message?: string;
    };
    expectedDiffOnly?: boolean;
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
    gates?: {
        passed: boolean;
        failed: string[];
        warnings: string[];
    };
}
export interface VerificationReportV2 {
    schema: "officegen.verify@2";
    version: 2;
    format: string;
    readiness: VerifyResult["readiness"];
    score: number;
    partial: boolean;
    gates: Record<VerificationGateName, VerificationGateProjection>;
    issues: Array<{
        code: string;
        severity: "info" | "warning" | "error";
        category: WarningCategory;
        message: string;
        gate?: VerificationGateName;
    }>;
    artifacts: Array<{
        artifactId: string;
        role: string;
        path?: string;
        format?: string;
        managed: boolean;
        exists?: boolean;
        sourceCommand?: string;
    }>;
    recommendedRepairs: VerifyResult["recommendedRepairs"];
}
type VerificationGateName = "schema" | "package" | "semantic" | "visual" | "native" | "security" | "accessibility" | "goal";
interface VerificationGateProjection {
    status: "pass" | "warning" | "fail" | "skipped";
    score?: number;
    summary?: Record<string, unknown>;
    issues: string[];
}
export declare function verify(input: InputLike, options?: VerifyOptions): Promise<VerifyResult>;
type WarningCategory = "quality" | "compatibility" | "security" | "environment";
export declare const verifyDocument: typeof verify;
export {};
