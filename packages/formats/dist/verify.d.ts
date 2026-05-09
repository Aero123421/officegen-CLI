import { type InputLike, type OfficegenConfig } from "./shared.js";
export interface VerifyOptions {
    native?: boolean;
    visual?: boolean;
    out?: string;
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
    artifacts: Record<string, unknown>;
}
export declare function verify(input: InputLike, options?: VerifyOptions): Promise<VerifyResult>;
export declare const verifyDocument: typeof verify;
