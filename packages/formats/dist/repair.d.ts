import { type EditOperation } from "./edit.js";
import { type DiagnoseIssue, type DiagnoseResult } from "./diagnose.js";
import { type InputLike, type OfficegenConfig } from "./shared.js";
export interface RepairOptions {
    out?: string;
    dryRun?: boolean;
    issues?: DiagnoseIssue[] | DiagnoseResult;
    config?: OfficegenConfig;
}
export type RepairTaxonomyCategory = "quality" | "compatibility" | "security" | "environment";
export type RepairTaxonomySeverity = "info" | "warning" | "error" | "critical";
export interface RepairFailureTaxonomyEntry {
    code: string;
    category: RepairTaxonomyCategory;
    severity: RepairTaxonomySeverity;
    autoRepairable: boolean;
    evidence: Array<{
        kind: string;
        message: string;
        stableObjectId?: string;
        issueCode?: string;
    }>;
    nextCommand: string;
}
export interface RepairPlanV2 {
    schema: "officegen.repairPlan@2";
    version: 2;
    target: string;
    input?: string;
    inputSha256: string;
    wouldWrite: boolean;
    planOnly: boolean;
    operations: EditOperation[];
    failureTaxonomy: RepairFailureTaxonomyEntry[];
    steps: Array<{
        id: string;
        command: string;
        dryRun: boolean;
        reason: string;
    }>;
    verify: {
        status: "not_run";
        requiredAfterRepair: boolean;
        command: string;
        readinessNote: string;
    };
}
export interface RepairResult {
    schema: "officegen.repair.result@1.2";
    format: string;
    inputSha256: string;
    applied: number;
    changed: boolean;
    out?: string;
    suggestedOps: EditOperation[];
    failureTaxonomy: RepairFailureTaxonomyEntry[];
    repairPlan: RepairPlanV2;
    readiness?: "warning";
    readinessNotes?: string[];
    postRepairVerify: RepairPlanV2["verify"];
    caveats: string[];
}
export declare function repair(input: InputLike, options?: RepairOptions): Promise<RepairResult>;
export declare const repairDocument: typeof repair;
