import { type EditOperation } from "./edit.js";
import { type DiagnoseIssue, type DiagnoseResult } from "./diagnose.js";
import { type InputLike } from "./shared.js";
export interface RepairOptions {
    out?: string;
    dryRun?: boolean;
    issues?: DiagnoseIssue[] | DiagnoseResult;
}
export interface RepairResult {
    schema: "officegen.repair.result@1.2";
    applied: number;
    changed: boolean;
    out?: string;
    suggestedOps: EditOperation[];
    caveats: string[];
}
export declare function repair(input: InputLike, options?: RepairOptions): Promise<RepairResult>;
export declare const repairDocument: typeof repair;
