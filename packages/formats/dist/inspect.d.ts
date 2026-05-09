import { type AgentSeparatedResult, type InputLike } from "./shared.js";
export type InspectDepth = "summary" | "shallow" | "full";
export interface InspectOptions {
    format?: "pptx" | "docx" | "xlsx" | "pdf" | "unknown";
    depth?: InspectDepth;
    include?: Array<"text" | "assets" | "relationships" | "rawPaths">;
}
export interface InspectResult extends AgentSeparatedResult<Record<string, unknown>> {
    schema: "officegen.inspect.result@1.2";
}
export declare function inspect(input: InputLike, options?: InspectOptions): Promise<InspectResult>;
export declare const inspectDocument: typeof inspect;
export declare const inspectOfficeFile: typeof inspect;
