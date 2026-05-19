import { type AgentSeparatedResult, type InputLike, type OfficegenConfig } from "./shared.js";
import { type BuildObjectGraphOptions, type ObjectGraph } from "./graphs/objectGraph.js";
export type InspectDepth = "summary" | "shallow" | "full";
export interface InspectOptions {
    format?: "pptx" | "docx" | "xlsx" | "pdf" | "unknown";
    depth?: InspectDepth;
    include?: Array<"text" | "assets" | "relationships" | "rawPaths">;
    structure?: boolean;
    sheet?: string;
    range?: string;
    config?: OfficegenConfig;
    emit?: "inspect" | "object-graph";
    includeObjectGraph?: boolean;
    objectGraph?: Pick<BuildObjectGraphOptions, "nodeOffset" | "nodeLimit" | "edgeOffset" | "edgeLimit">;
}
export interface InspectResult extends AgentSeparatedResult<Record<string, unknown>> {
    schema: "officegen.inspect.result@1.2";
    objectGraph?: ObjectGraph;
}
export declare function inspect(input: InputLike, options?: InspectOptions): Promise<InspectResult>;
export declare const inspectDocument: typeof inspect;
export declare const inspectOfficeFile: typeof inspect;
