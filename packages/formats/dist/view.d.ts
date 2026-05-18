import { type InspectResult } from "./inspect.js";
import { type ExportMode } from "./export.js";
import { type InputLike, type OfficegenConfig, type ObjectMapEntry } from "./shared.js";
export type ViewFormat = "svg" | "html" | "png" | "jpeg" | "jpg";
export interface ViewOptions {
    format?: ViewFormat;
    maxPages?: number;
    dpi?: number;
    mode?: ExportMode;
    timeoutMs?: number;
    config?: OfficegenConfig;
}
export interface ViewPage {
    page: number;
    stableObjectId: string;
    format: "svg" | "html" | "png" | "jpeg";
    content: string;
    bytes?: Uint8Array;
    width?: number;
    height?: number;
    renderer?: string;
    objectMap: ObjectMapEntry[];
}
export interface ViewResult {
    schema: "officegen.view.result@1.2";
    fidelity: "approximate" | "internal" | "native";
    caveats: string[];
    pages: ViewPage[];
    objectMap: ObjectMapEntry[];
    trusted: {
        sourceSchema: string;
        sourceFormat: string;
        generatedAt: string;
    };
    agentInstruction: string;
}
export declare function view(input: InputLike | InspectResult, options?: ViewOptions): Promise<ViewResult>;
export declare const viewDocument: typeof view;
