import { type InspectResult } from "./inspect.js";
import { type InputLike, type ObjectMapEntry } from "./shared.js";
export interface ViewOptions {
    format?: "svg" | "html";
    maxPages?: number;
}
export interface ViewPage {
    page: number;
    stableObjectId: string;
    format: "svg" | "html";
    content: string;
    objectMap: ObjectMapEntry[];
}
export interface ViewResult {
    schema: "officegen.view.result@1.2";
    fidelity: "approximate";
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
