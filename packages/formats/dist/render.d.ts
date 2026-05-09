import { type OfficegenConfig } from "@officegen/core";
export type RenderTarget = "pptx" | "docx" | "xlsx" | "pdf";
export interface DocumentIR {
    title?: string;
    kind?: string;
    targets?: string[];
    sections?: Array<{
        id?: string;
        title?: string;
        body?: string | string[];
        blocks?: Array<{
            type?: string;
            text?: string;
            rows?: Array<Record<string, unknown> | unknown[]>;
        }>;
        rows?: Array<Record<string, unknown> | unknown[]>;
        items?: string[];
    }>;
    slides?: Array<{
        title?: string;
        body?: string | string[];
    }>;
    sheets?: Array<{
        name?: string;
        rows?: Array<Record<string, unknown> | unknown[]>;
    }>;
}
export interface RenderOptions {
    out?: string;
    target?: string;
    config?: OfficegenConfig;
}
export interface RenderResult {
    schema: "officegen.render.result@1.2";
    target: RenderTarget;
    out?: string;
    bytes?: Uint8Array | Buffer;
    caveats: string[];
}
export declare function render(ir: DocumentIR, options?: RenderOptions): Promise<RenderResult>;
export declare const renderDocument: typeof render;
