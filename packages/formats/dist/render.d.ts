import { type OfficegenConfig } from "@officegen/core";
export type RenderTarget = "pptx" | "docx" | "xlsx" | "pdf";
interface BlockIR {
    type?: string;
    text?: string;
    title?: string;
    items?: string[];
    rows?: Array<Record<string, unknown> | unknown[]>;
    path?: string;
}
export interface DocumentIR {
    title?: string;
    kind?: string;
    targets?: string[];
    sections?: Array<{
        id?: string;
        title?: string;
        body?: string | string[];
        blocks?: BlockIR[];
        rows?: Array<Record<string, unknown> | unknown[]>;
        items?: string[];
    }>;
    design?: {
        colors?: Record<string, string>;
        typography?: Record<string, unknown>;
    };
    slides?: Array<{
        title?: string;
        body?: string | string[];
        blocks?: BlockIR[];
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
    diagnostics?: Array<Record<string, unknown>>;
}
export declare function render(ir: DocumentIR, options?: RenderOptions): Promise<RenderResult>;
export declare const renderDocument: typeof render;
export {};
