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
    objectId?: string;
    crop?: boolean;
    objectMapLimit?: number;
    objectMapOffset?: number;
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
export interface ViewCropArtifact {
    objectId: string;
    page: number;
    format: "svg" | "html";
    content: string;
    width: number;
    height: number;
    renderer: string;
    fidelity: "approximate" | "internal" | "native";
    metadata: ViewCropMetadata;
}
export interface ViewCropMetadata {
    requested: boolean;
    objectId?: string;
    status: "not_requested" | "created" | "object_not_found" | "bbox_unavailable";
    source: "objectMap" | "objectGraph" | "none";
    bbox?: [number, number, number, number];
    page?: number;
    padding: number;
    objectKind?: string;
    graphNodeId?: string;
}
export interface ViewCursor {
    objectMapOffset: number;
    objectMapLimit: number;
    objectMapReturned: number;
    objectMapTotal: number;
    hasMore: boolean;
    nextObjectMapOffset?: number;
}
export interface ViewResult {
    schema: "officegen.view.result@1.2";
    fidelity: "approximate" | "internal" | "native";
    renderer: {
        id: string;
        mode: ExportMode | "approximate";
        fidelity: "approximate" | "internal" | "native";
    };
    caveats: string[];
    pages: ViewPage[];
    crops: ViewCropArtifact[];
    crop: ViewCropMetadata;
    summary: Record<string, unknown>;
    cursor?: ViewCursor;
    nextActions: string[];
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
