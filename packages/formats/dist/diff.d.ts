import { type InputLike, type ObjectMapEntry, type OfficegenConfig } from "./shared.js";
import { type RasterPixelDiffResult, type VisualDiffStatus } from "./visualDiff.js";
export interface DiffOptions {
    config?: OfficegenConfig;
    visual?: boolean;
    native?: boolean;
    maxPages?: number;
    pixelThreshold?: number;
}
export interface DiffResult {
    schema: "officegen.diff.result@1.2";
    formatBefore: string;
    formatAfter: string;
    changed: boolean;
    summary: {
        addedObjects: number;
        removedObjects: number;
        changedTextObjects: number;
        changedGeometryObjects: number;
        changedSemanticObjects?: number;
        beforePages: number;
        afterPages: number;
        pageCountChanged: boolean;
        changedParts?: number;
        visualRegressionScore?: number;
    };
    semantic: {
        added: ObjectMapEntry[];
        removed: ObjectMapEntry[];
        changedText: Array<{
            stableObjectId: string;
            kind: string;
            before?: string;
            after?: string;
        }>;
        changedGeometry: Array<{
            stableObjectId: string;
            kind: string;
            beforeBbox?: [number, number, number, number];
            afterBbox?: [number, number, number, number];
            delta: {
                x: number;
                y: number;
                width: number;
                height: number;
            };
        }>;
        changedSemantic: Array<{
            stableObjectId: string;
            kind: string;
            changes: Array<"paragraph" | "bullet" | "numbering" | "run-format">;
            before?: Record<string, unknown>;
            after?: Record<string, unknown>;
        }>;
        partChanges?: Array<{
            path: string;
            kind: string;
            beforeHash?: string;
            afterHash?: string;
            status: "added" | "removed" | "changed";
        }>;
    };
    visual?: {
        status?: VisualDiffStatus;
        kind?: "approximate-string" | "pdf-byte-window" | "raster-pixel";
        fidelity: "approximate" | "native";
        pagesCompared: number;
        beforePages: number;
        afterPages: number;
        pageCountChanged: boolean;
        pageScores: Array<{
            page: number;
            score: number;
            beforeHash: string;
            afterHash: string;
            pixelDiff?: RasterPixelDiffResult;
        }>;
        renderer?: string;
        fallback?: boolean;
        message?: string;
    };
    caveats: string[];
}
export declare function diffDocuments(before: InputLike, after: InputLike, options?: DiffOptions): Promise<DiffResult>;
export declare const diff: typeof diffDocuments;
