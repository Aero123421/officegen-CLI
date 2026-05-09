import { type InputLike, type ObjectMapEntry, type OfficegenConfig } from "./shared.js";
export interface DiffOptions {
    config?: OfficegenConfig;
    visual?: boolean;
    maxPages?: number;
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
    };
    visual?: {
        fidelity: "approximate";
        pagesCompared: number;
        pageScores: Array<{
            page: number;
            score: number;
            beforeHash: string;
            afterHash: string;
        }>;
    };
    caveats: string[];
}
export declare function diffDocuments(before: InputLike, after: InputLike, options?: DiffOptions): Promise<DiffResult>;
export declare const diff: typeof diffDocuments;
