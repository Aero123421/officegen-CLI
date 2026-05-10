import { type InputLike, type OfficegenConfig } from "./shared.js";
export interface AssetInfo {
    schema: "officegen.asset.info@1.2";
    source?: string;
    mediaType: string;
    byteLength: number;
    sha256: string;
    width?: number;
    height?: number;
    trusted: false;
}
export interface ExtractAssetsOptions {
    outDir?: string;
    images?: boolean;
    config?: OfficegenConfig;
}
export interface ExtractAssetsResult {
    schema: "officegen.asset.extract.result@1.2";
    assets: Array<AssetInfo & {
        path: string;
        outPath?: string;
    }>;
    caveats: string[];
}
export interface EmbeddedAssetUsage {
    kind: "picture" | "chartEmbeddedWorkbook" | "worksheetDrawingImage" | "embeddedObject" | "relationship";
    partPath: string;
    relationshipId: string;
    relationshipType?: string;
    targetMode?: string;
    slide?: number;
    sheet?: number;
    story?: string;
}
export interface EmbeddedAssetInfo extends Omit<AssetInfo, "schema"> {
    schema: "officegen.asset.embedded.info@2.5";
    stableAssetId: string;
    zipPath: string;
    path: string;
    fileName: string;
    usageCount: number;
    usages: EmbeddedAssetUsage[];
    orphaned: boolean;
    replaceCommand: string;
    extractCommand: string;
    supportedActions: string[];
    limitation?: string;
    untrusted: true;
}
export interface InspectEmbeddedAssetsResult {
    schema: "officegen.asset.embedded.result@2.5";
    mode: "embedded";
    format: "pptx" | "docx" | "xlsx";
    trusted: {
        schema: "officegen.asset.embedded.trusted@2.5";
        format: "pptx" | "docx" | "xlsx";
        source?: string;
        summary: {
            assets: number;
            mediaAssets: number;
            embeddedObjects: number;
            usages: number;
            orphanedAssets: number;
            zipEntries: number;
        };
    };
    untrusted: {
        schema: "officegen.asset.embedded.untrusted@2.5";
        assets: EmbeddedAssetInfo[];
    };
    assets: EmbeddedAssetInfo[];
    caveats: string[];
}
export interface ReplaceAssetOptions {
    out?: string;
    assetPath: string;
    replacement: Uint8Array | Buffer;
    replacementPath?: string;
    allowMediaTypeChange?: boolean;
    config?: OfficegenConfig;
}
export declare function inspectAsset(input: InputLike): Promise<AssetInfo>;
export declare function inspectEmbeddedAssets(input: InputLike, options?: ExtractAssetsOptions): Promise<InspectEmbeddedAssetsResult>;
export declare function extractAssets(input: InputLike, options?: ExtractAssetsOptions): Promise<ExtractAssetsResult>;
export declare function replaceAsset(input: InputLike, options: ReplaceAssetOptions): Promise<{
    schema: "officegen.asset.replace.result@1.2";
    changed: boolean;
    out?: string;
    bytes?: Uint8Array;
    media: Record<string, unknown>;
    caveats: string[];
}>;
export declare const assetInspect: typeof inspectAsset;
export declare const assetInspectEmbedded: typeof inspectEmbeddedAssets;
export declare const assetExtract: typeof extractAssets;
export declare const assetReplace: typeof replaceAsset;
