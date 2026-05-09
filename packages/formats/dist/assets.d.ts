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
export interface ReplaceAssetOptions {
    out?: string;
    assetPath: string;
    replacement: Uint8Array | Buffer;
    replacementPath?: string;
    allowMediaTypeChange?: boolean;
    config?: OfficegenConfig;
}
export declare function inspectAsset(input: InputLike): Promise<AssetInfo>;
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
export declare const assetExtract: typeof extractAssets;
export declare const assetReplace: typeof replaceAsset;
