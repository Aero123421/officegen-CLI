export type VisualDiffStatus = "compared" | "skipped" | "blocked";
export interface RasterImageData {
    width: number;
    height: number;
    data: Uint8Array | Uint8ClampedArray;
    channels?: 3 | 4;
}
export interface RasterPixelDiffOptions {
    threshold?: number;
}
export interface RasterPixelDiffResult {
    kind: "raster-pixel";
    status: VisualDiffStatus;
    changed: boolean;
    beforeHash: string;
    afterHash: string;
    beforeWidth: number;
    beforeHeight: number;
    afterWidth: number;
    afterHeight: number;
    width: number;
    height: number;
    totalPixels: number;
    pixelsCompared: number;
    changedPixels: number;
    changedRatio: number;
    threshold: number;
    boundingBox?: {
        x: number;
        y: number;
        width: number;
        height: number;
    };
    message?: string;
}
export declare function comparePngPixels(before: Uint8Array, after: Uint8Array, options?: RasterPixelDiffOptions): Promise<RasterPixelDiffResult>;
export declare function decodePngRaster(bytes: Uint8Array): Promise<RasterImageData>;
export declare function compareRasterPixels(before: RasterImageData, after: RasterImageData, options?: RasterPixelDiffOptions): RasterPixelDiffResult;
