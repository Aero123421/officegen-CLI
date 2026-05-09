import type { OfficegenConfig, ZipSafetyReport } from "./types.js";
export declare function inspectZipSafety(input: Buffer | Uint8Array | ArrayBuffer, config: OfficegenConfig, options?: {
    depth?: number;
    compressionRatioLimit?: number;
}): Promise<ZipSafetyReport>;
