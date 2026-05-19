import JSZip from "jszip";
import type { OfficegenConfig, ZipSafetyReport } from "./types.js";
export interface ZipSafetyInspectOptions {
    depth?: number;
    compressionRatioLimit?: number;
    preloadedZip?: JSZip;
}
export declare function scanZipSafetyMetadata(input: Buffer | Uint8Array | ArrayBuffer, config: OfficegenConfig, options?: {
    depth?: number;
    compressionRatioLimit?: number;
}): ZipSafetyReport;
export declare function inspectZipSafety(input: Buffer | Uint8Array | ArrayBuffer, config: OfficegenConfig, options?: ZipSafetyInspectOptions): Promise<ZipSafetyReport>;
