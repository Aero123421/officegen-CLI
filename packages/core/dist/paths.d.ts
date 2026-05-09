import type { OfficegenConfig, PathValidationOptions, ValidatedPath } from "./types.js";
export declare function resolveOfficegenPath(config: OfficegenConfig, inputPath: string): string;
export declare function canonicalizePath(config: OfficegenConfig, inputPath: string): Promise<ValidatedPath>;
export declare function validatePath(config: OfficegenConfig, options: PathValidationOptions): Promise<ValidatedPath>;
export declare function ensureDirectory(filePath: string): Promise<void>;
