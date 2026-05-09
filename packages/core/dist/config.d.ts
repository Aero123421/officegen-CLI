import type { ConfigLoadOptions, FeatureName, OfficegenConfig, OfficegenConfigInput, OfficegenProfile } from "./types.js";
export declare const FEATURE_NAMES: FeatureName[];
export declare const BUILTIN_PROFILE_CONFIGS: Record<OfficegenProfile, OfficegenConfig>;
export declare function getBuiltinConfig(profile?: OfficegenProfile): OfficegenConfig;
export declare function normalizeFeatureConfig(config: OfficegenConfig): OfficegenConfig;
export declare function mergeConfig(base: OfficegenConfig, override?: OfficegenConfigInput): OfficegenConfig;
export declare function expandHome(inputPath: string): string;
export declare function loadConfig(options?: ConfigLoadOptions): Promise<OfficegenConfig>;
