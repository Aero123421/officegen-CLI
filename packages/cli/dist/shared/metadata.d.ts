import type { CapabilityEntry, FeatureKey } from "./types.js";
export declare const COMMAND_METADATA: CapabilityEntry[];
export declare function metadataFor(feature: FeatureKey): CapabilityEntry | undefined;
