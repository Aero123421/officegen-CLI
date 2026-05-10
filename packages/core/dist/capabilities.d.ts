import type { CapabilitiesDocument, CapabilityFeature, FeatureName, OfficegenConfig } from "./types.js";
export interface CommandSpec {
    feature: FeatureName;
    commands: readonly string[];
}
export declare const COMMAND_SPECS: readonly [CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec];
export declare const commandMap: Record<FeatureName, string[]>;
export declare function buildFeatureRegistry(config: OfficegenConfig): CapabilityFeature[];
export declare function computeCapabilitiesHash(config: OfficegenConfig, cliVersion?: string): string;
export declare function getVisibleCommands(config: OfficegenConfig, agent?: boolean): string[];
export declare function getCapabilities(config: OfficegenConfig, options?: {
    agent?: boolean;
    runInstructionsPath?: string;
}): CapabilitiesDocument;
export declare function isFeatureAvailable(config: OfficegenConfig, feature: FeatureName, agent?: boolean): boolean;
