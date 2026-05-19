import type { CapabilityEntry, FeatureKey } from "./types.js";
export interface OptionSpec {
    flag: string;
    value?: string;
    description: string;
}
export declare const GLOBAL_OPTION_SPECS: OptionSpec[];
export declare const COMMAND_OPTION_SPECS: Record<string, OptionSpec[]>;
export declare const COMMAND_METADATA: CapabilityEntry[];
export declare function metadataFor(feature: FeatureKey): CapabilityEntry | undefined;
export declare function effectiveOptionSpecsFor(commandGroup: string, subcommand?: string): OptionSpec[];
export declare function acceptedOptionSpecsFor(commandGroup: string, subcommand?: string): OptionSpec[];
export declare function effectiveOptionsFor(commandGroup: string, subcommand?: string): string[];
export declare function acceptedOptionsFor(commandGroup: string, subcommand?: string): string[];
export declare function optionSyntax(spec: OptionSpec): string;
