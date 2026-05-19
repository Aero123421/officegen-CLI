import type { CapabilitiesDocument, CapabilityContract, CapabilityFeature, FeatureName, OfficegenConfig } from "./types.js";
export interface CommandSpec {
    feature: FeatureName;
    commands: readonly string[];
}
export declare const COMMAND_SPECS: readonly [CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec];
export declare const commandMap: Record<FeatureName, string[]>;
export declare const RELEASE_CAPABILITY_CONTRACT_VERSION = "3.0.0-contract";
export declare const formatCapabilities: {
    readonly pptx: {
        readonly text: true;
        readonly lists: true;
        readonly tables: "scoped cell text edits";
        readonly images: "replace existing image shapes";
        readonly charts: "single-series chart data updates only";
        readonly smartArt: "unsupported";
        readonly designer: "limited; not a full PowerPoint designer";
    };
    readonly docx: {
        readonly text: "scoped paragraph/run replacement";
        readonly tables: "scoped table cell text edits";
        readonly comments: "add comment ranges where selectors resolve";
        readonly redlines: "tracked insert/delete/replace only";
        readonly styles: "scoped style application";
        readonly fullFidelityEditing: "limited; not a full Word/DTP/legal contract engine";
    };
    readonly xlsx: {
        readonly cells: true;
        readonly formulas: "guarded setFormula; no internal calculation engine";
        readonly tables: "scoped write/resize operations";
        readonly charts: "single-series chart data updates only";
        readonly pivots: "refresh flags only; no pivot layout/field/value editing";
        readonly slicers: "selection updates only; no slicer authoring/styling engine";
    };
    readonly pdf: {
        readonly generation: true;
        readonly inspect: "best-effort text preview plus page previews";
        readonly overlays: "text overlays and annotations only";
        readonly redaction: "unsupported; overlays do not physically remove underlying content";
        readonly contentRewrite: "unsupported";
    };
};
export declare const featureContracts: CapabilityContract[];
export declare const knownLimitations: string[];
export declare function buildFeatureRegistry(config: OfficegenConfig): CapabilityFeature[];
export declare function computeCapabilitiesHash(config: OfficegenConfig, cliVersion?: string): string;
export declare function getVisibleCommands(config: OfficegenConfig, agent?: boolean): string[];
export declare function getCapabilities(config: OfficegenConfig, options?: {
    agent?: boolean;
    runInstructionsPath?: string;
}): CapabilitiesDocument;
export declare function isFeatureAvailable(config: OfficegenConfig, feature: FeatureName, agent?: boolean): boolean;
