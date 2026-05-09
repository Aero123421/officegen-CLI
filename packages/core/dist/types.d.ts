export declare const OFFICEGEN_SCHEMA_VERSION = "1.2";
export declare const OFFICEGEN_CLI_VERSION = "1.2.2";
export declare const SCHEMA_REGISTRY_VERSION = "1.2.0";
export type JsonPrimitive = string | number | boolean | null;
export type JsonValue = JsonPrimitive | JsonValue[] | {
    [key: string]: JsonValue;
};
export type JsonObject = {
    [key: string]: JsonValue;
};
export type OfficegenProfile = "substrate" | "authoring" | "enterprise";
export type FeatureName = "capabilities" | "help" | "config" | "doctor" | "inspect" | "view" | "edit" | "render" | "scaffold" | "export" | "validate" | "diagnose" | "repair" | "run" | "asset" | "chart" | "diagram" | "schema" | "errors" | "template" | "design" | "layout" | "agent" | "mcp" | "renderer" | "plugin";
export type NetworkPolicy = "deny" | "allow";
export type ProcessPolicy = "deny" | "allow";
export type OutOfProjectPolicy = "deny" | "warn" | "allow";
export interface FeatureVisibility {
    enabled: boolean;
    visibleInHelp: boolean;
    visibleToAgents: boolean;
}
export type FeatureMap = Record<FeatureName, FeatureVisibility>;
export interface OfficegenConfigPaths {
    projectRoot: string;
    projectConfigDir: string;
    userConfigDir: string;
    defaultOutputDir: string;
    defaultRunsDir: string;
}
export interface OfficegenUntrustedInputLimits {
    maxInputFileBytes: number;
    maxZipEntries: number;
    maxZipExpandedBytes: number;
    maxSingleXmlPartBytes: number;
    maxRelationships: number;
    maxNestedZipDepth: number;
    xmlExternalEntities: "deny" | "allow";
    externalRelationships: "warn-and-drop-by-default" | "allow" | "deny";
    macros: "warn-and-preserve-only-if-requested" | "allow" | "deny";
    embeddedObjects: "warn" | "allow" | "deny";
    externalHyperlinks: "warn" | "allow" | "deny";
}
export interface OfficegenSecurityConfig {
    network: NetworkPolicy;
    externalProcess: ProcessPolicy;
    plugins: "disabled" | "enabled";
    renderers: "disabled" | "enabled";
    allowOverwrite: boolean;
    outOfProjectPolicy: OutOfProjectPolicy;
    allowAbsoluteInputPaths: boolean;
    allowAbsoluteOutputPaths: boolean;
    redactAbsolutePathsInJson: boolean;
    redactSecretsInJson: boolean;
    followSymlinks: boolean;
    allowHardlinks: boolean;
    trustedRoots: string[];
    untrustedInput: OfficegenUntrustedInputLimits;
}
export interface OfficegenAgentConfig {
    defaultJsonBudgetBytes: number;
    inspectDefaultDepth: "summary" | "full";
    largeOutputMode: "path-only" | "inline";
    requireCapabilitiesCheck: boolean;
}
export interface OfficegenConfig {
    version: "1.2";
    profile: OfficegenProfile;
    paths: OfficegenConfigPaths;
    features: FeatureMap;
    security: OfficegenSecurityConfig;
    agent: OfficegenAgentConfig;
}
export type PartialDeep<T> = {
    [K in keyof T]?: T[K] extends Array<infer U> ? U[] : T[K] extends object ? PartialDeep<T[K]> : T[K];
};
export type OfficegenConfigInput = PartialDeep<OfficegenConfig>;
export interface ConfigLoadOptions {
    cwd?: string;
    userConfigPath?: string;
    projectConfigPath?: string;
    overrides?: OfficegenConfigInput;
}
export type ErrorSeverity = "info" | "warning" | "error" | "critical";
export type OfficegenErrorCode = "FEATURE_DISABLED" | "FEATURE_HIDDEN_FROM_AGENT" | "UNKNOWN_COMMAND" | "UNKNOWN_OPTION" | "CAPABILITIES_STALE" | "INPUT_NOT_FOUND" | "SCHEMA_INVALID" | "SCHEMA_DEPRECATED" | "SCHEMA_MIGRATION_REQUIRED" | "SECURITY_PATH_OUTSIDE_ROOT" | "SECURITY_ABSOLUTE_OUT_DENIED" | "SECURITY_SYMLINK_DENIED" | "SECURITY_HARDLINK_DENIED" | "SECURITY_ZIP_BOMB_DETECTED" | "SECURITY_XML_ENTITY_DENIED" | "SECURITY_MACRO_DETECTED" | "PLUGIN_NOT_TRUSTED" | "PLUGIN_HASH_MISMATCH" | "PLUGIN_PERMISSION_DENIED" | "RENDERER_NOT_TRUSTED" | "SELECTOR_NOT_FOUND" | "SELECTOR_AMBIGUOUS" | "EDIT_TRANSACTION_FAILED" | "IDEMPOTENCY_REPLAY" | "TEXT_OVERFLOW" | "IMAGE_MISSING" | "ASSET_UNSUPPORTED_FORMAT" | "CHART_SPEC_INVALID" | "DIAGRAM_SPEC_INVALID" | "VIEW_FIDELITY_LOW" | "EXPORT_UNSUPPORTED";
export interface ErrorCatalogEntry {
    code: OfficegenErrorCode;
    category: string;
    severity: ErrorSeverity;
    message: string;
    typicalCause: string;
    suggestedOps: string[];
    nextSuggestedCommands: string[];
}
export interface OfficegenErrorPayload {
    code: OfficegenErrorCode;
    category: string;
    severity: ErrorSeverity;
    message: string;
    feature?: string;
    command?: string;
    details?: JsonValue;
}
export interface EnvelopeBase {
    schema: "officegen.envelope@1.2";
    ok: boolean;
    command?: string;
    runId?: string;
    cliVersion: string;
    capabilitiesHash?: string;
    pathsRedacted: boolean;
    warnings: JsonValue[];
    diagnostics: JsonValue[];
    artifacts: JsonValue[];
    nextSuggestedCommands: string[];
}
export interface SuccessEnvelope<T extends JsonValue = JsonObject> extends EnvelopeBase {
    ok: true;
    result: T;
}
export interface ErrorEnvelope extends EnvelopeBase {
    ok: false;
    error: OfficegenErrorPayload;
    availableCommands: string[];
}
export type JsonEnvelope<T extends JsonValue = JsonObject> = SuccessEnvelope<T> | ErrorEnvelope;
export interface CapabilityFeature {
    name: FeatureName;
    enabled: boolean;
    visibleInHelp: boolean;
    visibleToAgents: boolean;
    commands: string[];
    requires: FeatureName[];
}
export interface CapabilitiesDocument {
    schema: "officegen.capabilities@1.2";
    ok: true;
    profile: OfficegenProfile;
    capabilitiesHash: string;
    visibleCommands: string[];
    hiddenFromAgents: string[];
    disabled: string[];
    agentInstructionsPath: string;
    jsonBudgetBytes: number;
    nextSuggestedCommands: string[];
}
export interface PathValidationOptions {
    kind: "input" | "output";
    path: string;
    overwrite?: boolean;
    allowAbsoluteOut?: boolean;
    allowRoots?: string[];
}
export interface ValidatedPath {
    inputPath: string;
    absolutePath: string;
    realPath: string;
    existed: boolean;
    warnings: string[];
}
export interface RunFolder {
    runId: string;
    root: string;
    inputDir: string;
    irDir: string;
    opsDir: string;
    viewsDir: string;
    diagnosticsDir: string;
    outputDir: string;
    backupDir: string;
    logsDir: string;
    tracePath: string;
    runJsonPath: string;
    manifestPath: string;
}
export interface ManifestFileRecord {
    path: string;
    sha256?: string;
    trusted?: boolean;
    warnings?: string[];
    overwroteExisting?: boolean;
}
export interface RunManifest {
    schema: "officegen.manifest@1.2";
    runId: string;
    cliVersion: string;
    profile: OfficegenProfile;
    capabilitiesHash: string;
    inputs: ManifestFileRecord[];
    outputs: ManifestFileRecord[];
    security: {
        network: NetworkPolicy;
        externalProcess: ProcessPolicy;
        redactedPaths: boolean;
        macrosPreserved: boolean;
        externalRelationshipsDropped: boolean;
    };
    warnings: JsonValue[];
}
export type SchemaStability = "stable" | "experimental" | "deprecated" | "internal";
export interface SchemaRegistryEntry {
    id: string;
    schema: JsonObject;
    stability: SchemaStability;
    introducedIn: string;
    deprecated: boolean;
    feature?: FeatureName;
    visibleToAgents: boolean;
}
export interface RedactionRecord {
    kind: "absolute-path" | "secret-like-token";
    location: string;
    replacement: string;
}
export interface RedactionResult<T> {
    value: T;
    redactions: RedactionRecord[];
}
export interface ZipSafetyWarning {
    code: "ZIP_PATH_TRAVERSAL" | "ZIP_ENTRY_LIMIT_EXCEEDED" | "ZIP_EXPANDED_BYTES_EXCEEDED" | "ZIP_COMPRESSION_RATIO_EXCEEDED" | "ZIP_NESTED_ZIP_DETECTED" | "ZIP_XML_PART_TOO_LARGE" | "ZIP_XML_ENTITY_DENIED" | "ZIP_RELATIONSHIP_LIMIT_EXCEEDED" | "ZIP_EXTERNAL_RELATIONSHIP" | "ZIP_MACRO_DETECTED" | "ZIP_EMBEDDED_OBJECT";
    severity: ErrorSeverity;
    message: string;
    entry?: string;
}
export interface ZipSafetyReport {
    ok: boolean;
    entryCount: number;
    expandedBytes: number;
    compressedBytes: number;
    hasMacros: boolean;
    externalRelationships: string[];
    warnings: ZipSafetyWarning[];
}
