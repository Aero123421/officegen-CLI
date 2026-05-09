const REQUIRED_ERROR_CODES = [
    "FEATURE_DISABLED",
    "FEATURE_HIDDEN_FROM_AGENT",
    "UNKNOWN_COMMAND",
    "UNKNOWN_OPTION",
    "CAPABILITIES_STALE",
    "INPUT_NOT_FOUND",
    "SCHEMA_INVALID",
    "SCHEMA_DEPRECATED",
    "SCHEMA_MIGRATION_REQUIRED",
    "SECURITY_PATH_OUTSIDE_ROOT",
    "SECURITY_ABSOLUTE_OUT_DENIED",
    "SECURITY_SYMLINK_DENIED",
    "SECURITY_HARDLINK_DENIED",
    "SECURITY_ZIP_BOMB_DETECTED",
    "SECURITY_XML_ENTITY_DENIED",
    "SECURITY_MACRO_DETECTED",
    "PLUGIN_NOT_TRUSTED",
    "PLUGIN_HASH_MISMATCH",
    "PLUGIN_PERMISSION_DENIED",
    "RENDERER_NOT_TRUSTED",
    "SELECTOR_NOT_FOUND",
    "SELECTOR_AMBIGUOUS",
    "EDIT_TRANSACTION_FAILED",
    "IDEMPOTENCY_REPLAY",
    "TEXT_OVERFLOW",
    "IMAGE_MISSING",
    "ASSET_UNSUPPORTED_FORMAT",
    "CHART_SPEC_INVALID",
    "DIAGRAM_SPEC_INVALID",
    "VIEW_FIDELITY_LOW",
    "EXPORT_UNSUPPORTED"
];
const defaultCategory = (code) => {
    if (code.startsWith("FEATURE_") || code === "UNKNOWN_COMMAND" || code === "UNKNOWN_OPTION" || code === "CAPABILITIES_STALE") {
        return "capability";
    }
    if (code === "INPUT_NOT_FOUND")
        return "input";
    if (code.startsWith("SCHEMA_"))
        return "schema";
    if (code.startsWith("SECURITY_"))
        return "security";
    if (code.startsWith("PLUGIN_"))
        return "plugin";
    if (code.startsWith("RENDERER_"))
        return "renderer";
    if (code.startsWith("SELECTOR_"))
        return "edit.selector";
    if (code.startsWith("EDIT_") || code === "IDEMPOTENCY_REPLAY")
        return "edit";
    if (code === "TEXT_OVERFLOW" || code === "VIEW_FIDELITY_LOW")
        return "layout";
    if (code === "IMAGE_MISSING" || code === "ASSET_UNSUPPORTED_FORMAT")
        return "asset";
    if (code === "CHART_SPEC_INVALID")
        return "chart";
    if (code === "DIAGRAM_SPEC_INVALID")
        return "diagram";
    return "export";
};
const defaultSeverity = (code) => {
    if (code === "SECURITY_MACRO_DETECTED" || code === "VIEW_FIDELITY_LOW" || code === "TEXT_OVERFLOW") {
        return "warning";
    }
    return "error";
};
const messages = {
    FEATURE_DISABLED: "The feature is disabled by the active configuration.",
    FEATURE_HIDDEN_FROM_AGENT: "The feature is hidden from agent-visible capabilities.",
    UNKNOWN_COMMAND: "The command is not known to this Officegen CLI.",
    UNKNOWN_OPTION: "The option is not supported by this command.",
    CAPABILITIES_STALE: "The embedded capabilities hash does not match the active configuration.",
    INPUT_NOT_FOUND: "The input file was not found.",
    SCHEMA_INVALID: "The JSON document does not match the requested schema.",
    SCHEMA_DEPRECATED: "The requested schema is deprecated.",
    SCHEMA_MIGRATION_REQUIRED: "The JSON document requires schema migration.",
    SECURITY_PATH_OUTSIDE_ROOT: "The path resolves outside allowed roots.",
    SECURITY_ABSOLUTE_OUT_DENIED: "Absolute output paths are denied by policy.",
    SECURITY_SYMLINK_DENIED: "Symlink or reparse-point traversal is denied by policy.",
    SECURITY_HARDLINK_DENIED: "Writing to hardlinked files is denied by policy.",
    SECURITY_ZIP_BOMB_DETECTED: "The archive exceeds safe zip limits.",
    SECURITY_XML_ENTITY_DENIED: "XML entity declarations are denied by policy.",
    SECURITY_MACRO_DETECTED: "The input file contains VBA or macro parts.",
    PLUGIN_NOT_TRUSTED: "The plugin is not trusted.",
    PLUGIN_HASH_MISMATCH: "The plugin hash does not match the trust store.",
    PLUGIN_PERMISSION_DENIED: "The plugin requested a denied permission.",
    RENDERER_NOT_TRUSTED: "The renderer is not trusted.",
    SELECTOR_NOT_FOUND: "The selector matched no editable objects.",
    SELECTOR_AMBIGUOUS: "The selector matched multiple editable objects.",
    EDIT_TRANSACTION_FAILED: "The edit transaction failed and was rolled back.",
    IDEMPOTENCY_REPLAY: "The operation matched a previous idempotency key.",
    TEXT_OVERFLOW: "Text does not fit inside the target shape.",
    IMAGE_MISSING: "An expected image asset is missing.",
    ASSET_UNSUPPORTED_FORMAT: "The asset format is not supported.",
    CHART_SPEC_INVALID: "The chart specification is invalid.",
    DIAGRAM_SPEC_INVALID: "The diagram specification is invalid.",
    VIEW_FIDELITY_LOW: "The generated view has low fidelity.",
    EXPORT_UNSUPPORTED: "The requested export conversion is unsupported."
};
const suggestedOps = {
    TEXT_OVERFLOW: ["pptx.fitText", "pptx.setShapeText with shorter text", "layout.apply"],
    SELECTOR_AMBIGUOUS: ["edit --dry-run --resolve-selectors", "inspect --summary"],
    SELECTOR_NOT_FOUND: ["inspect --summary", "view --object-map"],
    CHART_SPEC_INVALID: ["schema validate --schema officegen.chart.vegalite-wrapper@1.2"],
    DIAGRAM_SPEC_INVALID: ["schema validate --schema officegen.diagram.spec@1.2"]
};
export const ERROR_CATALOG = Object.fromEntries(REQUIRED_ERROR_CODES.map((code) => [
    code,
    {
        code,
        category: defaultCategory(code),
        severity: defaultSeverity(code),
        message: messages[code] ?? code,
        typicalCause: messages[code] ?? code,
        suggestedOps: suggestedOps[code] ?? [],
        nextSuggestedCommands: ["officegen capabilities --agent --json", "officegen errors inspect " + code + " --json"]
    }
]));
export class OfficegenError extends Error {
    payload;
    constructor(code, message, details, overrides = {}) {
        const catalog = ERROR_CATALOG[code];
        super(message ?? catalog.message);
        this.name = "OfficegenError";
        this.payload = {
            code,
            category: overrides.category ?? catalog.category,
            severity: overrides.severity ?? catalog.severity,
            message: message ?? catalog.message,
            feature: overrides.feature,
            command: overrides.command,
            details: details ?? overrides.details
        };
    }
}
export function getRequiredErrorCodes() {
    return [...REQUIRED_ERROR_CODES];
}
export function listErrors() {
    return REQUIRED_ERROR_CODES.map((code) => ERROR_CATALOG[code]);
}
export function inspectError(code) {
    return ERROR_CATALOG[code];
}
export function createErrorPayload(code, options = {}) {
    const catalog = ERROR_CATALOG[code];
    return {
        code,
        category: options.category ?? catalog.category,
        severity: options.severity ?? catalog.severity,
        message: options.message ?? catalog.message,
        feature: options.feature,
        command: options.command,
        details: options.details
    };
}
export function assertKnownErrorCode(code) {
    if (!(code in ERROR_CATALOG)) {
        throw new OfficegenError("SCHEMA_INVALID", `Unknown Officegen error code: ${code}`);
    }
}
//# sourceMappingURL=errors.js.map