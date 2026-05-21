import type {
  ErrorCatalogEntry,
  ErrorSeverity,
  JsonValue,
  OfficegenErrorCode,
  OfficegenErrorPayload
} from "./types.js";

const REQUIRED_ERROR_CODES: OfficegenErrorCode[] = [
  "FEATURE_DISABLED",
  "FEATURE_NOT_IMPLEMENTED",
  "FEATURE_HIDDEN_FROM_AGENT",
  "UNKNOWN_COMMAND",
  "UNKNOWN_OPTION",
  "OPTION_NOT_EFFECTIVE",
  "CAPABILITIES_STALE",
  "INPUT_NOT_FOUND",
  "INPUT_PARSE_ERROR",
  "SCHEMA_INVALID",
  "SCHEMA_DEPRECATED",
  "SCHEMA_MIGRATION_REQUIRED",
  "SECURITY_PATH_OUTSIDE_ROOT",
  "SECURITY_ABSOLUTE_OUT_DENIED",
  "SECURITY_SYMLINK_DENIED",
  "SECURITY_HARDLINK_DENIED",
  "SECURITY_INPUT_TOO_LARGE",
  "SECURITY_EXTERNAL_PROCESS_DENIED",
  "SECURITY_ZIP_BOMB_DETECTED",
  "SECURITY_XML_ENTITY_DENIED",
  "SECURITY_MACRO_DETECTED",
  "SECURITY_RISKY_OOXML_DETECTED",
  "BENCHMARK_MANIFEST_PATH_DENIED",
  "PLUGIN_NOT_TRUSTED",
  "PLUGIN_HASH_MISMATCH",
  "PLUGIN_PERMISSION_DENIED",
  "RENDERER_NOT_TRUSTED",
  "SELECTOR_NOT_FOUND",
  "SELECTOR_AMBIGUOUS",
  "EDIT_TRANSACTION_FAILED",
  "EXPECTED_ARTIFACT_MISSING",
  "REPAIR_NO_SAFE_OPS",
  "RUN_STEP_FAILED",
  "TIMEOUT",
  "OOXML_VALIDATION_FAILED",
  "IDEMPOTENCY_REPLAY",
  "TEXT_OVERFLOW",
  "IMAGE_MISSING",
  "ASSET_UNSUPPORTED_FORMAT",
  "CHART_SPEC_INVALID",
  "DIAGRAM_SPEC_INVALID",
  "VIEW_FIDELITY_LOW",
  "RENDER_FONT_UNSUPPORTED",
  "TARGET_EXTENSION_MISMATCH",
  "UNSUPPORTED_FORMAT",
  "TEMPLATE_FILL_FAILED",
  "TEMPLATE_VALIDATE_FAILED",
  "DESIGN_NOT_INITIALIZED",
  "EXPORT_UNSUPPORTED"
];

const defaultCategory = (code: OfficegenErrorCode): string => {
  if (code.startsWith("FEATURE_") || code === "UNKNOWN_COMMAND" || code === "UNKNOWN_OPTION" || code === "OPTION_NOT_EFFECTIVE" || code === "CAPABILITIES_STALE") {
    return "capability";
  }
  if (code === "INPUT_NOT_FOUND" || code === "INPUT_PARSE_ERROR") return "input";
  if (code.startsWith("SCHEMA_")) return "schema";
  if (code.startsWith("SECURITY_")) return "security";
  if (code === "BENCHMARK_MANIFEST_PATH_DENIED") return "security";
  if (code.startsWith("PLUGIN_")) return "plugin";
  if (code.startsWith("RENDERER_")) return "renderer";
  if (code.startsWith("SELECTOR_")) return "edit.selector";
  if (code.startsWith("EDIT_") || code === "IDEMPOTENCY_REPLAY") return "edit";
  if (code === "EXPECTED_ARTIFACT_MISSING" || code === "REPAIR_NO_SAFE_OPS" || code === "RUN_STEP_FAILED" || code === "TIMEOUT" || code === "OOXML_VALIDATION_FAILED") return "runtime";
  if (code === "TEXT_OVERFLOW" || code === "VIEW_FIDELITY_LOW") return "layout";
  if (code === "IMAGE_MISSING" || code === "ASSET_UNSUPPORTED_FORMAT") return "asset";
  if (code === "TEMPLATE_FILL_FAILED" || code === "TEMPLATE_VALIDATE_FAILED") return "template";
  if (code === "DESIGN_NOT_INITIALIZED") return "design";
  if (code === "CHART_SPEC_INVALID") return "chart";
  if (code === "DIAGRAM_SPEC_INVALID") return "diagram";
  return "export";
};

const defaultSeverity = (code: OfficegenErrorCode): ErrorSeverity => {
  if (code === "SECURITY_MACRO_DETECTED" || code === "VIEW_FIDELITY_LOW" || code === "TEXT_OVERFLOW") {
    return "warning";
  }
  return "error";
};

const messages: Partial<Record<OfficegenErrorCode, string>> = {
  FEATURE_DISABLED: "The feature is disabled by the active configuration.",
  FEATURE_NOT_IMPLEMENTED: "The command or subcommand is not implemented.",
  FEATURE_HIDDEN_FROM_AGENT: "The feature is hidden from agent-visible capabilities.",
  UNKNOWN_COMMAND: "The command is not known to this Officegen CLI.",
  UNKNOWN_OPTION: "The option is not supported by this command.",
  OPTION_NOT_EFFECTIVE: "The option is accepted by the parser but is not effective for this command.",
  CAPABILITIES_STALE: "The embedded capabilities hash does not match the active configuration.",
  INPUT_NOT_FOUND: "The input file was not found.",
  INPUT_PARSE_ERROR: "The input file could not be parsed.",
  SCHEMA_INVALID: "The JSON document does not match the requested schema.",
  SCHEMA_DEPRECATED: "The requested schema is deprecated.",
  SCHEMA_MIGRATION_REQUIRED: "The JSON document requires schema migration.",
  SECURITY_PATH_OUTSIDE_ROOT: "The path resolves outside allowed roots.",
  SECURITY_ABSOLUTE_OUT_DENIED: "Absolute output paths are denied by policy.",
  SECURITY_SYMLINK_DENIED: "Symlink or reparse-point traversal is denied by policy.",
  SECURITY_HARDLINK_DENIED: "Writing to hardlinked files is denied by policy.",
  SECURITY_INPUT_TOO_LARGE: "The input file exceeds the configured size limit.",
  SECURITY_EXTERNAL_PROCESS_DENIED: "External process execution is denied by policy.",
  SECURITY_ZIP_BOMB_DETECTED: "The archive exceeds safe zip limits.",
  SECURITY_XML_ENTITY_DENIED: "XML entity declarations are denied by policy.",
  SECURITY_MACRO_DETECTED: "The input file contains VBA or macro parts.",
  SECURITY_RISKY_OOXML_DETECTED: "The Office package contains risky OOXML parts that are blocked for mutation.",
  BENCHMARK_MANIFEST_PATH_DENIED: "The benchmark manifest references a path outside the benchmark storage root.",
  PLUGIN_NOT_TRUSTED: "The plugin is not trusted.",
  PLUGIN_HASH_MISMATCH: "The plugin hash does not match the trust store.",
  PLUGIN_PERMISSION_DENIED: "The plugin requested a denied permission.",
  RENDERER_NOT_TRUSTED: "The renderer is not trusted.",
  SELECTOR_NOT_FOUND: "The selector matched no editable objects.",
  SELECTOR_AMBIGUOUS: "The selector matched multiple editable objects.",
  EDIT_TRANSACTION_FAILED: "The edit transaction failed and was rolled back.",
  EXPECTED_ARTIFACT_MISSING: "An expected output artifact was not created.",
  REPAIR_NO_SAFE_OPS: "No automatically safe repair operations were available.",
  RUN_STEP_FAILED: "A workflow run step failed.",
  TIMEOUT: "The operation exceeded its timeout budget.",
  OOXML_VALIDATION_FAILED: "The Office package failed OOXML validation after mutation.",
  IDEMPOTENCY_REPLAY: "The operation matched a previous idempotency key.",
  TEXT_OVERFLOW: "Text does not fit inside the target shape.",
  IMAGE_MISSING: "An expected image asset is missing.",
  ASSET_UNSUPPORTED_FORMAT: "The asset format is not supported.",
  CHART_SPEC_INVALID: "The chart specification is invalid.",
  DIAGRAM_SPEC_INVALID: "The diagram specification is invalid.",
  VIEW_FIDELITY_LOW: "The generated view has low fidelity.",
  RENDER_FONT_UNSUPPORTED: "The PDF renderer cannot encode the requested text with the active font.",
  TARGET_EXTENSION_MISMATCH: "The requested render target does not match the output file extension.",
  UNSUPPORTED_FORMAT: "The input format is not supported by this command.",
  TEMPLATE_FILL_FAILED: "Template fill could not create or validate the requested artifact.",
  TEMPLATE_VALIDATE_FAILED: "Template fill validation found unresolved or unsupported bindings.",
  DESIGN_NOT_INITIALIZED: "The requested design profile has not been initialized.",
  EXPORT_UNSUPPORTED: "The requested export conversion is unsupported."
};

const suggestedOps: Partial<Record<OfficegenErrorCode, string[]>> = {
  TEXT_OVERFLOW: ["pptx.fitText", "pptx.setShapeText with shorter text", "layout.apply"],
  SELECTOR_AMBIGUOUS: ["edit --dry-run --resolve-selectors", "inspect --summary"],
  SELECTOR_NOT_FOUND: ["inspect --summary", "view --object-map"],
  EXPECTED_ARTIFACT_MISSING: ["run --expected-artifacts", "inspect --depth summary"],
  REPAIR_NO_SAFE_OPS: ["diagnose --report-out diagnose.json", "edit --dry-run --resolve-selectors"],
  RUN_STEP_FAILED: ["run --log-jsonl run.jsonl --manifest run-manifest.json", "inspect --depth summary"],
  TIMEOUT: ["rerun with --timeout-ms <larger-ms>", "inspect with --fields/--range/--pages to narrow scope"],
  OOXML_VALIDATION_FAILED: ["verify --native --strict", "diagnose --report-out diagnose.json"],
  TEMPLATE_FILL_FAILED: ["template fill --validate-only", "inspect --depth shallow", "template apply-map"],
  TEMPLATE_VALIDATE_FAILED: ["template candidates <source> --agent --json", "inspect <source> --depth shallow --agent --json", "template apply-map --map corrected-map.json"],
  DESIGN_NOT_INITIALIZED: ["design init --name <name> --agent --json", "design capture <source.pptx> --name <name> --agent --json"],
  CHART_SPEC_INVALID: ["schema validate --schema officegen.chart.vegalite-wrapper@1.2"],
  DIAGRAM_SPEC_INVALID: ["schema validate --schema officegen.diagram.spec@1.2"]
};

export const ERROR_CATALOG: Record<OfficegenErrorCode, ErrorCatalogEntry> = Object.fromEntries(
  REQUIRED_ERROR_CODES.map((code) => [
    code,
    {
      code,
      category: defaultCategory(code),
      severity: defaultSeverity(code),
      message: messages[code] ?? code,
      typicalCause: messages[code] ?? code,
      suggestedOps: suggestedOps[code] ?? [],
      nextSuggestedCommands: ["officegen capabilities --agent --strict-json", "officegen errors inspect " + code + " --agent --strict-json"]
    }
  ])
) as Record<OfficegenErrorCode, ErrorCatalogEntry>;

export class OfficegenError extends Error {
  readonly payload: OfficegenErrorPayload;

  constructor(code: OfficegenErrorCode, message?: string, details?: JsonValue, overrides: Partial<OfficegenErrorPayload> = {}) {
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

export function getRequiredErrorCodes(): OfficegenErrorCode[] {
  return [...REQUIRED_ERROR_CODES];
}

export function listErrors(): ErrorCatalogEntry[] {
  return REQUIRED_ERROR_CODES.map((code) => ERROR_CATALOG[code]);
}

export function inspectError(code: OfficegenErrorCode): ErrorCatalogEntry {
  return ERROR_CATALOG[code];
}

export function createErrorPayload(
  code: OfficegenErrorCode,
  options: Partial<OfficegenErrorPayload> = {}
): OfficegenErrorPayload {
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

export function assertKnownErrorCode(code: string): asserts code is OfficegenErrorCode {
  if (!(code in ERROR_CATALOG)) {
    throw new OfficegenError("SCHEMA_INVALID", `Unknown Officegen error code: ${code}`);
  }
}
