import { OFFICEGEN_CLI_VERSION, redactJson, type JsonValue } from "@officegen/core";
import { availableCommands, nextSuggestedCommands } from "./context.js";
import { ENVELOPE_SCHEMA, type CliErrorPayload, type Envelope, type RuntimeContext } from "./types.js";

export function makeEnvelope(context: RuntimeContext, command: string, data: unknown, now: Date): Envelope {
  const result = redactForJson(data, context);
  const warnings = [...contextWarnings(context), ...extractArrayField(result, "warnings")];
  return {
    schema: ENVELOPE_SCHEMA,
    ok: true,
    command,
    runId: runId(now),
    cliVersion: OFFICEGEN_CLI_VERSION,
    capabilitiesHash: context.capabilitiesHash,
    pathsRedacted: true,
    result,
    warnings,
    diagnostics: extractArrayField(result, "diagnostics"),
    artifacts: extractArrayField(result, "artifacts"),
    availableCommands: availableCommands(context),
    nextSuggestedCommands: nextSuggestedCommands(context)
  };
}

export function makeErrorEnvelope(
  context: RuntimeContext,
  command: string,
  error: CliErrorPayload,
  now: Date
): Envelope {
  return {
    schema: ENVELOPE_SCHEMA,
    ok: false,
    command,
    runId: runId(now),
    cliVersion: OFFICEGEN_CLI_VERSION,
    capabilitiesHash: context.capabilitiesHash,
    pathsRedacted: true,
    error: normalizeCliError(error),
    warnings: contextWarnings(context),
    diagnostics: [],
    artifacts: [],
    availableCommands: availableCommands(context),
    nextSuggestedCommands: errorSuggestedCommands(context, error)
  };
}

export function writeResult(context: RuntimeContext, envelope: Envelope, writer: (text: string) => void): void {
  const safeEnvelope = redactForJson(envelope, context) as Envelope;
  if (context.json) {
    writer(JSON.stringify(applyAgentBudget(context, safeEnvelope), null, 2));
    return;
  }

  if (!safeEnvelope.ok) {
    writer(`${safeEnvelope.error?.code ?? "ERROR"}: ${safeEnvelope.error?.message ?? "Command failed"}`);
    return;
  }

  const summary = safeEnvelope.result && typeof safeEnvelope.result === "object" && "summary" in safeEnvelope.result
    ? String((safeEnvelope.result as { summary?: unknown }).summary)
    : `${safeEnvelope.command} completed. Use --json for the v1.2 envelope.`;
  writer(summary);
}

export function normalizeCliError(error: CliErrorPayload): CliErrorPayload {
  const security = error.code.startsWith("SECURITY_") || error.code.includes("TRUST") || error.code.includes("PLUGIN");
  const schema = error.code.startsWith("SCHEMA_") || error.code.includes("VALIDATION");
  const feature = error.code.startsWith("FEATURE_") || error.code === "UNKNOWN_COMMAND" || error.code === "UNKNOWN_OPTION";
  const input = error.code === "INPUT_NOT_FOUND";
  return {
    ...error,
    category: error.category ?? (security ? "security" : schema ? "schema" : feature ? "usage" : input ? "input" : "runtime"),
    severity: error.severity ?? (security ? "critical" : "error")
  };
}

export function redactForJson(value: unknown, context: RuntimeContext): unknown {
  return redactJson(value as JsonValue, context.config).value;
}

export function runId(now: Date): string {
  return now.toISOString().replace(/[:.]/g, "-");
}

function contextWarnings(context: RuntimeContext): unknown[] {
  return context.staleCapabilitiesWarning ? [context.staleCapabilitiesWarning] : [];
}

function extractArrayField(value: unknown, field: string): unknown[] {
  if (value && typeof value === "object" && Array.isArray((value as Record<string, unknown>)[field])) {
    return (value as Record<string, unknown>)[field] as unknown[];
  }
  if (field === "artifacts" && value && typeof value === "object") {
    const out = (value as Record<string, unknown>).out;
    return typeof out === "string" ? [{ path: out }] : [];
  }
  return [];
}

function applyAgentBudget(context: RuntimeContext, envelope: Envelope): Envelope {
  if (!context.agent || !context.jsonBudgetBytes) return envelope;
  const bytes = Buffer.byteLength(JSON.stringify(envelope, null, 2), "utf8");
  if (bytes <= context.jsonBudgetBytes) return envelope;

  const result = envelope.result;
  const resultSchema = result && typeof result === "object" ? (result as Record<string, unknown>).schema : undefined;
  const compact: Envelope = {
    ...envelope,
    result: {
      schema: "officegen.progressive-disclosure@1.2",
      status: "truncated",
      truncated: true,
      resultSchema,
      budgetBytes: context.jsonBudgetBytes,
      originalBytes: bytes,
      capabilitiesHash: context.capabilitiesHash,
      message: "JSON output exceeded the agent budget. Re-run with a narrower command or a larger --json-budget-bytes value."
    },
    warnings: [
      ...envelope.warnings,
      {
        code: "AGENT_JSON_BUDGET_EXCEEDED",
        severity: "warning",
        budgetBytes: context.jsonBudgetBytes,
        originalBytes: bytes
      }
    ],
    diagnostics: [],
    artifacts: []
  };

  return compact;
}

function errorSuggestedCommands(context: RuntimeContext, error: CliErrorPayload): string[] {
  const agent = context.agent ? " --agent" : "";
  const command = error.command ?? context.argv.slice(2).join(" ");
  const suggestions: string[] = [];
  if (error.code === "INPUT_PARSE_ERROR") {
    suggestions.push(`officegen schema validate <input.json> --schema officegen.ir.document@1.2${agent} --json`);
  } else if (error.code === "INPUT_NOT_FOUND") {
    suggestions.push(`officegen inspect <existing-file> --depth summary${agent} --json`);
  } else if (error.code === "SELECTOR_NOT_FOUND" || error.code === "SELECTOR_AMBIGUOUS") {
    suggestions.push(`officegen inspect <input> --depth summary${agent} --json`);
    suggestions.push("officegen view <input> --out .officegen/runs/view --json");
    suggestions.push("officegen edit <input> --ops ops.json --dry-run --resolve-selectors --agent --json");
  } else if (error.code === "FEATURE_DISABLED") {
    suggestions.push("officegen config show --json");
    suggestions.push(`OFFICEGEN_PROFILE=authoring officegen ${command}${agent} --json`);
  } else if (error.code === "ASSET_UNSUPPORTED_FORMAT") {
    suggestions.push("officegen asset inspect <replacement> --json");
    suggestions.push("officegen asset extract <input> --images --out .officegen/runs/assets --json");
  } else if (error.code === "EXPORT_UNSUPPORTED" || error.code === "TARGET_EXTENSION_MISMATCH") {
    suggestions.push("officegen export <input> --to pdf --mode fast --out output.pdf --json");
  } else if (error.code === "UNSUPPORTED_FORMAT") {
    suggestions.push("officegen inspect <input.pptx|input.docx|input.xlsx|input.pdf> --depth summary --agent --json");
    suggestions.push("officegen asset inspect <image-or-media-file> --json");
  } else if (error.code === "SCHEMA_INVALID") {
    suggestions.push("officegen schema get officegen.ir.document@1.2 --json");
    suggestions.push("officegen scaffold --kind pptx --title \"Draft\" --out draft.ir.json --json");
  }
  return [...new Set([...suggestions, ...nextSuggestedCommands(context)])];
}
