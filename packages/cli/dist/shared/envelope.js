import { OFFICEGEN_CLI_VERSION, redactJson } from "../../../core/dist/index.js";
import { availableCommands, nextSuggestedCommands } from "./context.js";
import { ENVELOPE_SCHEMA } from "./types.js";
export function makeEnvelope(context, command, data, now) {
    const result = redactForJson(data, context);
    const warnings = [...contextWarnings(context), ...extractArrayField(result, "warnings")];
    const artifacts = extractArrayField(result, "artifacts");
    const missingArtifact = artifacts.find((artifact) => artifact && typeof artifact === "object" && artifact.exists === false);
    if (missingArtifact) {
        return {
            schema: ENVELOPE_SCHEMA,
            ok: false,
            command,
            runId: runId(now),
            cliVersion: OFFICEGEN_CLI_VERSION,
            capabilitiesHash: context.capabilitiesHash,
            pathsRedacted: true,
            result,
            error: normalizeCliError({
                code: "EDIT_TRANSACTION_FAILED",
                command,
                message: "Expected output artifact was not created.",
                details: { artifacts }
            }),
            warnings,
            diagnostics: extractArrayField(result, "diagnostics"),
            artifacts,
            availableCommands: availableCommands(context),
            nextSuggestedCommands: errorSuggestedCommands(context, { code: "EDIT_TRANSACTION_FAILED", command, message: "Expected output artifact was not created.", details: { artifacts } })
        };
    }
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
        artifacts,
        availableCommands: availableCommands(context),
        nextSuggestedCommands: nextSuggestedCommands(context)
    };
}
export function makeErrorEnvelope(context, command, error, now) {
    const details = error.details ?? {};
    const artifacts = Array.isArray(details.artifacts) ? details.artifacts : [];
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
        artifacts,
        availableCommands: availableCommands(context),
        nextSuggestedCommands: errorSuggestedCommands(context, error)
    };
}
export function writeResult(context, envelope, writer) {
    const safeEnvelope = redactForJson(envelope, context);
    if (context.json) {
        writer(JSON.stringify(applyAgentBudget(context, safeEnvelope), null, 2));
        return;
    }
    if (!safeEnvelope.ok) {
        writer(`${safeEnvelope.error?.code ?? "ERROR"}: ${safeEnvelope.error?.message ?? "Command failed"}`);
        return;
    }
    const summary = safeEnvelope.result && typeof safeEnvelope.result === "object" && "summary" in safeEnvelope.result
        ? String(safeEnvelope.result.summary)
        : `${safeEnvelope.command} completed. Use --json for the v1.2 envelope.`;
    writer(summary);
}
export function normalizeCliError(error) {
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
export function redactForJson(value, context) {
    return redactJson(value, context.config).value;
}
export function runId(now) {
    return now.toISOString().replace(/[:.]/g, "-");
}
function contextWarnings(context) {
    return context.staleCapabilitiesWarning ? [context.staleCapabilitiesWarning] : [];
}
function extractArrayField(value, field) {
    if (value && typeof value === "object" && Array.isArray(value[field])) {
        return value[field];
    }
    if (field === "artifacts" && value && typeof value === "object") {
        const out = value.out;
        return typeof out === "string"
            ? [{
                    path: out,
                    exists: value.changed === false ? false : undefined,
                    kind: "output",
                    sourceCommand: value.kind
                }]
            : [];
    }
    return [];
}
function applyAgentBudget(context, envelope) {
    if (!context.agent || !context.jsonBudgetBytes)
        return envelope;
    const bytes = Buffer.byteLength(JSON.stringify(envelope, null, 2), "utf8");
    if (bytes <= context.jsonBudgetBytes)
        return envelope;
    const result = envelope.result;
    const resultSchema = result && typeof result === "object" ? result.schema : undefined;
    const artifacts = envelope.artifacts;
    const counts = result && typeof result === "object" ? summarizeCounts(result) : {};
    const compact = {
        ...envelope,
        result: {
            schema: "officegen.progressive-disclosure@1.2",
            status: "truncated",
            truncated: true,
            resultSchema,
            partialSummary: {
                resultSchema,
                counts,
                artifactCount: artifacts.length,
                originalBytes: bytes
            },
            budgetBytes: context.jsonBudgetBytes,
            originalBytes: bytes,
            capabilitiesHash: context.capabilitiesHash,
            message: "JSON output exceeded the agent budget. Re-run with a narrower command or a larger --json-budget-bytes value.",
            recommendedNarrowCommands: recommendedNarrowCommands(envelope.command, context)
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
        artifacts
    };
    return compact;
}
function summarizeCounts(value) {
    const counts = {};
    for (const [key, nested] of Object.entries(value)) {
        if (Array.isArray(nested))
            counts[key] = nested.length;
        if (nested && typeof nested === "object" && !Array.isArray(nested)) {
            const objectCounts = Object.entries(nested)
                .filter(([, item]) => Array.isArray(item))
                .map(([nestedKey, item]) => [`${key}.${nestedKey}`, item.length]);
            for (const [nestedKey, count] of objectCounts)
                counts[nestedKey] = count;
        }
    }
    return counts;
}
function errorSuggestedCommands(context, error) {
    const agent = context.agent ? " --agent" : "";
    const command = error.command ?? context.argv.slice(2).join(" ");
    const suggestions = [];
    if (error.code === "UNKNOWN_COMMAND") {
        const attempted = String(error.command ?? "");
        const alias = commandAliasSuggestion(attempted, context);
        if (alias)
            suggestions.push(alias);
    }
    else if (error.code === "INPUT_PARSE_ERROR") {
        suggestions.push(`officegen schema validate <input.json> --schema officegen.ir.document@1.2${agent} --json`);
    }
    else if (error.code === "INPUT_NOT_FOUND") {
        suggestions.push(`officegen inspect <existing-file> --depth summary${agent} --json`);
    }
    else if (error.code === "SELECTOR_NOT_FOUND" || error.code === "SELECTOR_AMBIGUOUS") {
        suggestions.push(`officegen inspect <input> --depth summary${agent} --json`);
        suggestions.push("officegen view <input> --out .officegen/runs/view --json");
        suggestions.push("officegen edit <input> --ops ops.json --dry-run --resolve-selectors --agent --json");
    }
    else if (error.code === "FEATURE_DISABLED") {
        suggestions.push("officegen config show --json");
        suggestions.push(`OFFICEGEN_PROFILE=authoring officegen ${command}${agent} --json`);
    }
    else if (error.code === "ASSET_UNSUPPORTED_FORMAT") {
        suggestions.push("officegen asset inspect <replacement> --json");
        suggestions.push("officegen asset extract <input> --images --out .officegen/runs/assets --json");
    }
    else if (error.code === "EXPORT_UNSUPPORTED" || error.code === "TARGET_EXTENSION_MISMATCH") {
        suggestions.push("officegen export <input> --to pdf --mode fast --out output.pdf --json");
    }
    else if (error.code === "SECURITY_ABSOLUTE_OUT_DENIED") {
        suggestions.push("officegen <command> <input> --out .officegen/outputs/output.ext --json");
        suggestions.push("officegen config show --json");
    }
    else if (error.code === "UNSUPPORTED_FORMAT") {
        suggestions.push("officegen inspect <input.pptx|input.docx|input.xlsx|input.pdf> --depth summary --agent --json");
        suggestions.push("officegen asset inspect <image-or-media-file> --json");
    }
    else if (error.code === "SCHEMA_INVALID") {
        suggestions.push("officegen schema get officegen.ir.document@1.2 --json");
        suggestions.push("officegen scaffold --kind pptx --title \"Draft\" --out draft.ir.json --json");
    }
    return [...new Set([...suggestions, ...nextSuggestedCommands(context)])];
}
function recommendedNarrowCommands(command, context) {
    const agent = context.agent ? " --agent" : "";
    if (command.startsWith("inspect")) {
        return [
            `officegen inspect <input> --depth summary --object-map-limit 50${agent} --json`,
            `officegen inspect <deck.pptx> --slides 1-5 --depth summary${agent} --json`,
            `officegen inspect <workbook.xlsx> --sheet Sheet1 --range A1:K40${agent} --json`,
            `officegen inspect <file> --fields schema,trusted,objectMap${agent} --json`
        ];
    }
    if (command.startsWith("view")) {
        return [
            "officegen view <input> --max-pages 3 --object-map-limit 50 --out .officegen/runs/view --json",
            "officegen view <pdf> --pages 1-3 --out .officegen/runs/pdf-view --json"
        ];
    }
    if (command.startsWith("verify")) {
        return [`officegen verify <input> --timeout-ms 60000${agent} --json`];
    }
    return [`officegen ${command || "<command>"} --json-budget-bytes ${Math.max((context.jsonBudgetBytes ?? 32768) * 2, 65536)}${agent} --json`];
}
function commandAliasSuggestion(command, context) {
    const agent = context.agent ? " --agent" : "";
    if (command.startsWith("schema fetch")) {
        const schemaId = command.split(/\s+/).slice(2).join(" ") || "officegen.ir.document@1.2";
        return `officegen schema get ${schemaId}${agent} --json`;
    }
    if (command === "schema")
        return `officegen schema list${agent} --json`;
    return undefined;
}
//# sourceMappingURL=envelope.js.map