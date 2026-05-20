import { OFFICEGEN_CLI_VERSION, redactJson, type JsonValue } from "@officegen/core";
import { availableCommands, nextSuggestedCommands } from "./context.js";
import { ENVELOPE_SCHEMA, RUNTIME_ENVELOPE_SCHEMA, type CliErrorPayload, type Envelope, type RuntimeContext } from "./types.js";

export function makeEnvelope(context: RuntimeContext, command: string, data: unknown, now: Date): Envelope {
  const result = redactForJson(data, context);
  const warnings = [...contextWarnings(context), ...extractArrayField(result, "warnings")];
  const diagnostics = extractArrayField(result, "diagnostics");
  const objective = evaluateObjective(context, command, result, extractArrayField(result, "artifacts"));
  if (!objective.ok) {
    const error = normalizeCliError({
      code: objective.code,
      command,
      message: objective.message,
      details: { ...objective.details, artifacts: objective.artifacts }
    });
    const nextActions = errorSuggestedCommands(context, error);
    return {
      schema: ENVELOPE_SCHEMA,
      runtimeEnvelope: RUNTIME_ENVELOPE_SCHEMA,
      ok: false,
      command,
      runId: runId(now),
      cliVersion: OFFICEGEN_CLI_VERSION,
      capabilitiesHash: context.capabilitiesHash,
      pathsRedacted: true,
      executionOk: true,
      objectiveOk: false,
      mutationStatus: objective.mutationStatus,
      artifactStatus: objective.artifactStatus,
      readiness: objective.readiness,
      partial: objective.partial,
      failureClass: failureClassFor({ error, result, readiness: objective.readiness, partial: objective.partial }),
      result,
      error,
      warnings,
      diagnostics,
      artifacts: objective.artifacts,
      availableCommands: availableCommands(context),
      nextSuggestedCommands: nextActions,
      nextActions
    };
  }
  const nextActions = nextSuggestedCommands(context);
  return {
    schema: ENVELOPE_SCHEMA,
    runtimeEnvelope: RUNTIME_ENVELOPE_SCHEMA,
    ok: true,
    command,
    runId: runId(now),
    cliVersion: OFFICEGEN_CLI_VERSION,
    capabilitiesHash: context.capabilitiesHash,
    pathsRedacted: true,
    executionOk: true,
    objectiveOk: objective.objectiveOk,
    mutationStatus: objective.mutationStatus,
    artifactStatus: objective.artifactStatus,
    readiness: objective.readiness,
    partial: objective.partial,
    failureClass: failureClassFor({ result, readiness: objective.readiness, partial: objective.partial }),
    result,
    warnings,
    diagnostics,
    artifacts: objective.artifacts,
    availableCommands: availableCommands(context),
    nextSuggestedCommands: nextActions,
    nextActions
  };
}

export function makeErrorEnvelope(
  context: RuntimeContext,
  command: string,
  error: CliErrorPayload,
  now: Date
): Envelope {
  const details = error.details ?? {};
  const artifacts = Array.isArray(details.artifacts) ? details.artifacts : [];
  const normalizedError = normalizeCliError(error);
  const nextActions = errorSuggestedCommands(context, normalizedError);
  return {
    schema: ENVELOPE_SCHEMA,
    runtimeEnvelope: RUNTIME_ENVELOPE_SCHEMA,
    ok: false,
    command,
    runId: runId(now),
    cliVersion: OFFICEGEN_CLI_VERSION,
    capabilitiesHash: context.capabilitiesHash,
    pathsRedacted: true,
    executionOk: false,
    objectiveOk: false,
    mutationStatus: "failed",
    artifactStatus: artifacts.some((artifact) => artifact && typeof artifact === "object" && (artifact as Record<string, unknown>).exists === false) ? "missing" : "not_expected",
    readiness: "blocked",
    partial: false,
    failureClass: failureClassFor({ error: normalizedError, readiness: "blocked", partial: false }),
    error: normalizedError,
    warnings: contextWarnings(context),
    diagnostics: [],
    artifacts,
    availableCommands: availableCommands(context),
    nextSuggestedCommands: nextActions,
    nextActions
  };
}

export function writeResult(context: RuntimeContext, envelope: Envelope, writer: (text: string) => void): void {
  const safeEnvelope = redactForJson(envelope, context) as Envelope;
  if (!safeEnvelope.ok && process.exitCode === undefined) process.exitCode = exitCodeForError(safeEnvelope.error?.code);
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

function failureClassFor(input: {
  error?: CliErrorPayload;
  result?: unknown;
  readiness?: Envelope["readiness"];
  partial?: boolean;
}): NonNullable<Envelope["failureClass"]> {
  const error = input.error ? normalizeCliError(input.error) : undefined;
  const code = error?.code ?? "";
  const record = asRecord(input.result);
  const readiness = input.readiness ?? readinessFor(record);
  const partial = input.partial === true || record.partial === true || record.truncated === true || record.status === "truncated";

  if (code.includes("UNSUPPORTED") || code === "FEATURE_NOT_IMPLEMENTED") return "unsupported";
  if (partial) return "partial";
  if (isDoctorRuntimeFailure(record, code)) return "runtime";
  if (!error) return "none";
  if (error.category === "security") return "security";
  if (error.category === "schema") return "schema";
  if (error.category === "input") return "input";
  if (error.category === "usage") return "usage";
  if (readiness === "blocked" || code === "VISUAL_DIFF_BLOCKED" || code === "RUN_STEP_FAILED" || code === "TIMEOUT") return "blocked";
  return "runtime";
}

function isDoctorRuntimeFailure(record: Record<string, unknown>, code: string): boolean {
  if (code === "RUNTIME_READINESS_FAILED") return true;
  if (record.schema !== "officegen.doctor@1.2") return false;
  return extractArrayField(record, "checks").some((check) => {
    const item = asRecord(check);
    return item.id === "node" && item.ok === false;
  });
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
    return typeof out === "string"
      ? [{
          path: out,
          exists: (value as Record<string, unknown>).changed === false ? false : undefined,
          kind: "output",
          sourceCommand: (value as Record<string, unknown>).kind
        }]
      : [];
  }
  return [];
}

interface ObjectiveEvaluation {
  ok: boolean;
  objectiveOk: boolean;
  code: string;
  message: string;
  details?: Record<string, unknown>;
  artifacts: unknown[];
  mutationStatus: NonNullable<Envelope["mutationStatus"]>;
  artifactStatus: NonNullable<Envelope["artifactStatus"]>;
  readiness: NonNullable<Envelope["readiness"]>;
  partial: boolean;
}

function evaluateObjective(context: RuntimeContext, command: string, result: unknown, initialArtifacts: unknown[]): ObjectiveEvaluation {
  const record = asRecord(result);
  const schema = typeof record.schema === "string" ? record.schema : "";
  const outArg = optionFromArgv(context.argv, "--out");
  const dryRun = hasFlagArg(context.argv, "--dry-run") || record.planOnly === true || record.dryRun === true;
  const artifacts = ensureRequestedArtifact(initialArtifacts, outArg, record, command, dryRun);
  const missingArtifact = artifacts.find((artifact) => artifact && typeof artifact === "object" && (artifact as Record<string, unknown>).exists === false) as Record<string, unknown> | undefined;
  const readiness = readinessFor(record);
  const partial = record.partial === true || record.truncated === true || record.status === "truncated";
  const defaultState: Omit<ObjectiveEvaluation, "ok" | "objectiveOk" | "code" | "message"> = {
    details: undefined,
    artifacts,
    mutationStatus: mutationStatusFor(command, record, dryRun),
    artifactStatus: missingArtifact ? "missing" : artifacts.length ? "complete" : "not_expected",
    readiness,
    partial
  };

  if (record.status === "not_implemented" || record.status === "wired") {
    return objectiveFailure(defaultState, "FEATURE_NOT_IMPLEMENTED", `${command} is not implemented as a mutating command.`, { status: record.status });
  }
  if (record.planOnly === true && outArg && isOfficeOutputPath(outArg)) {
    return objectiveFailure(defaultState, "FEATURE_NOT_IMPLEMENTED", `${command} produced a plan-only result for an Office output path.`, {
      out: outArg,
      planOnly: true,
      expectedResult: "No Office artifact was created; use a JSON plan output path or a supported mutating command."
    });
  }
  if (schema === "officegen.edit.result@1.2") {
    const errors = editRequiredFailures(record);
    const allowPartial = hasFlagArg(context.argv, "--allow-partial") || record.allowPartial === true;
    const applied = Number(record.applied ?? 0);
    const hasAppliedOp = applied > 0 || editHasAppliedOperation(record);
    if (errors.length) {
      const code = editErrorCode(errors);
      if (allowPartial && hasAppliedOp && !missingArtifact) {
        return {
          ok: true,
          objectiveOk: true,
          code: "",
          message: "",
          ...defaultState,
          readiness: "partial",
          partial: true
        };
      }
      return objectiveFailure({
        ...defaultState,
        mutationStatus: "failed",
        artifactStatus: missingArtifact ? "missing" : "not_expected",
        partial: defaultState.partial || hasAppliedOp
      }, code, editFailureMessage(code), { errors, allowPartial });
    }
    if (!dryRun && outArg && (record.changed === false || applied <= 0)) {
      return objectiveFailure(defaultState, "EDIT_TRANSACTION_FAILED", "Edit requested an output artifact but no operation was applied.", { changed: record.changed, applied: record.applied });
    }
  }
  const nestedEditResult = mutationEvidenceRecord(record);
  if (nestedEditResult !== record && (nestedEditResult.rolledBack === true || editRequiredFailures(nestedEditResult).length)) {
    const errors = editRequiredFailures(nestedEditResult);
    const code = errors.length ? editErrorCode(errors) : "EDIT_TRANSACTION_FAILED";
    return objectiveFailure({
      ...defaultState,
      mutationStatus: "failed",
      artifactStatus: missingArtifact ? "missing" : "not_expected",
      partial: defaultState.partial || editHasAppliedOperation(nestedEditResult)
    }, code, editFailureMessage(code), {
      rolledBack: nestedEditResult.rolledBack,
      errors,
      editResult: nestedEditResult
    });
  }
  if (schema === "officegen.repair.result@1.2" && !dryRun && outArg && (record.changed === false || Number(record.applied ?? 0) <= 0)) {
    return objectiveFailure(defaultState, "REPAIR_NO_SAFE_OPS", "No automatically safe repair operations were available.", { changed: record.changed, applied: record.applied });
  }
  if (missingArtifact) {
    return objectiveFailure(defaultState, "EXPECTED_ARTIFACT_MISSING", "Expected output artifact was not created.", { missingArtifact });
  }
  const resultError = asRecord(record.error);
  const successfulErrorInspection = schema === "officegen.error@1.2" && resultError.exitCode === undefined;
  if (typeof resultError.code === "string" && !successfulErrorInspection) {
    return objectiveFailure(defaultState, resultError.code, typeof resultError.message === "string" ? resultError.message : "Command objective failed.", { error: resultError });
  }
  if (schema.startsWith("officegen.run.") && Array.isArray(record.steps) && record.steps.some((step) => asRecord(step).ok === false)) {
    const failedStep = record.steps.map(asRecord).find((step) => step.ok === false);
    const stepError = asRecord(failedStep?.error);
    const code = typeof stepError.code === "string" ? stepError.code : "RUN_STEP_FAILED";
    const message = typeof stepError.message === "string" ? stepError.message : "One or more workflow steps failed.";
    return objectiveFailure(defaultState, code, message, { steps: record.steps, failedStep });
  }
  if (Array.isArray(record.missingExpectedArtifacts) && record.missingExpectedArtifacts.length) {
    return objectiveFailure(defaultState, "EXPECTED_ARTIFACT_MISSING", "One or more expected artifacts were not created.", { missingExpectedArtifacts: record.missingExpectedArtifacts });
  }
  if (schema === "officegen.verify.result@1.2") {
    const failedGates = verifyFailedGates(record);
    if (record.readiness === "blocked" || record.partial === true || failedGates.length) {
      return objectiveFailure(defaultState, record.partial === true ? "TIMEOUT" : "RUN_STEP_FAILED", "Verification did not reach a passing readiness state.", {
        readiness: record.readiness,
        partial: record.partial,
        failedGates
      });
    }
  }
  if (schema === "officegen.diff.result@1.2" && visualDiffBlocked(record)) {
    return objectiveFailure(defaultState, "VISUAL_DIFF_BLOCKED", "Visual diff was requested but could not run.", {
      visual: asRecord(record.visual)
    });
  }
  if (schema === "officegen.benchmark.run.result@2.5" || schema === "officegen.benchmark.run.result@2.3") {
    const count = Number(record.count ?? 0);
    const okCount = Number(record.okCount ?? 0);
    if (count === 0 || okCount === 0) {
      return objectiveFailure(defaultState, "RUN_STEP_FAILED", count === 0 ? "Benchmark manifest did not contain runnable documents." : "All benchmark documents failed.", {
        count,
        okCount,
        failureSummary: record.failureSummary
      });
    }
    if (okCount < count) {
      return objectiveFailure({
        ...defaultState,
        readiness: "warning",
        partial: true
      }, "RUN_STEP_FAILED", "Some benchmark documents failed.", {
        count,
        okCount,
        failureSummary: record.failureSummary
      });
    }
  }

  return {
    ok: true,
    objectiveOk: !partial && readiness !== "blocked",
    code: "",
    message: "",
    ...defaultState
  };
}

function objectiveFailure(
  state: Omit<ObjectiveEvaluation, "ok" | "objectiveOk" | "code" | "message">,
  code: string,
  message: string,
  details?: Record<string, unknown>
): ObjectiveEvaluation {
  return {
    ok: false,
    objectiveOk: false,
    code,
    message,
    ...state,
    details,
    mutationStatus: state.mutationStatus === "not_applicable" ? "failed" : state.mutationStatus,
    readiness: "blocked"
  };
}

function ensureRequestedArtifact(artifacts: unknown[], outArg: string | undefined, record: Record<string, unknown>, command: string, dryRun: boolean): unknown[] {
  if (!outArg || dryRun) return artifacts;
  if (artifacts.some((artifact) => artifact && typeof artifact === "object" && String((artifact as Record<string, unknown>).path ?? "") === outArg)) return artifacts;
  if (typeof record.out === "string") return artifacts;
  if (!isMutationCommand(command)) return artifacts;
  return [
    ...artifacts,
    { path: outArg, exists: false, kind: "output", sourceCommand: command, reason: "output artifact was requested but the result did not report an output path" }
  ];
}

function editErrorCode(errors: unknown[]): string {
  const reasons = errors.map((error) => String(asRecord(error).reason ?? asRecord(error).message ?? ""));
  if (reasons.some((reason) => reason.includes("stale-plan") || reason.includes("EDIT_STALE_PLAN"))) return "EDIT_TRANSACTION_FAILED";
  if (reasons.some((reason) => reason.includes("ambiguous"))) return "SELECTOR_AMBIGUOUS";
  if (reasons.some((reason) => reason.includes("not-found") || reason.includes("not found"))) return "SELECTOR_NOT_FOUND";
  return "EDIT_TRANSACTION_FAILED";
}

function editFailureMessage(code: string): string {
  if (code === "SELECTOR_AMBIGUOUS") return "Edit selector matched multiple objects.";
  if (code === "SELECTOR_NOT_FOUND") return "Edit selector matched no editable object.";
  return "Edit did not apply all required operations.";
}

function editRequiredFailures(record: Record<string, unknown>): unknown[] {
  const failures = [...extractArrayField(record, "errors"), ...extractArrayField(record, "opResults").filter(isRequiredEditFailure)];
  const seen = new Set<string>();
  return failures.filter((failure) => {
    const item = asRecord(failure);
    const key = `${String(item.operationIndex ?? "")}:${String(item.op ?? "")}:${String(item.reason ?? "")}:${String(item.message ?? "")}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function isRequiredEditFailure(value: unknown): boolean {
  const record = asRecord(value);
  if (record.applied !== false) return false;
  const reason = String(record.reason ?? "");
  return reason === "not-found"
    || reason === "ambiguous"
    || reason === "low-confidence"
    || reason === "unsupported"
    || reason === "validation-failed"
    || reason === "skipped-after-error"
    || reason === "stale-plan"
    || reason === "unsupported-selector";
}

function editHasAppliedOperation(record: Record<string, unknown>): boolean {
  return extractArrayField(record, "opResults").some((result) => asRecord(result).applied === true);
}

function mutationEvidenceRecord(record: Record<string, unknown>): Record<string, unknown> {
  if (record.changed !== undefined || record.applied !== undefined) return record;
  const editResult = asRecord(record.editResult);
  if (!Object.keys(editResult).length) return record;
  if (
    editResult.schema === "officegen.edit.result@1.2"
    || editResult.changed !== undefined
    || editResult.applied !== undefined
    || extractArrayField(editResult, "opResults").length
  ) {
    return editResult;
  }
  return record;
}

function isStatusOnlyMutationResult(record: Record<string, unknown>): boolean {
  const schema = typeof record.schema === "string" ? record.schema : "";
  return schema === "officegen.progressive-disclosure@1.2" || record.truncated === true || record.status === "truncated";
}

function wrapperMutationChanged(record: Record<string, unknown>): boolean {
  if (record.mutatesOffice !== true) return false;
  if (extractArrayField(record, "changedParts").length > 0) return true;
  const visualEffect = String(record.visualEffect ?? "");
  return visualEffect !== "" && visualEffect !== "none";
}

function wrapperMutationNoop(record: Record<string, unknown>): boolean {
  if (record.mutatesOffice !== true) return false;
  if (wrapperMutationChanged(record)) return false;
  return record.changedParts !== undefined || record.visualEffect !== undefined;
}

function mutationStatusFor(command: string, record: Record<string, unknown>, dryRun: boolean): NonNullable<Envelope["mutationStatus"]> {
  if (dryRun || record.planOnly === true) return "plan_only";
  if (!isMutationCommand(command)) return "not_applicable";
  if (isStatusOnlyMutationResult(record)) return "not_applicable";
  const evidence = mutationEvidenceRecord(record);
  if (evidence.rolledBack === true || editRequiredFailures(evidence).length) return "failed";
  if (evidence.changed === true || Number(evidence.applied ?? 0) > 0 || editHasAppliedOperation(evidence)) return "changed";
  if (wrapperMutationChanged(record)) return "changed";
  if (evidence.changed === false || (evidence.applied !== undefined && Number(evidence.applied) === 0)) return "noop";
  if (wrapperMutationNoop(record)) return "noop";
  return "not_applicable";
}

function readinessFor(record: Record<string, unknown>): NonNullable<Envelope["readiness"]> {
  const readiness = String(record.readiness ?? "");
  if (readiness === "pass" || readiness === "pass_with_environment_gap" || readiness === "warning" || readiness === "partial" || readiness === "blocked") return readiness;
  if (record.partial === true || record.status === "truncated") return "partial";
  return "pass";
}

function visualDiffBlocked(record: Record<string, unknown>): boolean {
  return asRecord(record.visual).status === "blocked";
}

function verifyFailedGates(record: Record<string, unknown>): string[] {
  const gates = asRecord(asRecord(record.verificationReport).gates);
  return Object.entries(gates)
    .filter(([, gate]) => asRecord(gate).status === "fail")
    .map(([name]) => name);
}

function isMutationCommand(command: string): boolean {
  const normalized = command.split(/\s+/).slice(0, 2).join(" ");
  return /^(edit|repair|render|export|asset replace|template fill|design apply|layout apply)/.test(normalized) || command.startsWith("edit") || command.startsWith("repair") || command.startsWith("render") || command.startsWith("export");
}

function isOfficeOutputPath(filePath: string): boolean {
  return /\.(pptx|docx|xlsx|pdf)$/i.test(filePath);
}

function optionFromArgv(argv: string[], name: string): string | undefined {
  const index = argv.indexOf(name);
  if (index < 0) return undefined;
  const value = argv[index + 1];
  return value && !value.startsWith("-") ? value : undefined;
}

function hasFlagArg(argv: string[], name: string): boolean {
  return argv.includes(name);
}

function asRecord(value: unknown): Record<string, unknown> {
  return value && typeof value === "object" && !Array.isArray(value) ? value as Record<string, unknown> : {};
}

function exitCodeForError(code: string | undefined): number {
  if (!code) return 1;
  if (code.startsWith("SECURITY_") || code === "BENCHMARK_MANIFEST_PATH_DENIED") return 4;
  if (code.startsWith("FEATURE_")) return 5;
  if (code === "UNKNOWN_COMMAND" || code === "UNKNOWN_OPTION" || code === "OPTION_NOT_EFFECTIVE") return 2;
  if (code === "INPUT_NOT_FOUND" || code === "INPUT_PARSE_ERROR") return 3;
  return 3;
}

function applyAgentBudget(context: RuntimeContext, envelope: Envelope): Envelope {
  if (!context.agent || !context.jsonBudgetBytes) return envelope;
  const bytes = Buffer.byteLength(JSON.stringify(envelope, null, 2), "utf8");
  if (bytes <= context.jsonBudgetBytes) return envelope;

  const result = envelope.result;
  const resultSchema = result && typeof result === "object" ? (result as Record<string, unknown>).schema : undefined;
  const artifacts = envelope.artifacts;
  const counts = result && typeof result === "object" ? summarizeCounts(result as Record<string, unknown>) : {};
  const compact: Envelope = {
    ...envelope,
    objectiveOk: envelope.objectiveOk,
    readiness: envelope.readiness,
    partial: envelope.partial,
    failureClass: envelope.failureClass,
    truncated: true,
    result: {
      schema: "officegen.progressive-disclosure@1.2",
      status: "truncated",
      truncated: true,
      resultSchema,
      partialSummary: {
        resultSchema,
        counts,
        artifactCount: artifacts.length,
        objectiveOk: envelope.objectiveOk,
        readiness: envelope.readiness,
        partial: envelope.partial,
        responseTruncated: true,
        responsePartialReason: "json_budget_truncated",
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
    artifacts,
    nextActions: recommendedNarrowCommands(envelope.command, context)
  };

  return compact;
}

function summarizeCounts(value: Record<string, unknown>): Record<string, number> {
  const counts: Record<string, number> = {};
  for (const [key, nested] of Object.entries(value)) {
    if (Array.isArray(nested)) counts[key] = nested.length;
    if (nested && typeof nested === "object" && !Array.isArray(nested)) {
      const objectCounts = Object.entries(nested as Record<string, unknown>)
        .filter(([, item]) => Array.isArray(item))
        .map(([nestedKey, item]) => [`${key}.${nestedKey}`, (item as unknown[]).length] as const);
      for (const [nestedKey, count] of objectCounts) counts[nestedKey] = count;
    }
  }
  return counts;
}

function errorSuggestedCommands(context: RuntimeContext, error: CliErrorPayload): string[] {
  const agent = context.agent ? " --agent" : "";
  const command = error.command ?? context.argv.slice(2).join(" ");
  const suggestions: string[] = [];
  if (error.code === "UNKNOWN_COMMAND") {
    const attempted = String(error.command ?? "");
    const alias = commandAliasSuggestion(attempted, context);
    if (alias) suggestions.push(alias);
  } else if (error.code === "INPUT_PARSE_ERROR") {
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
  } else if (error.code === "SECURITY_ABSOLUTE_OUT_DENIED") {
    suggestions.push("officegen <command> <input> --out .officegen/outputs/output.ext --json");
    suggestions.push("officegen config show --json");
  } else if (error.code === "UNSUPPORTED_FORMAT") {
    suggestions.push("officegen inspect <input.pptx|input.docx|input.xlsx|input.pdf> --depth summary --agent --json");
    suggestions.push("officegen asset inspect <image-or-media-file> --json");
  } else if (error.code === "SCHEMA_INVALID") {
    suggestions.push("officegen schema get officegen.ir.document@1.2 --json");
    suggestions.push("officegen scaffold --kind pptx --title \"Draft\" --out draft.ir.json --json");
  } else if (error.code === "EXPECTED_ARTIFACT_MISSING") {
    suggestions.push(`officegen run <workflow.json> --manifest .officegen/runs/run-manifest.json --log-jsonl .officegen/runs/events.jsonl${agent} --json`);
    suggestions.push(`officegen inspect <expected-output> --depth summary${agent} --json`);
  } else if (error.code === "RUN_STEP_FAILED" && command.startsWith("benchmark")) {
    suggestions.push("npm run benchmark:fetch");
    suggestions.push(`officegen benchmark run --manifest benchmarks/office-corpus/manifest.json${agent} --json --strict-json`);
    suggestions.push(`officegen benchmark compare <before.json> <after.json>${agent} --json --strict-json`);
  } else if (error.code === "REPAIR_NO_SAFE_OPS") {
    suggestions.push(`officegen diagnose <input> --report-out .officegen/runs/diagnose.json${agent} --json`);
    suggestions.push(`officegen edit <input> --ops suggested-ops.json --dry-run --resolve-selectors${agent} --json`);
  } else if (error.code === "DESIGN_NOT_INITIALIZED") {
    const detailSuggestions = Array.isArray(error.details?.nextSuggestedCommands) ? error.details.nextSuggestedCommands.map(String) : [];
    suggestions.push(...detailSuggestions);
    suggestions.push(`officegen design init --name <name>${agent} --json`);
    suggestions.push(`officegen design capture <source.pptx> --name <name>${agent} --json`);
  } else if (error.code === "TIMEOUT") {
    suggestions.push(`officegen ${command || "<command>"} --timeout-ms 120000${agent} --json`);
  }
  return [...new Set([...suggestions, ...nextSuggestedCommands(context)])];
}

function recommendedNarrowCommands(command: string, context: RuntimeContext): string[] {
  const agent = context.agent ? " --agent --strict-json" : "";
  if (command.startsWith("inspect")) {
    return [
      `officegen inspect <input> --depth summary --object-map-limit 50${agent} --json`,
      `officegen inspect <deck.pptx> --slides 1-5 --depth summary --object-map-limit 50${agent} --json`,
      `officegen inspect <workbook.xlsx> --sheet Sheet1 --range A1:K40${agent} --json`,
      `officegen inspect <file> --fields "schema,trusted,objectMap"${agent} --json`
    ];
  }
  if (command.startsWith("view")) {
    return [
      `officegen view <input> --max-pages 3 --object-map-limit 50 --out .officegen/runs/view${agent} --json`,
      `officegen view <pdf> --max-pages 3 --out .officegen/runs/pdf-view${agent} --json`
    ];
  }
  if (command.startsWith("verify")) {
    return [`officegen verify <input> --timeout-ms 60000${agent} --json`];
  }
  return [`officegen ${command || "<command>"} --json-budget-bytes ${Math.max((context.jsonBudgetBytes ?? 32768) * 2, 65536)}${agent} --json`];
}

function commandAliasSuggestion(command: string, context: RuntimeContext): string | undefined {
  const agent = context.agent ? " --agent" : "";
  if (command.startsWith("schema fetch")) {
    const schemaId = command.split(/\s+/).slice(2).join(" ") || "officegen.ir.document@1.2";
    return `officegen schema get ${schemaId}${agent} --json`;
  }
  if (command === "schema") return `officegen schema list${agent} --json`;
  return undefined;
}
