import { OFFICEGEN_CLI_VERSION } from "./types.js";
import { createErrorPayload } from "./errors.js";
import type {
  ErrorEnvelope,
  JsonObject,
  JsonValue,
  OfficegenErrorCode,
  OfficegenErrorPayload,
  SuccessEnvelope
} from "./types.js";

interface EnvelopeOptions {
  command?: string;
  runId?: string;
  capabilitiesHash?: string;
  pathsRedacted?: boolean;
  warnings?: JsonValue[];
  diagnostics?: JsonValue[];
  artifacts?: JsonValue[];
  nextSuggestedCommands?: string[];
  cliVersion?: string;
}

export function successEnvelope<T extends JsonValue = JsonObject>(
  result: T,
  options: EnvelopeOptions = {}
): SuccessEnvelope<T> {
  return {
    schema: "officegen.envelope@1.2",
    ok: true,
    command: options.command,
    runId: options.runId,
    cliVersion: options.cliVersion ?? OFFICEGEN_CLI_VERSION,
    capabilitiesHash: options.capabilitiesHash,
    pathsRedacted: options.pathsRedacted ?? true,
    result,
    warnings: options.warnings ?? [],
    diagnostics: options.diagnostics ?? [],
    artifacts: options.artifacts ?? [],
    nextSuggestedCommands: options.nextSuggestedCommands ?? []
  };
}

export function errorEnvelope(
  error: OfficegenErrorCode | OfficegenErrorPayload,
  options: EnvelopeOptions & { availableCommands?: string[]; errorOptions?: Partial<OfficegenErrorPayload> } = {}
): ErrorEnvelope {
  const payload = typeof error === "string" ? createErrorPayload(error, options.errorOptions) : error;
  return {
    schema: "officegen.envelope@1.2",
    ok: false,
    command: options.command,
    runId: options.runId,
    cliVersion: options.cliVersion ?? OFFICEGEN_CLI_VERSION,
    capabilitiesHash: options.capabilitiesHash,
    pathsRedacted: options.pathsRedacted ?? true,
    error: payload,
    availableCommands: options.availableCommands ?? [],
    warnings: options.warnings ?? [],
    diagnostics: options.diagnostics ?? [],
    artifacts: options.artifacts ?? [],
    nextSuggestedCommands: options.nextSuggestedCommands ?? ["officegen capabilities --agent --json"]
  };
}
