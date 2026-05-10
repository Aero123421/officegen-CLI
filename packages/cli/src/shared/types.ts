import type { FeatureName, JsonValue, OfficegenConfig } from "@officegen/core";

export const CLI_SPEC_VERSION = "1.2";
export const ENVELOPE_SCHEMA = "officegen.envelope@1.2";

export type Stability = "stable" | "experimental";
export type FeatureKey = FeatureName;

export interface CapabilityEntry {
  feature: FeatureKey;
  moduleId: string;
  commandGroup: string;
  description: string;
  stability: Stability;
  commands: string[];
  requires: FeatureKey[];
  security: {
    network: boolean;
    externalProcess: boolean;
  };
}

export interface ActiveCapability extends CapabilityEntry {
  enabled: boolean;
  visibleInHelp: boolean;
  visibleToAgents: boolean;
}

export interface RuntimeContext {
  argv: string[];
  cwd: string;
  agent: boolean;
  json: boolean;
  strictJson: boolean;
  config: OfficegenConfig;
  registry: ActiveCapability[];
  capabilitiesHash: string;
  jsonBudgetBytes?: number;
  staleCapabilitiesWarning?: JsonValue;
}

export interface CliErrorPayload {
  code: string;
  category?: string;
  severity?: "info" | "warning" | "error" | "critical";
  message: string;
  feature?: string;
  command?: string;
  details?: Record<string, unknown>;
}

export interface Envelope {
  schema: typeof ENVELOPE_SCHEMA;
  ok: boolean;
  command: string;
  runId: string;
  cliVersion: string;
  capabilitiesHash: string;
  pathsRedacted: boolean;
  result?: unknown;
  error?: CliErrorPayload;
  executionOk?: boolean;
  objectiveOk?: boolean;
  mutationStatus?: "changed" | "noop" | "plan_only" | "failed" | "not_applicable";
  artifactStatus?: "complete" | "missing" | "not_expected";
  readiness?: "pass" | "pass_with_environment_gap" | "warning" | "partial" | "blocked";
  partial?: boolean;
  truncated?: boolean;
  warnings: unknown[];
  diagnostics: unknown[];
  artifacts: unknown[];
  availableCommands: string[];
  nextSuggestedCommands: string[];
}

export class CliFailure extends Error {
  payload: CliErrorPayload;
  exitCode: number;

  constructor(payload: CliErrorPayload, exitCode = 1) {
    super(payload.message);
    this.payload = payload;
    this.exitCode = exitCode;
  }
}

export interface RunCliOptions {
  cwd?: string;
  stdout?: (text: string) => void;
  stderr?: (text: string) => void;
  env?: NodeJS.ProcessEnv;
  now?: Date;
}
