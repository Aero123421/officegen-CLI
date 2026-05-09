import { createHash } from "node:crypto";
import { FEATURE_NAMES } from "./config.js";
import { SCHEMA_REGISTRY_VERSION, OFFICEGEN_CLI_VERSION } from "./types.js";
import type { CapabilitiesDocument, CapabilityFeature, FeatureName, OfficegenConfig } from "./types.js";

const commandMap: Record<FeatureName, string[]> = {
  inspect: ["inspect"],
  view: ["view"],
  edit: ["edit"],
  render: ["render"],
  scaffold: ["scaffold"],
  export: ["export"],
  validate: ["validate"],
  diagnose: ["diagnose"],
  repair: ["repair"],
  asset: ["asset"],
  chart: ["chart"],
  diagram: ["diagram"],
  schema: ["schema"],
  errors: ["errors"],
  template: ["template"],
  design: ["design"],
  layout: ["layout"],
  agent: ["agent"],
  mcp: ["mcp"],
  renderer: ["renderer"],
  plugin: ["plugin"]
};

export function buildFeatureRegistry(config: OfficegenConfig): CapabilityFeature[] {
  return FEATURE_NAMES.map((name) => ({
    name,
    enabled: config.features[name].enabled,
    visibleInHelp: config.features[name].visibleInHelp,
    visibleToAgents: config.features[name].visibleToAgents,
    commands: commandMap[name],
    requires: []
  }));
}

function stableStringify(value: unknown): string {
  if (value === null || typeof value !== "object") return JSON.stringify(value);
  if (Array.isArray(value)) return `[${value.map(stableStringify).join(",")}]`;
  const record = value as Record<string, unknown>;
  return `{${Object.keys(record)
    .sort()
    .map((key) => `${JSON.stringify(key)}:${stableStringify(record[key])}`)
    .join(",")}}`;
}

export function computeCapabilitiesHash(config: OfficegenConfig, cliVersion = OFFICEGEN_CLI_VERSION): string {
  const payload = {
    cliVersion,
    profile: config.profile,
    features: config.features,
    security: config.security,
    agent: config.agent,
    schemaRegistryVersion: SCHEMA_REGISTRY_VERSION
  };
  return `sha256:${createHash("sha256").update(stableStringify(payload)).digest("hex")}`;
}

export function getVisibleCommands(config: OfficegenConfig, agent = false): string[] {
  return buildFeatureRegistry(config)
    .filter((feature) => feature.enabled && (agent ? feature.visibleToAgents : feature.visibleInHelp))
    .flatMap((feature) => feature.commands);
}

export function getCapabilities(config: OfficegenConfig, options: { agent?: boolean; runInstructionsPath?: string } = {}): CapabilitiesDocument {
  const registry = buildFeatureRegistry(config);
  const agent = options.agent ?? false;
  const visibleCommands = getVisibleCommands(config, agent);
  return {
    schema: "officegen.capabilities@1.2",
    ok: true,
    profile: config.profile,
    capabilitiesHash: computeCapabilitiesHash(config),
    visibleCommands,
    hiddenFromAgents: registry
      .filter((feature) => feature.enabled && feature.visibleInHelp && !feature.visibleToAgents)
      .map((feature) => feature.name),
    disabled: registry.filter((feature) => !feature.enabled).map((feature) => feature.name),
    agentInstructionsPath: options.runInstructionsPath ?? ".officegen/runs/current/agent-instructions.md",
    jsonBudgetBytes: config.agent.defaultJsonBudgetBytes,
    nextSuggestedCommands: visibleCommands.includes("schema")
      ? ["officegen help workflow edit-existing --agent --json", "officegen schema list --agent --json"]
      : ["officegen capabilities --agent --json"]
  };
}

export function isFeatureAvailable(config: OfficegenConfig, feature: FeatureName, agent = false): boolean {
  const visibility = config.features[feature];
  if (!visibility.enabled) return false;
  return agent ? visibility.visibleToAgents : visibility.visibleInHelp;
}
