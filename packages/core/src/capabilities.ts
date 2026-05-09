import { createHash } from "node:crypto";
import { FEATURE_NAMES } from "./config.js";
import { SCHEMA_REGISTRY_VERSION, OFFICEGEN_CLI_VERSION } from "./types.js";
import type { CapabilitiesDocument, CapabilityFeature, FeatureName, OfficegenConfig } from "./types.js";

const commandMap: Record<FeatureName, string[]> = {
  capabilities: ["capabilities"],
  help: ["help", "help workflow", "help error"],
  config: ["config show", "config set"],
  doctor: ["doctor"],
  inspect: ["inspect"],
  view: ["view"],
  edit: ["edit"],
  render: ["render"],
  scaffold: ["scaffold"],
  export: ["export"],
  validate: ["validate"],
  diagnose: ["diagnose"],
  repair: ["repair"],
  diff: ["diff"],
  run: ["run"],
  asset: ["asset add", "asset inspect", "asset extract", "asset replace"],
  chart: ["chart render"],
  diagram: ["diagram render"],
  schema: ["schema list", "schema get", "schema validate", "schema migrate"],
  errors: ["errors list", "errors inspect"],
  template: ["template list", "template inspect", "template candidates", "template create", "template apply-map", "template validate", "template fill"],
  design: ["design list", "design inspect", "design init", "design edit", "design update", "design validate", "design capture", "design apply"],
  layout: ["layout apply"],
  agent: ["agent install", "agent refresh"],
  mcp: ["mcp serve"],
  renderer: ["renderer list", "renderer inspect", "renderer trust"],
  plugin: ["plugin list", "plugin inspect", "plugin install", "plugin trust"]
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
    commandMap,
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
    nextSuggestedCommands: visibleCommands.some((command) => command === "schema" || command.startsWith("schema "))
      ? ["officegen help workflow edit-existing --agent --json", "officegen schema list --agent --json"]
      : ["officegen capabilities --agent --json"]
  };
}

export function isFeatureAvailable(config: OfficegenConfig, feature: FeatureName, agent = false): boolean {
  const visibility = config.features[feature];
  if (!visibility.enabled) return false;
  return agent ? visibility.visibleToAgents : visibility.visibleInHelp;
}
