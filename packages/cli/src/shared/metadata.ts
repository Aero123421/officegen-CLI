import type { CapabilityEntry, FeatureKey } from "./types.js";

export const COMMAND_METADATA: CapabilityEntry[] = [
  meta("capabilities", "Show enabled features and agent-visible commands", ["capabilities"]),
  meta("help", "Show human and agent workflow help", ["help", "help workflow", "help error"]),
  meta("config", "Inspect or update effective configuration", ["config show", "config set"]),
  meta("doctor", "Check the local runtime and project setup", ["doctor"]),
  meta("schema", "List, fetch, validate, and migrate schemas", ["schema list", "schema get", "schema validate", "schema migrate"]),
  meta("errors", "List and inspect machine-readable error codes", ["errors list", "errors inspect"]),
  core("inspect", "Inspect Office/PDF files and produce trusted metadata", ["inspect"]),
  core("view", "Create SVG/HTML previews and object maps", ["view"]),
  core("edit", "Apply declarative EditOps to existing files", ["edit"]),
  core("render", "Render document IR/specs into Office/PDF files", ["render"]),
  core("scaffold", "Create valid starter IR without an LLM", ["scaffold"]),
  core("export", "Export supported formats with explicit fidelity", ["export"]),
  core("validate", "Validate schemas, structure, and quality gates", ["validate"]),
  core("diagnose", "Detect problems in generated or existing files", ["diagnose"]),
  core("repair", "Repair files or produce a repair plan", ["repair"]),
  core("run", "Execute a multi-step workflow with run artifacts", ["run"]),
  core("asset", "Inspect, extract, and replace embedded media", ["asset add", "asset inspect", "asset extract", "asset replace"]),
  core("chart", "Render safe chart SVG assets", ["chart render"]),
  core("diagram", "Render safe diagram SVG assets", ["diagram render"]),
  optional("template", "Create, inspect, map, validate, and fill templates", [
    "template list",
    "template inspect",
    "template candidates",
    "template create",
    "template apply-map",
    "template validate",
    "template fill"
  ]),
  optional("design", "Capture, inspect, validate, and apply design knowledge", [
    "design list",
    "design inspect",
    "design init",
    "design edit",
    "design update",
    "design validate",
    "design capture",
    "design apply"
  ]),
  optional("layout", "Apply layout constraints", ["layout apply"]),
  optional("agent", "Install or refresh agent adapters", ["agent install", "agent refresh"]),
  optional("mcp", "MCP server", ["mcp serve"], true),
  optional("renderer", "Manage trusted external renderers", ["renderer list", "renderer inspect", "renderer trust"], true),
  optional("plugin", "Manage trusted plugins", ["plugin list", "plugin inspect", "plugin install", "plugin trust"], true)
];

export function metadataFor(feature: FeatureKey): CapabilityEntry | undefined {
  return COMMAND_METADATA.find((entry) => entry.feature === feature);
}

function meta(feature: FeatureKey, description: string, commands: string[]): CapabilityEntry {
  return entry(feature, description, commands, false, false);
}

function core(feature: FeatureKey, description: string, commands: string[]): CapabilityEntry {
  return entry(feature, description, commands, false, false);
}

function optional(feature: FeatureKey, description: string, commands: string[], externalProcess = false): CapabilityEntry {
  return entry(feature, description, commands, false, externalProcess);
}

function entry(
  feature: FeatureKey,
  description: string,
  commands: string[],
  network: boolean,
  externalProcess: boolean
): CapabilityEntry {
  return {
    feature,
    moduleId: `officegen.core.${feature}`,
    commandGroup: feature,
    description,
    stability: "stable",
    commands,
    requires: [],
    security: {
      network,
      externalProcess
    }
  };
}
