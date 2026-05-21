import path from "node:path";

import {
  OptionalCapabilities,
  OptionalContext,
  featureRoot,
  normalizeCapabilities,
  nowIso,
  requireFeature,
  slugify,
  untrustedContentWarning,
  writeJsonFile
} from "./common.js";

export interface AgentAdapterRecord {
  name: string;
  capabilities: OptionalCapabilities;
  adapterTextPath: string;
  generatedAt: string;
  warning: string;
}

export interface AgentInstallOptions extends OptionalContext {
  name: string;
  instructions?: string;
}

export type AgentRefreshOptions = AgentInstallOptions;

export async function installAgentAdapter(options: AgentInstallOptions): Promise<AgentAdapterRecord> {
  return writeAgentAdapter(options, "install");
}

export async function refreshAgentAdapter(options: AgentRefreshOptions): Promise<AgentAdapterRecord> {
  return writeAgentAdapter(options, "refresh");
}

export function generateAgentAdapterText(options: {
  name: string;
  capabilities?: OptionalCapabilities;
  instructions?: string;
  generatedAt?: string;
}): string {
  const capabilities = normalizeCapabilities(options.capabilities);
  const generatedAt = options.generatedAt ?? nowIso();
  const instructions = options.instructions?.trim() || "Use Officegen CLI optional modules only when the matching feature gate is present.";

  return [
    "---",
    `name: ${options.name}`,
    `generatedAt: ${generatedAt}`,
    `capabilitiesHash: ${capabilities.capabilitiesHash}`,
    `features: ${capabilities.features.join(",")}`,
    "---",
    "",
    `# Officegen CLI adapter: ${options.name}`,
    "",
    untrustedContentWarning,
    "",
    "## Capabilities",
    "",
    ...capabilities.features.map((feature) => `- ${feature}`),
    "",
    "## Instructions",
    "",
    instructions,
    ""
  ].join("\n");
}

async function writeAgentAdapter(
  options: AgentInstallOptions,
  operation: "install" | "refresh"
): Promise<AgentAdapterRecord> {
  const capabilities = requireFeature(options, "agent", `agent ${operation}`);
  const generatedAt = nowIso();
  const dir = featureRoot(options, "agent");
  const name = slugify(options.name);
  const adapterTextPath = path.join(dir, `${name}.adapter.md`);
  const recordPath = path.join(dir, `${name}.adapter.json`);
  const adapterText = generateAgentAdapterText({
    name,
    capabilities,
    instructions: options.instructions,
    generatedAt
  });

  await writeJsonFile(recordPath, {
    name,
    capabilities,
    adapterTextPath,
    generatedAt,
    warning: untrustedContentWarning
  });
  await writeTextFile(adapterTextPath, adapterText);

  return {
    name,
    capabilities,
    adapterTextPath,
    generatedAt,
    warning: untrustedContentWarning
  };
}

async function writeTextFile(filePath: string, text: string): Promise<void> {
  const { mkdir, writeFile } = await import("node:fs/promises");
  await mkdir(path.dirname(filePath), { recursive: true });
  await writeFile(filePath, text, "utf8");
}
