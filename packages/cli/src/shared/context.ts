import { computeCapabilitiesHash, loadConfig, type OfficegenConfigInput } from "@officegen/core";
import { hasFlag, optionValue, secondCommandToken } from "./argv.js";
import { COMMAND_METADATA } from "./metadata.js";
import type { ActiveCapability, CliErrorPayload, RuntimeContext } from "./types.js";

export async function createRuntimeContext(
  argv: string[],
  cwd: string,
  env: NodeJS.ProcessEnv
): Promise<RuntimeContext> {
  const config = await loadConfig({
    cwd,
    overrides: profileOverride(env)
  });
  const registry = buildActiveRegistry(config);
  const agent = hasFlag(argv, "--agent");
  const strictJson = hasFlag(argv, "--strict-json") || env.OFFICEGEN_STRICT_JSON === "1";
  const json = hasFlag(argv, "--json") || strictJson;
  const capabilitiesHash = computeCapabilitiesHash(config);
  const explicitJsonBudgetBytes = parsePositiveInt(optionValue(argv, "--json-budget-bytes") ?? env.OFFICEGEN_JSON_BUDGET_BYTES);
  const jsonBudgetBytes = explicitJsonBudgetBytes ?? (agent ? config.agent.defaultJsonBudgetBytes : undefined);
  const expectedCapabilitiesHash = optionValue(argv, "--capabilities-hash") ?? env.OFFICEGEN_CAPABILITIES_HASH;

  return {
    argv,
    cwd,
    agent,
    json,
    strictJson,
    config,
    registry,
    capabilitiesHash,
    jsonBudgetBytes,
    staleCapabilitiesWarning: staleWarning(expectedCapabilitiesHash, capabilitiesHash)
  };
}

export function buildActiveRegistry(config: RuntimeContext["config"]): ActiveCapability[] {
  return COMMAND_METADATA.map((entry) => {
    const configured = config.features[entry.feature];
    return {
      ...entry,
      enabled: configured.enabled,
      visibleInHelp: configured.visibleInHelp,
      visibleToAgents: configured.visibleToAgents
    };
  });
}

export function gateTopLevelCommand(command: string, context: RuntimeContext): CliErrorPayload | undefined {
  const entry = context.registry.find((candidate) => candidate.commandGroup === command);
  if (!entry) {
    const didYouMean = closestCommand(command, context);
    return {
      code: "UNKNOWN_COMMAND",
      command,
      message: `Unknown command: ${command}`,
      details: didYouMean ? {
        didYouMean,
        repairPlan: `Run ${didYouMean} with --agent --json, or use officegen help --agent --json to inspect the command surface.`
      } : undefined
    };
  }
  if (!entry.enabled) {
    if (entry.feature === "renderer" && secondCommandToken(context.argv) === "doctor") return undefined;
    return {
      code: "FEATURE_DISABLED",
      feature: entry.feature,
      command,
      message: `The ${entry.feature} feature is disabled by the active configuration.`
    };
  }
  if (context.agent && !entry.visibleToAgents) {
    return {
      code: "FEATURE_HIDDEN_FROM_AGENT",
      feature: entry.feature,
      command,
      message: `The ${entry.feature} feature is hidden from agents by the active configuration.`
    };
  }
  const second = secondCommandToken(context.argv);
  if (second && NO_POSITIONAL_LEAF_COMMANDS.has(entry.commandGroup)) {
    return {
      code: "UNKNOWN_COMMAND",
      command: `${command} ${second}`,
      message: `Unknown command: ${command} ${second}`
    };
  }
  if (entry.commandGroup === "help") return undefined;
  if (second && entry.commands.length > 1) {
    const allowed = new Set(entry.commands.map((registered) => registered.split(" ")[1]).filter(Boolean));
    if (allowed.size > 0 && !allowed.has(second)) {
      const alias = entry.commandGroup === "schema" && second === "fetch" ? "schema get" : undefined;
      return {
        code: "UNKNOWN_COMMAND",
        command: `${command} ${second}`,
        message: `Unknown command: ${command} ${second}`,
        details: alias ? {
          didYouMean: alias,
          repairPlan: `Use officegen ${alias} <schema-id> --agent --json.`
        } : undefined
      };
    }
  }
  return undefined;
}

function closestCommand(command: string, context: RuntimeContext): string | undefined {
  const commands = context.registry.flatMap((entry) => [entry.commandGroup, ...entry.commands]);
  if (command === "schemas") return "schema";
  if (command === "bench") return "benchmark";
  return commands.find((candidate) => candidate.startsWith(command) || command.startsWith(candidate));
}

const NO_POSITIONAL_LEAF_COMMANDS = new Set([
  "capabilities",
  "doctor",
  "scaffold"
]);

export function availableCommands(context: RuntimeContext): string[] {
  return [...new Set(context.registry
    .filter((entry) => entry.enabled)
    .concat(context.registry.filter((entry) => entry.feature === "renderer"))
    .filter((entry) => entry.visibleInHelp)
    .filter((entry) => !context.agent || entry.visibleToAgents)
    .map((entry) => entry.commandGroup))];
}

export function nextSuggestedCommands(context: RuntimeContext): string[] {
  const available = new Set(availableCommands(context));
  const suggestions = [
    context.agent ? "officegen capabilities --agent --json" : "officegen capabilities --json",
    context.agent ? "officegen help --agent --json" : "officegen help --json",
    context.agent ? "officegen schema list --agent --json" : "officegen schema list --json"
  ];
  return suggestions.filter((command) => available.has(command.split(" ")[1] ?? ""));
}

function profileOverride(env: NodeJS.ProcessEnv): OfficegenConfigInput | undefined {
  const profile = env.OFFICEGEN_PROFILE;
  if (profile === "substrate" || profile === "authoring" || profile === "enterprise") {
    return { profile };
  }
  return undefined;
}

function parsePositiveInt(value: string | undefined): number | undefined {
  if (!value) return undefined;
  const parsed = Number(value);
  return Number.isInteger(parsed) && parsed > 0 ? parsed : undefined;
}

function staleWarning(expected: string | undefined, actual: string): RuntimeContext["staleCapabilitiesWarning"] {
  if (!expected || expected === actual) return undefined;
  return {
    code: "CAPABILITIES_STALE",
    severity: "warning",
    message: "Agent adapter capabilities hash differs from the active CLI capabilities hash.",
    expected,
    actual
  };
}
