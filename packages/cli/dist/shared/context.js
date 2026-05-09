import { computeCapabilitiesHash, loadConfig } from "@officegen/core";
import { hasFlag, optionValue, secondCommandToken } from "./argv.js";
import { COMMAND_METADATA } from "./metadata.js";
export async function createRuntimeContext(argv, cwd, env) {
    const config = await loadConfig({
        cwd,
        overrides: profileOverride(env)
    });
    const registry = buildActiveRegistry(config);
    const agent = hasFlag(argv, "--agent");
    const json = hasFlag(argv, "--json");
    const capabilitiesHash = computeCapabilitiesHash(config);
    const explicitJsonBudgetBytes = parsePositiveInt(optionValue(argv, "--json-budget-bytes") ?? env.OFFICEGEN_JSON_BUDGET_BYTES);
    const jsonBudgetBytes = explicitJsonBudgetBytes ?? (agent ? config.agent.defaultJsonBudgetBytes : undefined);
    const expectedCapabilitiesHash = optionValue(argv, "--capabilities-hash") ?? env.OFFICEGEN_CAPABILITIES_HASH;
    return {
        argv,
        cwd,
        agent,
        json,
        config,
        registry,
        capabilitiesHash,
        jsonBudgetBytes,
        staleCapabilitiesWarning: staleWarning(expectedCapabilitiesHash, capabilitiesHash)
    };
}
export function buildActiveRegistry(config) {
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
export function gateTopLevelCommand(command, context) {
    const entry = context.registry.find((candidate) => candidate.commandGroup === command);
    if (!entry) {
        return {
            code: "UNKNOWN_COMMAND",
            command,
            message: `Unknown command: ${command}`
        };
    }
    if (!entry.enabled) {
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
    if (second && entry.commands.length > 1) {
        const allowed = new Set(entry.commands.map((registered) => registered.split(" ")[1]).filter(Boolean));
        if (allowed.size > 0 && !allowed.has(second)) {
            return {
                code: "UNKNOWN_COMMAND",
                command: `${command} ${second}`,
                message: `Unknown command: ${command} ${second}`
            };
        }
    }
    return undefined;
}
const NO_POSITIONAL_LEAF_COMMANDS = new Set([
    "capabilities",
    "doctor",
    "scaffold"
]);
export function availableCommands(context) {
    return context.registry
        .filter((entry) => entry.enabled)
        .filter((entry) => entry.visibleInHelp)
        .filter((entry) => !context.agent || entry.visibleToAgents)
        .map((entry) => entry.commandGroup);
}
export function nextSuggestedCommands(context) {
    const available = new Set(availableCommands(context));
    const suggestions = [
        context.agent ? "officegen capabilities --agent --json" : "officegen capabilities --json",
        context.agent ? "officegen help --agent --json" : "officegen help --json",
        context.agent ? "officegen schema list --agent --json" : "officegen schema list --json"
    ];
    return suggestions.filter((command) => available.has(command.split(" ")[1] ?? ""));
}
function profileOverride(env) {
    const profile = env.OFFICEGEN_PROFILE;
    if (profile === "substrate" || profile === "authoring" || profile === "enterprise") {
        return { profile };
    }
    return undefined;
}
function parsePositiveInt(value) {
    if (!value)
        return undefined;
    const parsed = Number(value);
    return Number.isInteger(parsed) && parsed > 0 ? parsed : undefined;
}
function staleWarning(expected, actual) {
    if (!expected || expected === actual)
        return undefined;
    return {
        code: "CAPABILITIES_STALE",
        severity: "warning",
        message: "Agent adapter capabilities hash differs from the active CLI capabilities hash.",
        expected,
        actual
    };
}
//# sourceMappingURL=context.js.map