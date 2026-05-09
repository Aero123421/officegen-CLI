import { readFile } from "node:fs/promises";
import { homedir } from "node:os";
import path from "node:path";
const featureNames = [
    "capabilities",
    "help",
    "config",
    "doctor",
    "inspect",
    "view",
    "edit",
    "render",
    "scaffold",
    "export",
    "validate",
    "verify",
    "diagnose",
    "repair",
    "diff",
    "run",
    "asset",
    "chart",
    "diagram",
    "schema",
    "errors",
    "template",
    "design",
    "layout",
    "agent",
    "mcp",
    "renderer",
    "plugin"
];
export const FEATURE_NAMES = [...featureNames];
function visibility(enabled) {
    return {
        enabled,
        visibleInHelp: enabled,
        visibleToAgents: enabled
    };
}
function featuresFromBooleans(values) {
    return Object.fromEntries(featureNames.map((name) => [name, visibility(values[name] ?? false)]));
}
const substrateBooleans = {
    capabilities: true,
    help: true,
    config: true,
    doctor: true,
    inspect: true,
    view: true,
    edit: true,
    render: true,
    scaffold: true,
    export: true,
    validate: true,
    verify: true,
    diagnose: true,
    repair: true,
    diff: true,
    run: true,
    asset: true,
    chart: true,
    diagram: true,
    schema: true,
    errors: true,
    template: true,
    design: true,
    layout: true,
    agent: true,
    mcp: false,
    renderer: false,
    plugin: false
};
const authoringBooleans = {
    ...substrateBooleans,
    template: true,
    design: true,
    layout: true,
    mcp: true
};
const enterpriseBooleans = {
    ...authoringBooleans,
    renderer: true,
    plugin: true
};
function baseConfig(profile, features) {
    return {
        version: "1.2",
        profile,
        paths: {
            projectRoot: ".",
            projectConfigDir: ".officegen",
            userConfigDir: "~/.officegen",
            defaultOutputDir: ".officegen/outputs",
            defaultRunsDir: ".officegen/runs"
        },
        features,
        security: {
            network: "deny",
            externalProcess: profile === "enterprise" ? "allow" : "deny",
            plugins: profile === "enterprise" ? "enabled" : "disabled",
            renderers: profile === "enterprise" ? "enabled" : "disabled",
            allowOverwrite: false,
            outOfProjectPolicy: "deny",
            allowAbsoluteInputPaths: true,
            allowAbsoluteOutputPaths: false,
            redactAbsolutePathsInJson: true,
            redactSecretsInJson: true,
            followSymlinks: false,
            allowHardlinks: false,
            trustedRoots: [".", ".officegen", "~/.officegen"],
            untrustedInput: {
                maxInputFileBytes: 104857600,
                maxZipEntries: 20000,
                maxZipExpandedBytes: 524288000,
                maxSingleXmlPartBytes: 52428800,
                maxRelationships: 50000,
                maxNestedZipDepth: 1,
                xmlExternalEntities: "deny",
                externalRelationships: "warn-and-drop-by-default",
                macros: "warn-and-preserve-only-if-requested",
                embeddedObjects: "warn",
                externalHyperlinks: "warn"
            }
        },
        agent: {
            defaultJsonBudgetBytes: 32768,
            inspectDefaultDepth: "summary",
            largeOutputMode: "path-only",
            requireCapabilitiesCheck: true
        }
    };
}
export const BUILTIN_PROFILE_CONFIGS = {
    substrate: baseConfig("substrate", featuresFromBooleans(substrateBooleans)),
    authoring: baseConfig("authoring", featuresFromBooleans(authoringBooleans)),
    enterprise: baseConfig("enterprise", featuresFromBooleans(enterpriseBooleans))
};
export function getBuiltinConfig(profile = "substrate") {
    return structuredClone(BUILTIN_PROFILE_CONFIGS[profile]);
}
function isRecord(value) {
    return typeof value === "object" && value !== null && !Array.isArray(value);
}
function deepMerge(base, override) {
    if (!isRecord(base) || !isRecord(override)) {
        return override === undefined ? base : override;
    }
    const out = { ...base };
    for (const [key, value] of Object.entries(override)) {
        if (value === undefined)
            continue;
        out[key] = isRecord(value) && isRecord(out[key]) ? deepMerge(out[key], value) : value;
    }
    return out;
}
export function normalizeFeatureConfig(config) {
    const normalized = structuredClone(config);
    for (const name of featureNames) {
        const feature = normalized.features[name];
        normalized.features[name] = {
            enabled: Boolean(feature?.enabled),
            visibleInHelp: Boolean(feature?.visibleInHelp && feature?.enabled),
            visibleToAgents: Boolean(feature?.visibleToAgents && feature?.enabled)
        };
    }
    normalized.security.plugins = normalized.features.plugin.enabled ? "enabled" : "disabled";
    normalized.security.renderers = normalized.features.renderer.enabled ? "enabled" : "disabled";
    return normalized;
}
export function mergeConfig(base, override) {
    if (!override)
        return normalizeFeatureConfig(base);
    const profile = override.profile ?? base.profile;
    const profileBase = profile === base.profile ? base : getBuiltinConfig(profile);
    return normalizeFeatureConfig(deepMerge(profileBase, override));
}
export function expandHome(inputPath) {
    if (inputPath === "~")
        return homedir();
    if (inputPath.startsWith("~/") || inputPath.startsWith("~\\")) {
        return path.join(homedir(), inputPath.slice(2));
    }
    return inputPath;
}
async function readJsonIfExists(filePath) {
    try {
        const raw = await readFile(filePath, "utf8");
        return JSON.parse(raw);
    }
    catch (error) {
        const nodeError = error;
        if (nodeError.code === "ENOENT")
            return undefined;
        throw error;
    }
}
export async function loadConfig(options = {}) {
    const cwd = options.cwd ?? process.cwd();
    const userConfigPath = options.userConfigPath ?? path.join(homedir(), ".officegen", "config.json");
    const projectConfigPath = options.projectConfigPath ?? path.join(cwd, ".officegen", "config.json");
    const userConfig = await readJsonIfExists(userConfigPath);
    const projectConfig = await readJsonIfExists(projectConfigPath);
    const profile = options.overrides?.profile ?? projectConfig?.profile ?? userConfig?.profile ?? "substrate";
    let config = getBuiltinConfig(profile);
    config = mergeConfig(config, userConfig);
    config = mergeConfig(config, projectConfig);
    config = mergeConfig(config, options.overrides);
    config.paths.projectRoot = path.resolve(cwd, expandHome(config.paths.projectRoot));
    config.paths.projectConfigDir = path.resolve(config.paths.projectRoot, expandHome(config.paths.projectConfigDir));
    config.paths.userConfigDir = path.resolve(expandHome(config.paths.userConfigDir));
    return config;
}
//# sourceMappingURL=config.js.map