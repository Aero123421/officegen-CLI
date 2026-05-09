import { readFile } from "node:fs/promises";
import { homedir } from "node:os";
import path from "node:path";
import type {
  ConfigLoadOptions,
  FeatureMap,
  FeatureName,
  FeatureVisibility,
  OfficegenConfig,
  OfficegenConfigInput,
  OfficegenProfile
} from "./types.js";

const featureNames: FeatureName[] = [
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
  "diagnose",
  "repair",
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

function visibility(enabled: boolean): FeatureVisibility {
  return {
    enabled,
    visibleInHelp: enabled,
    visibleToAgents: enabled
  };
}

function featuresFromBooleans(values: Partial<Record<FeatureName, boolean>>): FeatureMap {
  return Object.fromEntries(featureNames.map((name) => [name, visibility(values[name] ?? false)])) as FeatureMap;
}

const substrateBooleans: Record<FeatureName, boolean> = {
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
  diagnose: true,
  repair: true,
  run: true,
  asset: true,
  chart: true,
  diagram: true,
  schema: true,
  errors: true,
  template: false,
  design: false,
  layout: false,
  agent: true,
  mcp: false,
  renderer: false,
  plugin: false
};

const authoringBooleans: Record<FeatureName, boolean> = {
  ...substrateBooleans,
  template: true,
  design: true,
  layout: true,
  mcp: true
};

const enterpriseBooleans: Record<FeatureName, boolean> = {
  ...authoringBooleans,
  renderer: true,
  plugin: true
};

function baseConfig(profile: OfficegenProfile, features: FeatureMap): OfficegenConfig {
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
      externalProcess: "deny",
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
      defaultJsonBudgetBytes: 8192,
      inspectDefaultDepth: "summary",
      largeOutputMode: "path-only",
      requireCapabilitiesCheck: true
    }
  };
}

export const BUILTIN_PROFILE_CONFIGS: Record<OfficegenProfile, OfficegenConfig> = {
  substrate: baseConfig("substrate", featuresFromBooleans(substrateBooleans)),
  authoring: baseConfig("authoring", featuresFromBooleans(authoringBooleans)),
  enterprise: baseConfig("enterprise", featuresFromBooleans(enterpriseBooleans))
};

export function getBuiltinConfig(profile: OfficegenProfile = "substrate"): OfficegenConfig {
  return structuredClone(BUILTIN_PROFILE_CONFIGS[profile]);
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function deepMerge<T>(base: T, override: unknown): T {
  if (!isRecord(base) || !isRecord(override)) {
    return override === undefined ? base : (override as T);
  }
  const out: Record<string, unknown> = { ...base };
  for (const [key, value] of Object.entries(override)) {
    if (value === undefined) continue;
    out[key] = isRecord(value) && isRecord(out[key]) ? deepMerge(out[key], value) : value;
  }
  return out as T;
}

export function normalizeFeatureConfig(config: OfficegenConfig): OfficegenConfig {
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

export function mergeConfig(base: OfficegenConfig, override?: OfficegenConfigInput): OfficegenConfig {
  if (!override) return normalizeFeatureConfig(base);
  const profile = override.profile ?? base.profile;
  const profileBase = profile === base.profile ? base : getBuiltinConfig(profile);
  return normalizeFeatureConfig(deepMerge(profileBase, override));
}

export function expandHome(inputPath: string): string {
  if (inputPath === "~") return homedir();
  if (inputPath.startsWith("~/") || inputPath.startsWith("~\\")) {
    return path.join(homedir(), inputPath.slice(2));
  }
  return inputPath;
}

async function readJsonIfExists<T>(filePath: string): Promise<T | undefined> {
  try {
    const raw = await readFile(filePath, "utf8");
    return JSON.parse(raw) as T;
  } catch (error) {
    const nodeError = error as NodeJS.ErrnoException;
    if (nodeError.code === "ENOENT") return undefined;
    throw error;
  }
}

export async function loadConfig(options: ConfigLoadOptions = {}): Promise<OfficegenConfig> {
  const cwd = options.cwd ?? process.cwd();
  const userConfigPath = options.userConfigPath ?? path.join(homedir(), ".officegen", "config.json");
  const projectConfigPath = options.projectConfigPath ?? path.join(cwd, ".officegen", "config.json");
  const userConfig = await readJsonIfExists<OfficegenConfigInput>(userConfigPath);
  const projectConfig = await readJsonIfExists<OfficegenConfigInput>(projectConfigPath);
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
