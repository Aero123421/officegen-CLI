import path from "node:path";

import {
  OptionalContext,
  featureRoot,
  hashFile,
  listJsonFiles,
  nowIso,
  readJsonFile,
  requireFeature,
  sha256Json,
  slugify,
  writeJsonFile
} from "./common.js";

export interface PluginManifest {
  id: string;
  name?: string;
  version: string;
  entry?: string;
  capabilities?: string[];
  description?: string;
  sha256?: string;
  installedAt?: string;
}

export interface TrustPin {
  id: string;
  version: string;
  sha256: string;
  trustedAt: string;
}

export interface TrustStore {
  kind: "officegen.plugin.trust-store";
  pins: TrustPin[];
  updatedAt: string;
}

export interface PluginInstallOptions extends OptionalContext {
  manifest: PluginManifest;
  sourcePath?: string;
  trust?: boolean;
}

export interface PluginInspectOptions extends OptionalContext {
  id: string;
}

export async function installPlugin(options: PluginInstallOptions): Promise<PluginManifest> {
  requireFeature(options, "plugin", "plugin install");
  const id = slugify(options.manifest.id);
  const sourceHash = options.sourcePath ? await hashFile(path.resolve(options.cwd ?? process.cwd(), options.sourcePath)) : undefined;
  const manifest: PluginManifest = {
    ...options.manifest,
    id,
    sha256: options.manifest.sha256 ?? sourceHash ?? sha256Json(options.manifest),
    installedAt: nowIso()
  };

  await writeJsonFile(pluginPath(options, id), manifest);
  if (options.trust ?? true) {
    await pinPluginTrust(options, manifest);
  }
  return manifest;
}

export async function listPlugins(options: OptionalContext = {}): Promise<PluginManifest[]> {
  requireFeature(options, "plugin", "plugin list");
  const files = await listJsonFiles(featureRoot(options, "plugin"));
  const manifests = await Promise.all(
    files
      .filter((file) => !file.endsWith("trust-store.json"))
      .map((file) => readJsonFile<PluginManifest>(file))
  );
  return manifests.sort((left, right) => left.id.localeCompare(right.id));
}

export async function inspectPlugin(options: PluginInspectOptions): Promise<PluginManifest & { trusted: boolean }> {
  requireFeature(options, "plugin", "plugin inspect");
  const manifest = await readJsonFile<PluginManifest>(pluginPath(options, options.id));
  const trustStore = await readPluginTrustStore(options);
  const trusted = trustStore.pins.some(
    (pin) => pin.id === manifest.id && pin.version === manifest.version && pin.sha256 === manifest.sha256
  );
  return { ...manifest, trusted };
}

export async function readPluginTrustStore(options: OptionalContext = {}): Promise<TrustStore> {
  requireFeature(options, "plugin", "plugin trust-store");
  try {
    return await readJsonFile<TrustStore>(pluginTrustStorePath(options));
  } catch {
    return { kind: "officegen.plugin.trust-store", pins: [], updatedAt: nowIso() };
  }
}

async function pinPluginTrust(options: OptionalContext, manifest: PluginManifest): Promise<void> {
  const trustStore = await readPluginTrustStore(options);
  const pin: TrustPin = {
    id: manifest.id,
    version: manifest.version,
    sha256: manifest.sha256 ?? sha256Json(manifest),
    trustedAt: nowIso()
  };
  const pins = trustStore.pins.filter(
    (existing) => !(existing.id === pin.id && existing.version === pin.version)
  );
  pins.push(pin);
  await writeJsonFile(pluginTrustStorePath(options), {
    ...trustStore,
    pins: pins.sort((left, right) => left.id.localeCompare(right.id) || left.version.localeCompare(right.version)),
    updatedAt: nowIso()
  });
}

function pluginPath(context: OptionalContext, id: string): string {
  return path.join(featureRoot(context, "plugin"), `${slugify(id)}.json`);
}

function pluginTrustStorePath(context: OptionalContext): string {
  return path.join(featureRoot(context, "plugin"), "trust-store.json");
}
