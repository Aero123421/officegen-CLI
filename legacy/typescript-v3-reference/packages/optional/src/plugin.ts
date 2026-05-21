import path from "node:path";

import {
  OptionalContext,
  featureRoot,
  fileExists,
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
  installedSource?: {
    kind: "source-file" | "manifest";
    path?: string;
    sha256: string;
  };
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
  trust?: boolean | string;
  trustSha256?: string;
  expectedSha256?: string;
}

export interface PluginInspectOptions extends OptionalContext {
  id: string;
}

export async function installPlugin(options: PluginInstallOptions): Promise<PluginManifest> {
  requireFeature(options, "plugin", "plugin install");
  const id = slugify(options.manifest.id);
  const sourcePath = options.sourcePath ? path.resolve(options.cwd ?? process.cwd(), options.sourcePath) : undefined;
  const sourceHash = sourcePath ? await hashFile(sourcePath) : undefined;
  const actualSha256 = sourceHash ?? hashPluginManifest(options.manifest, id);
  const manifest: PluginManifest = {
    ...options.manifest,
    id,
    sha256: actualSha256,
    installedSource: {
      kind: sourcePath ? "source-file" : "manifest",
      path: sourcePath,
      sha256: actualSha256
    },
    installedAt: nowIso()
  };

  const expectedSha256 = parseExpectedTrustSha256(options);
  if (expectedSha256) {
    if (actualSha256 !== expectedSha256) {
      throw new Error(`Plugin trust hash mismatch: expected sha256:${expectedSha256}, got sha256:${actualSha256}.`);
    }
  } else if (options.trust === true) {
    throw new Error("Plugin install trust requires an explicit sha256:<hash> value.");
  }

  await writeJsonFile(pluginPath(options, id), manifest);
  if (expectedSha256) {
    await pinPluginTrust(options, manifest, actualSha256);
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
  const actualSha256 = await resolveInstalledPluginSha256(manifest);
  const trusted = trustStore.pins.some(
    (pin) => pin.id === manifest.id && pin.version === manifest.version && pin.sha256 === actualSha256
  );
  return { ...manifest, sha256: actualSha256, trusted };
}

export async function readPluginTrustStore(options: OptionalContext = {}): Promise<TrustStore> {
  requireFeature(options, "plugin", "plugin trust-store");
  const storePath = pluginTrustStorePath(options);
  if (!(await fileExists(storePath))) {
    return { kind: "officegen.plugin.trust-store", pins: [], updatedAt: nowIso() };
  }

  let trustStore: TrustStore;
  try {
    trustStore = await readJsonFile<TrustStore>(storePath);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    throw new Error(`Failed to read plugin trust store: ${storePath}: ${message}`);
  }
  if (trustStore.kind !== "officegen.plugin.trust-store" || !Array.isArray(trustStore.pins)) {
    throw new Error(`Invalid plugin trust store: ${storePath}`);
  }
  for (const pin of trustStore.pins) {
    if (!pin.id || !pin.version || !isSha256Hex(pin.sha256) || !pin.trustedAt) {
      throw new Error(`Invalid plugin trust store pin: ${storePath}`);
    }
  }
  return trustStore;
}

async function pinPluginTrust(options: OptionalContext, manifest: PluginManifest, actualSha256: string): Promise<void> {
  const trustStore = await readPluginTrustStore(options);
  const pin: TrustPin = {
    id: manifest.id,
    version: manifest.version,
    sha256: actualSha256,
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

function parseExpectedTrustSha256(options: PluginInstallOptions): string | undefined {
  const trustValue =
    typeof options.trust === "string"
      ? options.trust
      : options.trustSha256 ?? options.expectedSha256;
  if (options.trust === false || trustValue === undefined) return undefined;

  const normalized = trustValue.startsWith("sha256:") ? trustValue.slice("sha256:".length) : trustValue;
  if (!isSha256Hex(normalized)) {
    throw new Error("Plugin trust requires sha256:<64 hex characters>.");
  }
  return normalized.toLowerCase();
}

function hashPluginManifest(manifest: PluginManifest, id = slugify(manifest.id)): string {
  const { sha256, installedAt, installedSource, ...hashable } = manifest;
  void sha256;
  void installedAt;
  void installedSource;
  return sha256Json({ ...hashable, id });
}

async function resolveInstalledPluginSha256(manifest: PluginManifest): Promise<string> {
  if (manifest.installedSource?.kind === "source-file" && manifest.installedSource.path) {
    return hashFile(manifest.installedSource.path);
  }
  return hashPluginManifest(manifest, manifest.id);
}

function isSha256Hex(value: string | undefined): value is string {
  return typeof value === "string" && /^[a-fA-F0-9]{64}$/.test(value);
}

function pluginPath(context: OptionalContext, id: string): string {
  return path.join(featureRoot(context, "plugin"), `${slugify(id)}.json`);
}

function pluginTrustStorePath(context: OptionalContext): string {
  return path.join(featureRoot(context, "plugin"), "trust-store.json");
}
