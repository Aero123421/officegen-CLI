import path from "node:path";
import { nativeRendererDoctor } from "@officegen/formats";

import {
  OptionalContext,
  featureRoot,
  listJsonFiles,
  nowIso,
  readJsonFile,
  requireFeature,
  sha256Json,
  slugify,
  writeJsonFile
} from "./common.js";

export interface RendererManifest {
  id: string;
  name: string;
  version: string;
  formats: string[];
  capabilities?: string[];
  description?: string;
  sha256?: string;
}

export interface RendererTrustPin {
  id: string;
  version: string;
  sha256: string;
  trustedAt: string;
}

export interface RendererTrustStore {
  kind: "officegen.renderer.trust-store";
  pins: RendererTrustPin[];
  updatedAt: string;
}

export interface RendererInspectOptions extends OptionalContext {
  id: string;
}

export interface RendererTrustOptions extends RendererInspectOptions {
  sha256?: string;
}

const builtinRenderers: RendererManifest[] = [
  {
    id: "json-plan",
    name: "JSON Plan Renderer",
    version: "1.0.0",
    formats: ["json"],
    capabilities: ["plan:read"],
    description: "Writes renderer plans without mutating Office files."
  },
  {
    id: "html-preview",
    name: "HTML Preview Renderer",
    version: "1.0.0",
    formats: ["html"],
    capabilities: ["preview:write"],
    description: "Produces preview artifacts that can be handed to @officegen/formats."
  }
].map((renderer) => ({ ...renderer, sha256: sha256Json(renderer) }));

export async function listRenderers(options: OptionalContext = {}): Promise<RendererManifest[]> {
  requireFeature(options, "renderer", "renderer list");
  const files = await listJsonFiles(featureRoot(options, "renderer"));
  const custom = await Promise.all(
    files
      .filter((file) => !file.endsWith("trust-store.json"))
      .map((file) => readJsonFile<RendererManifest>(file))
  );
  return [...builtinRenderers, ...custom].sort((left, right) => left.id.localeCompare(right.id));
}

export async function inspectRenderer(options: RendererInspectOptions): Promise<RendererManifest & { trusted: boolean }> {
  requireFeature(options, "renderer", "renderer inspect");
  if (slugify(options.id) === "doctor") {
    return await doctorRenderer(options) as unknown as RendererManifest & { trusted: boolean };
  }
  const renderer = (await listRenderers(options)).find((candidate) => candidate.id === slugify(options.id));
  if (!renderer) {
    throw new Error(`Renderer not found: ${options.id}`);
  }

  const trustStore = await readRendererTrustStore(options);
  const trusted = trustStore.pins.some(
    (pin) => pin.id === renderer.id && pin.version === renderer.version && pin.sha256 === renderer.sha256
  );
  return { ...renderer, trusted };
}

export async function doctorRenderer(options: OptionalContext = {}): Promise<Record<string, unknown>> {
  requireFeature(options, "renderer", "renderer doctor");
  const doctor = await nativeRendererDoctor(undefined);
  return {
    ...doctor,
    trusted: true,
    guidance: [
      "Use OFFICEGEN_PROFILE=enterprise or an equivalent config enabling externalProcess/renderers for native export/verify.",
      "Windows Office COM is preferred when available; LibreOffice headless is the portable fallback."
    ]
  };
}

export async function trustRenderer(options: RendererTrustOptions): Promise<RendererTrustStore> {
  requireFeature(options, "renderer", "renderer trust");
  const renderer = await inspectRenderer(options);
  const sha256 = options.sha256 ?? renderer.sha256 ?? sha256Json(renderer);
  const trustStore = await readRendererTrustStore(options);
  const pins = trustStore.pins.filter(
    (pin) => !(pin.id === renderer.id && pin.version === renderer.version)
  );
  pins.push({
    id: renderer.id,
    version: renderer.version,
    sha256,
    trustedAt: nowIso()
  });

  const updated: RendererTrustStore = {
    kind: "officegen.renderer.trust-store",
    pins: pins.sort((left, right) => left.id.localeCompare(right.id) || left.version.localeCompare(right.version)),
    updatedAt: nowIso()
  };
  await writeJsonFile(rendererTrustStorePath(options), updated);
  return updated;
}

export async function registerRenderer(
  options: OptionalContext & { renderer: RendererManifest }
): Promise<RendererManifest> {
  requireFeature(options, "renderer", "renderer register");
  const renderer = {
    ...options.renderer,
    id: slugify(options.renderer.id),
    sha256: options.renderer.sha256 ?? sha256Json(options.renderer)
  };
  await writeJsonFile(path.join(featureRoot(options, "renderer"), `${renderer.id}.json`), renderer);
  return renderer;
}

export async function readRendererTrustStore(options: OptionalContext = {}): Promise<RendererTrustStore> {
  requireFeature(options, "renderer", "renderer trust-store");
  try {
    return await readJsonFile<RendererTrustStore>(rendererTrustStorePath(options));
  } catch {
    return { kind: "officegen.renderer.trust-store", pins: [], updatedAt: nowIso() };
  }
}

function rendererTrustStorePath(context: OptionalContext): string {
  return path.join(featureRoot(context, "renderer"), "trust-store.json");
}
