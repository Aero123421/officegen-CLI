import { OptionalContext } from "./common.js";
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
export declare function listRenderers(options?: OptionalContext): Promise<RendererManifest[]>;
export declare function inspectRenderer(options: RendererInspectOptions): Promise<RendererManifest & {
    trusted: boolean;
}>;
export declare function doctorRenderer(options?: OptionalContext): Promise<Record<string, unknown>>;
export declare function trustRenderer(options: RendererTrustOptions): Promise<RendererTrustStore>;
export declare function registerRenderer(options: OptionalContext & {
    renderer: RendererManifest;
}): Promise<RendererManifest>;
export declare function readRendererTrustStore(options?: OptionalContext): Promise<RendererTrustStore>;
