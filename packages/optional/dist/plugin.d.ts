import { OptionalContext } from "./common.js";
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
export declare function installPlugin(options: PluginInstallOptions): Promise<PluginManifest>;
export declare function listPlugins(options?: OptionalContext): Promise<PluginManifest[]>;
export declare function inspectPlugin(options: PluginInspectOptions): Promise<PluginManifest & {
    trusted: boolean;
}>;
export declare function readPluginTrustStore(options?: OptionalContext): Promise<TrustStore>;
