import { OptionalCapabilities, OptionalContext } from "./common.js";
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
export declare function installAgentAdapter(options: AgentInstallOptions): Promise<AgentAdapterRecord>;
export declare function refreshAgentAdapter(options: AgentRefreshOptions): Promise<AgentAdapterRecord>;
export declare function generateAgentAdapterText(options: {
    name: string;
    capabilities?: OptionalCapabilities;
    instructions?: string;
    generatedAt?: string;
}): string;
