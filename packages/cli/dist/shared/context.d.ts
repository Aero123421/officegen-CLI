import type { ActiveCapability, CliErrorPayload, RuntimeContext } from "./types.js";
export declare function createRuntimeContext(argv: string[], cwd: string, env: NodeJS.ProcessEnv): Promise<RuntimeContext>;
export declare function buildActiveRegistry(config: RuntimeContext["config"]): ActiveCapability[];
export declare function gateTopLevelCommand(command: string, context: RuntimeContext): CliErrorPayload | undefined;
export declare function availableCommands(context: RuntimeContext): string[];
export declare function nextSuggestedCommands(context: RuntimeContext): string[];
