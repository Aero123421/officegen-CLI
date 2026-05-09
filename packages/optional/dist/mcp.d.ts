import { OptionalContext } from "./common.js";
export interface JsonRpcRequest {
    jsonrpc?: "2.0";
    id?: string | number | null;
    method: string;
    params?: unknown;
}
export interface JsonRpcResponse {
    jsonrpc: "2.0";
    id: string | number | null;
    result?: unknown;
    error?: {
        code: number;
        message: string;
    };
}
export interface McpServeOptions extends OptionalContext {
    stdin?: NodeJS.ReadableStream;
    stdout?: NodeJS.WritableStream;
}
interface McpTool {
    name: string;
    description: string;
    inputSchema: {
        type: "object";
        properties: Record<string, unknown>;
    };
}
export declare function listMcpTools(options?: OptionalContext): McpTool[];
export declare function serveMcpStdio(options?: McpServeOptions): Promise<void>;
export declare function handleMcpRequest(options: OptionalContext, line: string): Promise<JsonRpcResponse | null>;
export {};
