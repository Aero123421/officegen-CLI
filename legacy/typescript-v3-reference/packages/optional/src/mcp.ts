import { createInterface } from "node:readline";
import { stdin as defaultStdin, stdout as defaultStdout } from "node:process";

import {
  OptionalContext,
  OptionalFeature,
  normalizeCapabilities,
  requireFeature,
  untrustedContentWarning
} from "./common.js";

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

const featureTools: Record<OptionalFeature, McpTool[]> = {
  agent: [
    {
      name: "officegen.agent.adapter",
      description: "Generate adapter text with capabilitiesHash and untrusted-content warning.",
      inputSchema: { type: "object", properties: { name: { type: "string" } } }
    }
  ],
  template: [
    {
      name: "officegen.template.list",
      description: "List optional template JSON records.",
      inputSchema: { type: "object", properties: {} }
    },
    {
      name: "officegen.template.validate",
      description: "Validate a template JSON record.",
      inputSchema: { type: "object", properties: { id: { type: "string" } } }
    }
  ],
  design: [
    {
      name: "officegen.design.list",
      description: "List optional design profiles.",
      inputSchema: { type: "object", properties: {} }
    }
  ],
  layout: [
    {
      name: "officegen.layout.apply",
      description: "Apply simple layout constraints to JSON boxes.",
      inputSchema: { type: "object", properties: { boxes: { type: "array" }, constraints: { type: "array" } } }
    }
  ],
  plugin: [
    {
      name: "officegen.plugin.list",
      description: "List installed plugin manifests and trust pins.",
      inputSchema: { type: "object", properties: {} }
    }
  ],
  renderer: [
    {
      name: "officegen.renderer.list",
      description: "List registered renderers gated by renderer capability.",
      inputSchema: { type: "object", properties: {} }
    }
  ],
  mcp: [
    {
      name: "officegen.capabilities",
      description: "Return optional feature gates and capabilitiesHash.",
      inputSchema: { type: "object", properties: {} }
    }
  ]
};

export function listMcpTools(options: OptionalContext = {}): McpTool[] {
  requireFeature(options, "mcp", "mcp tools/list");
  const capabilities = normalizeCapabilities(options.capabilities);
  return capabilities.features.flatMap((feature) => featureTools[feature] ?? []);
}

export async function serveMcpStdio(options: McpServeOptions = {}): Promise<void> {
  requireFeature(options, "mcp", "mcp serve --stdio");
  const input = options.stdin ?? defaultStdin;
  const output = options.stdout ?? defaultStdout;
  const lines = createInterface({ input });

  for await (const line of lines) {
    if (!line.trim()) {
      continue;
    }

    const response = await handleMcpRequest(options, line);
    if (response) {
      output.write(`${JSON.stringify(response)}\n`);
    }
  }
}

export async function handleMcpRequest(options: OptionalContext, line: string): Promise<JsonRpcResponse | null> {
  let request: JsonRpcRequest;
  try {
    request = JSON.parse(line) as JsonRpcRequest;
  } catch {
    return jsonRpcError(null, -32700, "Parse error");
  }

  const id = request.id ?? null;
  try {
    if (request.method === "initialize") {
      const capabilities = normalizeCapabilities(options.capabilities);
      return jsonRpcResult(id, {
        protocolVersion: "2024-11-05",
        serverInfo: { name: "officegen-optional", version: "0.1.0" },
        capabilities: {
          tools: {},
          officegen: capabilities,
          warning: untrustedContentWarning
        }
      });
    }

    if (request.method === "tools/list") {
      return jsonRpcResult(id, { tools: listMcpTools(options) });
    }

    if (request.method.startsWith("notifications/")) {
      return null;
    }

    return jsonRpcError(id, -32601, `Method not found: ${request.method}`);
  } catch (error) {
    return jsonRpcError(id, -32000, error instanceof Error ? error.message : String(error));
  }
}

function jsonRpcResult(id: string | number | null, result: unknown): JsonRpcResponse {
  return { jsonrpc: "2.0", id, result };
}

function jsonRpcError(id: string | number | null, code: number, message: string): JsonRpcResponse {
  return { jsonrpc: "2.0", id, error: { code, message } };
}
