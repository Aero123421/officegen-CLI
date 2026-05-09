import { createInterface } from "node:readline";
import { stdin as defaultStdin, stdout as defaultStdout } from "node:process";
import { normalizeCapabilities, requireFeature, untrustedContentWarning } from "./common.js";
const featureTools = {
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
export function listMcpTools(options = {}) {
    requireFeature(options, "mcp", "mcp tools/list");
    const capabilities = normalizeCapabilities(options.capabilities);
    return capabilities.features.flatMap((feature) => featureTools[feature] ?? []);
}
export async function serveMcpStdio(options = {}) {
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
export async function handleMcpRequest(options, line) {
    let request;
    try {
        request = JSON.parse(line);
    }
    catch {
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
    }
    catch (error) {
        return jsonRpcError(id, -32000, error instanceof Error ? error.message : String(error));
    }
}
function jsonRpcResult(id, result) {
    return { jsonrpc: "2.0", id, result };
}
function jsonRpcError(id, code, message) {
    return { jsonrpc: "2.0", id, error: { code, message } };
}
//# sourceMappingURL=mcp.js.map