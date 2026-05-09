import path from "node:path";
import type { JsonValue, OfficegenConfig, RedactionRecord, RedactionResult, RunFolder } from "./types.js";

const secretPatterns: Array<{ kind: "secret-like-token"; pattern: RegExp; replacement: string }> = [
  { kind: "secret-like-token", pattern: /-----BEGIN [A-Z ]*PRIVATE KEY-----[\s\S]*?-----END [A-Z ]*PRIVATE KEY-----/g, replacement: "<redacted:secret-like-token>" },
  { kind: "secret-like-token", pattern: /\bBearer\s+[A-Za-z0-9._~+/=-]{12,}\b/g, replacement: "Bearer <redacted:secret-like-token>" },
  { kind: "secret-like-token", pattern: /\b(?:sk|pk|api|key|token|secret)[-_]?[A-Za-z0-9]{16,}\b/gi, replacement: "<redacted:secret-like-token>" },
  { kind: "secret-like-token", pattern: /\b(?:password|pwd|secret|token|api[_-]?key)=([^;&\s]{6,})/gi, replacement: "$&" },
  { kind: "secret-like-token", pattern: /([?&](?:token|key|secret|signature)=)[^&#\s]+/gi, replacement: "$1<redacted:secret-like-token>" },
  { kind: "secret-like-token", pattern: /\b(?:Cookie|Set-Cookie):\s*[^\n\r]+/gi, replacement: "Cookie: <redacted:secret-like-token>" }
];

export function redactSecretsInText(text: string, location = "$"): RedactionResult<string> {
  let value = text;
  const redactions: RedactionRecord[] = [];
  for (const { pattern, replacement } of secretPatterns) {
    value = value.replace(pattern, (match, ...args: unknown[]) => {
      const finalReplacement = replacement === "$&" ? String(match).replace(/=.+$/, "=<redacted:secret-like-token>") : replacement.replace("$1", String(args[0] ?? ""));
      redactions.push({ kind: "secret-like-token", location, replacement: finalReplacement });
      return finalReplacement;
    });
  }
  return { value, redactions };
}

function normalize(filePath: string): string {
  return process.platform === "win32" ? filePath.toLowerCase() : filePath;
}

function pathReplacements(config: OfficegenConfig, run?: RunFolder): Array<{ root: string; label: string }> {
  const roots = [
    { root: path.resolve(config.paths.projectRoot), label: "<project>" },
    { root: path.resolve(config.paths.userConfigDir), label: "<userConfig>" }
  ];
  if (run) roots.push({ root: path.resolve(run.root), label: "<run>" });
  return roots.sort((a, b) => b.root.length - a.root.length);
}

export function redactPathsInText(text: string, config: OfficegenConfig, location = "$", run?: RunFolder): RedactionResult<string> {
  let value = text;
  const redactions: RedactionRecord[] = [];
  for (const { root, label } of pathReplacements(config, run)) {
    const escaped = root.replace(/[\\^$.*+?()[\]{}|]/g, "\\$&").replace(/\\\\/g, "[\\\\/]");
    const matcher = new RegExp(`${escaped}([\\\\/][^"'\\s,;)]*)?`, "gi");
    value = value.replace(matcher, (match) => {
      const suffix = match.slice(root.length).replace(/\\/g, "/");
      const replacement = `${label}${suffix}`;
      redactions.push({ kind: "absolute-path", location, replacement });
      return replacement;
    });
  }
  return { value, redactions };
}

function redactValue(value: JsonValue, config: OfficegenConfig, location: string, run?: RunFolder): RedactionResult<JsonValue> {
  if (typeof value === "string") {
    let text = value;
    const redactions: RedactionRecord[] = [];
    if (config.security.redactAbsolutePathsInJson) {
      const pathResult = redactPathsInText(text, config, location, run);
      text = pathResult.value;
      redactions.push(...pathResult.redactions);
    }
    if (config.security.redactSecretsInJson) {
      const secretResult = redactSecretsInText(text, location);
      text = secretResult.value;
      redactions.push(...secretResult.redactions);
    }
    return { value: text, redactions };
  }
  if (Array.isArray(value)) {
    const redactions: RedactionRecord[] = [];
    const items = value.map((item, index) => {
      const result = redactValue(item, config, `${location}[${index}]`, run);
      redactions.push(...result.redactions);
      return result.value;
    });
    return { value: items, redactions };
  }
  if (value && typeof value === "object") {
    const redactions: RedactionRecord[] = [];
    const out: Record<string, JsonValue> = {};
    for (const [key, child] of Object.entries(value)) {
      const result = redactValue(child, config, `${location}.${key}`, run);
      out[key] = result.value;
      redactions.push(...result.redactions);
    }
    return { value: out, redactions };
  }
  return { value, redactions: [] };
}

export function redactJson<T extends JsonValue>(value: T, config: OfficegenConfig, run?: RunFolder): RedactionResult<T> {
  const result = redactValue(value, config, "$", run);
  return { value: result.value as T, redactions: result.redactions };
}

export function isAbsolutePathRedactionNeeded(text: string): boolean {
  return path.isAbsolute(text) || /^[A-Za-z]:[\\/]/.test(text);
}

void normalize;
