import path from "node:path";
import { homedir } from "node:os";
const secretPatterns = [
    { kind: "secret-like-token", pattern: /-----BEGIN [A-Z ]*PRIVATE KEY-----[\s\S]*?-----END [A-Z ]*PRIVATE KEY-----/g, replacement: "<redacted:secret-like-token>" },
    { kind: "secret-like-token", pattern: /\bBearer\s+[A-Za-z0-9._~+/=-]{12,}\b/g, replacement: "Bearer <redacted:secret-like-token>" },
    { kind: "secret-like-token", pattern: /\b(?:sk|pk|api|key|token|secret)[-_]?[A-Za-z0-9]{16,}\b/gi, replacement: "<redacted:secret-like-token>" },
    { kind: "secret-like-token", pattern: /\b[A-Za-z0-9_.-]*(?:password|passwd|pwd|secret|token|api[_-]?key|access[_-]?key|private[_-]?key|client[_-]?secret)[A-Za-z0-9_.-]*\s*=\s*["']?([^"'\s;&]{6,})["']?/gi, replacement: "$&" },
    { kind: "secret-like-token", pattern: /([?&](?:access[_-]?token|api[_-]?key|token|key|secret|signature)=)[^&#\s]+/gi, replacement: "$1<redacted:secret-like-token>" },
    { kind: "secret-like-token", pattern: /\b(?:Cookie|Set-Cookie):\s*[^\n\r]+/gi, replacement: "Cookie: <redacted:secret-like-token>" }
];
export function redactSecretsInText(text, location = "$") {
    let value = text;
    const redactions = [];
    for (const { pattern, replacement } of secretPatterns) {
        value = value.replace(pattern, (match, ...args) => {
            const finalReplacement = replacement === "$&" ? String(match).replace(/=.*/, "=<redacted:secret-like-token>") : replacement.replace("$1", String(args[0] ?? ""));
            redactions.push({ kind: "secret-like-token", location, replacement: finalReplacement });
            return finalReplacement;
        });
    }
    return { value, redactions };
}
function expandHome(filePath) {
    return filePath === "~" || filePath.startsWith("~/") || filePath.startsWith("~\\")
        ? path.join(homedir(), filePath.slice(2))
        : filePath;
}
function pathReplacements(config, run) {
    const roots = [
        { root: config.paths.projectRoot, label: "<project>" },
        { root: path.resolve(expandHome(config.paths.projectRoot)), label: "<project>" },
        { root: config.paths.userConfigDir, label: "<userConfig>" },
        { root: path.resolve(expandHome(config.paths.userConfigDir)), label: "<userConfig>" },
        { root: homedir(), label: "<userHome>" },
        { root: path.resolve(homedir()), label: "<userHome>" }
    ];
    if (run)
        roots.push({ root: run.root, label: "<run>" }, { root: path.resolve(run.root), label: "<run>" });
    const seen = new Set();
    return roots
        .map(({ root, label }) => ({ root: root.replace(/[\\/]+$/, ""), label }))
        .filter(({ root }) => root.length > 0)
        .filter(({ root }) => {
        const key = root.replace(/\\/g, "/").toLowerCase();
        if (seen.has(key))
            return false;
        seen.add(key);
        return true;
    })
        .sort((a, b) => b.root.length - a.root.length);
}
function rootMatcher(root) {
    const escaped = root
        .replace(/\\/g, "/")
        .replace(/[.*+?^${}()|[\]\\]/g, "\\$&")
        .replace(/\//g, "[\\\\/]");
    return new RegExp(`${escaped}((?:[\\\\/][^"'\\s,;)]*)?)`, "gi");
}
export function redactPathsInText(text, config, location = "$", run) {
    let value = text;
    const redactions = [];
    for (const { root, label } of pathReplacements(config, run)) {
        value = value.replace(rootMatcher(root), (_match, suffix) => {
            const normalizedSuffix = String(suffix ?? "").replace(/\\/g, "/");
            const replacement = `${label}${normalizedSuffix}`;
            redactions.push({ kind: "absolute-path", location, replacement });
            return replacement;
        });
    }
    for (const uncMatch of [...value.matchAll(/\\\\[^\\/\s]+[\\/][^\\/\s]+(?:[\\/][^"'\s,;)]*)?/g)]) {
        const replacement = "<uncPath>";
        value = value.replace(uncMatch[0], replacement);
        redactions.push({ kind: "absolute-path", location, replacement });
    }
    for (const driveMatch of [...value.matchAll(/\b[A-Za-z]:[\\/][^"'\s,;)]*/g)]) {
        if (!driveMatch[0].startsWith("<")) {
            const replacement = "<absolutePath>";
            value = value.replace(driveMatch[0], replacement);
            redactions.push({ kind: "absolute-path", location, replacement });
        }
    }
    for (const homeMatch of [...value.matchAll(/(?<![\w:>/])~[\\/][^"'\s,;)]*/g)]) {
        const replacement = "<userHome>";
        value = value.replace(homeMatch[0], replacement);
        redactions.push({ kind: "absolute-path", location, replacement });
    }
    for (const posixMatch of [...value.matchAll(/(?<![\w:>/])\/(?!\/)[^"'\s,;)]+(?:\/[^"'\s,;)]+)+/g)]) {
        const raw = posixMatch[0];
        if (value.startsWith("#/"))
            continue;
        const prefix = raw.startsWith("/") ? "" : raw[0];
        const pathValue = prefix ? raw.slice(1) : raw;
        if (!pathValue.startsWith("<")) {
            const replacement = "<absolutePath>";
            redactions.push({ kind: "absolute-path", location, replacement });
            value = value.replace(pathValue, replacement);
        }
    }
    return { value, redactions };
}
function redactValue(value, config, location, run) {
    if (typeof value === "string") {
        let text = value;
        const redactions = [];
        if (config.security.redactAbsolutePathsInJson && !isMarkupPayload(text, location)) {
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
        const redactions = [];
        const items = value.map((item, index) => {
            const result = redactValue(item, config, `${location}[${index}]`, run);
            redactions.push(...result.redactions);
            return result.value;
        });
        return { value: items, redactions };
    }
    if (value && typeof value === "object") {
        const redactions = [];
        const out = {};
        for (const [key, child] of Object.entries(value)) {
            const result = redactValue(child, config, `${location}.${key}`, run);
            out[key] = result.value;
            redactions.push(...result.redactions);
        }
        return { value: out, redactions };
    }
    return { value, redactions: [] };
}
export function redactJson(value, config, run) {
    const result = redactValue(value, config, "$", run);
    return { value: result.value, redactions: result.redactions };
}
export function isAbsolutePathRedactionNeeded(text) {
    return path.isAbsolute(text) || /^[A-Za-z]:[\\/]/.test(text) || /^\\\\[^\\/]+[\\/][^\\/]+/.test(text) || /^~[\\/]/.test(text);
}
function isMarkupPayload(text, location) {
    const trimmed = text.trimStart();
    if (/\.svg$/i.test(location) && /^<svg\b/i.test(trimmed))
        return true;
    if (/\.html?$/i.test(location) && /^(?:<!doctype\s+html|<html\b)/i.test(trimmed))
        return true;
    if (/\.xml$/i.test(location) && /^<\?xml\b/i.test(trimmed))
        return true;
    return false;
}
//# sourceMappingURL=redaction.js.map