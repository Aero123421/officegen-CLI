import { promises as fs } from "node:fs";
import path from "node:path";
import { OfficegenError, validatePath } from "../../../core/dist/index.js";
import { createOptionalCapabilities } from "../../../optional/dist/index.js";
import { commandFromArgv, hasFlag, optionValue, positionalArgs } from "./argv.js";
import { CliFailure } from "./types.js";
export function requireInput(context, start, command) {
    const input = positionalArgs(context.argv, start)[0];
    if (!input) {
        throw new CliFailure({
            code: "SCHEMA_INVALID",
            command,
            message: `${command} requires an input path.`
        }, 2);
    }
    return input;
}
export function resolveCliPath(context, inputPath) {
    if (!inputPath)
        return path.resolve(context.cwd, inputPath);
    return path.isAbsolute(inputPath) ? inputPath : path.resolve(context.cwd, inputPath);
}
export async function validateInputPath(context, inputPath) {
    try {
        const validated = await validatePath(context.config, {
            kind: "input",
            path: inputPath
        });
        await enforceMaxInputFileBytes(context, validated.absolutePath, inputPath);
        return validated.absolutePath;
    }
    catch (error) {
        if (error instanceof OfficegenError) {
            throw pathFailure(context, error);
        }
        throw error;
    }
}
export async function validatedOutOption(context) {
    const out = optionValue(context.argv, "--out");
    return out ? validateOutputPath(context, out) : undefined;
}
export async function validateOutputPath(context, outputPath, options = {}) {
    try {
        const validated = await validatePath(context.config, {
            kind: "output",
            path: outputPath,
            overwrite: options.directory ? true : hasFlag(context.argv, "--overwrite")
        });
        return validated.absolutePath;
    }
    catch (error) {
        if (error instanceof OfficegenError) {
            throw new CliFailure({
                code: error.payload.code,
                category: error.payload.category,
                severity: error.payload.severity,
                command: commandFromArgv(context.argv),
                message: userFacingPathMessage(error.payload.code, error.payload.message),
                details: asRecord(error.payload.details)
            }, error.payload.code.startsWith("SECURITY_") ? 4 : 3);
        }
        throw error;
    }
}
export async function readInputFile(context, inputPath) {
    return fs.readFile(await validateInputPath(context, inputPath));
}
export async function readInputText(context, inputPath) {
    return fs.readFile(await validateInputPath(context, inputPath), "utf8");
}
export async function readInputJson(context, inputPath) {
    try {
        return JSON.parse(await readInputText(context, inputPath));
    }
    catch (error) {
        if (error instanceof SyntaxError) {
            throw new CliFailure({
                code: "INPUT_PARSE_ERROR",
                category: "input",
                severity: "error",
                command: commandFromArgv(context.argv),
                message: `Input JSON could not be parsed: ${error.message}`,
                details: { input: inputPath }
            }, 3);
        }
        throw error;
    }
}
export async function readInputJsonIfPresent(context, inputPath) {
    if (!inputPath)
        return {};
    return readInputJson(context, inputPath);
}
export function normalizeEditOperations(raw) {
    const record = asRecord(raw);
    const ops = Array.isArray(record.ops) ? record.ops : Array.isArray(raw) ? raw : [];
    return ops.map((op) => {
        const item = asRecord(op);
        if (typeof item.type === "string")
            return item;
        if (item.op === "pptx.setShapeText" || item.op === "docx.setParagraphText") {
            return {
                type: "setText",
                selector: asRecord(item.selector),
                text: String(item.text ?? item.value ?? "")
            };
        }
        if (item.op === "pptx.replaceText" || item.op === "docx.replaceText") {
            return {
                type: "replaceText",
                from: String(item.from ?? item.search ?? ""),
                to: String(item.to ?? item.text ?? "")
            };
        }
        if (item.op === "pdf.addTextOverlay") {
            return {
                type: "pdf.textOverlay",
                page: Number(item.page ?? 1),
                text: String(item.text ?? ""),
                x: Number(item.x ?? 72),
                y: Number(item.y ?? 72),
                size: Number(item.size ?? 12)
            };
        }
        return item;
    });
}
export function optionalContext(context) {
    const features = context.registry
        .filter((entry) => entry.enabled)
        .map((entry) => entry.feature)
        .filter((feature) => ["agent", "template", "design", "layout", "plugin", "renderer", "mcp"].includes(feature));
    return {
        cwd: context.cwd,
        capabilities: createOptionalCapabilities(features)
    };
}
export function asRecord(value) {
    return value && typeof value === "object" && !Array.isArray(value) ? value : {};
}
export function stringRecord(value) {
    return Object.fromEntries(Object.entries(asRecord(value)).map(([key, nested]) => [key, String(nested)]));
}
export function numberOption(context, name) {
    const value = optionValue(context.argv, name);
    if (!value)
        return undefined;
    const parsed = Number(value);
    return Number.isFinite(parsed) ? parsed : undefined;
}
export function schemaHiddenFromAgent(context, schema) {
    const hidden = new Set(context.registry.filter((entry) => !entry.visibleToAgents).map((entry) => entry.feature));
    return (schema.includes("template") && hidden.has("template")) || (schema.includes("design") && hidden.has("design"));
}
export function isNodeError(error) {
    return error instanceof Error && "code" in error;
}
async function enforceMaxInputFileBytes(context, filePath, inputPath) {
    const stats = await fs.stat(filePath);
    const maxInputFileBytes = context.config.security.untrustedInput.maxInputFileBytes;
    if (stats.size > maxInputFileBytes) {
        throw new CliFailure({
            code: "SECURITY_INPUT_TOO_LARGE",
            category: "security",
            severity: "error",
            command: commandFromArgv(context.argv),
            message: `Input file exceeds maxInputFileBytes (${stats.size} > ${maxInputFileBytes}).`,
            details: { input: inputPath, size: stats.size, maxInputFileBytes }
        }, 4);
    }
}
function pathFailure(context, error) {
    return new CliFailure({
        code: error.payload.code,
        category: error.payload.category,
        severity: error.payload.severity,
        command: commandFromArgv(context.argv),
        message: userFacingPathMessage(error.payload.code, error.payload.message),
        details: asRecord(error.payload.details)
    }, error.payload.code.startsWith("SECURITY_") ? 4 : 3);
}
function userFacingPathMessage(code, fallback) {
    if (code === "SECURITY_PATH_OUTSIDE_ROOT")
        return "Path must stay inside the configured trusted roots.";
    if (code === "SECURITY_ABSOLUTE_OUT_DENIED")
        return "Absolute output paths are denied by default.";
    if (code === "SECURITY_SYMLINK_DENIED")
        return "Refusing to write through a symlink.";
    if (code === "SECURITY_HARDLINK_DENIED")
        return "Refusing to overwrite a hardlinked file.";
    return fallback;
}
//# sourceMappingURL=io.js.map