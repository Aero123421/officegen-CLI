import { OFFICEGEN_CLI_VERSION, OfficegenError } from "../../core/dist/index.js";
import { createProgram, writeCommandHelp, writeNativeHelp } from "./commands/register.js";
import { commandFromArgv, getTopCommand, hasFlag, secondCommandToken } from "./shared/argv.js";
import { createRuntimeContext, gateTopLevelCommand } from "./shared/context.js";
import { makeEnvelope, makeErrorEnvelope, writeResult } from "./shared/envelope.js";
import { helpPayload } from "./commands/payloads.js";
import { CliFailure } from "./shared/types.js";
export async function runCli(argv, options = {}) {
    const cwd = options.cwd ?? process.cwd();
    const originalStdoutWrite = process.stdout.write.bind(process.stdout);
    const originalStderrWrite = process.stderr.write.bind(process.stderr);
    const stdout = options.stdout ?? ((text) => originalStdoutWrite(`${text}\n`));
    const stderr = options.stderr ?? ((text) => originalStderrWrite(`${text}\n`));
    const env = options.env ?? process.env;
    const now = options.now ?? new Date();
    const context = await createRuntimeContext(argv, cwd, env);
    const topCommand = getTopCommand(argv);
    const commandText = commandFromArgv(argv);
    if (!topCommand && (hasFlag(argv, "--version") || hasFlag(argv, "-V"))) {
        if (context.json) {
            writeResult(context, makeEnvelope(context, commandText || "version", { schema: "officegen.version.result@1.2", version: OFFICEGEN_CLI_VERSION }, now), stdout);
            return;
        }
        stdout(OFFICEGEN_CLI_VERSION);
        return;
    }
    if (!topCommand && (hasFlag(argv, "--help") || hasFlag(argv, "-h"))) {
        if (context.json) {
            writeResult(context, makeEnvelope(context, commandText || "help", helpPayload(context, []), now), stdout);
            return;
        }
        writeNativeHelp(context, stdout);
        return;
    }
    if (!topCommand && argv.slice(2).some((arg) => arg.startsWith("-"))) {
        await parseWithCommanderSafely(argv, context, stdout, stderr, now, commandText);
        return;
    }
    if (!topCommand) {
        writeResult(context, makeEnvelope(context, commandText || "help", helpPayload(context, []), now), stdout);
        return;
    }
    if (topCommand.startsWith("-")) {
        await parseWithCommanderSafely(argv, context, stdout, stderr, now, commandText);
        return;
    }
    const gateError = gateTopLevelCommand(topCommand, context);
    if (gateError) {
        process.exitCode = gateError.code === "UNKNOWN_COMMAND" ? 2 : 5;
        writeResult(context, makeErrorEnvelope(context, commandText, gateError, now), context.json ? stdout : stderr);
        return;
    }
    if (hasFlag(argv, "--help") || hasFlag(argv, "-h")) {
        if (context.json) {
            writeResult(context, makeEnvelope(context, commandText, helpPayload(context, [topCommand, secondCommandToken(argv)].filter(Boolean)), now), stdout);
        }
        else {
            writeCommandHelp(context, topCommand, secondCommandToken(argv), stdout);
        }
        return;
    }
    try {
        await parseWithCommander(argv, context, stdout, stderr, now);
    }
    catch (error) {
        const failure = toCliFailure(error, commandText);
        process.exitCode = failure.exitCode;
        writeResult(context, makeErrorEnvelope(context, commandText, failure.payload, now), context.json ? stdout : stderr);
    }
}
async function parseWithCommanderSafely(argv, context, stdout, stderr, now, commandText) {
    try {
        await parseWithCommander(argv, context, stdout, stderr, now);
    }
    catch (error) {
        const failure = toCliFailure(error, commandText);
        process.exitCode = failure.exitCode;
        writeResult(context, makeErrorEnvelope(context, commandText, failure.payload, now), context.json ? stdout : stderr);
    }
}
async function parseWithCommander(argv, context, stdout, stderr, now) {
    const program = createProgram(context, stdout, stderr, now);
    await program.parseAsync(argv, { from: "node" });
}
function toCliFailure(error, commandText) {
    if (error instanceof CliFailure)
        return error;
    if (error instanceof OfficegenError) {
        return new CliFailure({
            code: error.payload.code,
            category: error.payload.category,
            severity: error.payload.severity,
            command: error.payload.command ?? commandText,
            feature: error.payload.feature,
            message: error.payload.message,
            details: asRecord(error.payload.details)
        }, error.payload.code.startsWith("SECURITY_") ? 4 : 3);
    }
    if (error instanceof Error) {
        if (/unknown option/i.test(error.message)) {
            return new CliFailure({
                code: "UNKNOWN_OPTION",
                command: commandText,
                message: error.message
            }, 2);
        }
        if (isNodeError(error) && error.code === "ENOENT") {
            return new CliFailure({
                code: "INPUT_NOT_FOUND",
                command: commandText,
                message: error.message
            }, 3);
        }
        if (error instanceof SyntaxError || /JSON\.parse|Unexpected token|Unexpected end of JSON/i.test(error.message)) {
            return new CliFailure({
                code: "INPUT_PARSE_ERROR",
                command: commandText,
                message: error.message
            }, 3);
        }
        if (/WinAnsi cannot encode|cannot encode .* with WinAnsi/i.test(error.message)) {
            return new CliFailure({
                code: "RENDER_FONT_UNSUPPORTED",
                command: commandText,
                message: error.message
            }, 3);
        }
        if (/Unsupported inspect format/i.test(error.message)) {
            return new CliFailure({
                code: "UNSUPPORTED_FORMAT",
                command: commandText,
                message: error.message
            }, 3);
        }
        if (/Unsupported export|unsupported render target/i.test(error.message)) {
            return new CliFailure({
                code: "EXPORT_UNSUPPORTED",
                command: commandText,
                message: error.message
            }, 3);
        }
        const coded = /^([A-Z][A-Z0-9_]+):\s*(.*)$/.exec(error.message);
        if (coded) {
            return new CliFailure({
                code: coded[1] ?? "UNKNOWN_COMMAND",
                command: commandText,
                message: coded[2] ?? error.message
            }, 3);
        }
    }
    return new CliFailure({
        code: "UNKNOWN_COMMAND",
        command: commandText,
        message: error instanceof Error ? error.message : String(error)
    }, 2);
}
function asRecord(value) {
    return value && typeof value === "object" && !Array.isArray(value) ? value : undefined;
}
function isNodeError(error) {
    return error instanceof Error && "code" in error;
}
//# sourceMappingURL=program.js.map