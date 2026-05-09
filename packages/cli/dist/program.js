import { OFFICEGEN_CLI_VERSION, OfficegenError } from "@officegen/core";
import { createProgram, writeNativeHelp } from "./commands/register.js";
import { commandFromArgv, getTopCommand, hasFlag } from "./shared/argv.js";
import { createRuntimeContext, gateTopLevelCommand } from "./shared/context.js";
import { makeEnvelope, makeErrorEnvelope, writeResult } from "./shared/envelope.js";
import { helpPayload } from "./commands/payloads.js";
import { CliFailure } from "./shared/types.js";
export async function runCli(argv, options = {}) {
    const cwd = options.cwd ?? process.cwd();
    const stdout = options.stdout ?? ((text) => console.log(text));
    const stderr = options.stderr ?? ((text) => console.error(text));
    const env = options.env ?? process.env;
    const now = options.now ?? new Date();
    const context = await createRuntimeContext(argv, cwd, env);
    const topCommand = getTopCommand(argv);
    const commandText = commandFromArgv(argv);
    if (!topCommand && (hasFlag(argv, "--version") || hasFlag(argv, "-V"))) {
        stdout(OFFICEGEN_CLI_VERSION);
        return;
    }
    if (!topCommand && (hasFlag(argv, "--help") || hasFlag(argv, "-h"))) {
        writeNativeHelp(context, stdout);
        return;
    }
    if (!topCommand && argv.slice(2).some((arg) => arg.startsWith("-"))) {
        await parseWithCommander(argv, context, stdout, stderr, now);
        return;
    }
    if (!topCommand) {
        writeResult(context, makeEnvelope(context, commandText || "help", helpPayload(context, []), now), stdout);
        return;
    }
    if (topCommand.startsWith("-")) {
        await parseWithCommander(argv, context, stdout, stderr, now);
        return;
    }
    const gateError = gateTopLevelCommand(topCommand, context);
    if (gateError) {
        process.exitCode = gateError.code === "UNKNOWN_COMMAND" ? 2 : 5;
        writeResult(context, makeErrorEnvelope(context, commandText, gateError, now), context.json ? stdout : stderr);
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