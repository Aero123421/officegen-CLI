import { Command } from "commander";
import { OFFICEGEN_CLI_VERSION } from "../../../core/dist/index.js";
import { commandFromArgv, positionalArgs } from "../shared/argv.js";
import { makeEnvelope, writeResult } from "../shared/envelope.js";
import { COMMAND_METADATA, metadataFor } from "../shared/metadata.js";
import { agentPayload, assetPayload, capabilitiesPayload, chartPayload, configPayload, designPayload, diagnosePayload, diagramPayload, doctorPayload, editPayload, errorInspectPayload, errorsListPayload, exportPayload, groupPayload, helpPayload, inspectPayload, layoutPayload, mcpPayload, pluginPayload, renderPayload, rendererPayload, repairPayload, scaffoldPayload, schemaGetPayload, schemaListPayload, schemaMigratePayload, templatePayload, validatePayload, viewPayload, wiredPayload } from "./payloads.js";
const leafPayloads = {
    capabilities: capabilitiesPayload,
    help: (ctx) => helpPayload(ctx, positionalArgs(ctx.argv, 3)),
    doctor: doctorPayload,
    inspect: inspectPayload,
    view: viewPayload,
    edit: editPayload,
    render: renderPayload,
    scaffold: scaffoldPayload,
    export: exportPayload,
    validate: validatePayload,
    diagnose: diagnosePayload,
    repair: repairPayload,
    run: wiredPayload("run")
};
const groupPayloads = {
    asset: assetPayload,
    chart: chartPayload,
    diagram: diagramPayload,
    template: templatePayload,
    design: designPayload,
    layout: layoutPayload,
    agent: agentPayload,
    mcp: mcpPayload,
    renderer: rendererPayload,
    plugin: pluginPayload
};
export function createProgram(context, stdout, _stderr, now) {
    const program = new Command();
    suppressCommanderOutput(program);
    program
        .name("officegen")
        .description("Officegen CLI")
        .version(OFFICEGEN_CLI_VERSION)
        .helpCommand(false)
        .allowUnknownOption(false)
        .allowExcessArguments(true)
        .option("--json", "emit JSON envelope")
        .option("--agent", "filter output for agents")
        .option("--capabilities-hash <hash>", "expected active capabilities hash")
        .option("--json-budget-bytes <bytes>", "agent JSON output budget")
        .exitOverride();
    registerConfig(program, context, stdout, now);
    registerSchema(program, context, stdout, now);
    registerErrors(program, context, stdout, now);
    for (const entry of COMMAND_METADATA) {
        if (entry.feature === "config" || entry.feature === "schema" || entry.feature === "errors")
            continue;
        const leaf = leafPayloads[entry.feature];
        if (leaf) {
            registerLeaf(program, entry.feature, context, stdout, now, leaf);
            continue;
        }
        const group = groupPayloads[entry.feature] ?? groupPayload;
        registerGroup(program, entry.feature, context, stdout, now, group);
    }
    return program;
}
export function writeNativeHelp(context, stdout) {
    const enabled = context.registry
        .filter((entry) => entry.enabled && entry.visibleInHelp)
        .filter((entry) => !context.agent || entry.visibleToAgents);
    const disabled = context.registry.filter((entry) => !entry.enabled && entry.visibleInHelp);
    const lines = [
        "officegen - AI-friendly Office/PDF runtime",
        "",
        "Usage:",
        "  officegen <command> [options]",
        "",
        "Commands:",
        ...enabled.map((entry) => `  ${entry.commandGroup.padEnd(12)} ${entry.description}`),
        "",
        "Options:",
        "  --json                         emit JSON envelope",
        "  --agent                        filter output for agents",
        "  --capabilities-hash <hash>     warn if adapter capabilities are stale",
        "  --json-budget-bytes <bytes>    cap agent JSON output",
        "  -V, --version",
        "  -h, --help",
        "",
        "Run:",
        "  officegen capabilities --agent --json"
    ];
    if (disabled.length) {
        lines.splice(lines.length - 2, 0, "", "Disabled by current profile:", ...disabled.map((entry) => `  ${entry.commandGroup.padEnd(12)} disabled by profile ${context.config.profile}`));
    }
    stdout(lines.join("\n"));
}
function registerLeaf(program, feature, context, stdout, now, payloadFactory) {
    const metadata = metadataFor(feature);
    if (!metadata || !isCommandVisibleInNativeHelp(context, feature))
        return;
    program.addCommand(baseCommand(metadata.commandGroup, metadata.description).action(async () => {
        const payload = await payloadFactory(context);
        writeResult(context, makeEnvelope(context, commandFromArgv(context.argv), payload, now), stdout);
    }));
}
function registerGroup(program, feature, context, stdout, now, payloadFactory) {
    const metadata = metadataFor(feature);
    if (!metadata || !isCommandVisibleInNativeHelp(context, feature))
        return;
    const subcommands = metadata.commands.map((command) => command.split(" ")[1]).filter((value) => Boolean(value));
    const group = baseCommand(metadata.commandGroup, metadata.description).action(async () => {
        const payload = await payloadFactory(context);
        writeResult(context, makeEnvelope(context, commandFromArgv(context.argv), payload, now), stdout);
    });
    for (const subcommand of subcommands) {
        group.addCommand(baseCommand(subcommand, `${metadata.commandGroup} ${subcommand}`).action(async () => {
            const payload = await payloadFactory(context, subcommand);
            writeResult(context, makeEnvelope(context, commandFromArgv(context.argv), payload, now), stdout);
        }));
    }
    program.addCommand(group);
}
function registerConfig(program, context, stdout, now) {
    if (!isCommandVisibleInNativeHelp(context, "config"))
        return;
    const metadata = metadataFor("config");
    const config = baseCommand("config", metadata.description).action(async () => {
        writeResult(context, makeEnvelope(context, commandFromArgv(context.argv), configPayload(context), now), stdout);
    });
    config.addCommand(baseCommand("show", "show active config").action(async () => {
        writeResult(context, makeEnvelope(context, commandFromArgv(context.argv), configPayload(context), now), stdout);
    }));
    config.addCommand(baseCommand("set", "set config value").action(async () => {
        writeResult(context, makeEnvelope(context, commandFromArgv(context.argv), {
            schema: "officegen.config.result@1.2",
            status: "wired",
            message: "config set is registered; persistent writes are delegated to the core config API."
        }, now), stdout);
    }));
    program.addCommand(config);
}
function registerSchema(program, context, stdout, now) {
    if (!isCommandVisibleInNativeHelp(context, "schema"))
        return;
    const metadata = metadataFor("schema");
    const schema = baseCommand("schema", metadata.description).action(async () => {
        writeResult(context, makeEnvelope(context, commandFromArgv(context.argv), schemaListPayload(context), now), stdout);
    });
    schema.addCommand(baseCommand("list", "list schemas").action(async () => writeResult(context, makeEnvelope(context, commandFromArgv(context.argv), schemaListPayload(context), now), stdout)));
    schema.addCommand(baseCommand("get", "get schema").action(async () => writeResult(context, makeEnvelope(context, commandFromArgv(context.argv), schemaGetPayload(context), now), stdout)));
    schema.addCommand(baseCommand("validate", "validate schema").action(async () => writeResult(context, makeEnvelope(context, commandFromArgv(context.argv), await validatePayload(context), now), stdout)));
    schema.addCommand(baseCommand("migrate", "migrate schema").action(async () => writeResult(context, makeEnvelope(context, commandFromArgv(context.argv), await schemaMigratePayload(context), now), stdout)));
    program.addCommand(schema);
}
function registerErrors(program, context, stdout, now) {
    if (!isCommandVisibleInNativeHelp(context, "errors"))
        return;
    const metadata = metadataFor("errors");
    const errors = baseCommand("errors", metadata.description).action(async () => {
        writeResult(context, makeEnvelope(context, commandFromArgv(context.argv), errorsListPayload(), now), stdout);
    });
    errors.addCommand(baseCommand("list", "list errors").action(async () => writeResult(context, makeEnvelope(context, commandFromArgv(context.argv), errorsListPayload(), now), stdout)));
    errors.addCommand(baseCommand("inspect", "inspect error").action(async () => writeResult(context, makeEnvelope(context, commandFromArgv(context.argv), errorInspectPayload(context), now), stdout)));
    program.addCommand(errors);
}
function isCommandVisibleInNativeHelp(context, feature) {
    const entry = context.registry.find((candidate) => candidate.commandGroup === feature);
    return Boolean(entry?.enabled && entry.visibleInHelp && (!context.agent || entry.visibleToAgents));
}
function baseCommand(name, description) {
    const command = new Command(name);
    suppressCommanderOutput(command);
    return command
        .description(description)
        .allowUnknownOption(false)
        .allowExcessArguments(true)
        .argument("[args...]", "arguments")
        .option("--json", "emit JSON envelope")
        .option("--agent", "filter output for agents")
        .option("--capabilities-hash <hash>", "expected active capabilities hash")
        .option("--json-budget-bytes <bytes>", "agent JSON output budget")
        .option("--depth <depth>", "inspection depth")
        .option("--format <format>", "view/export format")
        .option("--max-pages <number>", "maximum pages")
        .option("--out <path>", "output path")
        .option("--overwrite", "allow overwriting an existing output")
        .option("--schema <id>", "schema id")
        .option("--kind <kind>", "document kind")
        .option("--title <title>", "document title")
        .option("--ops <path>", "edit operations JSON")
        .option("--dry-run", "resolve without writing")
        .option("--resolve-selectors", "include selector resolution details")
        .option("--to <format>", "export target format")
        .option("--mode <mode>", "export mode")
        .option("--issues <path>", "repair issues JSON")
        .option("--images", "extract image assets")
        .option("--asset <path>", "asset zip path")
        .option("--selector <selector>", "asset or object selector")
        .option("--name <name>", "template, design, plugin, or renderer name")
        .option("--map <path>", "template map JSON")
        .option("--data <path>", "template data JSON")
        .option("--sha256 <hash>", "expected sha256")
        .option("--trust <pin>", "trust pin")
        .option("--from <schema>", "source schema")
        .option("--target <target>", "adapter target")
        .option("--scope <scope>", "scope")
        .exitOverride();
}
function suppressCommanderOutput(command) {
    command.configureOutput({
        writeOut: () => undefined,
        writeErr: () => undefined
    });
}
//# sourceMappingURL=register.js.map