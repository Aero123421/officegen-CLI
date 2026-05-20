export const GLOBAL_OPTION_SPECS = [
    option("--json", undefined, "emit JSON envelope"),
    option("--agent", undefined, "filter output for agents"),
    option("--strict-json", undefined, "force JSON-only stdout for agent execution"),
    option("--capabilities-hash", "<hash>", "expected active capabilities hash"),
    option("--json-budget-bytes", "<bytes>", "progressive-disclosure threshold for agent JSON output")
];
const OPTION_SPECS = {
    asset: option("--asset", "<path>", "asset zip path"),
    allowPartial: option("--allow-partial", undefined, "allow best-effort edit output when at least one op succeeds"),
    crop: option("--crop", undefined, "write an object crop artifact for view --object"),
    data: option("--data", "<path>", "template data or design patch JSON"),
    denyOutsideOutputRoot: option("--deny-outside-output-root", undefined, "fail outputs outside --output-root"),
    depth: option("--depth", "<summary|shallow|full>", "inspection depth"),
    dpi: option("--dpi", "<dpi>", "raster render resolution for PNG/JPEG view output"),
    dryRun: option("--dry-run", undefined, "resolve without writing"),
    embedded: option("--embedded", undefined, "inspect embedded media in Office packages"),
    emit: option("--emit", "<inspect|object-graph>", "choose inspect payload shape"),
    expectedArtifacts: option("--expected-artifacts", "<path>", "JSON list of expected artifacts"),
    externalLinks: option("--external-links", undefined, "verify external links"),
    fields: option("--fields", "<csv>", "project selected top-level result fields"),
    format: option("--format", "<format>", "view/export format"),
    formulas: option("--formulas", undefined, "verify XLSX formulas"),
    from: option("--from", "<schema>", "source schema"),
    gates: option("--gates", "<path>", "verification gates JSON"),
    goal: option("--goal", "<path>", "natural-language goal or constraints file"),
    images: option("--images", undefined, "extract image assets"),
    inPlace: option("--in-place", undefined, "allow edit input --out input with a backup transaction"),
    input: option("--input", "<path>", "input file for integrated workflows"),
    issues: option("--issues", "<path>", "repair issues JSON"),
    kind: option("--kind", "<kind>", "document kind"),
    lock: option("--lock", "<path>", "lock file JSON"),
    logJsonl: option("--log-jsonl", "<path>", "write UTF-8 JSONL run log"),
    manifest: option("--manifest", "<path>", "write run artifact manifest or read benchmark manifest"),
    map: option("--map", "<path>", "template map JSON"),
    matchesOnly: option("--matches-only", undefined, "emit only selector matches for select output"),
    maxPages: option("--max-pages", "<number>", "maximum preview pages"),
    mode: option("--mode", "<fast|internal|native|proof>", "export/view/verify mode"),
    name: option("--name", "<name>", "template, design, plugin, renderer, or lock owner name"),
    namedRanges: option("--named-ranges", undefined, "verify XLSX named ranges"),
    native: option("--native", undefined, "use native renderer when enabled by policy"),
    noObjectMap: option("--no-object-map", undefined, "omit object map arrays from output"),
    object: option("--object", "<stableObjectId>", "stable object id for object-scoped output"),
    objectMapLimit: option("--object-map-limit", "<number>", "limit object map entries in output"),
    ops: option("--ops", "<path>", "edit operations JSON"),
    out: option("--out", "<path>", "output path"),
    outputRoot: option("--output-root", "<path>", "restrict outputs to a directory"),
    overwrite: option("--overwrite", undefined, "allow overwriting an existing output"),
    owner: option("--owner", "<agent>", "lock owner agent id; preferred over lock --agent <id>"),
    plan: option("--plan", undefined, "return a plan-only payload without writing"),
    profile: option("--profile", "<profile>", "critique or improve profile"),
    protectedSheets: option("--protected-sheets", undefined, "verify protected or hidden sheets"),
    range: option("--range", "<range>", "limit XLSX inspect to an A1 range"),
    reference: option("--reference", "<path>", "reference file path"),
    reportOut: option("--report-out", "<path>", "write JSON report"),
    resolveSelectors: option("--resolve-selectors", undefined, "include selector resolution details"),
    schema: option("--schema", "<id>", "schema id"),
    scope: option("--scope", "<scope>", "scope"),
    selector: option("--selector", "<selector>", "asset or object selector"),
    sha256: option("--sha256", "<hash>", "expected sha256"),
    sheet: option("--sheet", "<name>", "limit XLSX inspect to a sheet"),
    slides: option("--slides", "<range>", "limit PPTX inspect to slide numbers or ranges"),
    strategy: option("--strategy", "<strategy>", "design apply strategy: theme-only, inspired, or faithful"),
    structure: option("--structure", undefined, "include DOCX structure map"),
    summary: option("--summary", "<path>", "write run Markdown summary"),
    summaryOnly: option("--summary-only", undefined, "emit compact summaries for large candidate outputs"),
    target: option("--target", "<target>", "adapter or render target"),
    timeoutMs: option("--timeout-ms", "<ms>", "per-command timeout budget where supported"),
    title: option("--title", "<title>", "document title"),
    to: option("--to", "<format>", "export target format"),
    trust: option("--trust", "<pin>", "trust pin"),
    tx: option("--tx", "<path>", "transaction record JSON"),
    txOut: option("--tx-out", "<path>", "write transaction record JSON"),
    validateOnly: option("--validate-only", undefined, "validate without writing"),
    verify: option("--verify", "<list>", "verification gates for integrated workflows"),
    visual: option("--visual", undefined, "include approximate visual diff/regression output")
};
export const COMMAND_OPTION_SPECS = {
    "config set": [OPTION_SPECS.scope],
    "schema get": [OPTION_SPECS.schema],
    "schema fetch": [OPTION_SPECS.schema],
    "schema validate": [OPTION_SPECS.schema],
    "schema migrate": [OPTION_SPECS.from, OPTION_SPECS.to, OPTION_SPECS.out],
    inspect: [
        OPTION_SPECS.depth,
        OPTION_SPECS.structure,
        OPTION_SPECS.slides,
        OPTION_SPECS.sheet,
        OPTION_SPECS.range,
        OPTION_SPECS.fields,
        OPTION_SPECS.objectMapLimit,
        OPTION_SPECS.noObjectMap,
        OPTION_SPECS.emit,
        OPTION_SPECS.reportOut
    ],
    view: [
        OPTION_SPECS.format,
        OPTION_SPECS.maxPages,
        OPTION_SPECS.dpi,
        OPTION_SPECS.mode,
        OPTION_SPECS.timeoutMs,
        OPTION_SPECS.out,
        OPTION_SPECS.object,
        OPTION_SPECS.selector,
        OPTION_SPECS.crop,
        OPTION_SPECS.objectMapLimit,
        OPTION_SPECS.reportOut
    ],
    edit: [
        OPTION_SPECS.ops,
        OPTION_SPECS.out,
        OPTION_SPECS.dryRun,
        OPTION_SPECS.resolveSelectors,
        OPTION_SPECS.allowPartial,
        OPTION_SPECS.overwrite,
        OPTION_SPECS.inPlace,
        OPTION_SPECS.txOut,
        OPTION_SPECS.lock,
        OPTION_SPECS.scope,
        OPTION_SPECS.name
    ],
    render: [OPTION_SPECS.target, OPTION_SPECS.out, OPTION_SPECS.overwrite],
    scaffold: [OPTION_SPECS.kind, OPTION_SPECS.title, OPTION_SPECS.out, OPTION_SPECS.overwrite],
    export: [OPTION_SPECS.to, OPTION_SPECS.mode, OPTION_SPECS.out, OPTION_SPECS.overwrite, OPTION_SPECS.timeoutMs],
    validate: [OPTION_SPECS.schema],
    verify: [
        OPTION_SPECS.visual,
        OPTION_SPECS.native,
        OPTION_SPECS.mode,
        OPTION_SPECS.timeoutMs,
        OPTION_SPECS.gates,
        OPTION_SPECS.out,
        OPTION_SPECS.reportOut,
        OPTION_SPECS.formulas,
        OPTION_SPECS.namedRanges,
        OPTION_SPECS.externalLinks,
        OPTION_SPECS.protectedSheets
    ],
    repair: [OPTION_SPECS.issues, OPTION_SPECS.out, OPTION_SPECS.dryRun, OPTION_SPECS.plan, OPTION_SPECS.overwrite],
    diff: [OPTION_SPECS.visual, OPTION_SPECS.native, OPTION_SPECS.maxPages, OPTION_SPECS.out, OPTION_SPECS.reportOut],
    prepare: [
        OPTION_SPECS.reference,
        OPTION_SPECS.target,
        OPTION_SPECS.out,
        OPTION_SPECS.maxPages,
        OPTION_SPECS.format,
        OPTION_SPECS.dpi,
        OPTION_SPECS.timeoutMs,
        OPTION_SPECS.outputRoot,
        OPTION_SPECS.denyOutsideOutputRoot,
        OPTION_SPECS.reportOut
    ],
    manifest: [OPTION_SPECS.out, OPTION_SPECS.maxPages, OPTION_SPECS.reportOut],
    "manifest inspect": [OPTION_SPECS.reportOut],
    "manifest verify": [OPTION_SPECS.reportOut],
    select: [OPTION_SPECS.selector, OPTION_SPECS.matchesOnly, OPTION_SPECS.noObjectMap, OPTION_SPECS.reportOut],
    plan: [OPTION_SPECS.goal, OPTION_SPECS.out, OPTION_SPECS.reportOut],
    rollback: [OPTION_SPECS.tx, OPTION_SPECS.out, OPTION_SPECS.overwrite],
    lock: [OPTION_SPECS.scope, OPTION_SPECS.name, OPTION_SPECS.owner, OPTION_SPECS.out, OPTION_SPECS.reportOut],
    merge: [OPTION_SPECS.out, OPTION_SPECS.overwrite],
    run: [
        OPTION_SPECS.manifest,
        OPTION_SPECS.logJsonl,
        OPTION_SPECS.summary,
        OPTION_SPECS.outputRoot,
        OPTION_SPECS.expectedArtifacts,
        OPTION_SPECS.timeoutMs,
        OPTION_SPECS.reference,
        OPTION_SPECS.target,
        OPTION_SPECS.out,
        OPTION_SPECS.maxPages,
        OPTION_SPECS.format,
        OPTION_SPECS.dpi,
        OPTION_SPECS.verify,
        OPTION_SPECS.input,
        OPTION_SPECS.goal,
        OPTION_SPECS.visual,
        OPTION_SPECS.denyOutsideOutputRoot,
        OPTION_SPECS.reportOut
    ],
    critique: [OPTION_SPECS.profile, OPTION_SPECS.sheet, OPTION_SPECS.range, OPTION_SPECS.reportOut],
    improve: [OPTION_SPECS.dryRun, OPTION_SPECS.profile, OPTION_SPECS.reportOut],
    benchmark: [OPTION_SPECS.manifest, OPTION_SPECS.reportOut, OPTION_SPECS.timeoutMs],
    "benchmark run": [OPTION_SPECS.manifest, OPTION_SPECS.reportOut, OPTION_SPECS.timeoutMs],
    "benchmark compare": [OPTION_SPECS.reportOut],
    "run office-agent": [
        OPTION_SPECS.input,
        OPTION_SPECS.goal,
        OPTION_SPECS.out,
        OPTION_SPECS.target,
        OPTION_SPECS.manifest,
        OPTION_SPECS.logJsonl,
        OPTION_SPECS.summary,
        OPTION_SPECS.outputRoot,
        OPTION_SPECS.denyOutsideOutputRoot,
        OPTION_SPECS.reportOut
    ],
    asset: [OPTION_SPECS.embedded, OPTION_SPECS.images, OPTION_SPECS.out, OPTION_SPECS.asset, OPTION_SPECS.selector, OPTION_SPECS.overwrite],
    "asset inspect": [OPTION_SPECS.embedded],
    "asset extract": [OPTION_SPECS.images, OPTION_SPECS.out],
    "asset replace": [OPTION_SPECS.asset, OPTION_SPECS.selector, OPTION_SPECS.out, OPTION_SPECS.overwrite],
    "chart render": [OPTION_SPECS.out, OPTION_SPECS.overwrite],
    "diagram render": [OPTION_SPECS.out, OPTION_SPECS.overwrite],
    template: [
        OPTION_SPECS.name,
        OPTION_SPECS.map,
        OPTION_SPECS.data,
        OPTION_SPECS.out,
        OPTION_SPECS.validateOnly,
        OPTION_SPECS.summaryOnly,
        OPTION_SPECS.overwrite,
        OPTION_SPECS.reportOut
    ],
    "template candidates": [OPTION_SPECS.summaryOnly, OPTION_SPECS.reportOut],
    "template create": [OPTION_SPECS.name, OPTION_SPECS.reportOut],
    "template apply-map": [OPTION_SPECS.name, OPTION_SPECS.map, OPTION_SPECS.out, OPTION_SPECS.overwrite, OPTION_SPECS.reportOut],
    "template fill": [OPTION_SPECS.name, OPTION_SPECS.data, OPTION_SPECS.out, OPTION_SPECS.validateOnly, OPTION_SPECS.overwrite, OPTION_SPECS.reportOut],
    design: [OPTION_SPECS.name, OPTION_SPECS.out, OPTION_SPECS.strategy, OPTION_SPECS.data, OPTION_SPECS.overwrite, OPTION_SPECS.reportOut],
    "design capture": [OPTION_SPECS.name, OPTION_SPECS.reportOut],
    "design apply": [OPTION_SPECS.name, OPTION_SPECS.out, OPTION_SPECS.strategy, OPTION_SPECS.overwrite, OPTION_SPECS.reportOut],
    "design update": [OPTION_SPECS.name, OPTION_SPECS.data, OPTION_SPECS.reportOut],
    "design edit": [OPTION_SPECS.name, OPTION_SPECS.data, OPTION_SPECS.reportOut],
    "layout apply": [OPTION_SPECS.out, OPTION_SPECS.overwrite, OPTION_SPECS.reportOut],
    agent: [OPTION_SPECS.target, OPTION_SPECS.name],
    renderer: [OPTION_SPECS.sha256],
    "renderer trust": [OPTION_SPECS.sha256],
    plugin: [OPTION_SPECS.trust],
    "plugin install": [OPTION_SPECS.trust]
};
const COMMAND_ACCEPTED_EXTRA_OPTION_SPECS = {
    improve: [OPTION_SPECS.out]
};
export const COMMAND_METADATA = [
    meta("capabilities", "Show enabled features and agent-visible commands", ["capabilities"]),
    meta("help", "Show human and agent workflow help", ["help", "help workflow", "help error"]),
    meta("config", "Inspect or atomically update scoped user/project configuration", ["config show", "config set"]),
    meta("doctor", "Check the local runtime and project setup", ["doctor"]),
    meta("schema", "List, fetch, validate, and migrate schemas", ["schema list", "schema get", "schema fetch", "schema validate", "schema migrate"]),
    meta("errors", "List and inspect machine-readable error codes", ["errors list", "errors inspect"]),
    core("inspect", "Inspect Office/PDF files and produce trusted metadata", ["inspect"]),
    core("view", "Create SVG/HTML previews and object maps", ["view"]),
    core("edit", "Apply declarative EditOps to existing files", ["edit"]),
    core("render", "Render document IR/specs into Office/PDF files", ["render"]),
    core("scaffold", "Create valid starter IR without an LLM", ["scaffold"]),
    core("export", "Export supported formats with explicit fidelity", ["export"]),
    core("validate", "Validate schemas, structure, and quality gates", ["validate"]),
    core("verify", "Verify openability, native rendering, repair risk, and visual readiness", ["verify"]),
    core("diagnose", "Detect problems in generated or existing files", ["diagnose"]),
    core("repair", "Repair files or produce a repair plan", ["repair"]),
    core("diff", "Compare Office/PDF files semantically and with approximate visual regression scores", ["diff"]),
    core("prepare", "Prepare AI-readable reference/target artifacts for Office editing", ["prepare", "prepare reference"]),
    core("manifest", "Create, inspect, and verify artifact manifests", ["manifest", "manifest inspect", "manifest verify"]),
    core("select", "Resolve safe Office/PDF selectors with confidence", ["select"]),
    core("plan", "Turn goals into deterministic EditOps plans when possible", ["plan"]),
    core("rollback", "Restore artifacts from explicit officegen transaction records", ["rollback"]),
    core("lock", "Create and inspect scoped multi-agent edit locks", ["lock"]),
    core("merge", "Merge supported artifacts with conflict metadata", ["merge"]),
    core("run", "Execute a multi-step workflow or write an office-agent skeleton with run artifacts", ["run", "run prepare-reference", "run office-edit", "run office-agent"]),
    core("critique", "Lint business quality risks in PPTX/DOCX/XLSX files", ["critique"]),
    core("improve", "Produce dry-run improvement suggestions from critique findings", ["improve"]),
    core("benchmark", "Run and compare optional public corpus benchmark reviews", ["benchmark run", "benchmark compare"]),
    core("asset", "Inspect, extract, and replace embedded media", ["asset inspect", "asset extract", "asset replace"]),
    core("chart", "Render safe chart SVG assets", ["chart render"]),
    core("diagram", "Render safe diagram SVG assets", ["diagram render"]),
    optional("template", "Create, inspect, map, validate, and fill templates", [
        "template list",
        "template inspect",
        "template candidates",
        "template create",
        "template apply-map",
        "template validate",
        "template fill"
    ]),
    optional("design", "Capture, inspect, validate, and apply design knowledge", [
        "design list",
        "design inspect",
        "design init",
        "design edit",
        "design update",
        "design validate",
        "design capture",
        "design apply"
    ]),
    optional("layout", "Apply layout constraints", ["layout apply"]),
    optional("agent", "Install or refresh agent adapters", ["agent install", "agent refresh"]),
    optional("mcp", "MCP server", ["mcp serve"], true),
    optional("renderer", "Manage trusted external renderers", ["renderer list", "renderer inspect", "renderer trust", "renderer doctor"], true),
    optional("plugin", "Manage trusted plugins", ["plugin list", "plugin inspect", "plugin install", "plugin trust"], true)
];
export function metadataFor(feature) {
    return COMMAND_METADATA.find((entry) => entry.feature === feature);
}
export function effectiveOptionSpecsFor(commandGroup, subcommand) {
    const key = commandKey(commandGroup, subcommand);
    const direct = COMMAND_OPTION_SPECS[key] ?? COMMAND_OPTION_SPECS[commandGroup];
    if (direct)
        return uniqueOptionSpecs(direct);
    return uniqueOptionSpecs(Object.entries(COMMAND_OPTION_SPECS)
        .filter(([candidate]) => candidate.startsWith(`${commandGroup} `))
        .flatMap(([, specs]) => specs));
}
export function acceptedOptionSpecsFor(commandGroup, subcommand) {
    const key = commandKey(commandGroup, subcommand);
    return uniqueOptionSpecs([
        ...GLOBAL_OPTION_SPECS,
        ...effectiveOptionSpecsFor(commandGroup, subcommand),
        ...(COMMAND_ACCEPTED_EXTRA_OPTION_SPECS[key] ?? COMMAND_ACCEPTED_EXTRA_OPTION_SPECS[commandGroup] ?? [])
    ]);
}
export function effectiveOptionsFor(commandGroup, subcommand) {
    return effectiveOptionSpecsFor(commandGroup, subcommand).map((spec) => spec.flag);
}
export function acceptedOptionsFor(commandGroup, subcommand) {
    return acceptedOptionSpecsFor(commandGroup, subcommand).map((spec) => spec.flag);
}
export function optionSyntax(spec) {
    return spec.value ? `${spec.flag} ${spec.value}` : spec.flag;
}
function commandKey(commandGroup, subcommand) {
    return subcommand ? `${commandGroup} ${subcommand}` : commandGroup;
}
function option(flag, value, description) {
    return { flag, value, description };
}
function uniqueOptionSpecs(specs) {
    const seen = new Set();
    const result = [];
    for (const spec of specs) {
        if (seen.has(spec.flag))
            continue;
        seen.add(spec.flag);
        result.push(spec);
    }
    return result;
}
function meta(feature, description, commands) {
    return entry(feature, description, commands, false, false);
}
function core(feature, description, commands) {
    return entry(feature, description, commands, false, false);
}
function optional(feature, description, commands, externalProcess = false) {
    return entry(feature, description, commands, false, externalProcess);
}
function entry(feature, description, commands, network, externalProcess) {
    return {
        feature,
        moduleId: `officegen.core.${feature}`,
        commandGroup: feature,
        description,
        stability: "stable",
        commands,
        requires: [],
        security: {
            network,
            externalProcess
        }
    };
}
//# sourceMappingURL=metadata.js.map