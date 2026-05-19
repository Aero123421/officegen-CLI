import { createHash } from "node:crypto";
import { FEATURE_NAMES } from "./config.js";
import { SCHEMA_REGISTRY_VERSION, OFFICEGEN_CLI_VERSION } from "./types.js";
export const COMMAND_SPECS = [
    spec("capabilities", ["capabilities"]),
    spec("help", ["help", "help workflow", "help error"]),
    spec("config", ["config show", "config set"]),
    spec("doctor", ["doctor"]),
    spec("inspect", ["inspect"]),
    spec("view", ["view"]),
    spec("edit", ["edit"]),
    spec("render", ["render"]),
    spec("scaffold", ["scaffold"]),
    spec("export", ["export"]),
    spec("validate", ["validate"]),
    spec("verify", ["verify"]),
    spec("diagnose", ["diagnose"]),
    spec("repair", ["repair"]),
    spec("diff", ["diff"]),
    spec("prepare", ["prepare", "prepare reference"]),
    spec("manifest", ["manifest", "manifest inspect", "manifest verify"]),
    spec("select", ["select"]),
    spec("plan", ["plan"]),
    spec("rollback", ["rollback"]),
    spec("lock", ["lock"]),
    spec("merge", ["merge"]),
    spec("run", ["run", "run prepare-reference", "run office-edit"]),
    spec("critique", ["critique"]),
    spec("improve", ["improve"]),
    spec("benchmark", ["benchmark run", "benchmark compare"]),
    spec("asset", ["asset inspect", "asset extract", "asset replace"]),
    spec("chart", ["chart render"]),
    spec("diagram", ["diagram render"]),
    spec("schema", ["schema list", "schema get", "schema fetch", "schema validate", "schema migrate"]),
    spec("errors", ["errors list", "errors inspect"]),
    spec("template", ["template list", "template inspect", "template candidates", "template create", "template apply-map", "template validate", "template fill"]),
    spec("design", ["design list", "design inspect", "design init", "design edit", "design update", "design validate", "design capture", "design apply"]),
    spec("layout", ["layout apply"]),
    spec("agent", ["agent install", "agent refresh"]),
    spec("mcp", ["mcp serve"]),
    spec("renderer", ["renderer list", "renderer inspect", "renderer trust", "renderer doctor"]),
    spec("plugin", ["plugin list", "plugin inspect", "plugin install", "plugin trust"])
];
export const commandMap = Object.fromEntries(COMMAND_SPECS.map((entry) => [entry.feature, [...entry.commands]]));
export const RELEASE_CAPABILITY_CONTRACT_VERSION = "3.0.0-contract";
export const formatCapabilities = {
    pptx: {
        text: true,
        lists: true,
        tables: "scoped cell text edits",
        images: "replace existing image shapes",
        charts: "single-series chart data updates only",
        smartArt: "unsupported",
        designer: "limited; not a full PowerPoint designer"
    },
    docx: {
        text: "scoped paragraph/run replacement",
        tables: "scoped table cell text edits",
        comments: "add comment ranges where selectors resolve",
        redlines: "tracked insert/delete/replace only",
        styles: "scoped style application",
        fullFidelityEditing: "limited; not a full Word/DTP/legal contract engine"
    },
    xlsx: {
        cells: true,
        formulas: "guarded setFormula; no internal calculation engine",
        tables: "scoped write/resize operations",
        charts: "single-series chart data updates only",
        pivots: "refresh flags only; no pivot layout/field/value editing",
        slicers: "selection updates only; no slicer authoring/styling engine"
    },
    pdf: {
        generation: true,
        inspect: "best-effort text preview plus page previews",
        overlays: "text overlays and annotations only",
        redaction: "unsupported; overlays do not physically remove underlying content",
        contentRewrite: "unsupported"
    }
};
export const featureContracts = [
    {
        area: "PPTX SmartArt",
        formats: ["pptx"],
        support: "unsupported",
        summary: "SmartArt objects may be inspected as package content, but SmartArt creation or full editing is not implemented.",
        limitations: ["No SmartArt authoring, layout changes, node editing, or conversion engine is exposed as an EditOp."]
    },
    {
        area: "PPTX/XLSX charts",
        formats: ["pptx", "xlsx"],
        support: "limited",
        summary: "Chart data EditOps support single-series category/value updates only.",
        limitations: ["No multi-series chart engine.", "No secondary-axis updates.", "No combo-chart authoring or per-series chart-type updates."]
    },
    {
        area: "XLSX pivots and slicers",
        formats: ["xlsx"],
        support: "limited",
        summary: "Pivot support is refresh-flag oriented and slicer support is selection-only.",
        limitations: ["No pivot field/layout/value editing.", "No slicer creation, caption editing, style editing, or cache rebuild engine."]
    },
    {
        area: "PDF editing and redaction",
        formats: ["pdf"],
        support: "overlay-only",
        summary: "PDF mutation supports text overlays and annotations; physical redaction is unsupported.",
        limitations: ["Overlay text or rectangles are not redaction.", "Underlying PDF content is not removed, rewritten, or cryptographically scrubbed."]
    },
    {
        area: "Full-fidelity Office/PDF editing",
        formats: ["pptx", "docx", "xlsx", "pdf"],
        support: "limited",
        summary: "Officegen exposes scoped, auditable operations rather than complete Office application editing engines.",
        limitations: ["DOCX is not a full DTP/legal-contract authoring engine.", "XLSX does not recalculate or fully author chart/pivot/slicer models internally.", "PDF content rewriting is not implemented."]
    }
];
export const knownLimitations = featureContracts.flatMap((contract) => contract.limitations.map((limitation) => `${contract.area}: ${limitation}`));
function spec(feature, commands) {
    return { feature, commands };
}
export function buildFeatureRegistry(config) {
    return FEATURE_NAMES.map((name) => ({
        name,
        enabled: config.features[name].enabled,
        visibleInHelp: config.features[name].visibleInHelp,
        visibleToAgents: config.features[name].visibleToAgents,
        commands: [...commandMap[name]],
        requires: []
    }));
}
function stableStringify(value) {
    if (value === null || typeof value !== "object")
        return JSON.stringify(value);
    if (Array.isArray(value))
        return `[${value.map(stableStringify).join(",")}]`;
    const record = value;
    return `{${Object.keys(record)
        .sort()
        .map((key) => `${JSON.stringify(key)}:${stableStringify(record[key])}`)
        .join(",")}}`;
}
export function computeCapabilitiesHash(config, cliVersion = OFFICEGEN_CLI_VERSION) {
    const payload = {
        cliVersion,
        profile: config.profile,
        features: config.features,
        security: config.security,
        agent: config.agent,
        commandMap,
        releaseCapabilityContractVersion: RELEASE_CAPABILITY_CONTRACT_VERSION,
        formatCapabilities,
        featureContracts,
        schemaRegistryVersion: SCHEMA_REGISTRY_VERSION
    };
    return `sha256:${createHash("sha256").update(stableStringify(payload)).digest("hex")}`;
}
export function getVisibleCommands(config, agent = false) {
    return buildFeatureRegistry(config)
        .filter((feature) => feature.enabled && (agent ? feature.visibleToAgents : feature.visibleInHelp))
        .flatMap((feature) => feature.commands);
}
export function getCapabilities(config, options = {}) {
    const registry = buildFeatureRegistry(config);
    const agent = options.agent ?? false;
    const visibleCommands = getVisibleCommands(config, agent);
    return {
        schema: "officegen.capabilities@1.2",
        ok: true,
        profile: config.profile,
        capabilitiesHash: computeCapabilitiesHash(config),
        visibleCommands,
        hiddenFromAgents: registry
            .filter((feature) => feature.enabled && feature.visibleInHelp && !feature.visibleToAgents)
            .map((feature) => feature.name),
        disabled: registry.filter((feature) => !feature.enabled).map((feature) => feature.name),
        agentInstructionsPath: options.runInstructionsPath ?? ".officegen/runs/current/agent-instructions.md",
        jsonBudgetBytes: config.agent.defaultJsonBudgetBytes,
        featureContracts,
        formatCapabilities,
        knownLimitations,
        unsupportedNow: [
            "SmartArt creation and full SmartArt editing are unsupported.",
            "Multi-series, secondary-axis, and combo chart editing are unsupported; chart data ops are single-series only.",
            "PDF physical redaction and PDF content rewriting are unsupported; PDF mutation is overlay/annotation only.",
            "Complete DOCX, XLSX, PPTX, and PDF application-level editing engines are outside the portable CLI contract."
        ],
        nextSuggestedCommands: visibleCommands.some((command) => command === "schema" || command.startsWith("schema "))
            ? ["officegen help workflow edit-existing --agent --json", "officegen schema list --agent --json"]
            : ["officegen capabilities --agent --json"]
    };
}
export function isFeatureAvailable(config, feature, agent = false) {
    const visibility = config.features[feature];
    if (!visibility.enabled)
        return false;
    return agent ? visibility.visibleToAgents : visibility.visibleInHelp;
}
//# sourceMappingURL=capabilities.js.map