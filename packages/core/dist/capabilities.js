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
    spec("run", ["run", "run prepare-reference", "run office-edit", "run office-agent"]),
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
        area: "Configuration mutation",
        formats: ["json"],
        support: "limited",
        summary: "config set persists an allowlisted leaf value to either project or user config with an atomic JSON write.",
        limitations: ["Arbitrary config object rewrites are not exposed.", "Values are validated against the portable CLI config contract before writing."]
    },
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
export const runtimeProfiles = {
    "current-limited-v3.1": {
        id: "current-limited-v3.1",
        role: "current",
        runtime: "v3.1 portable CLI",
        summary: "Current v3.1 support is a limited, evidence-backed runtime profile. It includes runtime v2 projections and scoped Office/PDF operations, but it does not include perfect-runtime target capabilities.",
        capabilities: [
            {
                id: "runtime-v2-projections",
                area: "Runtime v2 projections",
                support: "supported",
                summary: "Runtime v2 envelope, selector, object graph, edit/patch plan, and verify projections are current supported contract surfaces.",
                evidence: [
                    "officegen.envelope@2",
                    "officegen.selectorResolution@2",
                    "officegen.objectGraph@2",
                    "officegen.editPlan@2",
                    "officegen.patchPlan@2",
                    "officegen.verify@2"
                ],
                gaps: []
            },
            {
                id: "office-agent-skeleton-evidence",
                area: "office-agent runtime skeleton",
                support: "supported",
                summary: "officegen run office-agent writes the 13-phase runtime-v2 skeleton and evidence manifest for release review.",
                evidence: ["officegen.office-agent.manifest@3.1", "officegen.office-agent.workflow@3.1", "officegen.office-agent.result@3.1"],
                gaps: ["Does not execute complete autonomous repair or prove final document readiness by itself."]
            },
            {
                id: "scoped-office-pdf-editops",
                area: "Scoped Office/PDF EditOps",
                support: "limited",
                summary: "Current edit operations are scoped, auditable mutations with dry-run and verification follow-up rather than full application-level engines.",
                evidence: ["officegen.edit.ops@1.2", "featureContracts", "formatCapabilities"],
                gaps: ["Multi-series charts, full pivot/slicer authoring, complete DOCX layout editing, and PDF content rewriting remain outside current support."]
            },
            {
                id: "smartart-editing",
                area: "PPTX SmartArt editing",
                support: "unsupported",
                summary: "SmartArt creation, node editing, layout changes, and conversion are unsupported in current-limited-v3.1.",
                evidence: ["unsupportedNow", "PPTX SmartArt feature contract"],
                gaps: ["Target profile requires SmartArt edit semantics and verification evidence."]
            },
            {
                id: "pdf-true-redaction",
                area: "PDF true redaction",
                support: "unsupported",
                summary: "PDF overlays and annotations are supported as additive marks only; physical redaction and content rewrite are unsupported in current-limited-v3.1.",
                evidence: ["unsupportedNow", "PDF editing and redaction feature contract", "pdf.textOverlay x-officegen-support"],
                gaps: ["Target profile requires removal or cryptographic scrubbing evidence for underlying PDF content."]
            },
            {
                id: "complete-autonomous-repair",
                area: "Complete autonomous repair",
                support: "unsupported",
                summary: "Current office-agent evidence is skeleton-only and does not claim complete autonomous repair.",
                evidence: ["office-agent caveats", "runtimeProjection=runtime-v2"],
                gaps: ["Target profile requires executed repair, verify, diff, and final readiness evidence."]
            }
        ]
    },
    "perfect-runtime-target": {
        id: "perfect-runtime-target",
        role: "target",
        runtime: "perfect runtime spec",
        summary: "Perfect runtime target describes desired capability, not current v3.1 support. Target-only entries must not be advertised as supported by current-limited-v3.1.",
        capabilities: [
            {
                id: "runtime-v2-projections",
                area: "Runtime v2 projections",
                support: "supported",
                summary: "Runtime v2 projections remain part of the target and are already supported by the current profile.",
                evidence: ["current-limited-v3.1/runtime-v2-projections"],
                gaps: []
            },
            {
                id: "smartart-editing",
                area: "PPTX SmartArt editing",
                support: "target-only",
                summary: "Target profile includes SmartArt authoring and editing, but current-limited-v3.1 marks it unsupported.",
                evidence: [],
                gaps: ["Implement SmartArt edit ops, safety checks, and verification evidence before moving to current support."]
            },
            {
                id: "pdf-true-redaction",
                area: "PDF true redaction",
                support: "target-only",
                summary: "Target profile includes true PDF redaction/content removal, but current-limited-v3.1 marks it unsupported.",
                evidence: [],
                gaps: ["Implement physical content removal/scrubbing and negative extraction tests before moving to current support."]
            },
            {
                id: "complete-autonomous-repair",
                area: "Complete autonomous repair",
                support: "target-only",
                summary: "Target profile includes end-to-end autonomous repair through final verify/diff readiness, not just skeleton evidence.",
                evidence: [],
                gaps: ["Attach executed edit, repair, verify, diff, and readiness artifacts."]
            },
            {
                id: "full-fidelity-office-engines",
                area: "Full-fidelity Office application engines",
                support: "target-only",
                summary: "Target profile includes richer Office behaviors such as multi-series chart editing, pivot/slicer authoring, and full DOCX layout/legal editing.",
                evidence: [],
                gaps: ["Promote each engine only after operation schemas, implementation, fixtures, and release evidence exist."]
            }
        ]
    }
};
export const specProfile = {
    currentProfileId: "current-limited-v3.1",
    targetProfileId: "perfect-runtime-target",
    runtimeProjection: "runtime-v2",
    truthfulnessPolicy: "Agents must treat current-limited-v3.1 as the only supported runtime profile. perfect-runtime-target is aspirational until a capability has current evidence.",
    currentEvidence: [
        "runtimeProfiles.current-limited-v3.1.capabilities",
        "featureContracts",
        "formatCapabilities",
        "unsupportedNow",
        "goal/v3.1.0-evidence-matrix.json"
    ],
    targetGapIds: ["smartart-editing", "pdf-true-redaction", "complete-autonomous-repair", "full-fidelity-office-engines"]
};
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
        runtimeProfiles,
        specProfile,
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
        runtimeProfiles,
        specProfile,
        knownLimitations,
        unsupportedNow: [
            "SmartArt creation and full SmartArt editing are unsupported.",
            "Multi-series, secondary-axis, and combo chart editing are unsupported; chart data ops are single-series only.",
            "PDF physical redaction and PDF content rewriting are unsupported; PDF mutation is overlay/annotation only.",
            "Direct edit commands write artifacts but do not prove final readiness; run verify after mutation before release.",
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