import type { CapabilitiesDocument, CapabilityContract, CapabilityFeature, FeatureName, OfficegenConfig } from "./types.js";
export interface CommandSpec {
    feature: FeatureName;
    commands: readonly string[];
}
export declare const COMMAND_SPECS: readonly [CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec, CommandSpec];
export declare const commandMap: Record<FeatureName, string[]>;
export declare const RELEASE_CAPABILITY_CONTRACT_VERSION = "3.0.0-contract";
export declare const formatCapabilities: {
    readonly pptx: {
        readonly text: true;
        readonly lists: true;
        readonly tables: "scoped cell text edits";
        readonly images: "replace existing image shapes";
        readonly charts: "single-series chart data updates only";
        readonly smartArt: "unsupported";
        readonly designer: "limited; not a full PowerPoint designer";
    };
    readonly docx: {
        readonly text: "scoped paragraph/run replacement";
        readonly tables: "scoped table cell text edits";
        readonly comments: "add comment ranges where selectors resolve";
        readonly redlines: "tracked insert/delete/replace only";
        readonly styles: "scoped style application";
        readonly fullFidelityEditing: "limited; not a full Word/DTP/legal contract engine";
    };
    readonly xlsx: {
        readonly cells: true;
        readonly formulas: "guarded setFormula; no internal calculation engine";
        readonly tables: "scoped write/resize operations";
        readonly charts: "single-series chart data updates only";
        readonly pivots: "refresh flags only; no pivot layout/field/value editing";
        readonly slicers: "selection updates only; no slicer authoring/styling engine";
    };
    readonly pdf: {
        readonly generation: true;
        readonly inspect: "best-effort text preview plus page previews";
        readonly overlays: "text overlays and annotations only";
        readonly redaction: "unsupported; overlays do not physically remove underlying content";
        readonly contentRewrite: "unsupported";
    };
};
export declare const featureContracts: CapabilityContract[];
export declare const runtimeProfiles: {
    readonly "current-limited-v3.1": {
        readonly id: "current-limited-v3.1";
        readonly role: "current";
        readonly runtime: "v3.1 portable CLI";
        readonly summary: "Current v3.1 support is a limited, evidence-backed runtime profile. It includes runtime v2 projections and scoped Office/PDF operations, but it does not include perfect-runtime target capabilities.";
        readonly capabilities: [{
            readonly id: "runtime-v2-projections";
            readonly area: "Runtime v2 projections";
            readonly support: "supported";
            readonly summary: "Runtime v2 envelope, selector, object graph, edit/patch plan, and verify projections are current supported contract surfaces.";
            readonly evidence: ["officegen.envelope@2", "officegen.selectorResolution@2", "officegen.objectGraph@2", "officegen.editPlan@2", "officegen.patchPlan@2", "officegen.verify@2"];
            readonly gaps: [];
        }, {
            readonly id: "office-agent-skeleton-evidence";
            readonly area: "office-agent runtime skeleton";
            readonly support: "supported";
            readonly summary: "officegen run office-agent writes the 13-phase runtime-v2 skeleton and evidence manifest for release review.";
            readonly evidence: ["officegen.office-agent.manifest@3.1", "officegen.office-agent.workflow@3.1", "officegen.office-agent.result@3.1"];
            readonly gaps: ["Does not execute complete autonomous repair or prove final document readiness by itself."];
        }, {
            readonly id: "scoped-office-pdf-editops";
            readonly area: "Scoped Office/PDF EditOps";
            readonly support: "limited";
            readonly summary: "Current edit operations are scoped, auditable mutations with dry-run and verification follow-up rather than full application-level engines.";
            readonly evidence: ["officegen.edit.ops@1.2", "featureContracts", "formatCapabilities"];
            readonly gaps: ["Multi-series charts, full pivot/slicer authoring, complete DOCX layout editing, and PDF content rewriting remain outside current support."];
        }, {
            readonly id: "smartart-editing";
            readonly area: "PPTX SmartArt editing";
            readonly support: "unsupported";
            readonly summary: "SmartArt creation, node editing, layout changes, and conversion are unsupported in current-limited-v3.1.";
            readonly evidence: ["unsupportedNow", "PPTX SmartArt feature contract"];
            readonly gaps: ["Target profile requires SmartArt edit semantics and verification evidence."];
        }, {
            readonly id: "pdf-true-redaction";
            readonly area: "PDF true redaction";
            readonly support: "unsupported";
            readonly summary: "PDF overlays and annotations are supported as additive marks only; physical redaction and content rewrite are unsupported in current-limited-v3.1.";
            readonly evidence: ["unsupportedNow", "PDF editing and redaction feature contract", "pdf.textOverlay x-officegen-support"];
            readonly gaps: ["Target profile requires removal or cryptographic scrubbing evidence for underlying PDF content."];
        }, {
            readonly id: "complete-autonomous-repair";
            readonly area: "Complete autonomous repair";
            readonly support: "unsupported";
            readonly summary: "Current office-agent evidence is skeleton-only and does not claim complete autonomous repair.";
            readonly evidence: ["office-agent caveats", "runtimeProjection=runtime-v2"];
            readonly gaps: ["Target profile requires executed repair, verify, diff, and final readiness evidence."];
        }];
    };
    readonly "perfect-runtime-target": {
        readonly id: "perfect-runtime-target";
        readonly role: "target";
        readonly runtime: "perfect runtime spec";
        readonly summary: "Perfect runtime target describes desired capability, not current v3.1 support. Target-only entries must not be advertised as supported by current-limited-v3.1.";
        readonly capabilities: [{
            readonly id: "runtime-v2-projections";
            readonly area: "Runtime v2 projections";
            readonly support: "supported";
            readonly summary: "Runtime v2 projections remain part of the target and are already supported by the current profile.";
            readonly evidence: ["current-limited-v3.1/runtime-v2-projections"];
            readonly gaps: [];
        }, {
            readonly id: "smartart-editing";
            readonly area: "PPTX SmartArt editing";
            readonly support: "target-only";
            readonly summary: "Target profile includes SmartArt authoring and editing, but current-limited-v3.1 marks it unsupported.";
            readonly evidence: [];
            readonly gaps: ["Implement SmartArt edit ops, safety checks, and verification evidence before moving to current support."];
        }, {
            readonly id: "pdf-true-redaction";
            readonly area: "PDF true redaction";
            readonly support: "target-only";
            readonly summary: "Target profile includes true PDF redaction/content removal, but current-limited-v3.1 marks it unsupported.";
            readonly evidence: [];
            readonly gaps: ["Implement physical content removal/scrubbing and negative extraction tests before moving to current support."];
        }, {
            readonly id: "complete-autonomous-repair";
            readonly area: "Complete autonomous repair";
            readonly support: "target-only";
            readonly summary: "Target profile includes end-to-end autonomous repair through final verify/diff readiness, not just skeleton evidence.";
            readonly evidence: [];
            readonly gaps: ["Attach executed edit, repair, verify, diff, and readiness artifacts."];
        }, {
            readonly id: "full-fidelity-office-engines";
            readonly area: "Full-fidelity Office application engines";
            readonly support: "target-only";
            readonly summary: "Target profile includes richer Office behaviors such as multi-series chart editing, pivot/slicer authoring, and full DOCX layout/legal editing.";
            readonly evidence: [];
            readonly gaps: ["Promote each engine only after operation schemas, implementation, fixtures, and release evidence exist."];
        }];
    };
};
export declare const specProfile: {
    readonly currentProfileId: "current-limited-v3.1";
    readonly targetProfileId: "perfect-runtime-target";
    readonly runtimeProjection: "runtime-v2";
    readonly truthfulnessPolicy: "Agents must treat current-limited-v3.1 as the only supported runtime profile. perfect-runtime-target is aspirational until a capability has current evidence.";
    readonly currentEvidence: ["runtimeProfiles.current-limited-v3.1.capabilities", "featureContracts", "formatCapabilities", "unsupportedNow", "goal/v3.1.0-evidence-matrix.json"];
    readonly targetGapIds: ["smartart-editing", "pdf-true-redaction", "complete-autonomous-repair", "full-fidelity-office-engines"];
};
export declare const knownLimitations: string[];
export declare function buildFeatureRegistry(config: OfficegenConfig): CapabilityFeature[];
export declare function computeCapabilitiesHash(config: OfficegenConfig, cliVersion?: string): string;
export declare function getVisibleCommands(config: OfficegenConfig, agent?: boolean): string[];
export declare function getCapabilities(config: OfficegenConfig, options?: {
    agent?: boolean;
    runInstructionsPath?: string;
}): CapabilitiesDocument;
export declare function isFeatureAvailable(config: OfficegenConfig, feature: FeatureName, agent?: boolean): boolean;
