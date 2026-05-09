import { OptionalContext, ValidationResult } from "./common.js";
export interface DesignProfile {
    id: string;
    name: string;
    version?: string;
    tokens: Record<string, unknown>;
    assets?: Record<string, unknown>;
    sourceCapture?: DesignSourceCapture | Record<string, unknown>;
    createdAt?: string;
    updatedAt?: string;
    hash?: string;
}
export interface DesignColorCandidate {
    value: string;
    count: number;
    sources: string[];
}
export interface DesignColorRoleCandidate {
    role: "background" | "text" | "accent";
    value: string;
    confidence: number;
    sources: string[];
    reason: string;
}
export interface DesignTextSizeBucket {
    minPt: number;
    maxPt: number;
    count: number;
}
export interface DesignBounds {
    x: number;
    y: number;
    width: number;
    height: number;
    unit: "ratio";
}
export interface DesignBBoxPattern {
    kind: "title" | "body" | "image";
    count: number;
    slides: number[];
    average: DesignBounds;
}
export interface DesignPreviewCandidate {
    stableObjectId: string;
    slide: number;
    title?: string;
    textSnippet?: string;
    shapeCount: number;
    pictureCount: number;
    slideType?: DesignSlideType;
    densityScore?: number;
    previewPath?: string;
    evidencePath?: string;
}
export interface DesignContextCandidate {
    key: string;
    value: unknown;
    confidence: number;
    source: string;
    untrusted?: true;
}
export interface DesignMapCandidate {
    field: string;
    stableObjectId: string;
    slide: number;
    text: string;
    confidence: number;
    untrusted?: true;
}
export interface TemplatePlaceholderCandidate {
    field: string;
    stableObjectId: string;
    slide: number;
    text?: string;
    name?: string;
    placeholderType?: string;
    objectKind?: "shape" | "picture" | "chart" | "diagram";
    bounds?: DesignBounds;
    confidence: number;
    source: string;
    untrusted?: true;
}
export interface NamedShapeCandidate {
    name: string;
    stableObjectId: string;
    slide: number;
    kind: "shape" | "picture" | "chart" | "diagram";
    text?: string;
    bounds?: DesignBounds;
    confidence: number;
    source: string;
    untrusted?: true;
}
export interface TemplateSchemaCandidate {
    name: string;
    type: "string" | "number" | "boolean" | "date" | "json" | "image" | "chartData" | "table" | "list";
    required: boolean;
    confidence: number;
    reason: string;
}
export interface TemplateMapSuggestion {
    schema: "officegen.template.map@1.2";
    mapping: Record<string, unknown>;
    confidence: number;
    candidateCount: number;
}
export type DesignSlideType = "title" | "title-body" | "section" | "image" | "chart" | "diagram" | "mixed" | "blank";
export interface DesignSlideSignal {
    stableObjectId: string;
    slide: number;
    title?: string;
    slideType: DesignSlideType;
    shapeCount: number;
    pictureCount: number;
    chartCount: number;
    diagramCount: number;
    textObjectCount: number;
    densityScore: number;
    titleBounds?: DesignBounds;
    bodyBounds?: DesignBounds;
    imageBounds: DesignBounds[];
    previewPath?: string;
    evidencePath?: string;
}
export interface PptxDesignArtifactPaths {
    contextPath?: string;
    evidencePath?: string;
    templateMapSuggestedPath?: string;
    schemaCandidatesPath?: string;
    previewPaths: string[];
}
export interface PptxDesignTrustEnvelope {
    trusted: {
        schema: "officegen.design.signals.trusted@1.2";
        format: "pptx";
        sourcePath: string;
        sha256: string;
        byteLength: number;
        generatedAt: string;
        summary: Record<string, unknown>;
    };
    untrusted: {
        schema: "officegen.design.signals.untrusted@1.2";
        slideTitles: string[];
        textSamples: string[];
        shapeNames: string[];
    };
    agentInstruction: string;
}
export interface PptxDesignSignals {
    metadata: {
        format: "pptx";
        slides: number;
        textObjects: number;
        assets: number;
        macros: number;
        byteLength: number;
        title?: string;
        creator?: string;
        created?: string;
        modified?: string;
    };
    colors: DesignColorCandidate[];
    colorRoleCandidates: DesignColorRoleCandidate[];
    textSizeDistribution: DesignTextSizeBucket[];
    bboxPatterns: DesignBBoxPattern[];
    slideSignals: DesignSlideSignal[];
    densityScore: number;
    chartPresence: {
        count: number;
        slides: number[];
    };
    diagramPresence: {
        count: number;
        slides: number[];
    };
    previewCandidates: DesignPreviewCandidate[];
    contextCandidates: DesignContextCandidate[];
    mapCandidates: DesignMapCandidate[];
    placeholderCandidates: TemplatePlaceholderCandidate[];
    namedShapeCandidates: NamedShapeCandidate[];
    schemaCandidates: TemplateSchemaCandidate[];
    templateMapSuggested: TemplateMapSuggestion;
    artifactPaths?: PptxDesignArtifactPaths;
    trust: PptxDesignTrustEnvelope;
}
export interface DesignSourceCapture {
    label: string;
    sourcePath: string;
    sha256: string;
    capturedAt: string;
    metadata?: PptxDesignSignals["metadata"];
    colors?: DesignColorCandidate[];
    colorRoleCandidates?: DesignColorRoleCandidate[];
    textSizeDistribution?: DesignTextSizeBucket[];
    bboxPatterns?: DesignBBoxPattern[];
    slideSignals?: DesignSlideSignal[];
    densityScore?: number;
    chartPresence?: PptxDesignSignals["chartPresence"];
    diagramPresence?: PptxDesignSignals["diagramPresence"];
    previewCandidates?: DesignPreviewCandidate[];
    contextCandidates?: DesignContextCandidate[];
    mapCandidates?: DesignMapCandidate[];
    placeholderCandidates?: TemplatePlaceholderCandidate[];
    namedShapeCandidates?: NamedShapeCandidate[];
    schemaCandidates?: TemplateSchemaCandidate[];
    templateMapSuggested?: TemplateMapSuggestion;
    artifactPaths?: PptxDesignArtifactPaths;
    trust?: PptxDesignTrustEnvelope;
}
export interface DesignInitOptions extends OptionalContext {
    id: string;
    name?: string;
}
export interface DesignInspectOptions extends OptionalContext {
    id: string;
}
export interface DesignUpdateOptions extends DesignInspectOptions {
    patch: Record<string, unknown>;
}
export interface DesignCaptureOptions extends DesignInspectOptions {
    sourcePath: string;
    label?: string;
}
export interface DesignApplyOptions extends DesignInspectOptions {
    targetPath?: string;
    outputPath?: string;
    strategy?: "theme-only" | "inspired" | "faithful";
}
export interface PptxDesignSignalOptions {
    cwd?: string;
    artifactsDir?: string;
}
export declare function initDesign(options: DesignInitOptions): Promise<DesignProfile>;
export declare function listDesigns(options?: OptionalContext): Promise<DesignProfile[]>;
export declare function inspectDesign(options: DesignInspectOptions): Promise<DesignProfile>;
export declare function updateDesign(options: DesignUpdateOptions): Promise<DesignProfile>;
export declare function captureDesign(options: DesignCaptureOptions): Promise<DesignProfile>;
export declare function applyDesign(options: DesignApplyOptions): Promise<Record<string, unknown>>;
export declare function validateDesign(options: DesignInspectOptions): Promise<ValidationResult>;
export declare function validateDesignProfile(design: DesignProfile): ValidationResult;
export declare function capturePptxDesignSignals(sourcePath: string, options?: PptxDesignSignalOptions): Promise<PptxDesignSignals | undefined>;
