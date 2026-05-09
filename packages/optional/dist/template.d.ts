import { OptionalContext, ValidationResult } from "./common.js";
import { type DesignContextCandidate, type DesignMapCandidate, type DesignPreviewCandidate, type NamedShapeCandidate, type PptxDesignSignals, type TemplatePlaceholderCandidate, type TemplateSchemaCandidate } from "./design.js";
export type TemplateFieldType = "string" | "number" | "boolean" | "date" | "json";
export type TemplateMappingValue = unknown;
export interface TemplateField {
    name: string;
    type?: TemplateFieldType;
    required?: boolean;
    description?: string;
    defaultValue?: unknown;
}
export interface TemplateDefinition {
    id: string;
    name: string;
    version?: string;
    description?: string;
    tags?: string[];
    fields?: TemplateField[];
    source?: {
        path?: string;
        format?: string;
        sha256?: string;
    };
    mapping?: Record<string, TemplateMappingValue>;
    sourceCapture?: {
        metadata?: PptxDesignSignals["metadata"];
        artifactPaths?: PptxDesignSignals["artifactPaths"];
        placeholderCandidates?: TemplatePlaceholderCandidate[];
        namedShapeCandidates?: NamedShapeCandidate[];
        schemaCandidates?: TemplateSchemaCandidate[];
        templateMapSuggested?: PptxDesignSignals["templateMapSuggested"];
        trust?: PptxDesignSignals["trust"];
    };
    requiredCapabilities?: string[];
    createdAt?: string;
    updatedAt?: string;
    hash?: string;
}
export interface TemplateCandidate {
    template: TemplateDefinition;
    score: number;
    reasons: string[];
    sourceMetadata?: PptxDesignSignals["metadata"];
    previewCandidates?: DesignPreviewCandidate[];
    contextCandidates?: DesignContextCandidate[];
    mapCandidates?: DesignMapCandidate[];
    placeholderCandidates?: TemplatePlaceholderCandidate[];
    namedShapeCandidates?: NamedShapeCandidate[];
    schemaCandidates?: TemplateSchemaCandidate[];
    templateMapSuggested?: PptxDesignSignals["templateMapSuggested"];
    artifactPaths?: PptxDesignSignals["artifactPaths"];
    trust?: PptxDesignSignals["trust"];
    generatedFromSource?: boolean;
}
export interface TemplateCreateOptions extends OptionalContext {
    template: Omit<TemplateDefinition, "createdAt" | "updatedAt" | "hash">;
    sourcePath?: string;
}
export interface TemplateQueryOptions extends OptionalContext {
    query?: string;
    tags?: string[];
    fields?: string[];
    sourcePath?: string;
}
export interface TemplateIdOptions extends OptionalContext {
    id: string;
}
export interface TemplateApplyMapOptions extends TemplateIdOptions {
    mapping: Record<string, TemplateMappingValue>;
    outputPath?: string;
}
export interface TemplateFillOptions extends TemplateIdOptions {
    values: Record<string, unknown>;
    outputPath?: string;
}
export declare function createTemplate(options: TemplateCreateOptions): Promise<TemplateDefinition>;
export declare function listTemplates(options?: OptionalContext): Promise<TemplateDefinition[]>;
export declare function inspectTemplate(options: TemplateIdOptions): Promise<TemplateDefinition>;
export declare function templateCandidates(options?: TemplateQueryOptions): Promise<TemplateCandidate[]>;
export declare function applyTemplateMap(options: TemplateApplyMapOptions): Promise<Record<string, unknown>>;
export declare function fillTemplate(options: TemplateFillOptions): Promise<Record<string, unknown>>;
export declare function validateTemplate(options: TemplateIdOptions): Promise<ValidationResult>;
export declare function validateTemplateDefinition(template: TemplateDefinition): ValidationResult;
