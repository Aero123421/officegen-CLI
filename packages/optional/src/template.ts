import path from "node:path";

import {
  OptionalContext,
  ValidationResult,
  featureRoot,
  hashFile,
  listJsonFiles,
  nowIso,
  readJsonFile,
  requireFeature,
  sha256Json,
  slugify,
  validation,
  writeJsonFile
} from "./common.js";
import {
  capturePptxDesignSignals,
  type DesignContextCandidate,
  type DesignMapCandidate,
  type DesignPreviewCandidate,
  type NamedShapeCandidate,
  type PptxDesignSignals,
  type TemplatePlaceholderCandidate,
  type TemplateSchemaCandidate
} from "./design.js";

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

export async function createTemplate(options: TemplateCreateOptions): Promise<TemplateDefinition> {
  requireFeature(options, "template", "template create");
  const now = nowIso();
  const id = slugify(options.template.id);
  const sourcePath = options.sourcePath ?? options.template.source?.path;
  const resolvedSourcePath = sourcePath ? path.resolve(options.cwd ?? process.cwd(), sourcePath) : undefined;
  const sourceSignals = resolvedSourcePath
    ? await capturePptxDesignSignals(resolvedSourcePath, {
        cwd: options.cwd,
        artifactsDir: path.join(featureRoot(options, "template"), "captures", id, slugify(path.basename(resolvedSourcePath, path.extname(resolvedSourcePath))))
      })
    : undefined;
  const inferredFields = sourceSignals?.schemaCandidates.map((field) => ({
    name: field.name,
    type: field.type,
    required: field.required,
    description: `${field.reason}; confidence ${field.confidence}`
  }));
  const template: TemplateDefinition = {
    ...options.template,
    id,
    fields: options.template.fields && options.template.fields.length > 0 ? options.template.fields : inferredFields ?? options.template.fields,
    mapping: options.template.mapping ?? sourceSignals?.templateMapSuggested.mapping,
    source: resolvedSourcePath
      ? {
          path: resolvedSourcePath,
          format: path.extname(resolvedSourcePath).replace(/^\./, "").toLowerCase() || options.template.source?.format,
          sha256: await hashFile(resolvedSourcePath)
        }
      : options.template.source,
    sourceCapture: sourceSignals
      ? {
          metadata: sourceSignals.metadata,
          artifactPaths: sourceSignals.artifactPaths,
          placeholderCandidates: sourceSignals.placeholderCandidates,
          namedShapeCandidates: sourceSignals.namedShapeCandidates,
          schemaCandidates: sourceSignals.schemaCandidates,
          templateMapSuggested: sourceSignals.templateMapSuggested,
          trust: sourceSignals.trust
        }
      : options.template.sourceCapture,
    createdAt: now,
    updatedAt: now
  };
  template.hash = sha256Json({ ...template, hash: undefined });
  assertTemplateValid(template);

  await writeJsonFile(templatePath(options, template.id), template);
  return template;
}

export async function listTemplates(options: OptionalContext = {}): Promise<TemplateDefinition[]> {
  requireFeature(options, "template", "template list");
  const files = await listJsonFiles(featureRoot(options, "template"));
  const templates = await Promise.all(files.map((file) => readJsonFile<TemplateDefinition>(file)));
  return templates.sort((left, right) => left.id.localeCompare(right.id));
}

export async function inspectTemplate(options: TemplateIdOptions): Promise<TemplateDefinition> {
  requireFeature(options, "template", "template inspect");
  return readJsonFile<TemplateDefinition>(templatePath(options, options.id));
}

export async function templateCandidates(options: TemplateQueryOptions = {}): Promise<TemplateCandidate[]> {
  const templates = await listTemplates(options);
  const query = options.query?.trim().toLowerCase();
  const tags = new Set((options.tags ?? []).map((tag) => tag.toLowerCase()));
  const fields = new Set((options.fields ?? []).map((field) => field.toLowerCase()));
  const resolvedSourcePath = options.sourcePath ? path.resolve(options.cwd ?? process.cwd(), options.sourcePath) : undefined;
  const sourceSignals = resolvedSourcePath
    ? await capturePptxDesignSignals(resolvedSourcePath, {
        cwd: options.cwd,
        artifactsDir: path.join(
          featureRoot(options, "template"),
          "candidates",
          slugify(path.basename(resolvedSourcePath, path.extname(resolvedSourcePath)))
        )
      })
    : undefined;

  const registryCandidates = templates
    .map((template) => {
      const reasons: string[] = [];
      let score = 0;
      const mapCandidates = sourceSignals ? matchTemplateMapCandidates(template, sourceSignals) : undefined;
      const searchable = [template.id, template.name, template.description, ...(template.tags ?? [])]
        .filter(Boolean)
        .join(" ")
        .toLowerCase();

      if (query && searchable.includes(query)) {
        score += 5;
        reasons.push("query");
      }

      const matchedTags = (template.tags ?? []).filter((tag) => tags.has(tag.toLowerCase()));
      if (matchedTags.length > 0) {
        score += matchedTags.length * 3;
        reasons.push(`tags:${matchedTags.join(",")}`);
      }

      const matchedFields = (template.fields ?? []).filter((field) => fields.has(field.name.toLowerCase()));
      if (matchedFields.length > 0) {
        score += matchedFields.length * 2;
        reasons.push(`fields:${matchedFields.map((field) => field.name).join(",")}`);
      }

      if (sourceSignals) {
        score += 1;
        reasons.push("source:pptx");
        if (mapCandidates && mapCandidates.length > 0) {
          score += mapCandidates.length;
          reasons.push(`map:${mapCandidates.map((candidate) => candidate.field).join(",")}`);
        }
      }

      if (!query && tags.size === 0 && fields.size === 0 && !sourceSignals) {
        score = 1;
        reasons.push("available");
      }

      return {
        template,
        score,
        reasons,
        ...(sourceSignals
          ? {
              sourceMetadata: sourceSignals.metadata,
              previewCandidates: sourceSignals.previewCandidates,
              contextCandidates: sourceSignals.contextCandidates,
              mapCandidates,
              placeholderCandidates: sourceSignals.placeholderCandidates,
              namedShapeCandidates: sourceSignals.namedShapeCandidates,
              schemaCandidates: sourceSignals.schemaCandidates,
              templateMapSuggested: sourceSignals.templateMapSuggested,
              artifactPaths: sourceSignals.artifactPaths,
              trust: sourceSignals.trust
            }
          : {})
      };
    })
    .filter((candidate) => candidate.score > 0);

  const sourceDerivedCandidate = sourceSignals && resolvedSourcePath ? makeSourceDerivedTemplateCandidate(resolvedSourcePath, sourceSignals) : undefined;

  return [...(sourceDerivedCandidate ? [sourceDerivedCandidate] : []), ...registryCandidates]
    .sort((left, right) => right.score - left.score || left.template.id.localeCompare(right.template.id));
}

export async function applyTemplateMap(options: TemplateApplyMapOptions): Promise<Record<string, unknown>> {
  requireFeature(options, "template", "template apply-map");
  const template = await inspectTemplate(options);
  const plan = {
    kind: "officegen.template.apply-map",
    planOnly: true,
    generatedAt: nowIso(),
    templateId: template.id,
    templateHash: template.hash,
    mapping: options.mapping,
    note: "Direct Office file edits are delegated to @officegen/formats."
  };

  await writeJsonFile(
    options.outputPath ?? path.join(featureRoot(options, "template"), "runs", `${slugify(template.id)}.map.json`),
    plan
  );
  return plan;
}

export async function fillTemplate(options: TemplateFillOptions): Promise<Record<string, unknown>> {
  requireFeature(options, "template", "template fill");
  const template = await inspectTemplate(options);
  const result = validateTemplateValues(template, options.values);
  if (!result.ok) {
    throw new Error(`Template fill failed validation: ${result.errors.join("; ")}`);
  }

  const filled = {
    kind: "officegen.template.fill",
    planOnly: true,
    generatedAt: nowIso(),
    templateId: template.id,
    templateHash: template.hash,
    values: options.values,
    note: "This JSON is a fill plan for @officegen/formats."
  };

  await writeJsonFile(
    options.outputPath ?? path.join(featureRoot(options, "template"), "runs", `${slugify(template.id)}.fill.json`),
    filled
  );
  return filled;
}

export async function validateTemplate(options: TemplateIdOptions): Promise<ValidationResult> {
  requireFeature(options, "template", "template validate");
  return validateTemplateDefinition(await inspectTemplate(options));
}

export function validateTemplateDefinition(template: TemplateDefinition): ValidationResult {
  const errors: string[] = [];
  const warnings: string[] = [];

  if (!template.id?.trim()) errors.push("id is required");
  if (!template.name?.trim()) errors.push("name is required");
  for (const field of template.fields ?? []) {
    if (!field.name?.trim()) errors.push("field.name is required");
    if (field.type && !["string", "number", "boolean", "date", "json"].includes(field.type)) {
      errors.push(`unsupported field type: ${field.type}`);
    }
  }
  if (!template.fields || template.fields.length === 0) {
    warnings.push("template has no fields");
  }

  return validation(errors.length === 0, errors, warnings);
}

function validateTemplateValues(template: TemplateDefinition, values: Record<string, unknown>): ValidationResult {
  const errors: string[] = [];
  for (const field of template.fields ?? []) {
    const value = values[field.name] ?? field.defaultValue;
    if (field.required && value === undefined) {
      errors.push(`missing required field: ${field.name}`);
      continue;
    }
    if (value !== undefined && field.type && field.type !== "json" && !matchesFieldType(value, field.type)) {
      errors.push(`field ${field.name} expected ${field.type}`);
    }
  }

  return validation(errors.length === 0, errors);
}

function matchesFieldType(value: unknown, type: TemplateFieldType): boolean {
  if (type === "date") return typeof value === "string" && !Number.isNaN(Date.parse(value));
  return typeof value === type;
}

function assertTemplateValid(template: TemplateDefinition): void {
  const result = validateTemplateDefinition(template);
  if (!result.ok) {
    throw new Error(`Invalid template: ${result.errors.join("; ")}`);
  }
}

function matchTemplateMapCandidates(
  template: TemplateDefinition,
  sourceSignals: PptxDesignSignals
): DesignMapCandidate[] {
  const templateFields = new Set((template.fields ?? []).map((field) => normalizeField(field.name)).filter(Boolean));
  if (templateFields.size === 0) return sourceSignals.mapCandidates.slice(0, 8);

  const matched: DesignMapCandidate[] = [];
  for (const candidate of sourceSignals.mapCandidates) {
    const normalizedCandidate = normalizeField(candidate.field);
    if (!normalizedCandidate) continue;
    if (
      templateFields.has(normalizedCandidate) ||
      [...templateFields].some((field) => field.includes(normalizedCandidate) || normalizedCandidate.includes(field))
    ) {
      matched.push(candidate);
    }
    if (matched.length >= 8) break;
  }
  return matched;
}

function makeSourceDerivedTemplateCandidate(sourcePath: string, sourceSignals: PptxDesignSignals): TemplateCandidate {
  const name = path.basename(sourcePath, path.extname(sourcePath));
  const fields = sourceSignals.schemaCandidates.map((field) => ({
    name: field.name,
    type: field.type,
    required: field.required,
    description: `${field.reason}; confidence ${field.confidence}`
  }));
  const template: TemplateDefinition = {
    id: slugify(`suggested-${name}`),
    name: `Suggested template from ${name}`,
    version: "0.1.0",
    description: "Generated from sourcePath PPTX analysis. Review untrusted text before use.",
    tags: ["pptx", "source-derived"],
    fields,
    source: {
      path: sourcePath,
      format: "pptx",
      sha256: sourceSignals.trust.trusted.sha256
    },
    mapping: sourceSignals.templateMapSuggested.mapping,
    sourceCapture: {
      metadata: sourceSignals.metadata,
      artifactPaths: sourceSignals.artifactPaths,
      placeholderCandidates: sourceSignals.placeholderCandidates,
      namedShapeCandidates: sourceSignals.namedShapeCandidates,
      schemaCandidates: sourceSignals.schemaCandidates,
      templateMapSuggested: sourceSignals.templateMapSuggested,
      trust: sourceSignals.trust
    },
    createdAt: sourceSignals.trust.trusted.generatedAt,
    updatedAt: sourceSignals.trust.trusted.generatedAt
  };
  template.hash = sha256Json({ ...template, hash: undefined });
  return {
    template,
    score: 100 + sourceSignals.schemaCandidates.length + sourceSignals.placeholderCandidates.length,
    reasons: ["source:pptx-analysis", "template-map:suggested", "schema:candidates"],
    sourceMetadata: sourceSignals.metadata,
    previewCandidates: sourceSignals.previewCandidates,
    contextCandidates: sourceSignals.contextCandidates,
    mapCandidates: sourceSignals.mapCandidates,
    placeholderCandidates: sourceSignals.placeholderCandidates,
    namedShapeCandidates: sourceSignals.namedShapeCandidates,
    schemaCandidates: sourceSignals.schemaCandidates,
    templateMapSuggested: sourceSignals.templateMapSuggested,
    artifactPaths: sourceSignals.artifactPaths,
    trust: sourceSignals.trust,
    generatedFromSource: true
  };
}

function normalizeField(value: string): string {
  return value.toLowerCase().replace(/[^a-z0-9]+/g, "");
}

function templatePath(context: OptionalContext, id: string): string {
  return path.join(featureRoot(context, "template"), `${slugify(id)}.json`);
}
