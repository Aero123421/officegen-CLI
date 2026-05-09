import path from "node:path";

import {
  OptionalContext,
  ValidationResult,
  featureRoot,
  listJsonFiles,
  nowIso,
  readJsonFile,
  requireFeature,
  sha256Json,
  slugify,
  validation,
  writeJsonFile
} from "./common.js";

export type TemplateFieldType = "string" | "number" | "boolean" | "date" | "json";

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
  mapping?: Record<string, string>;
  requiredCapabilities?: string[];
  createdAt?: string;
  updatedAt?: string;
  hash?: string;
}

export interface TemplateCandidate {
  template: TemplateDefinition;
  score: number;
  reasons: string[];
}

export interface TemplateCreateOptions extends OptionalContext {
  template: Omit<TemplateDefinition, "createdAt" | "updatedAt" | "hash">;
}

export interface TemplateQueryOptions extends OptionalContext {
  query?: string;
  tags?: string[];
  fields?: string[];
}

export interface TemplateIdOptions extends OptionalContext {
  id: string;
}

export interface TemplateApplyMapOptions extends TemplateIdOptions {
  mapping: Record<string, string>;
  outputPath?: string;
}

export interface TemplateFillOptions extends TemplateIdOptions {
  values: Record<string, unknown>;
  outputPath?: string;
}

export async function createTemplate(options: TemplateCreateOptions): Promise<TemplateDefinition> {
  requireFeature(options, "template", "template create");
  const now = nowIso();
  const template: TemplateDefinition = {
    ...options.template,
    id: slugify(options.template.id),
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

  return templates
    .map((template) => {
      const reasons: string[] = [];
      let score = 0;
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

      if (!query && tags.size === 0 && fields.size === 0) {
        score = 1;
        reasons.push("available");
      }

      return { template, score, reasons };
    })
    .filter((candidate) => candidate.score > 0)
    .sort((left, right) => right.score - left.score || left.template.id.localeCompare(right.template.id));
}

export async function applyTemplateMap(options: TemplateApplyMapOptions): Promise<Record<string, unknown>> {
  requireFeature(options, "template", "template apply-map");
  const template = await inspectTemplate(options);
  const plan = {
    kind: "officegen.template.apply-map",
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

function templatePath(context: OptionalContext, id: string): string {
  return path.join(featureRoot(context, "template"), `${slugify(id)}.json`);
}
