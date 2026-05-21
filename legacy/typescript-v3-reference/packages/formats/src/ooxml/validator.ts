import { XMLValidator } from "fast-xml-parser";
import JSZip from "jszip";
import { type InputLike, normalizeInput, readZipText, sortedZipFiles } from "../shared.js";
import { parseRelationships, relationshipTarget, type Relationship } from "./relationships.js";

export type OoxmlFormat = "pptx" | "docx" | "xlsx";
export type OoxmlValidationSeverity = "info" | "warning" | "error";
export type OoxmlRiskyPartKind = "macro" | "embeddedObject" | "externalRelationship";

export interface OoxmlValidationIssue {
  code: string;
  severity: OoxmlValidationSeverity;
  message: string;
  path?: string;
  relationshipId?: string;
  target?: string;
  resolvedTarget?: string;
  line?: number;
  column?: number;
}

export interface OoxmlRiskyPart {
  kind: OoxmlRiskyPartKind;
  path: string;
  message: string;
  relationshipId?: string;
  target?: string;
  targetMode?: string;
}

export interface OoxmlRelationshipRecord extends Relationship {
  path: string;
  resolvedTarget?: string;
  external: boolean;
}

export interface OoxmlValidationResult {
  schema: "officegen.ooxml.validation@1";
  ok: boolean;
  format: OoxmlFormat | "unknown";
  issues: OoxmlValidationIssue[];
  riskyParts: OoxmlRiskyPart[];
  relationships: OoxmlRelationshipRecord[];
  summary: {
    entries: number;
    xmlParts: number;
    relationshipParts: number;
    riskyParts: number;
  };
}

export interface ValidateOoxmlOptions {
  format?: OoxmlFormat | "unknown";
}

export async function validateOoxml(input: InputLike, options: ValidateOoxmlOptions = {}): Promise<OoxmlValidationResult> {
  const normalized = await normalizeInput(input, options.format ?? "unknown");
  const zip = await JSZip.loadAsync(normalized.bytes, { checkCRC32: false });
  const format = options.format ?? (isOoxmlFormat(normalized.format) ? normalized.format : "unknown");
  return validateOoxmlZip(zip, { format });
}

export async function validateOoxmlZip(zip: JSZip, options: ValidateOoxmlOptions = {}): Promise<OoxmlValidationResult> {
  const paths = sortedZipFiles(zip);
  const pathSet = new Set(paths);
  const issues: OoxmlValidationIssue[] = [];
  const relationships: OoxmlRelationshipRecord[] = [];

  for (const requiredPath of ["[Content_Types].xml", "_rels/.rels"]) {
    if (!pathSet.has(requiredPath)) {
      issues.push({
        code: "OOXML_MISSING_REQUIRED_PART",
        severity: "error",
        path: requiredPath,
        message: `Required OPC part is missing: ${requiredPath}.`
      });
    }
  }
  for (const requiredPath of requiredFormatParts(options.format ?? "unknown")) {
    if (!pathSet.has(requiredPath)) {
      issues.push({
        code: "OOXML_MISSING_FORMAT_MAIN_PART",
        severity: "error",
        path: requiredPath,
        message: `Required ${options.format} part is missing: ${requiredPath}.`
      });
    }
  }

  const xmlPaths = paths.filter(isXmlPartPath);
  for (const path of xmlPaths) {
    const xml = await readZipText(zip, path);
    const validation = XMLValidator.validate(xml ?? "");
    if (validation !== true) {
      issues.push({
        code: "OOXML_XML_NOT_WELL_FORMED",
        severity: "error",
        path,
        line: validation.err.line,
        column: validation.err.col,
        message: `XML is not well-formed in ${path}: ${validation.err.code} ${validation.err.msg}`
      });
    }
  }

  const relsPaths = paths.filter((path) => path.endsWith(".rels"));
  for (const path of relsPaths) {
    const xml = await readZipText(zip, path);
    if (!xml || XMLValidator.validate(xml) !== true) continue;
    const seenIds = new Set<string>();
    const duplicateIds = new Set<string>();
    const baseDir = relationshipBaseDir(path);

    for (const rel of parseRelationships(xml)) {
      if (seenIds.has(rel.id) && !duplicateIds.has(rel.id)) {
        duplicateIds.add(rel.id);
        issues.push({
          code: "OOXML_DUPLICATE_RELATIONSHIP_ID",
          severity: "error",
          path,
          relationshipId: rel.id,
          message: `Duplicate relationship id ${rel.id} in ${path}.`
        });
      }
      seenIds.add(rel.id);

      const external = isExternalRelationship(rel);
      const resolvedTarget = external ? undefined : resolveInternalRelationshipTarget(baseDir, rel.target);
      relationships.push({ ...rel, path, resolvedTarget, external });

      if (external) {
        issues.push({
          code: "OOXML_EXTERNAL_RELATIONSHIP_TARGET",
          severity: "warning",
          path,
          relationshipId: rel.id,
          target: rel.target,
          message: `Relationship ${rel.id} in ${path} targets an external resource.`
        });
        continue;
      }

      if (resolvedTarget && !pathSet.has(resolvedTarget)) {
        issues.push({
          code: "OOXML_MISSING_INTERNAL_RELATIONSHIP_TARGET",
          severity: "error",
          path,
          relationshipId: rel.id,
          target: rel.target,
          resolvedTarget,
          message: `Relationship ${rel.id} in ${path} points to missing internal target ${resolvedTarget}.`
        });
      }
    }
  }

  const riskyParts = collectOoxmlRiskyParts(paths, relationships);
  const errors = issues.filter((issue) => issue.severity === "error").length;

  return {
    schema: "officegen.ooxml.validation@1",
    ok: errors === 0,
    format: options.format ?? "unknown",
    issues,
    riskyParts,
    relationships,
    summary: {
      entries: paths.length,
      xmlParts: xmlPaths.length,
      relationshipParts: relsPaths.length,
      riskyParts: riskyParts.length
    }
  };
}

function requiredFormatParts(format: OoxmlFormat | "unknown"): string[] {
  if (format === "pptx") return ["ppt/presentation.xml"];
  if (format === "docx") return ["word/document.xml"];
  if (format === "xlsx") return ["xl/workbook.xml"];
  return [];
}

export async function detectOoxmlRiskyParts(input: InputLike, options: ValidateOoxmlOptions = {}): Promise<OoxmlRiskyPart[]> {
  return (await validateOoxml(input, options)).riskyParts;
}

export function collectOoxmlRiskyParts(paths: Iterable<string>, relationships: Iterable<OoxmlRelationshipRecord> = []): OoxmlRiskyPart[] {
  const riskyParts: OoxmlRiskyPart[] = [];

  for (const path of paths) {
    if (/(^|\/)vbaProject\.bin$/i.test(path)) {
      riskyParts.push({
        kind: "macro",
        path,
        message: `VBA project part detected: ${path}.`
      });
    } else if (/(^|\/)embeddings\/[^/]+$/i.test(path)) {
      riskyParts.push({
        kind: "embeddedObject",
        path,
        message: `Embedded object part detected: ${path}.`
      });
    }
  }

  for (const rel of relationships) {
    if (!isExternalRelationship(rel)) continue;
    riskyParts.push({
      kind: "externalRelationship",
      path: rel.path,
      relationshipId: rel.id,
      target: rel.target,
      targetMode: rel.targetMode,
      message: `External relationship detected: ${rel.path} ${rel.id}.`
    });
  }

  return riskyParts;
}

function isOoxmlFormat(format: string): format is OoxmlFormat {
  return format === "pptx" || format === "docx" || format === "xlsx";
}

function isXmlPartPath(path: string): boolean {
  return path.endsWith(".xml") || path.endsWith(".rels");
}

function relationshipBaseDir(relsPath: string): string {
  if (relsPath === "_rels/.rels") return "";
  return relsPath.replace(/\/_rels\/[^/]+\.rels$/, "");
}

function resolveInternalRelationshipTarget(baseDir: string, target: string): string | undefined {
  const cleanTarget = target.replace(/\\/g, "/").split("#", 1)[0]?.split("?", 1)[0] ?? "";
  if (!cleanTarget) return undefined;
  return relationshipTarget(baseDir, cleanTarget);
}

function isExternalRelationship(rel: Pick<Relationship, "target" | "targetMode">): boolean {
  return /^external$/i.test(rel.targetMode ?? "") || /^[A-Za-z][A-Za-z0-9+.-]*:/.test(rel.target);
}
