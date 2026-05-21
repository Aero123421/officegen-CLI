import JSZip from "jszip";
import { XMLBuilder, XMLParser } from "fast-xml-parser";
import { parseRelationships, relationshipTarget, nextRelationshipId, type Relationship } from "./relationships.js";

export type PackageGraphFormat = "pptx" | "docx" | "xlsx" | "unknown";
export type RelationshipTargetMode = "Internal" | "External" | string;
export type PackageSecurityFlagKind = "macro" | "embeddedObject" | "externalRelationship";
export type PackageGraphIssueSeverity = "warning" | "error";

export interface PackageGraphOptions {
  format?: PackageGraphFormat;
}

export interface PackagePart {
  path: string;
  contentType?: string;
  relationshipOwner?: string;
  relationships: PackageRelationship[];
  unknown: boolean;
  securityFlags: PackageSecurityFlag[];
}

export interface PackageRelationship extends Relationship {
  ownerPath: string;
  path: string;
  external: boolean;
  resolvedTarget?: string;
}

export interface PackageSecurityFlag {
  kind: PackageSecurityFlagKind;
  path: string;
  relationshipId?: string;
  target?: string;
  message: string;
}

export interface PackageGraphIssue {
  code: string;
  severity: PackageGraphIssueSeverity;
  message: string;
  path?: string;
  relationshipId?: string;
  target?: string;
  resolvedTarget?: string;
}

export interface PackageGraphValidationResult {
  ok: boolean;
  issues: PackageGraphIssue[];
  securityFlags: PackageSecurityFlag[];
  summary: {
    parts: number;
    relationships: number;
    unknownParts: number;
    externalRelationships: number;
  };
}

interface ContentTypesXml {
  Types?: {
    Default?: ContentTypeDefault | ContentTypeDefault[];
    Override?: ContentTypeOverride | ContentTypeOverride[];
    [key: string]: unknown;
  };
}

interface ContentTypeDefault {
  "@_Extension"?: string;
  "@_ContentType"?: string;
  [key: string]: unknown;
}

interface ContentTypeOverride {
  "@_PartName"?: string;
  "@_ContentType"?: string;
  [key: string]: unknown;
}

const CONTENT_TYPES_PATH = "[Content_Types].xml";
const RELATIONSHIP_CONTENT_TYPE = "application/vnd.openxmlformats-package.relationships+xml";
const CONTENT_TYPES_XMLNS = "http://schemas.openxmlformats.org/package/2006/content-types";
const RELATIONSHIPS_XMLNS = "http://schemas.openxmlformats.org/package/2006/relationships";

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  preserveOrder: false,
  parseTagValue: false,
  parseAttributeValue: false,
  trimValues: false
});

const builder = new XMLBuilder({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  suppressEmptyNode: true,
  format: false
});

export class PackageGraph {
  readonly zip: JSZip;
  readonly format: PackageGraphFormat;

  private readonly contentTypeDefaults = new Map<string, string>();
  private readonly contentTypeOverrides = new Map<string, string>();
  private readonly relationshipsByOwner = new Map<string, PackageRelationship[]>();
  private readonly relationshipOwnersByPath = new Map<string, string>();
  private readonly parts = new Map<string, PackagePart>();

  private constructor(zip: JSZip, format: PackageGraphFormat) {
    this.zip = zip;
    this.format = format;
  }

  static async fromZip(zip: JSZip, options: PackageGraphOptions = {}): Promise<PackageGraph> {
    const graph = new PackageGraph(zip, options.format ?? "unknown");
    await graph.reindex();
    return graph;
  }

  listParts(): PackagePart[] {
    return [...this.parts.values()].sort((a, b) => a.path.localeCompare(b.path));
  }

  getPart(path: string): PackagePart | undefined {
    return this.parts.get(normalizePartPath(path));
  }

  async readText(path: string): Promise<string | undefined> {
    const file = this.zip.file(normalizePartPath(path));
    return file ? file.async("string") : undefined;
  }

  async writeText(path: string, value: string): Promise<void> {
    const normalized = normalizePartPath(path);
    this.zip.file(normalized, value);
    this.upsertPart(normalized);
  }

  resolveRelationship(ownerPath: string, rId: string): PackageRelationship | undefined {
    return this.relationshipsByOwner.get(normalizeOwnerPath(ownerPath))?.find((rel) => rel.id === rId);
  }

  async ensureRelationship(
    ownerPath: string,
    type: string,
    target: string,
    mode?: RelationshipTargetMode
  ): Promise<PackageRelationship> {
    const normalizedOwner = normalizeOwnerPath(ownerPath);
    const rels = this.relationshipsByOwner.get(normalizedOwner) ?? [];
    const normalizedTargetMode = normalizeTargetMode(mode);
    const existing = rels.find((rel) => rel.type === type && rel.target === target && normalizeTargetMode(rel.targetMode) === normalizedTargetMode);
    if (existing) return existing;

    const relsPath = relationshipPathForOwner(normalizedOwner);
    const existingXml = await this.readText(relsPath);
    const id = nextRelationshipId(existingXml ?? relationshipsXml(rels));
    const rel = this.decorateRelationship(normalizedOwner, relsPath, { id, type, target, targetMode: mode });
    const next = [...rels, rel];

    this.relationshipsByOwner.set(normalizedOwner, next);
    this.relationshipOwnersByPath.set(relsPath, normalizedOwner);
    this.zip.file(relsPath, relationshipsXml(next));
    this.upsertPart(relsPath);
    this.ensureContentTypeDefault("rels", RELATIONSHIP_CONTENT_TYPE);
    this.refreshPartRelationships(normalizedOwner);
    return rel;
  }

  async ensureContentTypeOverride(partName: string, contentType: string): Promise<void> {
    const normalized = normalizePartPath(partName);
    this.contentTypeOverrides.set(normalized, contentType);
    this.zip.file(CONTENT_TYPES_PATH, this.contentTypesXml());
    this.upsertPart(CONTENT_TYPES_PATH);
    const part = this.parts.get(normalized);
    if (part) part.contentType = contentType;
  }

  validate(): PackageGraphValidationResult {
    const issues: PackageGraphIssue[] = [];
    const paths = new Set(this.parts.keys());
    if (!paths.has(CONTENT_TYPES_PATH)) {
      issues.push({
        code: "OPC_MISSING_CONTENT_TYPES",
        severity: "error",
        path: CONTENT_TYPES_PATH,
        message: "OPC package is missing [Content_Types].xml."
      });
    }
    if (!paths.has("_rels/.rels")) {
      issues.push({
        code: "OPC_MISSING_ROOT_RELATIONSHIPS",
        severity: "error",
        path: "_rels/.rels",
        message: "OPC package is missing the root relationships part."
      });
    }

    for (const rel of this.allRelationships()) {
      if (rel.external) continue;
      if (rel.resolvedTarget && !paths.has(rel.resolvedTarget)) {
        issues.push({
          code: "OPC_MISSING_RELATIONSHIP_TARGET",
          severity: "error",
          path: rel.path,
          relationshipId: rel.id,
          target: rel.target,
          resolvedTarget: rel.resolvedTarget,
          message: `Relationship ${rel.id} in ${rel.path} points to missing internal target ${rel.resolvedTarget}.`
        });
      }
    }

    const securityFlags = this.securityFlags();
    return {
      ok: !issues.some((issue) => issue.severity === "error"),
      issues,
      securityFlags,
      summary: {
        parts: this.parts.size,
        relationships: this.allRelationships().length,
        unknownParts: this.listParts().filter((part) => part.unknown).length,
        externalRelationships: this.allRelationships().filter((rel) => rel.external).length
      }
    };
  }

  private async reindex(): Promise<void> {
    this.contentTypeDefaults.clear();
    this.contentTypeOverrides.clear();
    this.relationshipsByOwner.clear();
    this.relationshipOwnersByPath.clear();
    this.parts.clear();

    await this.readContentTypes();

    for (const path of zipFilePaths(this.zip).filter((item) => item.endsWith(".rels"))) {
      const xml = await this.readText(path);
      const ownerPath = ownerPathForRelationshipPath(path);
      this.relationshipOwnersByPath.set(path, ownerPath);
      this.relationshipsByOwner.set(ownerPath, parseRelationships(xml).map((rel) => this.decorateRelationship(ownerPath, path, rel)));
    }

    for (const path of zipFilePaths(this.zip)) {
      this.upsertPart(path);
    }
  }

  private async readContentTypes(): Promise<void> {
    const xml = await this.readText(CONTENT_TYPES_PATH);
    if (!xml) return;
    const parsed = parser.parse(xml) as ContentTypesXml;
    for (const item of ensureArray(parsed.Types?.Default)) {
      const extension = normalizeExtension(item["@_Extension"]);
      const contentType = item["@_ContentType"];
      if (extension && contentType) this.contentTypeDefaults.set(extension, contentType);
    }
    for (const item of ensureArray(parsed.Types?.Override)) {
      const partName = item["@_PartName"];
      const contentType = item["@_ContentType"];
      if (partName && contentType) this.contentTypeOverrides.set(normalizePartPath(partName), contentType);
    }
  }

  private upsertPart(path: string): void {
    const normalized = normalizePartPath(path);
    const owner = this.relationshipOwnersByPath.get(normalized);
    const relationships = this.relationshipsByOwner.get(normalized) ?? [];
    this.parts.set(normalized, {
      path: normalized,
      contentType: this.contentTypeFor(normalized),
      relationshipOwner: owner,
      relationships,
      unknown: isUnknownPart(normalized, this.format),
      securityFlags: this.securityFlagsForPath(normalized)
    });
  }

  private refreshPartRelationships(ownerPath: string): void {
    const part = this.parts.get(ownerPath);
    if (part) part.relationships = this.relationshipsByOwner.get(ownerPath) ?? [];
  }

  private contentTypeFor(path: string): string | undefined {
    if (path.endsWith(".rels")) return this.contentTypeOverrides.get(path) ?? this.contentTypeDefaults.get("rels") ?? RELATIONSHIP_CONTENT_TYPE;
    const override = this.contentTypeOverrides.get(path);
    if (override) return override;
    const extension = extensionForPath(path);
    return extension ? this.contentTypeDefaults.get(extension) : undefined;
  }

  private decorateRelationship(ownerPath: string, relsPath: string, rel: Relationship): PackageRelationship {
    const external = isExternalRelationship(rel);
    return {
      ...rel,
      ownerPath,
      path: relsPath,
      external,
      resolvedTarget: external ? undefined : resolveRelationshipTargetForOwner(ownerPath, rel.target)
    };
  }

  private contentTypesXml(): string {
    const defaults = [...this.contentTypeDefaults.entries()]
      .sort(([left], [right]) => left.localeCompare(right))
      .map(([extension, contentType]) => ({ "@_Extension": extension, "@_ContentType": contentType }));
    const overrides = [...this.contentTypeOverrides.entries()]
      .sort(([left], [right]) => left.localeCompare(right))
      .map(([path, contentType]) => ({ "@_PartName": `/${path}`, "@_ContentType": contentType }));
    return builder.build({
      Types: {
        "@_xmlns": CONTENT_TYPES_XMLNS,
        Default: defaults,
        Override: overrides
      }
    });
  }

  private ensureContentTypeDefault(extension: string, contentType: string): void {
    if (this.contentTypeDefaults.get(extension) === contentType) return;
    this.contentTypeDefaults.set(extension, contentType);
    this.zip.file(CONTENT_TYPES_PATH, this.contentTypesXml());
    this.upsertPart(CONTENT_TYPES_PATH);
  }

  private allRelationships(): PackageRelationship[] {
    return [...this.relationshipsByOwner.values()].flat();
  }

  private securityFlags(): PackageSecurityFlag[] {
    return this.listParts().flatMap((part) => part.securityFlags);
  }

  private securityFlagsForPath(path: string): PackageSecurityFlag[] {
    const flags: PackageSecurityFlag[] = [];
    if (/(^|\/)vbaProject\.bin$/i.test(path)) {
      flags.push({
        kind: "macro",
        path,
        message: `VBA project part detected: ${path}.`
      });
    }
    if (/(^|\/)embeddings\/[^/]+$/i.test(path)) {
      flags.push({
        kind: "embeddedObject",
        path,
        message: `Embedded object part detected: ${path}.`
      });
    }
    for (const rel of this.allRelationships().filter((item) => item.path === path && item.external)) {
      flags.push({
        kind: "externalRelationship",
        path,
        relationshipId: rel.id,
        target: rel.target,
        message: `External relationship detected: ${path} ${rel.id}.`
      });
    }
    return flags;
  }
}

function zipFilePaths(zip: JSZip): string[] {
  return Object.keys(zip.files)
    .filter((path) => !zip.files[path]?.dir)
    .map(normalizePartPath)
    .sort((left, right) => left.localeCompare(right));
}

function normalizePartPath(path: string): string {
  return path.replace(/\\/g, "/").replace(/^\/+/, "");
}

function normalizeOwnerPath(path: string): string {
  const normalized = normalizePartPath(path);
  return normalized === "." ? "" : normalized;
}

function normalizeExtension(extension: string | undefined): string | undefined {
  return extension?.replace(/^\./, "").toLowerCase();
}

function extensionForPath(path: string): string | undefined {
  const name = path.split("/").at(-1) ?? "";
  const dot = name.lastIndexOf(".");
  return dot >= 0 ? name.slice(dot + 1).toLowerCase() : undefined;
}

function relationshipPathForOwner(ownerPath: string): string {
  if (!ownerPath) return "_rels/.rels";
  const parts = ownerPath.split("/");
  const fileName = parts.pop() ?? "";
  return [...parts, "_rels", `${fileName}.rels`].filter(Boolean).join("/");
}

function ownerPathForRelationshipPath(path: string): string {
  if (path === "_rels/.rels") return "";
  const match = /^(.*)\/_rels\/([^/]+)\.rels$/.exec(path);
  if (!match) return "";
  return [match[1], match[2]].filter(Boolean).join("/");
}

function relationshipBaseDir(ownerPath: string): string {
  const index = ownerPath.lastIndexOf("/");
  return index >= 0 ? ownerPath.slice(0, index) : "";
}

function resolveRelationshipTargetForOwner(ownerPath: string, target: string): string | undefined {
  const cleanTarget = target.replace(/\\/g, "/").split("#", 1)[0]?.split("?", 1)[0] ?? "";
  if (!cleanTarget) return undefined;
  return relationshipTarget(relationshipBaseDir(ownerPath), cleanTarget);
}

function isExternalRelationship(rel: Pick<Relationship, "target" | "targetMode">): boolean {
  return /^external$/i.test(rel.targetMode ?? "") || /^[A-Za-z][A-Za-z0-9+.-]*:/.test(rel.target);
}

function relationshipsXml(rels: Relationship[]): string {
  const body = rels.map((rel) => {
    const attrs = [
      `Id="${escapeXml(rel.id)}"`,
      `Type="${escapeXml(rel.type)}"`,
      `Target="${escapeXml(rel.target)}"`,
      rel.targetMode ? `TargetMode="${escapeXml(rel.targetMode)}"` : undefined
    ].filter(Boolean);
    return `<Relationship ${attrs.join(" ")}/>`;
  });
  return `<Relationships xmlns="${RELATIONSHIPS_XMLNS}">${body.join("")}</Relationships>`;
}

function escapeXml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function ensureArray<T>(value: T | T[] | undefined): T[] {
  if (value === undefined) return [];
  return Array.isArray(value) ? value : [value];
}

function normalizeTargetMode(mode: string | undefined): string | undefined {
  return mode ? mode.toLowerCase() : undefined;
}

function isUnknownPart(path: string, format: PackageGraphFormat): boolean {
  if (path === CONTENT_TYPES_PATH || path.endsWith(".rels") || path.startsWith("docProps/")) return false;
  if (format === "pptx") return !path.startsWith("ppt/");
  if (format === "docx") return !path.startsWith("word/");
  if (format === "xlsx") return !path.startsWith("xl/");
  return false;
}
