import JSZip from "jszip";
import { type Relationship } from "./relationships.js";
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
export declare class PackageGraph {
    readonly zip: JSZip;
    readonly format: PackageGraphFormat;
    private readonly contentTypeDefaults;
    private readonly contentTypeOverrides;
    private readonly relationshipsByOwner;
    private readonly relationshipOwnersByPath;
    private readonly parts;
    private constructor();
    static fromZip(zip: JSZip, options?: PackageGraphOptions): Promise<PackageGraph>;
    listParts(): PackagePart[];
    getPart(path: string): PackagePart | undefined;
    readText(path: string): Promise<string | undefined>;
    writeText(path: string, value: string): Promise<void>;
    resolveRelationship(ownerPath: string, rId: string): PackageRelationship | undefined;
    ensureRelationship(ownerPath: string, type: string, target: string, mode?: RelationshipTargetMode): Promise<PackageRelationship>;
    ensureContentTypeOverride(partName: string, contentType: string): Promise<void>;
    validate(): PackageGraphValidationResult;
    private reindex;
    private readContentTypes;
    private upsertPart;
    private refreshPartRelationships;
    private contentTypeFor;
    private decorateRelationship;
    private contentTypesXml;
    private ensureContentTypeDefault;
    private allRelationships;
    private securityFlags;
    private securityFlagsForPath;
}
