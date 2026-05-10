import JSZip from "jszip";
import { type InputLike } from "../shared.js";
import { type Relationship } from "./relationships.js";
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
export declare function validateOoxml(input: InputLike, options?: ValidateOoxmlOptions): Promise<OoxmlValidationResult>;
export declare function validateOoxmlZip(zip: JSZip, options?: ValidateOoxmlOptions): Promise<OoxmlValidationResult>;
export declare function detectOoxmlRiskyParts(input: InputLike, options?: ValidateOoxmlOptions): Promise<OoxmlRiskyPart[]>;
export declare function collectOoxmlRiskyParts(paths: Iterable<string>, relationships?: Iterable<OoxmlRelationshipRecord>): OoxmlRiskyPart[];
