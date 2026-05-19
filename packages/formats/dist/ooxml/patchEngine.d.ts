import { type SourceFingerprint, type SourceSpan } from "./sourceSpan.js";
export type XmlPatch = ReplaceXmlPatch | InsertXmlPatch | DeleteXmlPatch;
export interface ReplaceXmlPatch {
    type: "replace";
    span: SourceSpan;
    value: string;
    fingerprint?: SourceFingerprint;
}
export interface InsertXmlPatch {
    type: "insert";
    offset: number;
    value: string;
    fingerprint?: SourceFingerprint;
}
export interface DeleteXmlPatch {
    type: "delete";
    span: SourceSpan;
    fingerprint?: SourceFingerprint;
}
export interface PatchEngineOptions {
    validateWellFormed?: boolean;
}
export declare class PatchEngineError extends Error {
    readonly code: "PATCH_OFFSET_OUT_OF_RANGE" | "PATCH_STALE_FINGERPRINT" | "PATCH_OVERLAP" | "PATCH_NOT_WELL_FORMED";
    constructor(code: "PATCH_OFFSET_OUT_OF_RANGE" | "PATCH_STALE_FINGERPRINT" | "PATCH_OVERLAP" | "PATCH_NOT_WELL_FORMED", message: string);
}
export declare function applyXmlPatches(source: string, patches: XmlPatch[], options?: PatchEngineOptions): string;
