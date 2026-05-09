import type { ErrorCatalogEntry, JsonValue, OfficegenErrorCode, OfficegenErrorPayload } from "./types.js";
export declare const ERROR_CATALOG: Record<OfficegenErrorCode, ErrorCatalogEntry>;
export declare class OfficegenError extends Error {
    readonly payload: OfficegenErrorPayload;
    constructor(code: OfficegenErrorCode, message?: string, details?: JsonValue, overrides?: Partial<OfficegenErrorPayload>);
}
export declare function getRequiredErrorCodes(): OfficegenErrorCode[];
export declare function listErrors(): ErrorCatalogEntry[];
export declare function inspectError(code: OfficegenErrorCode): ErrorCatalogEntry;
export declare function createErrorPayload(code: OfficegenErrorCode, options?: Partial<OfficegenErrorPayload>): OfficegenErrorPayload;
export declare function assertKnownErrorCode(code: string): asserts code is OfficegenErrorCode;
