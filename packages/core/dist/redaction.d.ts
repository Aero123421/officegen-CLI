import type { JsonValue, OfficegenConfig, RedactionResult, RunFolder } from "./types.js";
export declare function redactSecretsInText(text: string, location?: string): RedactionResult<string>;
export declare function redactPathsInText(text: string, config: OfficegenConfig, location?: string, run?: RunFolder): RedactionResult<string>;
export declare function redactJson<T extends JsonValue>(value: T, config: OfficegenConfig, run?: RunFolder): RedactionResult<T>;
export declare function isAbsolutePathRedactionNeeded(text: string): boolean;
