import { Ajv, type ErrorObject } from "ajv/dist/ajv.js";
import type { OfficegenConfig, SchemaRegistryEntry } from "./types.js";
export interface SchemaOneOfDiagnostic {
    instancePath: string;
    schemaPath: string;
    bestMatch?: {
        schemaPath: string;
        op?: string;
        score: number;
    };
    missing: string[];
    unexpected: string[];
    expectedTypes: Record<string, string[]>;
}
export type SchemaValidationResult = {
    ok: true;
} | {
    ok: false;
    errors: ErrorObject[];
    diagnostics?: SchemaOneOfDiagnostic[];
};
export declare class SchemaRegistry {
    readonly ajv: Ajv;
    private readonly entriesById;
    private readonly validators;
    constructor(registryEntries?: SchemaRegistryEntry[]);
    list(options?: {
        agent?: boolean;
        config?: OfficegenConfig;
    }): SchemaRegistryEntry[];
    get(id: string): SchemaRegistryEntry | undefined;
    validate(id: string, value: unknown, options?: {
        diagnostics?: boolean;
    }): SchemaValidationResult;
}
export declare const defaultSchemaRegistry: SchemaRegistry;
export declare function listSchemas(options?: {
    agent?: boolean;
    config?: OfficegenConfig;
}): SchemaRegistryEntry[];
export declare function getSchema(id: string): SchemaRegistryEntry | undefined;
export declare function validateSchema(id: string, value: unknown, options?: {
    diagnostics?: boolean;
}): SchemaValidationResult;
export declare function compactSchemaErrors(errors: ErrorObject[], diagnostics?: SchemaOneOfDiagnostic[]): ErrorObject[];
