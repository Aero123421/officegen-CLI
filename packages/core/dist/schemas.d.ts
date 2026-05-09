import { Ajv, type ErrorObject } from "ajv/dist/ajv.js";
import type { OfficegenConfig, SchemaRegistryEntry } from "./types.js";
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
    validate(id: string, value: unknown): {
        ok: true;
    } | {
        ok: false;
        errors: ErrorObject[];
    };
}
export declare const defaultSchemaRegistry: SchemaRegistry;
export declare function listSchemas(options?: {
    agent?: boolean;
    config?: OfficegenConfig;
}): SchemaRegistryEntry[];
export declare function getSchema(id: string): SchemaRegistryEntry | undefined;
export declare function validateSchema(id: string, value: unknown): {
    ok: true;
} | {
    ok: false;
    errors: ErrorObject[];
};
