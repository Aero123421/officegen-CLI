import { listTemplates } from "@officegen/optional";
import type { EditOperation } from "@officegen/formats";
import { type RuntimeContext } from "./types.js";
export declare function requireInput(context: RuntimeContext, start: number, command: string): string;
export declare function resolveCliPath(context: RuntimeContext, inputPath: string): string;
export declare function validatedOutOption(context: RuntimeContext): Promise<string | undefined>;
export declare function validateOutputPath(context: RuntimeContext, outputPath: string, options?: {
    directory?: boolean;
}): Promise<string>;
export declare function readJson(filePath: string): Promise<unknown>;
export declare function readJsonIfPresent(filePath: string): Promise<unknown>;
export declare function copyJsonIfPresent(cwd: string, input: string, out: string): Promise<void>;
export declare function normalizeEditOperations(raw: unknown): EditOperation[];
export declare function optionalContext(context: RuntimeContext): Parameters<typeof listTemplates>[0];
export declare function asRecord(value: unknown): Record<string, unknown>;
export declare function stringRecord(value: unknown): Record<string, string>;
export declare function numberOption(context: RuntimeContext, name: string): number | undefined;
export declare function schemaHiddenFromAgent(context: RuntimeContext, schema: string): boolean;
export declare function isNodeError(error: unknown): error is NodeJS.ErrnoException;
