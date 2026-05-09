import { type CliErrorPayload, type Envelope, type RuntimeContext } from "./types.js";
export declare function makeEnvelope(context: RuntimeContext, command: string, data: unknown, now: Date): Envelope;
export declare function makeErrorEnvelope(context: RuntimeContext, command: string, error: CliErrorPayload, now: Date): Envelope;
export declare function writeResult(context: RuntimeContext, envelope: Envelope, writer: (text: string) => void): void;
export declare function normalizeCliError(error: CliErrorPayload): CliErrorPayload;
export declare function redactForJson(value: unknown, context: RuntimeContext): unknown;
export declare function runId(now: Date): string;
