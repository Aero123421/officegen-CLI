import type { ErrorEnvelope, JsonObject, JsonValue, OfficegenErrorCode, OfficegenErrorPayload, SuccessEnvelope } from "./types.js";
interface EnvelopeOptions {
    command?: string;
    runId?: string;
    capabilitiesHash?: string;
    pathsRedacted?: boolean;
    warnings?: JsonValue[];
    diagnostics?: JsonValue[];
    artifacts?: JsonValue[];
    nextSuggestedCommands?: string[];
    cliVersion?: string;
}
export declare function successEnvelope<T extends JsonValue = JsonObject>(result: T, options?: EnvelopeOptions): SuccessEnvelope<T>;
export declare function errorEnvelope(error: OfficegenErrorCode | OfficegenErrorPayload, options?: EnvelopeOptions & {
    availableCommands?: string[];
    errorOptions?: Partial<OfficegenErrorPayload>;
}): ErrorEnvelope;
export {};
