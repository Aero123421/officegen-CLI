import { type NamedOperation, type OperationFormat, type OperationLookupResult, type OperationName, type OperationRegistration, type OperationResult, type RegisteredOperation } from "./types.js";
export declare class DuplicateOperationRegistrationError extends Error {
    readonly format: OperationFormat;
    readonly opName: OperationName;
    readonly code = "DUPLICATE_OPERATION_REGISTRATION";
    constructor(format: OperationFormat, opName: OperationName);
}
export declare class UnsupportedOperationError extends Error {
    readonly format: OperationFormat;
    readonly opName: OperationName | undefined;
    readonly code = "UNSUPPORTED_OPERATION";
    constructor(format: OperationFormat, opName: OperationName | undefined, reason?: "missing-name" | "unsupported");
}
export declare class OperationRegistry<TOperation extends NamedOperation = NamedOperation, TResult extends OperationResult = OperationResult, TPart = Uint8Array | string, TShared = unknown> {
    private readonly operations;
    register(registration: OperationRegistration<TOperation, TResult, TPart, TShared>): this;
    has(format: OperationFormat, opName: OperationName): boolean;
    get(format: OperationFormat, opName: OperationName): RegisteredOperation<TOperation, TResult, TPart, TShared> | undefined;
    lookup(format: OperationFormat, operation: TOperation | OperationName): OperationLookupResult<TOperation, TResult, TPart, TShared>;
    require(format: OperationFormat, operation: TOperation | OperationName): RegisteredOperation<TOperation, TResult, TPart, TShared>;
    list(format?: OperationFormat): RegisteredOperation<TOperation, TResult, TPart, TShared>[];
}
export declare function createOperationRegistry<TOperation extends NamedOperation = NamedOperation, TResult extends OperationResult = OperationResult, TPart = Uint8Array | string, TShared = unknown>(): OperationRegistry<TOperation, TResult, TPart, TShared>;
export declare function operationName(operation: NamedOperation | OperationName): OperationName | undefined;
export declare function operationKey(format: OperationFormat, opName: OperationName): string;
