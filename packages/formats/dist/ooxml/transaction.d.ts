export type MaybePromise<T> = T | Promise<T>;
export type EditTransactionMode = "atomic" | "best-effort";
export interface PartSnapshot<TPart> {
    path: string;
    existed: boolean;
    value?: TPart;
}
export interface PartJournalEntry<TPart> extends PartSnapshot<TPart> {
    sequence: number;
}
export interface EditPartStore<TPart> {
    readPart(path: string): MaybePromise<TPart | undefined>;
    writePart(path: string, value: TPart): MaybePromise<void>;
    deletePart?(path: string): MaybePromise<void>;
}
export interface EditTransactionOptions<TPart> {
    atomic?: boolean;
    continueOnError?: boolean;
    clonePart?: (value: TPart) => TPart;
}
export interface TransactionOperationResult<T = unknown> {
    operationIndex: number;
    applied: boolean;
    rolledBack: boolean;
    value?: T;
    error?: unknown;
}
export interface EditTransactionCommitResult {
    committed: true;
    journaledParts: number;
}
export interface EditTransactionRollbackResult {
    rolledBack: true;
    restoredParts: number;
    errors: Array<{
        path: string;
        error: unknown;
    }>;
}
export declare class EditTransactionError extends Error {
    readonly code: string;
    constructor(message: string, code: string);
}
export declare class EditTransaction<TPart = Uint8Array | string> {
    private readonly store;
    private readonly options;
    private readonly journal;
    private readonly clonePart;
    private closed;
    private nextSequence;
    constructor(store: EditPartStore<TPart>, options?: EditTransactionOptions<TPart>);
    get mode(): EditTransactionMode;
    get atomic(): boolean;
    get continueOnError(): boolean;
    get closedForWrites(): boolean;
    get journaledParts(): number;
    snapshot(): PartJournalEntry<TPart>[];
    snapshotPart(path: string): Promise<PartSnapshot<TPart>>;
    readPart(path: string): Promise<TPart | undefined>;
    writePart(path: string, value: TPart): Promise<void>;
    deletePart(path: string): Promise<void>;
    run<T>(operationIndex: number, operation: () => MaybePromise<T>): Promise<TransactionOperationResult<T>>;
    commit(): Promise<EditTransactionCommitResult>;
    rollback(): Promise<EditTransactionRollbackResult>;
    private assertWritable;
}
export declare function createEditTransaction<TPart>(store: EditPartStore<TPart>, options?: EditTransactionOptions<TPart>): EditTransaction<TPart>;
