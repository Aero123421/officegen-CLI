import type { MaybePromise, EditPartStore, EditTransaction } from "../transaction.js";

export type OperationFormat = "pptx" | "docx" | "xlsx" | "pdf" | string;

export type OperationName = string;

export interface NamedOperation {
  op?: OperationName;
  type?: OperationName;
  [key: string]: unknown;
}

export interface OperationExecutionOptions {
  atomic?: boolean;
  continueOnError?: boolean;
  dryRun?: boolean;
}

export interface OperationContext<TPart = Uint8Array | string, TShared = unknown> {
  format: OperationFormat;
  transaction: EditTransaction<TPart>;
  store: EditPartStore<TPart>;
  options: OperationExecutionOptions;
  shared?: TShared;
}

export interface OperationResult {
  applied: boolean;
  changed?: boolean;
  skipped?: boolean;
  reason?: "unsupported" | "not-found" | "validation-failed" | "skipped-after-error" | string;
  message?: string;
}

export type OperationHandler<
  TOperation extends NamedOperation = NamedOperation,
  TResult extends OperationResult = OperationResult,
  TPart = Uint8Array | string,
  TShared = unknown
> = (operation: TOperation, context: OperationContext<TPart, TShared>) => MaybePromise<TResult>;

export interface OperationRegistration<
  TOperation extends NamedOperation = NamedOperation,
  TResult extends OperationResult = OperationResult,
  TPart = Uint8Array | string,
  TShared = unknown
> {
  format: OperationFormat;
  opName: OperationName;
  handler: OperationHandler<TOperation, TResult, TPart, TShared>;
  description?: string;
}

export interface RegisteredOperation<
  TOperation extends NamedOperation = NamedOperation,
  TResult extends OperationResult = OperationResult,
  TPart = Uint8Array | string,
  TShared = unknown
> extends OperationRegistration<TOperation, TResult, TPart, TShared> {
  key: string;
}

export interface OperationLookupSupported<
  TOperation extends NamedOperation = NamedOperation,
  TResult extends OperationResult = OperationResult,
  TPart = Uint8Array | string,
  TShared = unknown
> {
  supported: true;
  format: OperationFormat;
  opName: OperationName;
  registration: RegisteredOperation<TOperation, TResult, TPart, TShared>;
}

export interface OperationLookupUnsupported {
  supported: false;
  format: OperationFormat;
  opName?: OperationName;
  reason: "missing-name" | "unsupported";
  message: string;
}

export type OperationLookupResult<
  TOperation extends NamedOperation = NamedOperation,
  TResult extends OperationResult = OperationResult,
  TPart = Uint8Array | string,
  TShared = unknown
> =
  | OperationLookupSupported<TOperation, TResult, TPart, TShared>
  | OperationLookupUnsupported;
