import {
  type NamedOperation,
  type OperationFormat,
  type OperationLookupResult,
  type OperationName,
  type OperationRegistration,
  type OperationResult,
  type RegisteredOperation
} from "./types.js";

export class DuplicateOperationRegistrationError extends Error {
  readonly code = "DUPLICATE_OPERATION_REGISTRATION";

  constructor(
    public readonly format: OperationFormat,
    public readonly opName: OperationName
  ) {
    super(`Operation handler already registered for ${format}.${opName}.`);
    this.name = "DuplicateOperationRegistrationError";
  }
}

export class UnsupportedOperationError extends Error {
  readonly code = "UNSUPPORTED_OPERATION";

  constructor(
    public readonly format: OperationFormat,
    public readonly opName: OperationName | undefined,
    reason: "missing-name" | "unsupported" = opName ? "unsupported" : "missing-name"
  ) {
    super(reason === "missing-name"
      ? `Operation for ${format} is missing an op/type name.`
      : `Unsupported operation for ${format}: ${opName}.`);
    this.name = "UnsupportedOperationError";
  }
}

export class OperationRegistry<
  TOperation extends NamedOperation = NamedOperation,
  TResult extends OperationResult = OperationResult,
  TPart = Uint8Array | string,
  TShared = unknown
> {
  private readonly operations = new Map<string, RegisteredOperation<TOperation, TResult, TPart, TShared>>();

  register(registration: OperationRegistration<TOperation, TResult, TPart, TShared>): this {
    const format = normalizeFormat(registration.format);
    const opName = normalizeOperationName(registration.opName);
    const key = operationKey(format, opName);
    if (this.operations.has(key)) {
      throw new DuplicateOperationRegistrationError(format, opName);
    }
    this.operations.set(key, {
      ...registration,
      format,
      opName,
      key
    });
    return this;
  }

  has(format: OperationFormat, opName: OperationName): boolean {
    return this.operations.has(operationKey(normalizeFormat(format), normalizeOperationName(opName)));
  }

  get(format: OperationFormat, opName: OperationName): RegisteredOperation<TOperation, TResult, TPart, TShared> | undefined {
    return this.operations.get(operationKey(normalizeFormat(format), normalizeOperationName(opName)));
  }

  lookup(format: OperationFormat, operation: TOperation | OperationName): OperationLookupResult<TOperation, TResult, TPart, TShared> {
    const normalizedFormat = normalizeFormat(format);
    const opName = operationName(operation);
    if (!opName) {
      return {
        supported: false,
        format: normalizedFormat,
        reason: "missing-name",
        message: `Operation for ${normalizedFormat} is missing an op/type name.`
      };
    }

    const registration = this.get(normalizedFormat, opName);
    if (!registration) {
      return {
        supported: false,
        format: normalizedFormat,
        opName,
        reason: "unsupported",
        message: `Unsupported operation for ${normalizedFormat}: ${opName}.`
      };
    }

    return {
      supported: true,
      format: normalizedFormat,
      opName,
      registration
    };
  }

  require(format: OperationFormat, operation: TOperation | OperationName): RegisteredOperation<TOperation, TResult, TPart, TShared> {
    const result = this.lookup(format, operation);
    if (!result.supported) {
      throw new UnsupportedOperationError(result.format, result.opName, result.reason);
    }
    return result.registration;
  }

  list(format?: OperationFormat): RegisteredOperation<TOperation, TResult, TPart, TShared>[] {
    const normalizedFormat = format === undefined ? undefined : normalizeFormat(format);
    return [...this.operations.values()]
      .filter((registration) => normalizedFormat === undefined || registration.format === normalizedFormat)
      .sort((left, right) => left.key.localeCompare(right.key));
  }
}

export function createOperationRegistry<
  TOperation extends NamedOperation = NamedOperation,
  TResult extends OperationResult = OperationResult,
  TPart = Uint8Array | string,
  TShared = unknown
>(): OperationRegistry<TOperation, TResult, TPart, TShared> {
  return new OperationRegistry<TOperation, TResult, TPart, TShared>();
}

export function operationName(operation: NamedOperation | OperationName): OperationName | undefined {
  if (typeof operation === "string") {
    return normalizeOperationName(operation);
  }
  const name = typeof operation.op === "string" ? operation.op : operation.type;
  return typeof name === "string" ? normalizeOperationName(name) : undefined;
}

export function operationKey(format: OperationFormat, opName: OperationName): string {
  return `${normalizeFormat(format)}\u0000${normalizeOperationName(opName)}`;
}

function normalizeFormat(format: OperationFormat): OperationFormat {
  const normalized = format.trim().toLowerCase();
  if (!normalized) {
    throw new Error("Operation format must not be empty.");
  }
  return normalized;
}

function normalizeOperationName(opName: OperationName): OperationName {
  const normalized = opName.trim();
  if (!normalized) {
    throw new Error("Operation name must not be empty.");
  }
  return normalized;
}
