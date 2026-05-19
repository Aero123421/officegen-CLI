export class DuplicateOperationRegistrationError extends Error {
    format;
    opName;
    code = "DUPLICATE_OPERATION_REGISTRATION";
    constructor(format, opName) {
        super(`Operation handler already registered for ${format}.${opName}.`);
        this.format = format;
        this.opName = opName;
        this.name = "DuplicateOperationRegistrationError";
    }
}
export class UnsupportedOperationError extends Error {
    format;
    opName;
    code = "UNSUPPORTED_OPERATION";
    constructor(format, opName, reason = opName ? "unsupported" : "missing-name") {
        super(reason === "missing-name"
            ? `Operation for ${format} is missing an op/type name.`
            : `Unsupported operation for ${format}: ${opName}.`);
        this.format = format;
        this.opName = opName;
        this.name = "UnsupportedOperationError";
    }
}
export class OperationRegistry {
    operations = new Map();
    register(registration) {
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
    has(format, opName) {
        return this.operations.has(operationKey(normalizeFormat(format), normalizeOperationName(opName)));
    }
    get(format, opName) {
        return this.operations.get(operationKey(normalizeFormat(format), normalizeOperationName(opName)));
    }
    lookup(format, operation) {
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
    require(format, operation) {
        const result = this.lookup(format, operation);
        if (!result.supported) {
            throw new UnsupportedOperationError(result.format, result.opName, result.reason);
        }
        return result.registration;
    }
    list(format) {
        const normalizedFormat = format === undefined ? undefined : normalizeFormat(format);
        return [...this.operations.values()]
            .filter((registration) => normalizedFormat === undefined || registration.format === normalizedFormat)
            .sort((left, right) => left.key.localeCompare(right.key));
    }
}
export function createOperationRegistry() {
    return new OperationRegistry();
}
export function operationName(operation) {
    if (typeof operation === "string") {
        return normalizeOperationName(operation);
    }
    const name = typeof operation.op === "string" ? operation.op : operation.type;
    return typeof name === "string" ? normalizeOperationName(name) : undefined;
}
export function operationKey(format, opName) {
    return `${normalizeFormat(format)}\u0000${normalizeOperationName(opName)}`;
}
function normalizeFormat(format) {
    const normalized = format.trim().toLowerCase();
    if (!normalized) {
        throw new Error("Operation format must not be empty.");
    }
    return normalized;
}
function normalizeOperationName(opName) {
    const normalized = opName.trim();
    if (!normalized) {
        throw new Error("Operation name must not be empty.");
    }
    return normalized;
}
//# sourceMappingURL=registry.js.map