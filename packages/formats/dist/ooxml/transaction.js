export class EditTransactionError extends Error {
    code;
    constructor(message, code) {
        super(message);
        this.code = code;
        this.name = "EditTransactionError";
    }
}
export class EditTransaction {
    store;
    options;
    journal = new Map();
    clonePart;
    closed = false;
    nextSequence = 0;
    constructor(store, options = {}) {
        this.store = store;
        this.options = options;
        this.clonePart = options.clonePart ?? clonePartValue;
    }
    get mode() {
        return this.options.atomic === false ? "best-effort" : "atomic";
    }
    get atomic() {
        return this.mode === "atomic";
    }
    get continueOnError() {
        return this.options.continueOnError === true;
    }
    get closedForWrites() {
        return this.closed;
    }
    get journaledParts() {
        return this.journal.size;
    }
    snapshot() {
        return [...this.journal.values()]
            .sort((left, right) => left.sequence - right.sequence)
            .map((entry) => ({
            ...entry,
            value: entry.existed ? this.clonePart(entry.value) : undefined
        }));
    }
    async snapshotPart(path) {
        const normalizedPath = normalizePartPath(path);
        const existing = this.journal.get(normalizedPath);
        if (existing) {
            return {
                path: existing.path,
                existed: existing.existed,
                value: existing.existed ? this.clonePart(existing.value) : undefined
            };
        }
        const current = await this.store.readPart(normalizedPath);
        const entry = {
            path: normalizedPath,
            existed: current !== undefined,
            value: current === undefined ? undefined : this.clonePart(current),
            sequence: this.nextSequence++
        };
        this.journal.set(normalizedPath, entry);
        return {
            path: entry.path,
            existed: entry.existed,
            value: entry.existed ? this.clonePart(entry.value) : undefined
        };
    }
    async readPart(path) {
        return this.store.readPart(normalizePartPath(path));
    }
    async writePart(path, value) {
        this.assertWritable();
        const normalizedPath = normalizePartPath(path);
        await this.snapshotPart(normalizedPath);
        await this.store.writePart(normalizedPath, value);
    }
    async deletePart(path) {
        this.assertWritable();
        if (!this.store.deletePart) {
            throw new EditTransactionError("The part store does not support deleting parts.", "PART_DELETE_UNSUPPORTED");
        }
        const normalizedPath = normalizePartPath(path);
        await this.snapshotPart(normalizedPath);
        await this.store.deletePart(normalizedPath);
    }
    async run(operationIndex, operation) {
        this.assertWritable();
        try {
            const value = await operation();
            return {
                operationIndex,
                applied: true,
                rolledBack: false,
                value
            };
        }
        catch (error) {
            let rolledBack = false;
            if (this.atomic && !this.continueOnError) {
                await this.rollback();
                rolledBack = true;
            }
            return {
                operationIndex,
                applied: false,
                rolledBack,
                error
            };
        }
    }
    async commit() {
        this.assertWritable();
        const journaledParts = this.journal.size;
        this.journal.clear();
        this.closed = true;
        return { committed: true, journaledParts };
    }
    async rollback() {
        if (this.closed) {
            return { rolledBack: true, restoredParts: 0, errors: [] };
        }
        const entries = [...this.journal.values()].sort((left, right) => right.sequence - left.sequence);
        const errors = [];
        let restoredParts = 0;
        for (const entry of entries) {
            try {
                if (entry.existed) {
                    await this.store.writePart(entry.path, this.clonePart(entry.value));
                }
                else if (this.store.deletePart) {
                    await this.store.deletePart(entry.path);
                }
                else {
                    throw new EditTransactionError(`Cannot remove newly-created part ${entry.path}; the part store does not support deleting parts.`, "PART_DELETE_UNSUPPORTED");
                }
                restoredParts += 1;
            }
            catch (error) {
                errors.push({ path: entry.path, error });
            }
        }
        if (errors.length === 0) {
            this.journal.clear();
            this.closed = true;
        }
        return { rolledBack: true, restoredParts, errors };
    }
    assertWritable() {
        if (this.closed) {
            throw new EditTransactionError("The edit transaction is already closed.", "TRANSACTION_CLOSED");
        }
    }
}
export function createEditTransaction(store, options = {}) {
    return new EditTransaction(store, options);
}
function normalizePartPath(path) {
    const normalized = path.replace(/\\/g, "/").replace(/^\/+/, "");
    if (!normalized) {
        throw new EditTransactionError("Part path must not be empty.", "EMPTY_PART_PATH");
    }
    return normalized;
}
function clonePartValue(value) {
    if (value instanceof Uint8Array) {
        return new Uint8Array(value);
    }
    if (value instanceof ArrayBuffer) {
        return value.slice(0);
    }
    return value;
}
//# sourceMappingURL=transaction.js.map