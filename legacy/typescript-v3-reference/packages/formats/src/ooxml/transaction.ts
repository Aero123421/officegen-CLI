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
  errors: Array<{ path: string; error: unknown }>;
}

export class EditTransactionError extends Error {
  constructor(message: string, public readonly code: string) {
    super(message);
    this.name = "EditTransactionError";
  }
}

export class EditTransaction<TPart = Uint8Array | string> {
  private readonly journal = new Map<string, PartJournalEntry<TPart>>();
  private readonly clonePart: (value: TPart) => TPart;
  private closed = false;
  private nextSequence = 0;

  constructor(
    private readonly store: EditPartStore<TPart>,
    private readonly options: EditTransactionOptions<TPart> = {}
  ) {
    this.clonePart = options.clonePart ?? clonePartValue;
  }

  get mode(): EditTransactionMode {
    return this.options.atomic === false ? "best-effort" : "atomic";
  }

  get atomic(): boolean {
    return this.mode === "atomic";
  }

  get continueOnError(): boolean {
    return this.options.continueOnError === true;
  }

  get closedForWrites(): boolean {
    return this.closed;
  }

  get journaledParts(): number {
    return this.journal.size;
  }

  snapshot(): PartJournalEntry<TPart>[] {
    return [...this.journal.values()]
      .sort((left, right) => left.sequence - right.sequence)
      .map((entry) => ({
        ...entry,
        value: entry.existed ? this.clonePart(entry.value as TPart) : undefined
      }));
  }

  async snapshotPart(path: string): Promise<PartSnapshot<TPart>> {
    const normalizedPath = normalizePartPath(path);
    const existing = this.journal.get(normalizedPath);
    if (existing) {
      return {
        path: existing.path,
        existed: existing.existed,
        value: existing.existed ? this.clonePart(existing.value as TPart) : undefined
      };
    }

    const current = await this.store.readPart(normalizedPath);
    const entry: PartJournalEntry<TPart> = {
      path: normalizedPath,
      existed: current !== undefined,
      value: current === undefined ? undefined : this.clonePart(current),
      sequence: this.nextSequence++
    };
    this.journal.set(normalizedPath, entry);
    return {
      path: entry.path,
      existed: entry.existed,
      value: entry.existed ? this.clonePart(entry.value as TPart) : undefined
    };
  }

  async readPart(path: string): Promise<TPart | undefined> {
    return this.store.readPart(normalizePartPath(path));
  }

  async writePart(path: string, value: TPart): Promise<void> {
    this.assertWritable();
    const normalizedPath = normalizePartPath(path);
    await this.snapshotPart(normalizedPath);
    await this.store.writePart(normalizedPath, value);
  }

  async deletePart(path: string): Promise<void> {
    this.assertWritable();
    if (!this.store.deletePart) {
      throw new EditTransactionError(
        "The part store does not support deleting parts.",
        "PART_DELETE_UNSUPPORTED"
      );
    }
    const normalizedPath = normalizePartPath(path);
    await this.snapshotPart(normalizedPath);
    await this.store.deletePart(normalizedPath);
  }

  async run<T>(operationIndex: number, operation: () => MaybePromise<T>): Promise<TransactionOperationResult<T>> {
    this.assertWritable();
    try {
      const value = await operation();
      return {
        operationIndex,
        applied: true,
        rolledBack: false,
        value
      };
    } catch (error) {
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

  async commit(): Promise<EditTransactionCommitResult> {
    this.assertWritable();
    const journaledParts = this.journal.size;
    this.journal.clear();
    this.closed = true;
    return { committed: true, journaledParts };
  }

  async rollback(): Promise<EditTransactionRollbackResult> {
    if (this.closed) {
      return { rolledBack: true, restoredParts: 0, errors: [] };
    }

    const entries = [...this.journal.values()].sort((left, right) => right.sequence - left.sequence);
    const errors: Array<{ path: string; error: unknown }> = [];
    let restoredParts = 0;

    for (const entry of entries) {
      try {
        if (entry.existed) {
          await this.store.writePart(entry.path, this.clonePart(entry.value as TPart));
        } else if (this.store.deletePart) {
          await this.store.deletePart(entry.path);
        } else {
          throw new EditTransactionError(
            `Cannot remove newly-created part ${entry.path}; the part store does not support deleting parts.`,
            "PART_DELETE_UNSUPPORTED"
          );
        }
        restoredParts += 1;
      } catch (error) {
        errors.push({ path: entry.path, error });
      }
    }

    if (errors.length === 0) {
      this.journal.clear();
      this.closed = true;
    }

    return { rolledBack: true, restoredParts, errors };
  }

  private assertWritable(): void {
    if (this.closed) {
      throw new EditTransactionError(
        "The edit transaction is already closed.",
        "TRANSACTION_CLOSED"
      );
    }
  }
}

export function createEditTransaction<TPart>(
  store: EditPartStore<TPart>,
  options: EditTransactionOptions<TPart> = {}
): EditTransaction<TPart> {
  return new EditTransaction(store, options);
}

function normalizePartPath(path: string): string {
  const normalized = path.replace(/\\/g, "/").replace(/^\/+/, "");
  if (!normalized) {
    throw new EditTransactionError("Part path must not be empty.", "EMPTY_PART_PATH");
  }
  return normalized;
}

function clonePartValue<TPart>(value: TPart): TPart {
  if (value instanceof Uint8Array) {
    return new Uint8Array(value) as TPart;
  }
  if (value instanceof ArrayBuffer) {
    return value.slice(0) as TPart;
  }
  return value;
}
