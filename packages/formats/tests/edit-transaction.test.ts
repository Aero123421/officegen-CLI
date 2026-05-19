import { describe, expect, it } from "vitest";
import { EditTransaction, type EditPartStore } from "../src/ooxml/transaction.js";

describe("EditTransaction", () => {
  it("snapshots parts and rolls back writes, creates, and deletes", async () => {
    const store = new MemoryPartStore<string>([
      ["ppt/slides/slide1.xml", "<slide>old</slide>"],
      ["ppt/slides/slide2.xml", "<slide>remove</slide>"]
    ]);
    const transaction = new EditTransaction(store);

    await transaction.writePart("ppt/slides/slide1.xml", "<slide>new</slide>");
    await transaction.writePart("ppt/slides/slide3.xml", "<slide>created</slide>");
    await transaction.deletePart("ppt/slides/slide2.xml");

    expect(await store.readPart("ppt/slides/slide1.xml")).toBe("<slide>new</slide>");
    expect(await store.readPart("ppt/slides/slide2.xml")).toBeUndefined();
    expect(await store.readPart("ppt/slides/slide3.xml")).toBe("<slide>created</slide>");

    const rollback = await transaction.rollback();

    expect(rollback).toMatchObject({ rolledBack: true, restoredParts: 3, errors: [] });
    expect(await store.readPart("ppt/slides/slide1.xml")).toBe("<slide>old</slide>");
    expect(await store.readPart("ppt/slides/slide2.xml")).toBe("<slide>remove</slide>");
    expect(await store.readPart("ppt/slides/slide3.xml")).toBeUndefined();
  });

  it("keeps the first snapshot when a part is written multiple times", async () => {
    const store = new MemoryPartStore<string>([["word/document.xml", "before"]]);
    const transaction = new EditTransaction(store);

    await transaction.writePart("word/document.xml", "during");
    await transaction.writePart("word/document.xml", "after");
    await transaction.rollback();

    expect(await store.readPart("word/document.xml")).toBe("before");
    expect(transaction.journaledParts).toBe(0);
  });

  it("rolls back atomic run failures", async () => {
    const store = new MemoryPartStore<string>([["xl/workbook.xml", "before"]]);
    const transaction = new EditTransaction(store, { atomic: true });

    const result = await transaction.run(0, async () => {
      await transaction.writePart("xl/workbook.xml", "changed");
      throw new Error("boom");
    });

    expect(result.applied).toBe(false);
    expect(result.rolledBack).toBe(true);
    expect(await store.readPart("xl/workbook.xml")).toBe("before");
  });

  it("leaves rollback to the caller when continueOnError is enabled", async () => {
    const store = new MemoryPartStore<string>([["xl/workbook.xml", "before"]]);
    const transaction = new EditTransaction(store, { atomic: true, continueOnError: true });

    const result = await transaction.run(0, async () => {
      await transaction.writePart("xl/workbook.xml", "changed");
      throw new Error("boom");
    });

    expect(result.applied).toBe(false);
    expect(result.rolledBack).toBe(false);
    expect(await store.readPart("xl/workbook.xml")).toBe("changed");

    await transaction.rollback();
    expect(await store.readPart("xl/workbook.xml")).toBe("before");
  });

  it("does not auto-rollback failed best-effort operations", async () => {
    const store = new MemoryPartStore<string>([["word/document.xml", "before"]]);
    const transaction = new EditTransaction(store, { atomic: false });

    const result = await transaction.run(1, async () => {
      await transaction.writePart("word/document.xml", "changed");
      throw new Error("keep going");
    });

    expect(transaction.mode).toBe("best-effort");
    expect(result).toMatchObject({ operationIndex: 1, applied: false, rolledBack: false });
    expect(await store.readPart("word/document.xml")).toBe("changed");
  });

  it("clones Uint8Array snapshots so later mutations cannot corrupt rollback", async () => {
    const original = new Uint8Array([1, 2, 3]);
    const replacement = new Uint8Array([4, 5, 6]);
    const store = new MemoryPartStore<Uint8Array>([["ppt/media/image1.png", original]]);
    const transaction = new EditTransaction(store);

    await transaction.snapshotPart("ppt/media/image1.png");
    original[0] = 9;
    await transaction.writePart("ppt/media/image1.png", replacement);
    await transaction.rollback();

    expect([...(await store.readPart("ppt/media/image1.png") ?? [])]).toEqual([1, 2, 3]);
  });
});

class MemoryPartStore<TPart> implements EditPartStore<TPart> {
  private readonly parts = new Map<string, TPart>();

  constructor(entries: Array<[string, TPart]> = []) {
    for (const [path, value] of entries) {
      this.parts.set(path, value);
    }
  }

  readPart(path: string): TPart | undefined {
    return this.parts.get(path);
  }

  writePart(path: string, value: TPart): void {
    this.parts.set(path, value);
  }

  deletePart(path: string): void {
    this.parts.delete(path);
  }
}
