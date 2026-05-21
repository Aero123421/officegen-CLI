import { describe, expect, it } from "vitest";
import {
  DuplicateOperationRegistrationError,
  OperationRegistry,
  UnsupportedOperationError,
  operationName
} from "../src/ooxml/operations/registry.js";
import type { NamedOperation, OperationContext, OperationResult } from "../src/ooxml/operations/types.js";

describe("OperationRegistry", () => {
  it("registers handlers and resolves them by format and op name", async () => {
    const registry = new OperationRegistry<TestOperation, OperationResult>();
    const handler = async (operation: TestOperation, _context: OperationContext): Promise<OperationResult> => ({
      applied: true,
      changed: operation.text.length > 0
    });

    registry.register({ format: "pptx", opName: "pptx.setText", handler });

    expect(registry.has("PPTX", "pptx.setText")).toBe(true);
    expect(registry.get("pptx", "pptx.setText")?.handler).toBe(handler);
    expect(registry.list("pptx").map((entry) => entry.opName)).toEqual(["pptx.setText"]);
  });

  it("looks up op and type based operations", () => {
    const registry = new OperationRegistry();
    const handler = () => ({ applied: true });
    registry
      .register({ format: "docx", opName: "docx.setHeader", handler })
      .register({ format: "docx", opName: "replaceText", handler });

    expect(registry.lookup("docx", { op: "docx.setHeader" })).toMatchObject({
      supported: true,
      format: "docx",
      opName: "docx.setHeader"
    });
    expect(registry.lookup("docx", { type: "replaceText" })).toMatchObject({
      supported: true,
      format: "docx",
      opName: "replaceText"
    });
    expect(operationName({ type: "replaceText" })).toBe("replaceText");
  });

  it("returns unsupported lookup results and throws from require", () => {
    const registry = new OperationRegistry();

    expect(registry.lookup("xlsx", { op: "xlsx.nope" })).toEqual({
      supported: false,
      format: "xlsx",
      opName: "xlsx.nope",
      reason: "unsupported",
      message: "Unsupported operation for xlsx: xlsx.nope."
    });
    expect(registry.lookup("xlsx", {})).toEqual({
      supported: false,
      format: "xlsx",
      reason: "missing-name",
      message: "Operation for xlsx is missing an op/type name."
    });
    expect(() => registry.require("xlsx", { op: "xlsx.nope" })).toThrow(UnsupportedOperationError);
  });

  it("rejects duplicate registrations for the same format and op name", () => {
    const registry = new OperationRegistry();
    const handler = () => ({ applied: true });

    registry.register({ format: "pptx", opName: "pptx.setText", handler });

    expect(() => registry.register({ format: "PPTX", opName: "pptx.setText", handler }))
      .toThrow(DuplicateOperationRegistrationError);
  });
});

interface TestOperation extends NamedOperation {
  op: "pptx.setText";
  text: string;
}
