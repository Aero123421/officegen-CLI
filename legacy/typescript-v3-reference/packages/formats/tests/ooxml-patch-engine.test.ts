import { describe, expect, it } from "vitest";
import { applyXmlPatches, PatchEngineError } from "../src/ooxml/patchEngine.js";
import { createSourceFingerprint } from "../src/ooxml/sourceSpan.js";
import { buildTokenIndex } from "../src/ooxml/tokenIndex.js";

describe("OOXML patch engine", () => {
  it("replaces a token span while preserving surrounding XML byte-for-byte", () => {
    const xml = "<root><!--keep--><w:t id=\"1\">Old</w:t><![CDATA[x]]><w:br/></root>";
    const index = buildTokenIndex(xml);
    const oldText = index.textRuns.find((run) => run.text === "Old");
    expect(oldText).toBeDefined();

    const result = applyXmlPatches(xml, [{
      type: "replace",
      span: oldText!.span,
      value: "New",
      fingerprint: createSourceFingerprint(xml, oldText!.span)
    }]);

    expect(result).toBe("<root><!--keep--><w:t id=\"1\">New</w:t><![CDATA[x]]><w:br/></root>");
    expect(result.slice(0, result.indexOf("New"))).toBe(xml.slice(0, xml.indexOf("Old")));
    expect(result.slice(result.indexOf("New") + 3)).toBe(xml.slice(xml.indexOf("Old") + 3));
  });

  it("applies delete and insert patches in offset order", () => {
    const xml = "<root><w:t>One</w:t><w:br/></root>";
    const index = buildTokenIndex(xml);
    const br = index.findElementsByName("w:br")[0];
    const root = index.findElementsByName("root")[0];
    expect(br).toBeDefined();
    expect(root?.closeTagSpan).toBeDefined();

    const result = applyXmlPatches(xml, [
      { type: "insert", offset: root!.closeTagSpan!.start, value: "<w:t>Two</w:t>" },
      { type: "delete", span: br!.span, fingerprint: createSourceFingerprint(xml, br!.span) }
    ]);

    expect(result).toBe("<root><w:t>One</w:t><w:t>Two</w:t></root>");
  });

  it("rejects overlapping patches", () => {
    const xml = "<root><w:t>Old</w:t></root>";
    const text = buildTokenIndex(xml).textRuns[0];
    expect(text).toBeDefined();

    expect(() => applyXmlPatches(xml, [
      { type: "replace", span: text!.span, value: "New" },
      { type: "delete", span: text!.span }
    ])).toThrow(PatchEngineError);
  });

  it("rejects stale fingerprints", () => {
    const xml = "<root><w:t>Old</w:t></root>";
    const text = buildTokenIndex(xml).textRuns[0];
    const fingerprint = createSourceFingerprint(xml, text!.span);
    const changed = "<root><w:t>New</w:t></root>";
    const changedText = buildTokenIndex(changed).textRuns[0];

    expect(() => applyXmlPatches(changed, [{
      type: "replace",
      span: changedText!.span,
      value: "Later",
      fingerprint
    }])).toThrow(/stale/i);
  });

  it("rejects patches that produce malformed XML", () => {
    const xml = "<root><w:t>Old</w:t></root>";
    const text = buildTokenIndex(xml).textRuns[0];

    expect(() => applyXmlPatches(xml, [{
      type: "replace",
      span: text!.span,
      value: "<bad>"
    }])).toThrow(/not well-formed/i);
  });

  it("rejects patches that produce unescaped entity references", () => {
    const xml = "<root><w:t>Old</w:t></root>";
    const text = buildTokenIndex(xml).textRuns[0];

    expect(() => applyXmlPatches(xml, [{
      type: "replace",
      span: text!.span,
      value: "Tom & Jerry"
    }])).toThrow(/not well-formed/i);
  });
});
