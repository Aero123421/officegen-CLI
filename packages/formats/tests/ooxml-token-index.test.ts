import { describe, expect, it } from "vitest";
import { createSourceFingerprint, sliceSource } from "../src/ooxml/sourceSpan.js";
import { buildTokenIndex } from "../src/ooxml/tokenIndex.js";

describe("OOXML token index", () => {
  it("indexes elements, attributes, text, CDATA, comments, PI, and empty elements without rewriting source", () => {
    const xml = "<?xml version=\"1.0\"?><w:document xmlns:w=\"urn:w\" xmlns:a='urn:a'><!--keep--><?mso test?><w:body><a:t data-id=\"1\">Alpha</a:t><![CDATA[<raw>&text]]><w:br/><p:t>Beta</p:t></w:body></w:document>";

    const index = buildTokenIndex(xml);

    expect(index.findElementsByName("w:document")).toHaveLength(1);
    expect(index.findElementsByLocalName("t").map((element) => element.name)).toEqual(["a:t", "p:t"]);

    const document = index.findElementsByName("w:document")[0];
    expect(document?.prefix).toBe("w");
    expect(document?.localName).toBe("document");
    expect(document ? index.sourceFor(document.openTagSpan) : "").toBe("<w:document xmlns:w=\"urn:w\" xmlns:a='urn:a'>");

    const xmlnsA = index.findAttributesByName("xmlns:a")[0];
    expect(xmlnsA?.quote).toBe("'");
    expect(xmlnsA ? index.sourceFor(xmlnsA.rawValueSpan) : "").toBe("'urn:a'");
    expect(xmlnsA ? index.sourceFor(xmlnsA.valueSpan) : "").toBe("urn:a");

    const textRuns = index.textRuns.map((run) => ({ text: run.text, cdata: run.cdata, source: index.sourceFor(run.span) }));
    expect(textRuns).toEqual([
      { text: "Alpha", cdata: false, source: "Alpha" },
      { text: "<raw>&text", cdata: true, source: "<![CDATA[<raw>&text]]>" },
      { text: "Beta", cdata: false, source: "Beta" }
    ]);

    const empty = index.findElementsByName("w:br")[0];
    expect(empty?.selfClosing).toBe(true);
    expect(empty ? index.sourceFor(empty.span) : "").toBe("<w:br/>");
  });

  it("keeps byte offsets distinct from JavaScript character offsets", () => {
    const xml = "<w:t>日本語</w:t>";
    const index = buildTokenIndex(xml);
    const text = index.textRuns[0];

    expect(text?.text).toBe("日本語");
    expect(text ? text.span.end - text.span.start : 0).toBe(9);
    expect(text ? text.span.charEnd - text.span.charStart : 0).toBe(3);
    expect(text ? sliceSource(xml, text.span) : "").toBe("日本語");

    const fingerprint = text ? createSourceFingerprint(xml, text.span) : undefined;
    expect(fingerprint?.byteLength).toBe(9);
  });
});
