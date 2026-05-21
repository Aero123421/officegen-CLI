import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { detectOoxmlRiskyParts, validateOoxml } from "../src/index.js";

describe("OOXML validator", () => {
  it("detects a broken internal relationship target", async () => {
    const zip = minimalPptxZip();
    zip.file("ppt/_rels/presentation.xml.rels", relationships([
      { id: "rId1", target: "slides/missing-slide.xml" }
    ]));

    const result = await validateOoxml(await zipBytes(zip), { format: "pptx" });

    expect(result.ok).toBe(false);
    expect(result.issues).toEqual(expect.arrayContaining([
      expect.objectContaining({
        code: "OOXML_MISSING_INTERNAL_RELATIONSHIP_TARGET",
        path: "ppt/_rels/presentation.xml.rels",
        relationshipId: "rId1",
        resolvedTarget: "ppt/slides/missing-slide.xml"
      })
    ]));
  });

  it("detects duplicate relationship ids in a rels part", async () => {
    const zip = minimalPptxZip();
    zip.file("ppt/slides/slide1.xml", "<p:sld/>");
    zip.file("ppt/_rels/presentation.xml.rels", relationships([
      { id: "rId1", target: "slides/slide1.xml" },
      { id: "rId1", target: "slides/slide1.xml" }
    ]));

    const result = await validateOoxml(await zipBytes(zip), { format: "pptx" });

    expect(result.ok).toBe(false);
    expect(result.issues).toEqual(expect.arrayContaining([
      expect.objectContaining({
        code: "OOXML_DUPLICATE_RELATIONSHIP_ID",
        path: "ppt/_rels/presentation.xml.rels",
        relationshipId: "rId1"
      })
    ]));
  });

  it("detects external relationship targets and reports them as risky parts", async () => {
    const zip = minimalPptxZip();
    zip.file("ppt/slides/slide1.xml", "<p:sld/>");
    zip.file("ppt/_rels/presentation.xml.rels", relationships([
      { id: "rId1", target: "slides/slide1.xml" },
      { id: "rId2", target: "https://example.test/image.png", targetMode: "External" }
    ]));

    const result = await validateOoxml(await zipBytes(zip), { format: "pptx" });

    expect(result.ok).toBe(true);
    expect(result.issues).toEqual(expect.arrayContaining([
      expect.objectContaining({
        code: "OOXML_EXTERNAL_RELATIONSHIP_TARGET",
        severity: "warning",
        relationshipId: "rId2",
        target: "https://example.test/image.png"
      })
    ]));
    expect(result.riskyParts).toEqual(expect.arrayContaining([
      expect.objectContaining({
        kind: "externalRelationship",
        path: "ppt/_rels/presentation.xml.rels",
        relationshipId: "rId2"
      })
    ]));
  });

  it("detects malformed XML parts", async () => {
    const zip = minimalPptxZip();
    zip.file("ppt/slides/slide1.xml", "<p:sld><p:cSld></p:sld>");
    zip.file("ppt/_rels/presentation.xml.rels", relationships([
      { id: "rId1", target: "slides/slide1.xml" }
    ]));

    const result = await validateOoxml(await zipBytes(zip), { format: "pptx" });

    expect(result.ok).toBe(false);
    expect(result.issues).toEqual(expect.arrayContaining([
      expect.objectContaining({
        code: "OOXML_XML_NOT_WELL_FORMED",
        path: "ppt/slides/slide1.xml"
      })
    ]));
  });

  it("detects macro and embedded object risky parts", async () => {
    const zip = new JSZip();
    zip.file("[Content_Types].xml", "<Types/>");
    zip.file("_rels/.rels", relationships([
      { id: "rId1", target: "word/document.xml" }
    ]));
    zip.file("word/document.xml", "<w:document/>");
    zip.file("word/vbaProject.bin", new Uint8Array([1, 2, 3]));
    zip.file("word/embeddings/oleObject1.bin", new Uint8Array([4, 5, 6]));

    const riskyParts = await detectOoxmlRiskyParts(await zipBytes(zip), { format: "docx" });

    expect(riskyParts).toEqual(expect.arrayContaining([
      expect.objectContaining({ kind: "macro", path: "word/vbaProject.bin" }),
      expect.objectContaining({ kind: "embeddedObject", path: "word/embeddings/oleObject1.bin" })
    ]));
  });

  it("detects missing format-specific main parts", async () => {
    const zip = new JSZip();
    zip.file("[Content_Types].xml", "<Types/>");
    zip.file("_rels/.rels", relationships([
      { id: "rId1", target: "word/missing.xml" }
    ]));

    const result = await validateOoxml(await zipBytes(zip), { format: "docx" });

    expect(result.ok).toBe(false);
    expect(result.issues).toEqual(expect.arrayContaining([
      expect.objectContaining({
        code: "OOXML_MISSING_FORMAT_MAIN_PART",
        path: "word/document.xml"
      })
    ]));
  });
});

interface RelationshipFixture {
  id: string;
  target: string;
  targetMode?: string;
}

function minimalPptxZip(): JSZip {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", "<Types/>");
  zip.file("_rels/.rels", relationships([
    { id: "rId1", target: "ppt/presentation.xml" }
  ]));
  zip.file("ppt/presentation.xml", "<p:presentation/>");
  return zip;
}

function relationships(items: RelationshipFixture[]): string {
  return [
    "<Relationships>",
    ...items.map((item) => `<Relationship Id="${item.id}" Target="${item.target}"${item.targetMode ? ` TargetMode="${item.targetMode}"` : ""}/>`),
    "</Relationships>"
  ].join("");
}

function zipBytes(zip: JSZip): Promise<Uint8Array> {
  return zip.generateAsync({ type: "uint8array" });
}
