import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { PackageGraph } from "../src/ooxml/packageGraph.js";

describe("PackageGraph", () => {
  it("indexes unknown parts and preserves them when editing known text parts", async () => {
    const zip = minimalPptxZip();
    zip.file("vendor/private.bin", new Uint8Array([7, 8, 9]));

    const graph = await PackageGraph.fromZip(zip, { format: "pptx" });

    expect(graph.getPart("vendor/private.bin")).toEqual(expect.objectContaining({
      path: "vendor/private.bin",
      unknown: true
    }));

    await graph.writeText("ppt/presentation.xml", "<p:presentation updated=\"1\"/>");
    const roundTripped = await JSZip.loadAsync(await zipBytes(zip));

    expect(await roundTripped.file("ppt/presentation.xml")?.async("string")).toBe("<p:presentation updated=\"1\"/>");
    expect(await roundTripped.file("vendor/private.bin")?.async("uint8array")).toEqual(new Uint8Array([7, 8, 9]));
  });

  it("resolves internal relationship targets from the owning part", async () => {
    const zip = minimalPptxZip();

    const graph = await PackageGraph.fromZip(zip, { format: "pptx" });
    const rel = graph.resolveRelationship("ppt/presentation.xml", "rId1");

    expect(rel).toEqual(expect.objectContaining({
      ownerPath: "ppt/presentation.xml",
      path: "ppt/_rels/presentation.xml.rels",
      target: "slides/slide1.xml",
      resolvedTarget: "ppt/slides/slide1.xml",
      external: false
    }));
    expect(graph.validate().ok).toBe(true);
  });

  it("writes content type overrides without publishing a public API", async () => {
    const zip = minimalPptxZip();
    const graph = await PackageGraph.fromZip(zip, { format: "pptx" });

    await graph.writeText("ppt/slides/slide2.xml", "<p:sld/>");
    await graph.ensureContentTypeOverride(
      "/ppt/slides/slide2.xml",
      "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
    );

    expect(graph.getPart("ppt/slides/slide2.xml")).toEqual(expect.objectContaining({
      contentType: "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
    }));
    expect(await graph.readText("[Content_Types].xml")).toContain('PartName="/ppt/slides/slide2.xml"');
  });

  it("classifies external relationships and records security flags", async () => {
    const zip = minimalPptxZip();
    const graph = await PackageGraph.fromZip(zip, { format: "pptx" });

    const rel = await graph.ensureRelationship(
      "ppt/presentation.xml",
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
      "https://example.test/deck",
      "External"
    );

    expect(rel).toEqual(expect.objectContaining({
      external: true,
      resolvedTarget: undefined,
      targetMode: "External"
    }));
    expect(graph.resolveRelationship("ppt/presentation.xml", rel.id)).toEqual(expect.objectContaining({
      external: true,
      target: "https://example.test/deck"
    }));

    const validation = graph.validate();
    expect(validation.ok).toBe(true);
    expect(validation.summary.externalRelationships).toBe(1);
    expect(validation.securityFlags).toEqual(expect.arrayContaining([
      expect.objectContaining({
        kind: "externalRelationship",
        path: "ppt/_rels/presentation.xml.rels",
        relationshipId: rel.id,
        target: "https://example.test/deck"
      })
    ]));
  });
});

interface RelationshipFixture {
  id: string;
  type?: string;
  target: string;
  targetMode?: string;
}

function minimalPptxZip(): JSZip {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", [
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
    '<Default Extension="xml" ContentType="application/xml"/>',
    '<Default Extension="bin" ContentType="application/octet-stream"/>',
    '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>',
    '<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>',
    "</Types>"
  ].join(""));
  zip.file("_rels/.rels", relationships([
    {
      id: "rId1",
      type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
      target: "ppt/presentation.xml"
    }
  ]));
  zip.file("ppt/presentation.xml", "<p:presentation/>");
  zip.file("ppt/slides/slide1.xml", "<p:sld/>");
  zip.file("ppt/_rels/presentation.xml.rels", relationships([
    {
      id: "rId1",
      type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
      target: "slides/slide1.xml"
    }
  ]));
  return zip;
}

function relationships(items: RelationshipFixture[]): string {
  return [
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    ...items.map((item) => [
      `<Relationship Id="${item.id}"`,
      ` Type="${item.type ?? "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"}"`,
      ` Target="${item.target}"`,
      item.targetMode ? ` TargetMode="${item.targetMode}"` : "",
      "/>"
    ].join("")),
    "</Relationships>"
  ].join("");
}

function zipBytes(zip: JSZip): Promise<Uint8Array> {
  return zip.generateAsync({ type: "uint8array" });
}
