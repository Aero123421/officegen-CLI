import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { inspect } from "../src/inspect.js";

describe("DOCX story/run graph inspection", () => {
  it("exposes stories, paragraphs, split runs, and text tokens without changing paragraph objectMap entries", async () => {
    const inspected = await inspectDocx([
      '<w:p><w:r><w:t>Al</w:t></w:r><w:r><w:t>pha</w:t></w:r></w:p>'
    ]);
    const storyGraph = inspected.untrusted.storyGraph as any;
    const runGraph = inspected.untrusted.runGraph as any;
    const firstParagraph = inspected.untrusted.paragraphs[0] as any;
    const paragraphRuns = runGraph.nodes.filter((node: any) => node.type === "run" && node.paragraphId === firstParagraph.stableObjectId);
    const textTokens = runGraph.nodes.filter((node: any) => node.type === "text" && node.paragraphId === firstParagraph.stableObjectId);

    expect(storyGraph.graphVersion).toBe("officegen.docx.storyGraph@0.1");
    expect(storyGraph.stories[0]).toMatchObject({ kind: "document", paragraphCount: 1, runCount: 2, textTokenCount: 2 });
    expect(paragraphRuns).toHaveLength(2);
    expect(textTokens.map((node: any) => node.text)).toEqual(["Al", "pha"]);
    expect(inspected.objectMap.find((entry) => entry.kind === "paragraph" && entry.text === "Alpha")).toMatchObject({
      selectorHints: { story: "document", partKind: "body" }
    });
  });

  it("detects DOCX hyperlinks and bookmarks in the run graph", async () => {
    const inspected = await inspectDocx([
      '<w:p><w:hyperlink r:id="rId5"><w:r><w:t>Link</w:t></w:r></w:hyperlink></w:p>',
      '<w:p><w:bookmarkStart w:id="1" w:name="Here"/><w:r><w:t>Bookmarked</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>'
    ]);
    const nodes = (inspected.untrusted.runGraph as any).nodes;
    const hyperlink = nodes.find((node: any) => node.type === "hyperlink");
    const bookmarks = nodes.filter((node: any) => node.type === "bookmark");

    expect(hyperlink).toMatchObject({ text: "Link", attrs: { relationshipId: "rId5" } });
    expect(bookmarks.map((node: any) => node.attrs.marker)).toEqual(["start", "end"]);
    expect(bookmarks[0]).toMatchObject({ attrs: { id: "1", name: "Here" } });
  });

  it("detects DOCX comment ranges and comment story content", async () => {
    const inspected = await inspectDocx(
      [
        '<w:p><w:commentRangeStart w:id="7"/><w:r><w:t>Commented</w:t></w:r><w:commentRangeEnd w:id="7"/><w:r><w:commentReference w:id="7"/></w:r></w:p>'
      ],
      '<w:comments><w:comment w:id="7" w:author="QA"><w:p><w:r><w:t>Review note</w:t></w:r></w:p></w:comment></w:comments>'
    );
    const storyGraph = inspected.untrusted.storyGraph as any;
    const nodes = (inspected.untrusted.runGraph as any).nodes;

    expect(storyGraph.stories.map((story: any) => story.kind)).toEqual(["document", "comments"]);
    expect(nodes.filter((node: any) => node.type === "commentRange").map((node: any) => node.attrs.marker)).toEqual(["start", "end"]);
    expect(nodes.find((node: any) => node.type === "commentReference")).toMatchObject({ attrs: { id: "7" } });
    expect(nodes.find((node: any) => node.type === "comment")).toMatchObject({ text: "Review note", attrs: { id: "7", author: "QA" } });
  });

  it("detects DOCX revisions, content controls, and field codes", async () => {
    const inspected = await inspectDocx([
      '<w:p><w:fldChar w:fldCharType="begin"/><w:r><w:instrText> DATE </w:instrText></w:r><w:fldChar w:fldCharType="end"/></w:p>',
      '<w:p><w:ins w:id="9" w:author="QA"><w:r><w:t>Inserted</w:t></w:r></w:ins><w:del w:id="10" w:author="QA"><w:r><w:delText>Deleted</w:delText></w:r></w:del></w:p>',
      '<w:p><w:sdt><w:sdtPr><w:tag w:val="client"/></w:sdtPr><w:sdtContent><w:r><w:t>Client</w:t></w:r></w:sdtContent></w:sdt></w:p>'
    ]);
    const nodes = (inspected.untrusted.runGraph as any).nodes;
    const fields = nodes.filter((node: any) => node.type === "field");
    const revisions = nodes.filter((node: any) => node.type === "revision");

    expect(fields.some((node: any) => node.attrs.fldCharType === "begin")).toBe(true);
    expect(fields.some((node: any) => node.text === " DATE " && node.attrs.fieldKind === "instrText")).toBe(true);
    expect(revisions.map((node: any) => node.attrs.revisionType)).toEqual(["ins", "del"]);
    expect(revisions.map((node: any) => node.text)).toEqual(["Inserted", "Deleted"]);
    expect(nodes.find((node: any) => node.type === "contentControl")).toMatchObject({ text: "Client" });
  });
});

async function inspectDocx(paragraphs: string[], commentsXml?: string) {
  const zip = new JSZip();
  zip.file("word/document.xml", `<w:document><w:body>${paragraphs.join("")}</w:body></w:document>`);
  if (commentsXml) zip.file("word/comments.xml", commentsXml);
  return inspect({ data: await zip.generateAsync({ type: "uint8array" }), format: "docx" }, { depth: "full" });
}
