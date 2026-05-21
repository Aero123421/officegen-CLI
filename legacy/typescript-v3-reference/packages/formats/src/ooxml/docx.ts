import type JSZip from "jszip";
import type { ObjectMapEntry } from "../shared.js";
import { decodeXmlEntities, readZipText, sortedZipFiles, stableHashId } from "../shared.js";
import { localText, paragraphXml, preview, replaceNthBlock, setFirstTextInBlock, xmlAttr } from "./xml.js";
import { parseRelationships, relationshipTarget } from "./relationships.js";

export const DOCX_STORY_GRAPH_VERSION = "officegen.docx.storyGraph@0.1" as const;
export const DOCX_RUN_GRAPH_VERSION = "officegen.docx.runGraph@0.1" as const;

export type DocxStoryKind = "document" | "header" | "footer" | "comments";
export type DocxRunGraphNodeType =
  | "story"
  | "paragraph"
  | "run"
  | "text"
  | "field"
  | "hyperlink"
  | "bookmark"
  | "commentRange"
  | "commentReference"
  | "comment"
  | "contentControl"
  | "revision";

export interface DocxStoryNode {
  storyId: string;
  kind: DocxStoryKind;
  partKind: DocxParagraph["partKind"];
  sourcePath: string;
  index: number;
  paragraphIds: string[];
  paragraphCount: number;
  runCount: number;
  textTokenCount: number;
  markers: Record<string, number>;
  untrusted: true;
}

export interface DocxStoryGraph {
  graphVersion: typeof DOCX_STORY_GRAPH_VERSION;
  stories: DocxStoryNode[];
  edges: Array<{ from: string; to: string; relation: "contains"; untrusted: true }>;
  summary: Record<string, number>;
  untrusted: true;
}

export interface DocxRunGraphNode {
  nodeId: string;
  stableObjectId: string;
  type: DocxRunGraphNodeType;
  storyId: string;
  storyKind: DocxStoryKind;
  paragraphId?: string;
  runId?: string;
  sourcePath: string;
  index?: number;
  text?: string;
  textPreview?: string;
  attrs?: Record<string, string | undefined>;
  untrusted: true;
}

export interface DocxRunGraph {
  graphVersion: typeof DOCX_RUN_GRAPH_VERSION;
  nodes: DocxRunGraphNode[];
  edges: Array<{ from: string; to: string; relation: "contains" | "next"; untrusted: true }>;
  summary: Record<string, number>;
  untrusted: true;
}

export interface DocxParagraph {
  stableObjectId: string;
  index: number;
  text: string;
  sourcePath: string;
  partKind: "body" | "header" | "footer" | "comment";
  storyId?: string;
  storyKind?: DocxStoryKind;
  runCount?: number;
  textTokenCount?: number;
  markers?: Record<string, number>;
  untrusted: true;
}

export async function inspectParagraphs(zip: JSZip): Promise<{ paragraphs: DocxParagraph[]; objectMap: ObjectMapEntry[]; storyGraph: DocxStoryGraph; runGraph: DocxRunGraph }> {
  const paths = sortedZipFiles(zip);
  const docxParts = [
    "word/document.xml",
    ...paths.filter((path) => /^word\/header\d+\.xml$/i.test(path)).sort(),
    ...paths.filter((path) => /^word\/footer\d+\.xml$/i.test(path)).sort(),
    ...paths.filter((path) => /^word\/comments\.xml$/i.test(path)).sort()
  ];
  const paragraphs: DocxParagraph[] = [];
  const extraObjectMap: ObjectMapEntry[] = [];
  const storyGraph = emptyStoryGraph();
  const runGraph = emptyRunGraph();
  for (const partPath of docxParts) {
    const xml = (await readZipText(zip, partPath)) ?? "";
    const partKind = docxPartKind(partPath);
    const storyKind = docxStoryKind(partPath);
    const storyId = stableHashId("docx", "story", storyKind, partPath);
    const storyNode: DocxStoryNode = {
      storyId,
      kind: storyKind,
      partKind,
      sourcePath: partPath,
      index: storyGraph.stories.length + 1,
      paragraphIds: [],
      paragraphCount: 0,
      runCount: 0,
      textTokenCount: 0,
      markers: {},
      untrusted: true
    };
    storyGraph.stories.push(storyNode);
    addRunNode(runGraph, {
      nodeId: storyId,
      stableObjectId: storyId,
      type: "story",
      storyId,
      storyKind,
      sourcePath: partPath,
      index: storyNode.index,
      untrusted: true
    });
    for (const [index, match] of [...xml.matchAll(/<w:p\b[\s\S]*?<\/w:p>/g)].entries()) {
      const paragraphXml = match[0];
      const stableObjectId = stableHashId("docx", partKind, "paragraph", `${partPath}#${index + 1}`);
      const paragraphGraph = inspectParagraphGraph(paragraphXml, {
        partPath,
        partKind,
        storyKind,
        storyId,
        stableObjectId,
        paragraphIndex: index + 1
      });
      paragraphs.push({
        stableObjectId,
        index: index + 1,
        text: localText(paragraphXml, "t").join(""),
        sourcePath: partPath,
        partKind,
        storyId,
        storyKind,
        runCount: paragraphGraph.runCount,
        textTokenCount: paragraphGraph.textTokenCount,
        markers: paragraphGraph.markers,
        untrusted: true as const
      });
      storyNode.paragraphIds.push(stableObjectId);
      storyNode.paragraphCount += 1;
      storyNode.runCount += paragraphGraph.runCount;
      storyNode.textTokenCount += paragraphGraph.textTokenCount;
      mergeMarkerCounts(storyNode.markers, paragraphGraph.markers);
      storyGraph.edges.push({ from: storyId, to: stableObjectId, relation: "contains", untrusted: true });
      appendRunGraph(runGraph, paragraphGraph.runGraph);
    }
    for (const commentMatch of xml.matchAll(/<w:comment\b([^>]*)>[\s\S]*?<\/w:comment>/g)) {
      const attrs = commentMatch[1] ?? "";
      const commentId = xmlAttr(attrs, "w:id") ?? xmlAttr(attrs, "id");
      const nodeId = stableHashId("docx", "comment", commentId ?? String(commentMatch.index ?? 0), partPath);
      addRunNode(runGraph, {
        nodeId,
        stableObjectId: nodeId,
        type: "comment",
        storyId,
        storyKind,
        sourcePath: partPath,
        text: docxText(commentMatch[0]),
        textPreview: preview(docxText(commentMatch[0])),
        attrs: {
          id: commentId,
          author: xmlAttr(attrs, "w:author") ?? xmlAttr(attrs, "author")
        },
        untrusted: true
      });
      addRunEdge(runGraph, storyId, nodeId, "contains");
      increment(storyNode.markers, "comment");
    }
    extraObjectMap.push(...inspectDocxTables(xml, partPath, partKind));
    extraObjectMap.push(...await inspectDocxImages(zip, xml, partPath, partKind));
  }
  storyGraph.summary = summarizeStories(storyGraph);
  runGraph.summary = summarizeRunGraph(runGraph);
  const objectMap: ObjectMapEntry[] = paragraphs
    .filter((paragraph) => paragraph.text)
    .map((paragraph) => ({
      stableObjectId: paragraph.stableObjectId,
      kind: "paragraph",
      text: paragraph.text,
      textPreview: preview(paragraph.text),
      sourcePath: paragraph.sourcePath,
      xmlPath: paragraph.sourcePath,
      bounds: { x: 72, y: 72 + (paragraph.index - 1) * 28, width: 468, height: 24 },
      bbox: [72, 72 + (paragraph.index - 1) * 28, 468, 24],
      selectorHints: { paragraph: paragraph.index, partKind: paragraph.partKind, story: paragraph.storyKind, storyId: paragraph.storyId, sourcePath: paragraph.sourcePath, textPreview: preview(paragraph.text) },
      editableOps: ["setText", "docx.insertParagraphAfter", "docx.addComment", "docx.addRedline"],
      trust: { level: "untrusted", reason: "document-content" },
      untrusted: true
    }));
  return { paragraphs, objectMap: [...objectMap, ...extraObjectMap], storyGraph, runGraph };
}

function emptyStoryGraph(): DocxStoryGraph {
  return {
    graphVersion: DOCX_STORY_GRAPH_VERSION,
    stories: [],
    edges: [],
    summary: {},
    untrusted: true
  };
}

function emptyRunGraph(): DocxRunGraph {
  return {
    graphVersion: DOCX_RUN_GRAPH_VERSION,
    nodes: [],
    edges: [],
    summary: {},
    untrusted: true
  };
}

function inspectParagraphGraph(
  xml: string,
  context: {
    partPath: string;
    partKind: DocxParagraph["partKind"];
    storyKind: DocxStoryKind;
    storyId: string;
    stableObjectId: string;
    paragraphIndex: number;
  }
): { runGraph: DocxRunGraph; runCount: number; textTokenCount: number; markers: Record<string, number> } {
  const runGraph = emptyRunGraph();
  const paragraphId = context.stableObjectId;
  const markers: Record<string, number> = {};
  addRunNode(runGraph, {
    nodeId: paragraphId,
    stableObjectId: paragraphId,
    type: "paragraph",
    storyId: context.storyId,
    storyKind: context.storyKind,
    paragraphId,
    sourcePath: context.partPath,
    index: context.paragraphIndex,
    text: docxText(xml),
    textPreview: preview(docxText(xml)),
    untrusted: true
  });
  addRunEdge(runGraph, context.storyId, paragraphId, "contains");

  let runCount = 0;
  let textTokenCount = 0;
  let previousRunId: string | undefined;
  for (const [runIndex, runMatch] of [...xml.matchAll(/<w:r\b[\s\S]*?<\/w:r>/g)].entries()) {
    runCount += 1;
    const runXml = runMatch[0];
    const runId = stableHashId("docx", context.partKind, "run", `${context.partPath}#${context.paragraphIndex}.${runIndex + 1}`);
    addRunNode(runGraph, {
      nodeId: runId,
      stableObjectId: runId,
      type: "run",
      storyId: context.storyId,
      storyKind: context.storyKind,
      paragraphId,
      runId,
      sourcePath: context.partPath,
      index: runIndex + 1,
      text: docxText(runXml),
      textPreview: preview(docxText(runXml)),
      untrusted: true
    });
    addRunEdge(runGraph, paragraphId, runId, "contains");
    if (previousRunId) addRunEdge(runGraph, previousRunId, runId, "next");
    previousRunId = runId;

    for (const [tokenIndex, tokenMatch] of [...runXml.matchAll(/<w:(t|delText|instrText)\b[^>]*>([\s\S]*?)<\/w:\1>/g)].entries()) {
      textTokenCount += 1;
      const tokenKind = tokenMatch[1] === "instrText" ? "fieldCode" : tokenMatch[1] === "delText" ? "deletedText" : "text";
      const text = decodeXmlEntities(tokenMatch[2] ?? "");
      const textId = stableHashId("docx", context.partKind, "text", `${context.partPath}#${context.paragraphIndex}.${runIndex + 1}.${tokenIndex + 1}`);
      addRunNode(runGraph, {
        nodeId: textId,
        stableObjectId: textId,
        type: "text",
        storyId: context.storyId,
        storyKind: context.storyKind,
        paragraphId,
        runId,
        sourcePath: context.partPath,
        index: tokenIndex + 1,
        text,
        textPreview: preview(text),
        attrs: { tokenKind },
        untrusted: true
      });
      addRunEdge(runGraph, runId, textId, "contains");
    }
  }

  const structuralNodes = inspectParagraphStructuralNodes(xml, context, paragraphId, markers);
  appendRunGraph(runGraph, structuralNodes);
  return { runGraph, runCount, textTokenCount, markers };
}

function inspectParagraphStructuralNodes(
  xml: string,
  context: {
    partPath: string;
    partKind: DocxParagraph["partKind"];
    storyKind: DocxStoryKind;
    storyId: string;
    paragraphIndex: number;
  },
  paragraphId: string,
  markers: Record<string, number>
): DocxRunGraph {
  const graph = emptyRunGraph();
  const matches: Array<{ index: number; node: DocxRunGraphNode }> = [];
  const add = (type: DocxRunGraphNodeType, key: string, index: number, attrs: Record<string, string | undefined> = {}, text?: string) => {
    increment(markers, type);
    const nodeId = stableHashId("docx", context.partKind, type, `${context.partPath}#${context.paragraphIndex}.${key}.${index}`);
    matches.push({
      index,
      node: {
        nodeId,
        stableObjectId: nodeId,
        type,
        storyId: context.storyId,
        storyKind: context.storyKind,
        paragraphId,
        sourcePath: context.partPath,
        text,
        textPreview: preview(text),
        attrs,
        untrusted: true
      }
    });
  };

  for (const match of xml.matchAll(/<w:hyperlink\b([^>]*)>[\s\S]*?<\/w:hyperlink>/g)) {
    const attrs = match[1] ?? "";
    add("hyperlink", xmlAttr(attrs, "r:id") ?? xmlAttr(attrs, "w:anchor") ?? "hyperlink", match.index ?? 0, {
      relationshipId: xmlAttr(attrs, "r:id"),
      anchor: xmlAttr(attrs, "w:anchor"),
      history: xmlAttr(attrs, "w:history")
    }, docxText(match[0]));
  }
  for (const match of xml.matchAll(/<w:fldSimple\b([^>]*)>[\s\S]*?<\/w:fldSimple>/g)) {
    const attrs = match[1] ?? "";
    add("field", xmlAttr(attrs, "w:instr") ?? "fldSimple", match.index ?? 0, {
      fieldKind: "fldSimple",
      instruction: xmlAttr(attrs, "w:instr")
    }, docxText(match[0]));
  }
  for (const match of xml.matchAll(/<w:fldChar\b([^>]*?)(?:\/>|><\/w:fldChar>)/g)) {
    const attrs = match[1] ?? "";
    add("field", xmlAttr(attrs, "w:fldCharType") ?? "fldChar", match.index ?? 0, {
      fieldKind: "fldChar",
      fldCharType: xmlAttr(attrs, "w:fldCharType")
    });
  }
  for (const match of xml.matchAll(/<w:instrText\b[^>]*>([\s\S]*?)<\/w:instrText>/g)) {
    add("field", "instrText", match.index ?? 0, { fieldKind: "instrText" }, decodeXmlEntities(match[1] ?? ""));
  }
  for (const match of xml.matchAll(/<w:bookmark(Start|End)\b([^>]*?)(?:\/>|><\/w:bookmark\1>)/g)) {
    const attrs = match[2] ?? "";
    add("bookmark", `${match[1]}:${xmlAttr(attrs, "w:id") ?? ""}`, match.index ?? 0, {
      marker: match[1]?.toLowerCase(),
      id: xmlAttr(attrs, "w:id"),
      name: xmlAttr(attrs, "w:name")
    });
  }
  for (const match of xml.matchAll(/<w:commentRange(Start|End)\b([^>]*?)(?:\/>|><\/w:commentRange\1>)/g)) {
    const attrs = match[2] ?? "";
    add("commentRange", `${match[1]}:${xmlAttr(attrs, "w:id") ?? ""}`, match.index ?? 0, {
      marker: match[1]?.toLowerCase(),
      id: xmlAttr(attrs, "w:id")
    });
  }
  for (const match of xml.matchAll(/<w:commentReference\b([^>]*?)(?:\/>|><\/w:commentReference>)/g)) {
    const attrs = match[1] ?? "";
    add("commentReference", xmlAttr(attrs, "w:id") ?? "commentReference", match.index ?? 0, {
      id: xmlAttr(attrs, "w:id")
    });
  }
  for (const match of xml.matchAll(/<w:sdt\b[\s\S]*?<\/w:sdt>/g)) {
    const block = match[0];
    const sdtProps = contentControlProps(block);
    add("contentControl", sdtProps.tag ?? sdtProps.alias ?? "sdt", match.index ?? 0, sdtProps, docxText(block));
  }
  for (const match of xml.matchAll(/<w:(ins|del)\b([^>]*)>[\s\S]*?<\/w:\1>/g)) {
    const attrs = match[2] ?? "";
    add("revision", `${match[1]}:${xmlAttr(attrs, "w:id") ?? ""}`, match.index ?? 0, {
      revisionType: match[1],
      id: xmlAttr(attrs, "w:id"),
      author: xmlAttr(attrs, "w:author"),
      date: xmlAttr(attrs, "w:date")
    }, docxText(match[0]));
  }

  matches.sort((left, right) => left.index - right.index);
  let previousId: string | undefined;
  for (const item of matches) {
    addRunNode(graph, item.node);
    addRunEdge(graph, paragraphId, item.node.nodeId, "contains");
    if (previousId) addRunEdge(graph, previousId, item.node.nodeId, "next");
    previousId = item.node.nodeId;
  }
  return graph;
}

function addRunNode(graph: DocxRunGraph, node: DocxRunGraphNode): void {
  graph.nodes.push(node);
}

function addRunEdge(graph: DocxRunGraph, from: string, to: string, relation: "contains" | "next"): void {
  graph.edges.push({ from, to, relation, untrusted: true });
}

function appendRunGraph(target: DocxRunGraph, source: DocxRunGraph): void {
  target.nodes.push(...source.nodes);
  target.edges.push(...source.edges);
}

function mergeMarkerCounts(target: Record<string, number>, source: Record<string, number>): void {
  for (const [key, value] of Object.entries(source)) target[key] = (target[key] ?? 0) + value;
}

function increment(target: Record<string, number>, key: string): void {
  target[key] = (target[key] ?? 0) + 1;
}

function summarizeStories(graph: DocxStoryGraph): Record<string, number> {
  return {
    stories: graph.stories.length,
    paragraphs: graph.stories.reduce((count, story) => count + story.paragraphCount, 0),
    runs: graph.stories.reduce((count, story) => count + story.runCount, 0),
    textTokens: graph.stories.reduce((count, story) => count + story.textTokenCount, 0)
  };
}

function summarizeRunGraph(graph: DocxRunGraph): Record<string, number> {
  const summary: Record<string, number> = { nodes: graph.nodes.length, edges: graph.edges.length };
  for (const node of graph.nodes) summary[node.type] = (summary[node.type] ?? 0) + 1;
  return summary;
}

function docxText(xml: string): string {
  return [...xml.matchAll(/<w:(?:t|delText|instrText)\b[^>]*>([\s\S]*?)<\/w:(?:t|delText|instrText)>/g)]
    .map((match) => decodeXmlEntities(match[1] ?? ""))
    .join("");
}

function contentControlProps(xml: string): Record<string, string | undefined> {
  const tagAttrs = /<w:tag\b([^>]*)/.exec(xml)?.[1] ?? "";
  const aliasAttrs = /<w:alias\b([^>]*)/.exec(xml)?.[1] ?? "";
  return {
    tag: xmlAttr(tagAttrs, "w:val") ?? xmlAttr(tagAttrs, "val"),
    alias: xmlAttr(aliasAttrs, "w:val") ?? xmlAttr(aliasAttrs, "val")
  };
}

function inspectDocxTables(xml: string, partPath: string, partKind: DocxParagraph["partKind"]): ObjectMapEntry[] {
  const entries: ObjectMapEntry[] = [];
  let tableIndex = 0;
  let cellIndex = 0;
  for (const tableMatch of xml.matchAll(/<w:tbl\b[\s\S]*?<\/w:tbl>/g)) {
    tableIndex += 1;
    const table = tableMatch[0];
    entries.push({
      stableObjectId: stableHashId("docx", partKind, "table", `${partPath}#${tableIndex}`),
      kind: "table",
      label: `Table ${tableIndex}`,
      sourcePath: partPath,
      xmlPath: partPath,
      selectorHints: { story: partKind, sourcePath: partPath, table: tableIndex },
      editableOps: ["docx.setTableCellText"],
      trust: { level: "untrusted", reason: "document-content" },
      untrusted: true
    });
    let rowIndex = 0;
    for (const rowMatch of table.matchAll(/<w:tr\b[\s\S]*?<\/w:tr>/g)) {
      rowIndex += 1;
      let columnIndex = 0;
      for (const cellMatch of rowMatch[0].matchAll(/<w:tc\b[\s\S]*?<\/w:tc>/g)) {
        columnIndex += 1;
        cellIndex += 1;
        const text = localText(cellMatch[0], "t").join("");
        entries.push({
          stableObjectId: stableHashId("docx", partKind, "tableCell", `${partPath}#${cellIndex}`),
          kind: "tableCell",
          text,
          textPreview: preview(text),
          sourcePath: partPath,
          xmlPath: partPath,
          bounds: { x: 72 + (columnIndex - 1) * 144, y: 120 + (rowIndex - 1) * 28, width: 144, height: 28 },
          bbox: [72 + (columnIndex - 1) * 144, 120 + (rowIndex - 1) * 28, 144, 28],
          selectorHints: { story: partKind, sourcePath: partPath, table: tableIndex, row: rowIndex, column: columnIndex, cell: cellIndex },
          editableOps: ["setText", "docx.setTableCellText"],
          trust: { level: "untrusted", reason: "document-content" },
          untrusted: true
        });
      }
    }
  }
  return entries;
}

async function inspectDocxImages(zip: JSZip, xml: string, partPath: string, partKind: DocxParagraph["partKind"]): Promise<ObjectMapEntry[]> {
  const relsPath = partPath.replace(/^word\//, "word/_rels/") + ".rels";
  const rels = parseRelationships((await readZipText(zip, relsPath)) ?? "");
  let imageIndex = 0;
  return [...xml.matchAll(/<w:drawing\b[\s\S]*?<\/w:drawing>/g)].map((match) => {
    imageIndex += 1;
    const relationshipId = /(?:r:embed|r:link)="([^"]+)"/.exec(match[0])?.[1];
    const rel = rels.find((item) => item.id === relationshipId);
    const assetPath = rel ? relationshipTarget("word", rel.target) : undefined;
    return {
      stableObjectId: stableHashId("docx", partKind, "image", `${partPath}#${imageIndex}`),
      kind: "image",
      label: `Image ${imageIndex}`,
      sourcePath: partPath,
      xmlPath: partPath,
      selectorHints: { story: partKind, sourcePath: partPath, relationshipId, assetPath },
      editableOps: ["asset.replace", "docx.replaceImage"],
      media: { relationshipId, assetPath },
      trust: { level: "untrusted" as const, reason: "document-content" },
      untrusted: true as const
    };
  });
}

export function setParagraphText(xml: string, ordinal: number, text: string): { changed: boolean; matched: boolean; xml: string } {
  return replaceNthBlock(xml, /<w:p\b[\s\S]*?<\/w:p>/g, ordinal, (paragraph) => setFirstTextInBlock(paragraph, "w:t", text));
}

export function insertParagraphAfter(xml: string, ordinal: number, text: string): { changed: boolean; matched: boolean; xml: string } {
  return replaceNthBlock(xml, /<w:p\b[\s\S]*?<\/w:p>/g, ordinal, (paragraph) => `${paragraph}${paragraphXml(text, "w")}`);
}

export function replaceOrCreateHeaderFooter(xml: string | undefined, kind: "header" | "footer", text: string): string {
  const root = kind === "header" ? "w:hdr" : "w:ftr";
  const ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  const content = `${paragraphXml(text, "w")}`;
  if (!xml) return `<${root} ${ns}>${content}</${root}>`;
  if (new RegExp(`<${root}\\b[\\s\\S]*?<\\/${root}>`).test(xml)) {
    return xml.replace(new RegExp(`(<${root}\\b[^>]*>)[\\s\\S]*?(<\\/${root}>)`), `$1${content}$2`);
  }
  return `<${root} ${ns}>${content}</${root}>`;
}

export function commentXml(id: number, author: string, text: string, date = new Date()): string {
  return `<w:comment w:id="${id}" w:author="${escapeXmlAttr(author)}" w:date="${date.toISOString()}">${paragraphXml(text, "w")}</w:comment>`;
}

export function insertedParagraphXml(text: string, author = "officegen", date = new Date(), revisionId = 1): string {
  return `<w:p><w:ins w:author="${escapeXmlAttr(author)}" w:date="${date.toISOString()}" w:id="${revisionId}"><w:r><w:t>${escapeXmlTextLocal(text)}</w:t></w:r></w:ins></w:p>`;
}

function docxPartKind(path: string): DocxParagraph["partKind"] {
  if (/^word\/header/i.test(path)) return "header";
  if (/^word\/footer/i.test(path)) return "footer";
  if (/^word\/comments/i.test(path)) return "comment";
  return "body";
}

function docxStoryKind(path: string): DocxStoryKind {
  if (/^word\/header/i.test(path)) return "header";
  if (/^word\/footer/i.test(path)) return "footer";
  if (/^word\/comments/i.test(path)) return "comments";
  return "document";
}

function escapeXmlAttr(value: string): string {
  return escapeXmlTextLocal(value).replace(/"/g, "&quot;").replace(/'/g, "&apos;");
}

function escapeXmlTextLocal(value: string): string {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}
