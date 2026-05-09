import type JSZip from "jszip";
import type { ObjectMapEntry } from "../shared.js";
import { readZipText, sortedZipFiles, stableHashId } from "../shared.js";
import { localText, paragraphXml, preview, replaceNthBlock, setFirstTextInBlock } from "./xml.js";

export interface DocxParagraph {
  stableObjectId: string;
  index: number;
  text: string;
  sourcePath: string;
  partKind: "body" | "header" | "footer" | "comment";
  untrusted: true;
}

export async function inspectParagraphs(zip: JSZip): Promise<{ paragraphs: DocxParagraph[]; objectMap: ObjectMapEntry[] }> {
  const paths = sortedZipFiles(zip);
  const docxParts = [
    "word/document.xml",
    ...paths.filter((path) => /^word\/header\d+\.xml$/i.test(path)).sort(),
    ...paths.filter((path) => /^word\/footer\d+\.xml$/i.test(path)).sort(),
    ...paths.filter((path) => /^word\/comments\.xml$/i.test(path)).sort()
  ];
  const paragraphs: DocxParagraph[] = [];
  for (const partPath of docxParts) {
    const xml = (await readZipText(zip, partPath)) ?? "";
    const partKind = docxPartKind(partPath);
    for (const [index, match] of [...xml.matchAll(/<w:p\b[\s\S]*?<\/w:p>/g)].entries()) {
      paragraphs.push({
        stableObjectId: stableHashId("docx", partKind, "paragraph", `${partPath}#${index + 1}`),
        index: index + 1,
        text: localText(match[0], "t").join(""),
        sourcePath: partPath,
        partKind,
        untrusted: true as const
      });
    }
  }
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
      selectorHints: { paragraph: paragraph.index, partKind: paragraph.partKind, sourcePath: paragraph.sourcePath, textPreview: preview(paragraph.text) },
      editableOps: ["setText", "docx.insertParagraphAfter", "docx.addComment", "docx.addRedline"],
      trust: { level: "untrusted", reason: "document-content" },
      untrusted: true
    }));
  return { paragraphs, objectMap };
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

function escapeXmlAttr(value: string): string {
  return escapeXmlTextLocal(value).replace(/"/g, "&quot;").replace(/'/g, "&apos;");
}

function escapeXmlTextLocal(value: string): string {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}
