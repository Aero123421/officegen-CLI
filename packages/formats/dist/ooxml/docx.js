import { makeStableObjectId, readZipText } from "../shared.js";
import { localText, paragraphXml, preview, replaceNthBlock, setFirstTextInBlock } from "./xml.js";
export async function inspectParagraphs(zip) {
    const documentXml = (await readZipText(zip, "word/document.xml")) ?? "";
    const paragraphs = [...documentXml.matchAll(/<w:p\b[\s\S]*?<\/w:p>/g)].map((match, index) => ({
        stableObjectId: makeStableObjectId("docx", "body", "paragraph", index + 1),
        index: index + 1,
        text: localText(match[0], "t").join(""),
        sourcePath: "word/document.xml",
        untrusted: true
    }));
    const objectMap = paragraphs
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
        selectorHints: { paragraph: paragraph.index, textPreview: preview(paragraph.text) },
        trust: { level: "untrusted", reason: "document-content" },
        untrusted: true
    }));
    return { paragraphs, objectMap };
}
export function setParagraphText(xml, ordinal, text) {
    return replaceNthBlock(xml, /<w:p\b[\s\S]*?<\/w:p>/g, ordinal, (paragraph) => setFirstTextInBlock(paragraph, "w:t", text));
}
export function insertParagraphAfter(xml, ordinal, text) {
    return replaceNthBlock(xml, /<w:p\b[\s\S]*?<\/w:p>/g, ordinal, (paragraph) => `${paragraph}${paragraphXml(text, "w")}`);
}
//# sourceMappingURL=docx.js.map