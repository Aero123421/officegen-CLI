import {
  type InputLike,
  isOfficeFormat,
  loadZip,
  normalizeInput,
  readZipText,
  replaceAllLiteral,
  writeOutput,
  zipToBytes
} from "./shared.js";
import { PDFDocument, rgb, StandardFonts } from "pdf-lib";

export type EditOperation =
  | { type: "replaceText"; from: string; to: string; selector?: { stableObjectId?: string } }
  | { type: "setText"; text: string; selector: { stableObjectId?: string; contains?: string } }
  | { type: "pdf.textOverlay"; page: number; text: string; x: number; y: number; size?: number; color?: string }
  | { type: "pdf.annotation"; page: number; text: string; x: number; y: number; width?: number; height?: number };

export interface EditOptions {
  out?: string;
  dryRun?: boolean;
  format?: "pptx" | "docx" | "xlsx" | "pdf" | "unknown";
}

export interface EditResult {
  schema: "officegen.edit.result@1.2";
  format: string;
  changed: boolean;
  applied: number;
  skipped: number;
  out?: string;
  bytes?: Uint8Array;
  caveats: string[];
}

export async function edit(input: InputLike, operations: EditOperation[], options: EditOptions = {}): Promise<EditResult> {
  const normalized = await normalizeInput(input, options.format ?? "unknown");
  if (isOfficeFormat(normalized.format)) return editOfficeXml(normalized, operations, options);
  if (normalized.format === "pdf") return editPdf(normalized, operations, options);
  throw new Error(`Unsupported edit format: ${normalized.format}`);
}

export const editDocument = edit;

async function editOfficeXml(
  input: Awaited<ReturnType<typeof normalizeInput>>,
  operations: EditOperation[],
  options: EditOptions
): Promise<EditResult> {
  const zip = await loadZip(input);
  const editablePaths = Object.keys(zip.files)
    .filter((path) => !zip.files[path]?.dir)
    .filter((path) =>
      input.format === "pptx"
        ? /^ppt\/slides\/slide\d+\.xml$/i.test(path)
        : input.format === "docx"
          ? /^word\/(document|header\d+|footer\d+)\.xml$/i.test(path)
          : /^xl\/(worksheets\/sheet\d+|sharedStrings)\.xml$/i.test(path)
    )
    .sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));

  let applied = 0;
  let skipped = 0;

  for (const path of editablePaths) {
    let xml = (await readZipText(zip, path)) ?? "";
    const before = xml;
    for (const op of operations) {
      if (op.type === "replaceText") {
        if (!op.from) {
          skipped += 1;
          continue;
        }
        const next = replaceAllLiteral(xml, escapeXmlText(op.from), escapeXmlText(op.to));
        if (next !== xml) applied += 1;
        xml = next;
      } else if (op.type === "setText") {
        const stableTarget = op.selector.stableObjectId
          ? replaceByStableObjectId(input.format as "pptx" | "docx" | "xlsx", path, xml, op.selector.stableObjectId, op.text)
          : undefined;
        if (stableTarget?.changed) {
          xml = stableTarget.xml;
          applied += 1;
        } else if (stableTarget?.matchedPath === false) {
          continue;
        } else if (stableTarget) {
          skipped += 1;
        } else if (op.selector.contains && xml.includes(escapeXmlText(op.selector.contains))) {
          const contains = op.selector.contains;
          xml = replaceAllLiteral(xml, escapeXmlText(contains), escapeXmlText(op.text));
          applied += 1;
        } else {
          skipped += 1;
        }
      } else {
        skipped += 1;
      }
    }
    if (xml !== before && !options.dryRun) zip.file(path, xml);
  }

  const bytes = options.dryRun ? input.bytes : await zipToBytes(zip);
  if (!options.dryRun) await writeOutput(options.out, bytes);
  return {
    schema: "officegen.edit.result@1.2",
    format: input.format,
    changed: applied > 0,
    applied,
    skipped,
    out: options.dryRun ? undefined : options.out,
    bytes: options.out ? undefined : bytes,
    caveats: [
      "Office edit is XML text replacement only; rich text run boundaries can prevent a match.",
      "Prefer stableObjectId from inspect for selection; this MVP also supports literal text matching."
    ]
  };
}

async function editPdf(
  input: Awaited<ReturnType<typeof normalizeInput>>,
  operations: EditOperation[],
  options: EditOptions
): Promise<EditResult> {
  const pdf = await PDFDocument.load(input.bytes, { ignoreEncryption: true });
  const font = await pdf.embedFont(StandardFonts.Helvetica);
  let applied = 0;
  let skipped = 0;

  for (const op of operations) {
    if (op.type === "pdf.textOverlay") {
      const page = pdf.getPage(op.page - 1);
      if (!page) {
        skipped += 1;
        continue;
      }
      page.drawText(op.text, {
        x: op.x,
        y: op.y,
        size: op.size ?? 12,
        font,
        color: parseRgb(op.color)
      });
      applied += 1;
    } else if (op.type === "pdf.annotation") {
      const page = pdf.getPage(op.page - 1);
      if (!page) {
        skipped += 1;
        continue;
      }
      page.drawRectangle({
        x: op.x,
        y: op.y,
        width: op.width ?? 160,
        height: op.height ?? 48,
        borderColor: rgb(0.91, 0.59, 0.12),
        borderWidth: 1,
        color: rgb(1, 0.96, 0.82),
        opacity: 0.9
      });
      page.drawText(op.text, { x: op.x + 6, y: op.y + (op.height ?? 48) - 18, size: 10, font, color: rgb(0.2, 0.2, 0.2) });
      applied += 1;
    } else {
      skipped += 1;
    }
  }

  const bytes = options.dryRun ? input.bytes : await pdf.save();
  if (!options.dryRun) await writeOutput(options.out, bytes);
  return {
    schema: "officegen.edit.result@1.2",
    format: "pdf",
    changed: applied > 0,
    applied,
    skipped,
    out: options.dryRun ? undefined : options.out,
    bytes: options.out ? undefined : bytes,
    caveats: ["PDF edit is additive; existing text/content is not removed in the MVP."]
  };
}

function escapeXmlText(value: string): string {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function replaceByStableObjectId(
  format: "pptx" | "docx" | "xlsx",
  path: string,
  xml: string,
  stableObjectId: string,
  text: string
): { changed: boolean; matchedPath: boolean; xml: string } | undefined {
  const parts = stableObjectId.split(":");
  if (parts[0] !== format || parts.length < 4) return undefined;
  const scope = parts[1] ?? "";
  const kind = parts[2] ?? "";
  const ordinal = Number(parts[3]);
  if (!Number.isInteger(ordinal) || ordinal < 1) return undefined;

  if (format === "pptx" && kind === "text") {
    const slideNo = Number(scope.replace(/^s/, ""));
    if (path !== `ppt/slides/slide${slideNo}.xml`) return { changed: false, matchedPath: false, xml };
    return replaceNthTextTag(xml, /<a:t([^>]*)>([\s\S]*?)<\/a:t>/g, ordinal, text);
  }

  if (format === "docx" && kind === "paragraph" && path === "word/document.xml") {
    return replaceNthBlock(xml, /<w:p[\s\S]*?<\/w:p>/g, ordinal, (paragraph) => {
      let used = false;
      return paragraph.replace(/<w:t([^>]*)>([\s\S]*?)<\/w:t>/g, (_match, attrs: string) => {
        if (used) return `<w:t${attrs}></w:t>`;
        used = true;
        return `<w:t${attrs}>${escapeXmlText(text)}</w:t>`;
      });
    });
  }

  if (format === "xlsx" && kind === "cell") {
    const sheetNo = Number(scope.replace(/^s/, ""));
    if (path !== `xl/worksheets/sheet${sheetNo}.xml`) return { changed: false, matchedPath: false, xml };
    return replaceNthBlock(xml, /<c([^>]*)>[\s\S]*?<\/c>/g, ordinal, (cell) => {
      if (/<v>[\s\S]*?<\/v>/.test(cell)) return cell.replace(/<v>[\s\S]*?<\/v>/, `<v>${escapeXmlText(text)}</v>`);
      return cell.replace(/<\/c>$/, `<v>${escapeXmlText(text)}</v></c>`);
    });
  }

  return undefined;
}

function replaceNthTextTag(
  xml: string,
  pattern: RegExp,
  ordinal: number,
  text: string
): { changed: boolean; matchedPath: boolean; xml: string } {
  let index = 0;
  let changed = false;
  const next = xml.replace(pattern, (match, attrs: string) => {
    index += 1;
    if (index !== ordinal) return match;
    changed = true;
    return `<a:t${attrs}>${escapeXmlText(text)}</a:t>`;
  });
  return { changed, matchedPath: true, xml: next };
}

function replaceNthBlock(
  xml: string,
  pattern: RegExp,
  ordinal: number,
  replacer: (block: string) => string
): { changed: boolean; matchedPath: boolean; xml: string } {
  let index = 0;
  let changed = false;
  const next = xml.replace(pattern, (match) => {
    index += 1;
    if (index !== ordinal) return match;
    changed = true;
    return replacer(match);
  });
  return { changed, matchedPath: true, xml: next };
}

function parseRgb(hex?: string): ReturnType<typeof rgb> {
  if (!hex || !/^#[0-9a-f]{6}$/i.test(hex)) return rgb(0, 0, 0);
  const r = Number.parseInt(hex.slice(1, 3), 16) / 255;
  const g = Number.parseInt(hex.slice(3, 5), 16) / 255;
  const b = Number.parseInt(hex.slice(5, 7), 16) / 255;
  return rgb(r, g, b);
}
