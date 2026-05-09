import {
  AGENT_UNTRUSTED_INSTRUCTION,
  type AgentSeparatedResult,
  type InputLike,
  type ObjectMapEntry,
  extractXmlTexts,
  loadZip,
  makeStableObjectId,
  normalizeInput,
  readZipText,
  sortedZipFiles,
  stripXmlTags,
  trustedMeta
} from "./shared.js";
import { PDFDocument } from "pdf-lib";

export type InspectDepth = "summary" | "shallow" | "full";

export interface InspectOptions {
  format?: "pptx" | "docx" | "xlsx" | "pdf" | "unknown";
  depth?: InspectDepth;
  include?: Array<"text" | "assets" | "relationships" | "rawPaths">;
}

export interface InspectResult extends AgentSeparatedResult<Record<string, unknown>> {
  schema: "officegen.inspect.result@1.2";
}

export async function inspect(input: InputLike, options: InspectOptions = {}): Promise<InspectResult> {
  const normalized = await normalizeInput(input, options.format ?? "unknown");
  if (normalized.format === "pptx") return inspectPptx(normalized, options);
  if (normalized.format === "docx") return inspectDocx(normalized, options);
  if (normalized.format === "xlsx") return inspectXlsx(normalized, options);
  if (normalized.format === "pdf") return inspectPdf(normalized);
  throw new Error(`Unsupported inspect format: ${normalized.format}`);
}

export const inspectDocument = inspect;
export const inspectOfficeFile = inspect;

async function inspectPptx(input: Awaited<ReturnType<typeof normalizeInput>>, options: InspectOptions): Promise<InspectResult> {
  const zip = await loadZip(input);
  const paths = sortedZipFiles(zip);
  const slidePaths = paths.filter((path) => /^ppt\/slides\/slide\d+\.xml$/i.test(path)).sort(naturalSort);
  const mediaPaths = paths.filter((path) => /^ppt\/media\//i.test(path));
  const objectMap: ObjectMapEntry[] = [];
  const slides = [];

  for (const [slideIndex, slidePath] of slidePaths.entries()) {
    const xml = (await readZipText(zip, slidePath)) ?? "";
    const texts = extractXmlTexts(xml, "t");
    const slideNo = slideIndex + 1;
    const textObjects = texts.map((text, textIndex) => {
      const entry: ObjectMapEntry = {
        stableObjectId: makeStableObjectId("pptx", `s${String(slideNo).padStart(3, "0")}`, "text", textIndex + 1),
        kind: "text",
        text,
        sourcePath: slidePath,
        xmlPath: slidePath,
        untrusted: true
      };
      objectMap.push(entry);
      return entry;
    });
    const shapeCount = (xml.match(/<p:sp[\s>]/g) ?? []).length;
    const pictureCount = (xml.match(/<p:pic[\s>]/g) ?? []).length;
    slides.push({
      stableObjectId: makeStableObjectId("pptx", "deck", "slide", slideNo),
      index: slideNo,
      sourcePath: slidePath,
      text: texts.join("\n"),
      textObjects,
      shapeCount,
      pictureCount,
      untrusted: true
    });
  }

  const macros = paths.filter((path) => /vbaProject\.bin$/i.test(path));
  return {
    schema: "officegen.inspect.result@1.2",
    trusted: trustedMeta(
      "officegen.inspect.result@1.2",
      input,
      {
        slides: slidePaths.length,
        textObjects: objectMap.length,
        assets: mediaPaths.length,
        macros: macros.length,
        zipEntries: paths.length
      },
      ["PPTX inspect is zip/XML based; animation and theme resolution are summarized only."]
    ),
    untrusted: {
      slides,
      assets: mediaPaths.map((path, index) => ({
        stableObjectId: makeStableObjectId("pptx", "deck", "asset", index + 1),
        path,
        fileName: path.split("/").pop(),
        untrusted: true
      })),
      ...(options.depth === "full" || options.include?.includes("rawPaths") ? { rawPaths: paths } : {})
    },
    objectMap,
    agentInstruction: AGENT_UNTRUSTED_INSTRUCTION
  };
}

async function inspectDocx(input: Awaited<ReturnType<typeof normalizeInput>>, options: InspectOptions): Promise<InspectResult> {
  const zip = await loadZip(input);
  const paths = sortedZipFiles(zip);
  const documentXml = (await readZipText(zip, "word/document.xml")) ?? "";
  const paragraphs = [...documentXml.matchAll(/<w:p[\s\S]*?<\/w:p>/g)].map((match, index) => {
    const texts = extractXmlTexts(match[0], "t");
    return {
      stableObjectId: makeStableObjectId("docx", "body", "paragraph", index + 1),
      index: index + 1,
      text: texts.join(""),
      untrusted: true
    };
  });
  const objectMap: ObjectMapEntry[] = paragraphs
    .filter((paragraph) => paragraph.text)
    .map((paragraph) => ({
      stableObjectId: paragraph.stableObjectId,
      kind: "paragraph",
      text: paragraph.text,
      sourcePath: "word/document.xml",
      xmlPath: "word/document.xml",
      untrusted: true
    }));
  const mediaPaths = paths.filter((path) => /^word\/media\//i.test(path));
  const macros = paths.filter((path) => /vbaProject\.bin$/i.test(path));

  return {
    schema: "officegen.inspect.result@1.2",
    trusted: trustedMeta(
      "officegen.inspect.result@1.2",
      input,
      {
        paragraphs: paragraphs.length,
        textObjects: objectMap.length,
        assets: mediaPaths.length,
        macros: macros.length,
        zipEntries: paths.length
      },
      ["DOCX inspect reads main document XML; headers, footers, fields, and styles are summarized."]
    ),
    untrusted: {
      paragraphs,
      assets: mediaPaths.map((path, index) => ({
        stableObjectId: makeStableObjectId("docx", "body", "asset", index + 1),
        path,
        fileName: path.split("/").pop(),
        untrusted: true
      })),
      ...(options.depth === "full" || options.include?.includes("rawPaths") ? { rawPaths: paths } : {})
    },
    objectMap,
    agentInstruction: AGENT_UNTRUSTED_INSTRUCTION
  };
}

async function inspectXlsx(input: Awaited<ReturnType<typeof normalizeInput>>, options: InspectOptions): Promise<InspectResult> {
  const zip = await loadZip(input);
  const paths = sortedZipFiles(zip);
  const sheetPaths = paths.filter((path) => /^xl\/worksheets\/sheet\d+\.xml$/i.test(path)).sort(naturalSort);
  const sharedStringsXml = (await readZipText(zip, "xl/sharedStrings.xml")) ?? "";
  const sharedStrings = [...sharedStringsXml.matchAll(/<si[\s\S]*?<\/si>/g)].map((match) =>
    extractXmlTexts(match[0], "t").join("")
  );
  const objectMap: ObjectMapEntry[] = [];
  const sheets = [];

  for (const [sheetIndex, sheetPath] of sheetPaths.entries()) {
    const xml = (await readZipText(zip, sheetPath)) ?? "";
    const cells = [...xml.matchAll(/<c([^>]*)>([\s\S]*?)<\/c>/g)].map((match, cellIndex) => {
      const attrs = match[1] ?? "";
      const body = match[2] ?? "";
      const ref = /r="([^"]+)"/.exec(attrs)?.[1] ?? `cell${cellIndex + 1}`;
      const type = /t="([^"]+)"/.exec(attrs)?.[1];
      const raw = extractXmlTexts(body, "v")[0] ?? extractXmlTexts(body, "t")[0] ?? stripXmlTags(body);
      const value = type === "s" ? sharedStrings[Number(raw)] ?? raw : raw;
      const stableObjectId = makeStableObjectId("xlsx", `s${String(sheetIndex + 1).padStart(3, "0")}`, "cell", cellIndex + 1);
      const entry: ObjectMapEntry = {
        stableObjectId,
        kind: "cell",
        label: ref,
        text: value,
        sourcePath: sheetPath,
        xmlPath: sheetPath,
        untrusted: true
      };
      objectMap.push(entry);
      return {
        stableObjectId,
        ref,
        value,
        untrusted: true
      };
    });
    sheets.push({
      stableObjectId: makeStableObjectId("xlsx", "workbook", "sheet", sheetIndex + 1),
      index: sheetIndex + 1,
      sourcePath: sheetPath,
      cells,
      untrusted: true
    });
  }

  const macros = paths.filter((path) => /vbaProject\.bin$/i.test(path));
  return {
    schema: "officegen.inspect.result@1.2",
    trusted: trustedMeta(
      "officegen.inspect.result@1.2",
      input,
      {
        sheets: sheetPaths.length,
        cells: objectMap.length,
        sharedStrings: sharedStrings.length,
        macros: macros.length,
        zipEntries: paths.length
      },
      ["XLSX inspect reads cached cell values; formulas, styles, and charts are summarized."]
    ),
    untrusted: {
      sheets,
      ...(options.depth === "full" || options.include?.includes("rawPaths") ? { rawPaths: paths } : {})
    },
    objectMap,
    agentInstruction: AGENT_UNTRUSTED_INSTRUCTION
  };
}

async function inspectPdf(input: Awaited<ReturnType<typeof normalizeInput>>): Promise<InspectResult> {
  const pdf = await PDFDocument.load(input.bytes, { ignoreEncryption: true });
  const pages = pdf.getPages().map((page, index) => {
    const size = page.getSize();
    return {
      stableObjectId: makeStableObjectId("pdf", "document", "page", index + 1),
      index: index + 1,
      width: size.width,
      height: size.height,
      untrusted: true
    };
  });
  return {
    schema: "officegen.inspect.result@1.2",
    trusted: trustedMeta(
      "officegen.inspect.result@1.2",
      input,
      { pages: pages.length },
      ["PDF text extraction is not implemented in the MVP inspect path."]
    ),
    untrusted: { pages },
    objectMap: [],
    agentInstruction: AGENT_UNTRUSTED_INSTRUCTION
  };
}

function naturalSort(a: string, b: string): number {
  return a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" });
}

