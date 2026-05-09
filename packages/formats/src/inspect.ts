import {
  AGENT_UNTRUSTED_INSTRUCTION,
  type AgentSeparatedResult,
  type InputLike,
  getLoadedZipSafetyReport,
  loadZip,
  makeStableObjectId,
  normalizeInput,
  sortedZipFiles,
  trustedMeta,
  type OfficegenConfig,
  zipSafetyCaveats
} from "./shared.js";
import { inspectParagraphs } from "./ooxml/docx.js";
import { inspectSlides } from "./ooxml/pptx.js";
import { inspectSheets } from "./ooxml/xlsx.js";
import { PDFDocument } from "pdf-lib";

export type InspectDepth = "summary" | "shallow" | "full";

export interface InspectOptions {
  format?: "pptx" | "docx" | "xlsx" | "pdf" | "unknown";
  depth?: InspectDepth;
  include?: Array<"text" | "assets" | "relationships" | "rawPaths">;
  config?: OfficegenConfig;
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
  const zip = await loadZip(input, { zipSafety: { config: options.config } });
  const paths = sortedZipFiles(zip);
  const mediaPaths = paths.filter((path) => /^ppt\/media\//i.test(path));
  const { slides, objectMap } = await inspectSlides(zip);

  const macros = paths.filter((path) => /vbaProject\.bin$/i.test(path));
  return {
    schema: "officegen.inspect.result@1.2",
    trusted: trustedMeta(
      "officegen.inspect.result@1.2",
      input,
      {
        slides: slides.length,
        textObjects: objectMap.length,
        assets: mediaPaths.length,
        macros: macros.length,
        zipEntries: paths.length
      },
      ["PPTX inspect is zip/XML based; animation and theme resolution are summarized only.", ...zipSafetyCaveats(getLoadedZipSafetyReport(zip))]
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
  const zip = await loadZip(input, { zipSafety: { config: options.config } });
  const paths = sortedZipFiles(zip);
  const { paragraphs, objectMap } = await inspectParagraphs(zip);
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
      ["DOCX inspect reads main document XML; headers, footers, fields, and styles are summarized.", ...zipSafetyCaveats(getLoadedZipSafetyReport(zip))]
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
  const zip = await loadZip(input, { zipSafety: { config: options.config } });
  const paths = sortedZipFiles(zip);
  const { sheets, objectMap, sharedStrings } = await inspectSheets(zip);

  const macros = paths.filter((path) => /vbaProject\.bin$/i.test(path));
  return {
    schema: "officegen.inspect.result@1.2",
    trusted: trustedMeta(
      "officegen.inspect.result@1.2",
      input,
      {
        sheets: sheets.length,
        cells: objectMap.length,
        sharedStrings: sharedStrings.length,
        macros: macros.length,
        zipEntries: paths.length
      },
      ["XLSX inspect reads cached cell values; formulas, styles, and charts are summarized.", ...zipSafetyCaveats(getLoadedZipSafetyReport(zip))]
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
