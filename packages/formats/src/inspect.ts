import {
  AGENT_UNTRUSTED_INSTRUCTION,
  type AgentSeparatedResult,
  type InputLike,
  getLoadedZipSafetyReport,
  loadZip,
  makeStableObjectId,
  normalizeInput,
  readZipText,
  sortedZipFiles,
  trustedMeta,
  type OfficegenConfig,
  zipSafetyCaveats
} from "./shared.js";
import { inspectParagraphs } from "./ooxml/docx.js";
import { inspectSlides } from "./ooxml/pptx.js";
import { inspectSheets } from "./ooxml/xlsx.js";
import { buildObjectGraph, type BuildObjectGraphOptions, type ObjectGraph, type ObjectGraphRiskFlag } from "./graphs/objectGraph.js";
import { inspectPdfObjectGraph } from "./pdf/objectGraph.js";
import { PDFDocument } from "pdf-lib";

export type InspectDepth = "summary" | "shallow" | "full";

export interface InspectOptions {
  format?: "pptx" | "docx" | "xlsx" | "pdf" | "unknown";
  depth?: InspectDepth;
  include?: Array<"text" | "assets" | "relationships" | "rawPaths">;
  structure?: boolean;
  sheet?: string;
  range?: string;
  config?: OfficegenConfig;
  emit?: "inspect" | "object-graph";
  includeObjectGraph?: boolean;
  objectGraph?: Pick<BuildObjectGraphOptions, "nodeOffset" | "nodeLimit" | "edgeOffset" | "edgeLimit">;
}

export interface InspectResult extends AgentSeparatedResult<Record<string, unknown>> {
  schema: "officegen.inspect.result@1.2";
  objectGraph?: ObjectGraph;
}

export async function inspect(input: InputLike, options: InspectOptions = {}): Promise<InspectResult> {
  const normalized = await normalizeInput(input, options.format ?? "unknown");
  if (normalized.format === "pptx") return inspectPptx(normalized, options);
  if (normalized.format === "docx") return inspectDocx(normalized, options);
  if (normalized.format === "xlsx") return inspectXlsx(normalized, options);
  if (normalized.format === "pdf") return inspectPdf(normalized, options);
  throw new Error(`Unsupported inspect format: ${normalized.format}`);
}

export const inspectDocument = inspect;
export const inspectOfficeFile = inspect;

async function inspectPptx(input: Awaited<ReturnType<typeof normalizeInput>>, options: InspectOptions): Promise<InspectResult> {
  const zip = await loadZip(input, { zipSafety: { config: options.config } });
  const paths = sortedZipFiles(zip);
  const mediaPaths = paths.filter((path) => /^ppt\/media\//i.test(path));
  const { slides, objectMap } = await inspectSlides(zip);
  const summaryDepth = options.depth === "summary";
  const themePaths = paths.filter((path) => /^ppt\/theme\/theme\d+\.xml$/i.test(path));
  const masterPaths = paths.filter((path) => /^ppt\/slideMasters\/slideMaster\d+\.xml$/i.test(path));
  const layoutPaths = paths.filter((path) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/i.test(path));
  const chartPaths = paths.filter((path) => /^ppt\/charts\/chart\d+\.xml$/i.test(path));
  const slidePayload = summaryDepth
    ? slides.map((slide) => ({
        stableObjectId: slide.stableObjectId,
        index: slide.index,
        sourcePath: slide.sourcePath,
        textPreview: slide.text.slice(0, 300),
        textObjectCount: slide.textObjects.length,
        shapeCount: slide.shapeCount,
        pictureCount: slide.pictureCount,
        chartCount: slide.chartCount,
        untrusted: true
      }))
    : slides;

  const macros = paths.filter((path) => /vbaProject\.bin$/i.test(path));
  return withObjectGraph({
    schema: "officegen.inspect.result@1.2",
    trusted: trustedMeta(
      "officegen.inspect.result@1.2",
      input,
      {
        slides: slides.length,
        textObjects: objectMap.length,
        semanticTextObjects: objectMap.filter((entry) => (entry as { semantic?: unknown }).semantic).length,
        assets: mediaPaths.length,
        charts: chartPaths.length,
        masters: masterPaths.length,
        layouts: layoutPaths.length,
        themes: themePaths.length,
        macros: macros.length,
        zipEntries: paths.length
      },
      ["PPTX inspect is zip/XML based; animation and theme resolution are summarized only.", ...zipSafetyCaveats(getLoadedZipSafetyReport(zip))]
    ),
    untrusted: {
      slides: slidePayload,
      assets: mediaPaths.map((path, index) => ({
        stableObjectId: makeStableObjectId("pptx", "deck", "asset", index + 1),
        path,
        fileName: path.split("/").pop(),
        untrusted: true
      })),
      designInventory: {
        themes: themePaths,
        masters: masterPaths,
        layouts: layoutPaths,
        charts: chartPaths,
        placeholders: objectMap
          .filter((entry) => entry.selectorHints?.placeholder)
          .slice(0, summaryDepth ? 40 : undefined)
          .map((entry) => ({
            stableObjectId: entry.stableObjectId,
            slide: entry.selectorHints?.slide,
            placeholder: entry.selectorHints?.placeholder,
            label: entry.label,
            untrusted: true
          }))
      },
      ...(options.depth === "full" || options.include?.includes("rawPaths") ? { rawPaths: paths } : {})
    },
    objectMap: summaryDepth ? compactObjectMap(objectMap, 25) : objectMap,
    agentInstruction: AGENT_UNTRUSTED_INSTRUCTION
  }, objectMap, input, options, riskFlagsFromMacros(macros, "PPTX_MACROS_PRESENT"));
}

function docxStructureObjectMap(structureMap: Record<string, unknown> | undefined): InspectResult["objectMap"] {
  if (!structureMap) return [];
  const entries: InspectResult["objectMap"] = [];
  for (const [index, header] of ((structureMap.headerFooterVariants as Record<string, string[]> | undefined)?.headers ?? []).entries()) {
    entries.push({
      stableObjectId: makeStableObjectId("docx", "structure", "header", index + 1),
      kind: "header",
      label: header,
      xmlPath: header,
      selectorHints: { headerPath: header },
      editableOps: ["docx.setHeader", "docx.headerFooter.setText"],
      trust: { level: "untrusted", reason: "document-content" },
      untrusted: true
    });
  }
  for (const [index, footer] of ((structureMap.headerFooterVariants as Record<string, string[]> | undefined)?.footers ?? []).entries()) {
    entries.push({
      stableObjectId: makeStableObjectId("docx", "structure", "footer", index + 1),
      kind: "footer",
      label: footer,
      xmlPath: footer,
      selectorHints: { footerPath: footer },
      editableOps: ["docx.setFooter", "docx.headerFooter.setText"],
      trust: { level: "untrusted", reason: "document-content" },
      untrusted: true
    });
  }
  for (const [index, style] of ((structureMap.styles as string[] | undefined) ?? []).entries()) {
    entries.push({
      stableObjectId: makeStableObjectId("docx", "structure", "style", index + 1),
      kind: "style",
      label: style,
      xmlPath: "word/styles.xml",
      selectorHints: { styleId: style },
      editableOps: ["docx.setStyle", "docx.applyStyle"],
      trust: { level: "untrusted", reason: "style-definition" },
      untrusted: true
    });
  }
  return entries;
}

async function inspectDocx(input: Awaited<ReturnType<typeof normalizeInput>>, options: InspectOptions): Promise<InspectResult> {
  const zip = await loadZip(input, { zipSafety: { config: options.config } });
  const paths = sortedZipFiles(zip);
  const { paragraphs, objectMap, storyGraph, runGraph } = await inspectParagraphs(zip);
  const mediaPaths = paths.filter((path) => /^word\/media\//i.test(path));
  const headerPaths = paths.filter((path) => /^word\/header\d+\.xml$/i.test(path));
  const footerPaths = paths.filter((path) => /^word\/footer\d+\.xml$/i.test(path));
  const commentPaths = paths.filter((path) => /^word\/comments\.xml$/i.test(path));
  const stylePaths = paths.filter((path) => /^word\/styles\.xml$/i.test(path));
  const macros = paths.filter((path) => /vbaProject\.bin$/i.test(path));
  const summaryDepth = options.depth === "summary";
  const structureMap = options.structure ? await inspectDocxStructure(zip, paths) : undefined;

  const fullObjectMap = [...objectMap, ...docxStructureObjectMap(structureMap)];
  return withObjectGraph({
    schema: "officegen.inspect.result@1.2",
    trusted: trustedMeta(
      "officegen.inspect.result@1.2",
      input,
      {
        paragraphs: paragraphs.length,
        textObjects: objectMap.length,
        assets: mediaPaths.length,
        headers: headerPaths.length,
        footers: footerPaths.length,
        comments: commentPaths.length,
        styles: stylePaths.length,
        macros: macros.length,
        zipEntries: paths.length
      },
      ["DOCX inspect reads main document XML; headers, footers, fields, and styles are summarized.", ...zipSafetyCaveats(getLoadedZipSafetyReport(zip))]
    ),
    untrusted: {
      paragraphs: summaryDepth ? paragraphs.slice(0, 50).map((paragraph) => ({ ...paragraph, text: paragraph.text.slice(0, 300) })) : paragraphs,
      storyGraph,
      runGraph,
      documentParts: {
        headers: headerPaths,
        footers: footerPaths,
        comments: commentPaths,
        styles: stylePaths
      },
      ...(structureMap ? { structureMap } : {}),
      assets: mediaPaths.map((path, index) => ({
        stableObjectId: makeStableObjectId("docx", "body", "asset", index + 1),
        path,
        fileName: path.split("/").pop(),
        untrusted: true
      })),
      ...(options.depth === "full" || options.include?.includes("rawPaths") ? { rawPaths: paths } : {})
    },
    objectMap: summaryDepth ? compactObjectMap(fullObjectMap, 35) : fullObjectMap,
    agentInstruction: AGENT_UNTRUSTED_INSTRUCTION
  }, fullObjectMap, input, options, riskFlagsFromMacros(macros, "DOCX_MACROS_PRESENT"));
}

async function inspectXlsx(input: Awaited<ReturnType<typeof normalizeInput>>, options: InspectOptions): Promise<InspectResult> {
  const zip = await loadZip(input, { zipSafety: { config: options.config } });
  const paths = sortedZipFiles(zip);
  const { sheets, objectMap, sharedStrings } = await inspectSheets(zip);
  const workbookXml = (await readZipText(zip, "xl/workbook.xml")) ?? "";
  const sheetNames = readWorkbookSheetNames(workbookXml);
  const namedSheets = sheets.map((sheet, index) => ({
    ...sheet,
    name: sheetNames[index] ?? `Sheet${index + 1}`
  }));
  const namedObjectMap = objectMap.map((entry) => {
    const sheetIndex = typeof entry.selectorHints?.sheet === "number"
      ? entry.selectorHints.sheet
      : sheetIndexFromWorksheetPath(entry.sourcePath ?? entry.xmlPath);
    if (!sheetIndex) return entry;
    return {
      ...entry,
      selectorHints: {
        ...entry.selectorHints,
        sheetName: sheetNames[sheetIndex - 1] ?? `Sheet${sheetIndex}`
      }
    };
  });
  const scopedSheets = scopeSheets(namedSheets, options.sheet, options.range);
  const scopedObjectMap = scopeObjectMap(namedObjectMap, options.sheet, options.range);

  const macros = paths.filter((path) => /vbaProject\.bin$/i.test(path));
  const summaryDepth = options.depth === "summary";
  const worksheetXml = await Promise.all(paths
    .filter((path) => /^xl\/worksheets\/sheet\d+\.xml$/i.test(path))
    .map(async (path) => (await readZipText(zip, path)) ?? ""));
  const formulaCount = worksheetXml.reduce((count, xml) => count + (xml.match(/<f\b/g) ?? []).length, 0);
  const tablePaths = paths.filter((path) => /^xl\/tables\//i.test(path));
  const chartPaths = paths.filter((path) => /^xl\/charts\//i.test(path));
  const pivotPaths = paths.filter((path) => /^xl\/pivotTables\//i.test(path));
  const slicerPaths = paths.filter((path) => /^xl\/slicers\//i.test(path) || /^xl\/slicerCaches\//i.test(path));
  const definedNames = await readDefinedNames(zip);
  const workbookMap = await inspectWorkbookMap(zip, paths, worksheetXml, definedNames);
  const cellCount = scopedSheets.reduce((count, sheet) => count + sheet.cells.length, 0);
  const sheetSummaries = summaryDepth
    ? scopedSheets.map((sheet) => ({
        stableObjectId: sheet.stableObjectId,
        index: sheet.index,
        sourcePath: sheet.sourcePath,
        cellCount: sheet.cells.length,
        usedRange: usedRangeFromCells(sheet.cells.map((cell: Record<string, string>) => cell.ref)),
        previewCells: sheet.cells.slice(0, 20).map((cell: Record<string, string>) => ({
          stableObjectId: cell.stableObjectId,
          ref: cell.ref,
          valuePreview: cell.value.slice(0, 120),
          sourcePath: cell.sourcePath,
          untrusted: true
        }))
      }))
    : namedSheets;
  return withObjectGraph({
    schema: "officegen.inspect.result@1.2",
    trusted: trustedMeta(
      "officegen.inspect.result@1.2",
      input,
      {
        sheets: sheets.length,
        cells: cellCount,
        sharedStrings: sharedStrings.length,
        formulas: formulaCount,
        tables: tablePaths.length,
        charts: chartPaths.length,
        pivotTables: pivotPaths.length,
        slicers: slicerPaths.length,
        definedNames: definedNames.length,
        macros: macros.length,
        zipEntries: paths.length
      },
      ["XLSX inspect reads cached cell values; formulas, styles, and charts are summarized.", ...zipSafetyCaveats(getLoadedZipSafetyReport(zip))]
    ),
    untrusted: {
      sheets: sheetSummaries,
        workbookObjects: {
          tables: tablePaths,
          charts: chartPaths,
          pivotTables: pivotPaths,
          slicers: slicerPaths,
          definedNames
        },
        workbookMap,
        scope: {
          sheet: options.sheet,
          range: options.range,
          scoped: Boolean(options.sheet || options.range)
        },
      ...(options.depth === "full" || options.include?.includes("rawPaths") ? { rawPaths: paths } : {})
    },
    objectMap: summaryDepth ? compactObjectMap(scopedObjectMap, 50) : scopedObjectMap,
    agentInstruction: AGENT_UNTRUSTED_INSTRUCTION
  }, scopedObjectMap, input, options, riskFlagsFromMacros(macros, "XLSX_MACROS_PRESENT"));
}

async function inspectPdf(input: Awaited<ReturnType<typeof normalizeInput>>, options: InspectOptions): Promise<InspectResult> {
  const pdf = await PDFDocument.load(input.bytes, { ignoreEncryption: true });
  const summaryDepth = options.depth === "summary";
  const pageSizes = pdf.getPages().map((page) => page.getSize());
  const graph = await inspectPdfObjectGraph(input.bytes, pageSizes);
  const pageText = groupPdfTextByPage(graph.textBlocks);
  const pages = pageSizes.map((size, index) => {
    const text = pageText.get(index + 1) ?? "";
    return {
      stableObjectId: makeStableObjectId("pdf", "document", "page", index + 1),
      index: index + 1,
      width: size.width,
      height: size.height,
      textBlockCount: graph.textBlocks.filter((block) => block.page === index + 1).length,
      annotationCount: graph.annotations.filter((annotation) => annotation.page === index + 1).length,
      textPreview: text.slice(0, summaryDepth ? 300 : 1000),
      untrusted: true
    };
  });
  const objectMap = pdfObjectMap(pages, graph, summaryDepth);
  const metadata = {
    title: pdf.getTitle(),
    author: pdf.getAuthor(),
    subject: pdf.getSubject(),
    keywords: pdf.getKeywords(),
    creator: pdf.getCreator(),
    producer: pdf.getProducer(),
    creationDate: pdf.getCreationDate()?.toISOString(),
    modificationDate: pdf.getModificationDate()?.toISOString(),
    ...graph.metadata
  };
  return withObjectGraph({
    schema: "officegen.inspect.result@1.2",
    trusted: trustedMeta(
      "officegen.inspect.result@1.2",
      input,
      {
        pages: pages.length,
        textBlocks: graph.textBlocks.length,
        annotations: graph.annotations.length,
        images: graph.scan.imageObjects,
        embeddedFiles: graph.scan.embeddedFiles,
        encrypted: graph.scan.encrypted,
        unsupportedFilters: graph.scan.unsupportedFilters.length,
        riskFlags: graph.riskFlags.length
      },
      [
        "PDF text extraction is best-effort; scanned/image PDFs should be reviewed through page preview artifacts.",
        "PDF redact operations are intentionally blocked; overlay text or rectangles do not remove underlying content.",
        ...(graph.scan.encrypted ? ["PDF_ENCRYPTED: inspect is allowed for risk reporting, but PDF mutation/export operations are blocked by default."] : []),
        ...graph.caveats
      ]
    ),
    untrusted: {
      pages,
      pdfGraph: {
        pageCount: graph.pageCount,
        textBlocks: summaryDepth ? graph.textBlocks.slice(0, 40).map((block) => ({ ...block, text: block.text.slice(0, 240) })) : graph.textBlocks,
        annotations: summaryDepth ? graph.annotations.slice(0, 40) : graph.annotations,
        metadata,
        scan: graph.scan,
        riskFlags: graph.riskFlags
      },
      ...(summaryDepth ? {} : { text: pages.map((page) => pageText.get(page.index) ?? "") }),
      qualityWarnings: graph.textBlocks.length
        ? []
        : [{
            code: "PDF_TEXT_BLOCKS_ZERO",
            severity: "warning",
            message: "No extractable PDF text was found.",
            aiVisionRecommended: true,
            previewCommand: "officegen view input.pdf --out .officegen/runs/pdf-view --json",
            doctorCommand: "officegen renderer doctor --json"
          }]
    },
    objectMap: summaryDepth ? compactObjectMap(objectMap, 60) : objectMap,
    agentInstruction: AGENT_UNTRUSTED_INSTRUCTION
  }, objectMap, input, options, graph.riskFlags);
}

function withObjectGraph(
  result: InspectResult,
  fullObjectMap: InspectResult["objectMap"],
  input: Awaited<ReturnType<typeof normalizeInput>>,
  options: InspectOptions,
  riskFlags: ObjectGraphRiskFlag[] = []
): InspectResult {
  if (!options.includeObjectGraph && options.emit !== "object-graph") return result;
  return {
    ...result,
    objectGraph: buildObjectGraph(fullObjectMap, {
      format: input.format,
      inputPath: input.path,
      inputSha256: result.trusted.sha256,
      nodeOffset: options.objectGraph?.nodeOffset,
      nodeLimit: options.objectGraph?.nodeLimit,
      edgeOffset: options.objectGraph?.edgeOffset,
      edgeLimit: options.objectGraph?.edgeLimit,
      riskFlags
    })
  };
}

function riskFlagsFromMacros(paths: string[], code: string): ObjectGraphRiskFlag[] {
  return paths.length
    ? [{
        code,
        severity: "warning",
        message: "Macro project content was detected; inspect treats macros as package content only.",
        source: paths[0]
      }]
    : [];
}

function groupPdfTextByPage(textBlocks: Array<{ page: number; text: string }>): Map<number, string> {
  const pages = new Map<number, string[]>();
  for (const block of textBlocks) {
    const text = block.text.trim();
    if (!text) continue;
    pages.set(block.page, [...(pages.get(block.page) ?? []), text]);
  }
  return new Map([...pages.entries()].map(([page, blocks]) => [page, blocks.join(" ").replace(/\s+/g, " ").trim()]));
}

function pdfObjectMap(
  pages: Array<{ stableObjectId: string; index: number; width: number; height: number; textPreview: string }>,
  graph: Awaited<ReturnType<typeof inspectPdfObjectGraph>>,
  summaryDepth: boolean
): InspectResult["objectMap"] {
  const entries: InspectResult["objectMap"] = [];
  for (const block of graph.textBlocks) {
    const width = block.width ?? Math.max(24, block.text.length * 6);
    const height = block.height ?? 12;
    const x = block.x ?? 24;
    const y = block.y ?? (56 + (block.index - 1) * 18);
    entries.push({
      stableObjectId: makeStableObjectId("pdf", `page-${String(block.page).padStart(4, "0")}`, "text", block.index),
      kind: "pdfText",
      text: summaryDepth ? undefined : block.text,
      textPreview: block.text.slice(0, 240),
      bounds: { x, y, width, height },
      bbox: [x, y, width, height],
      selectorHints: { page: block.page, textBlock: block.index, source: block.source },
      editableOps: ["pdf.textOverlay", "pdf.annotation"],
      trust: { level: "untrusted", reason: "document-content" },
      untrusted: true
    });
  }
  for (const annotation of graph.annotations) {
    const [x1 = 24, y1 = 56, x2 = x1 + 120, y2 = y1 + 32] = annotation.rect ?? [];
    entries.push({
      stableObjectId: makeStableObjectId("pdf", `page-${String(annotation.page).padStart(4, "0")}`, "annotation", annotation.index),
      kind: "pdfAnnotation",
      label: annotation.subtype,
      text: summaryDepth ? undefined : annotation.contents,
      textPreview: annotation.contents?.slice(0, 240),
      bounds: { x: x1, y: y1, width: Math.abs(x2 - x1), height: Math.abs(y2 - y1) },
      bbox: [x1, y1, Math.abs(x2 - x1), Math.abs(y2 - y1)],
      selectorHints: { page: annotation.page, annotation: annotation.index, subtype: annotation.subtype },
      trust: { level: "untrusted", reason: "document-content" },
      untrusted: true
    });
  }
  entries.push(...pages.map((page) => ({
    stableObjectId: page.stableObjectId,
    kind: "pdfPage",
    label: `Page ${page.index}`,
    textPreview: page.textPreview,
    bounds: { x: 0, y: 0, width: page.width, height: page.height },
    bbox: [0, 0, page.width, page.height] as [number, number, number, number],
    selectorHints: { page: page.index },
    editableOps: ["pdf.textOverlay", "pdf.annotation"],
    trust: { level: "untrusted" as const, reason: "document-content" },
    untrusted: true as const
  })));
  return entries;
}

function extractPdfTextPreview(bytes: Uint8Array): { pages: string[]; imageRefs: number; caveats: string[] } {
  const raw = Buffer.from(bytes).toString("latin1");
  const pageChunks = raw.split(/\/Type\s*\/Page\b/g).slice(1);
  const chunks = pageChunks.length ? pageChunks : [raw];
  const pages = chunks.map((chunk) => extractPdfTextFromChunk(chunk).slice(0, 8000));
  const imageRefs = (raw.match(/\/Subtype\s*\/Image\b/g) ?? []).length;
  const caveats = [];
  if (/\/Filter\s*\/(?:FlateDecode|DCTDecode|JPXDecode|LZWDecode)/.test(raw)) {
    caveats.push("Some PDF streams are compressed or image-based; text preview may be incomplete.");
  }
  if (!pages.some(Boolean)) {
    caveats.push("No plain text operators were found; use page preview artifacts or native PDF tooling for scanned/compressed PDFs.");
  }
  return { pages, imageRefs, caveats };
}

function extractPdfTextFromChunk(chunk: string): string {
  const strings: string[] = [];
  for (const match of chunk.matchAll(/\((?:\\.|[^\\)])*\)\s*Tj/g)) {
    strings.push(decodePdfLiteral(match[0].replace(/\)\s*Tj$/, "").slice(1)));
  }
  for (const match of chunk.matchAll(/\[((?:\s*\((?:\\.|[^\\)])*\)\s*-?\d*)+)\]\s*TJ/g)) {
    const array = match[1] ?? "";
    for (const item of array.matchAll(/\((?:\\.|[^\\)])*\)/g)) strings.push(decodePdfLiteral(item[0].slice(1, -1)));
  }
  return strings.join(" ").replace(/\s+/g, " ").trim();
}

function decodePdfLiteral(value: string): string {
  return value
    .replace(/\\([nrtbf()\\])/g, (_match, code: string) => ({ n: "\n", r: "\r", t: "\t", b: "\b", f: "\f", "(": "(", ")": ")", "\\": "\\" })[code] ?? code)
    .replace(/\\([0-7]{1,3})/g, (_match, octal: string) => String.fromCharCode(Number.parseInt(octal, 8)))
    .replace(/\)$/g, "");
}

function scopeSheets(sheets: Array<Record<string, any>>, sheetName: string | undefined, range: string | undefined): Array<Record<string, any>> {
  const bounds = parseA1Range(range);
  return sheets
    .filter((sheet) => !sheetName || String(sheet.name ?? sheet.sourcePath ?? "").toLowerCase().includes(sheetName.toLowerCase()))
    .map((sheet) => ({
      ...sheet,
      cells: Array.isArray(sheet.cells)
        ? sheet.cells.filter((cell: Record<string, unknown>) => !bounds || cellInRange(String(cell.ref ?? ""), bounds))
        : []
    }));
}

function scopeObjectMap(objectMap: InspectResult["objectMap"], sheetName: string | undefined, range: string | undefined): InspectResult["objectMap"] {
  const bounds = parseA1Range(range);
  return objectMap.filter((entry) => {
    if (sheetName) {
      const sheetCandidates = [
        entry.selectorHints?.sheetName,
        entry.selectorHints?.sheet,
        entry.sourcePath
      ].map((value) => String(value ?? "").toLowerCase());
      if (!sheetCandidates.some((value) => value.includes(sheetName.toLowerCase()))) return false;
    }
    if (!bounds) return true;
    const ref = String(entry.selectorHints?.ref ?? entry.label ?? "");
    return cellInRange(ref, bounds);
  });
}

function readWorkbookSheetNames(workbookXml: string): string[] {
  return [...workbookXml.matchAll(/<sheet\b([^>]*)/g)].map((match, index) =>
    decodeXmlAttr(/\bname="([^"]+)"/.exec(match[1] ?? "")?.[1] ?? `Sheet${index + 1}`)
  );
}

function sheetIndexFromWorksheetPath(path: string | undefined): number | undefined {
  const match = /^xl\/worksheets\/sheet(\d+)\.xml$/i.exec(path ?? "");
  return match ? Number(match[1]) : undefined;
}

function parseA1Range(range: string | undefined): { minCol: number; maxCol: number; minRow: number; maxRow: number } | undefined {
  if (!range) return undefined;
  const match = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i.exec(range.trim());
  if (!match) return undefined;
  const left = { col: columnIndex(match[1] ?? "A"), row: Number(match[2]) };
  const right = { col: columnIndex(match[3] ?? "A"), row: Number(match[4]) };
  return {
    minCol: Math.min(left.col, right.col),
    maxCol: Math.max(left.col, right.col),
    minRow: Math.min(left.row, right.row),
    maxRow: Math.max(left.row, right.row)
  };
}

function cellInRange(ref: string, bounds: ReturnType<typeof parseA1Range>): boolean {
  if (!bounds) return true;
  const match = /^([A-Z]+)(\d+)$/i.exec(ref);
  if (!match) return false;
  const col = columnIndex(match[1] ?? "A");
  const row = Number(match[2]);
  return col >= bounds.minCol && col <= bounds.maxCol && row >= bounds.minRow && row <= bounds.maxRow;
}

async function inspectWorkbookMap(
  zip: Awaited<ReturnType<typeof loadZip>>,
  paths: string[],
  worksheetXml: string[],
  definedNames: Array<{ name: string; ref: string; untrusted: true }>
): Promise<Record<string, unknown>> {
  const workbookXml = (await readZipText(zip, "xl/workbook.xml")) ?? "";
  const sheetStates = [...workbookXml.matchAll(/<sheet\b([^>]*)/g)].map((match, index) => {
    const attrs = match[1] ?? "";
    return {
      index: index + 1,
      name: decodeXmlAttr(/\bname="([^"]+)"/.exec(attrs)?.[1] ?? `Sheet${index + 1}`),
      hidden: /\bstate="hidden"/i.test(attrs),
      veryHidden: /\bstate="veryHidden"/i.test(attrs),
      role: inferSheetRole(index + 1, worksheetXml[index] ?? "")
    };
  });
  return {
    sheets: sheetStates,
    formulas: worksheetXml.map((xml, index) => ({
      sheetIndex: index + 1,
      count: (xml.match(/<f\b/g) ?? []).length,
      samples: [...xml.matchAll(/<c\b[^>]*\br="([^"]+)"[\s\S]*?<f\b[^>]*>([\s\S]*?)<\/f>/g)]
        .slice(0, 12)
        .map((match) => ({ ref: match[1], formula: decodeXmlEntities(match[2] ?? ""), untrusted: true }))
    })),
    inputCells: worksheetXml.map((xml, index) => ({
      sheetIndex: index + 1,
      samples: [...xml.matchAll(/<c\b[^>]*\br="([^"]+)"(?![\s\S]*?<f\b)[\s\S]*?<\/c>/g)]
        .slice(0, 20)
        .map((match) => ({ ref: match[1], untrusted: true }))
    })),
    validations: paths.filter((file) => /^xl\/worksheets\/sheet\d+\.xml$/i.test(file)).map((file, index) => ({
      sheetIndex: index + 1,
      count: (worksheetXml[index]?.match(/<dataValidation\b/g) ?? []).length
    })),
    protectedSheets: worksheetXml.map((xml, index) => ({ sheetIndex: index + 1, protected: /<sheetProtection\b/i.test(xml) })).filter((entry) => entry.protected),
    namedRanges: definedNames,
    externalLinks: paths.filter((file) => /^xl\/externalLinks\//i.test(file)),
    tables: paths.filter((file) => /^xl\/tables\//i.test(file)),
    charts: paths.filter((file) => /^xl\/charts\//i.test(file)),
    pivotTables: paths.filter((file) => /^xl\/pivotTables\//i.test(file)),
    slicers: paths.filter((file) => /^xl\/slicers\//i.test(file) || /^xl\/slicerCaches\//i.test(file))
  };
}

async function inspectDocxStructure(zip: Awaited<ReturnType<typeof loadZip>>, paths: string[]): Promise<Record<string, unknown>> {
  const documentXml = (await readZipText(zip, "word/document.xml")) ?? "";
  const stylesXml = (await readZipText(zip, "word/styles.xml")) ?? "";
  const commentsXml = (await readZipText(zip, "word/comments.xml")) ?? "";
  const headerPaths = paths.filter((file) => /^word\/header\d+\.xml$/i.test(file));
  const footerPaths = paths.filter((file) => /^word\/footer\d+\.xml$/i.test(file));
  const paragraphs = [...documentXml.matchAll(/<w:p\b[\s\S]*?<\/w:p>/g)];
  return {
    headingTree: paragraphs
      .map((match, index) => {
        const block = match[0];
        const style = /<w:pStyle\b[^>]*\bw:val="([^"]+)"/.exec(block)?.[1];
        const heading = style && /^Heading/i.test(style);
        return heading ? { index: index + 1, style, text: extractDocxText(block).slice(0, 200), untrusted: true } : undefined;
      })
      .filter(Boolean),
    sections: (documentXml.match(/<w:sectPr\b/g) ?? []).length,
    headerFooterVariants: { headers: headerPaths, footers: footerPaths },
    tables: (documentXml.match(/<w:tbl\b/g) ?? []).length,
    fields: [...documentXml.matchAll(/<w:fldChar\b[^>]*|<w:instrText\b[^>]*>([\s\S]*?)<\/w:instrText>/g)].slice(0, 40).map((match) => ({ text: decodeXmlEntities(match[1] ?? match[0]), untrusted: true })),
    contentControls: (documentXml.match(/<w:sdt\b/g) ?? []).length,
    comments: (commentsXml.match(/<w:comment\b/g) ?? []).length,
    trackedChanges: (documentXml.match(/<w:(ins|del)\b/g) ?? []).length,
    fillablePlaceholders: [...documentXml.matchAll(/\{\{\s*([^}]+)\s*\}\}/g)].slice(0, 40).map((match) => ({ field: match[1], untrusted: true })),
    styles: [...stylesXml.matchAll(/<w:style\b[^>]*\bw:styleId="([^"]+)"/g)].slice(0, 80).map((match) => match[1])
  };
}

function inferSheetRole(index: number, xml: string): string {
  if ((xml.match(/<f\b/g) ?? []).length > 20) return "model";
  if ((xml.match(/<dataValidation\b/g) ?? []).length > 0) return "input";
  if (index === 1) return "primary";
  return "support";
}

function extractDocxText(xml: string): string {
  return [...xml.matchAll(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g)].map((match) => decodeXmlEntities(match[1] ?? "")).join("").replace(/\s+/g, " ").trim();
}

function decodeXmlAttr(value: string): string {
  return decodeXmlEntities(value);
}

function decodeXmlEntities(value: string): string {
  return value
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, "\"")
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

function compactObjectMap(objectMap: InspectResult["objectMap"], limit: number): InspectResult["objectMap"] {
  return objectMap.slice(0, limit).map((entry) => ({
    stableObjectId: entry.stableObjectId,
    kind: entry.kind,
    label: entry.label,
    textPreview: entry.textPreview,
    selectorHints: entry.selectorHints,
    editableOps: entry.editableOps,
    media: entry.media,
    trust: entry.trust,
    untrusted: true
  }));
}

function usedRangeFromCells(refs: string[]): string | undefined {
  const cells = refs
    .map((ref) => /^([A-Z]+)(\d+)$/i.exec(ref))
    .filter((match): match is RegExpExecArray => Boolean(match))
    .map((match) => ({ col: columnIndex(match[1] ?? "A"), row: Number(match[2]) }));
  if (!cells.length) return undefined;
  const minCol = Math.min(...cells.map((cell) => cell.col));
  const maxCol = Math.max(...cells.map((cell) => cell.col));
  const minRow = Math.min(...cells.map((cell) => cell.row));
  const maxRow = Math.max(...cells.map((cell) => cell.row));
  return `${columnName(minCol)}${minRow}:${columnName(maxCol)}${maxRow}`;
}

function columnIndex(name: string): number {
  let value = 0;
  for (const char of name.toUpperCase()) value = value * 26 + (char.charCodeAt(0) - 64);
  return value || 1;
}

function columnName(index: number): string {
  let value = index;
  let name = "";
  while (value > 0) {
    value -= 1;
    name = String.fromCharCode(65 + (value % 26)) + name;
    value = Math.floor(value / 26);
  }
  return name || "A";
}

async function readDefinedNames(zip: Awaited<ReturnType<typeof loadZip>>): Promise<Array<{ name: string; ref: string; untrusted: true }>> {
  const workbookXml = (await readZipText(zip, "xl/workbook.xml")) ?? "";
  return [...workbookXml.matchAll(/<definedName\b([^>]*)>([\s\S]*?)<\/definedName>/g)]
    .map((match) => {
      const attrs = match[1] ?? "";
      const name = /\bname="([^"]+)"/.exec(attrs)?.[1] ?? "";
      const ref = (match[2] ?? "").replace(/&apos;/g, "'").replace(/&amp;/g, "&");
      return name ? { name, ref, untrusted: true as const } : undefined;
    })
    .filter((item): item is { name: string; ref: string; untrusted: true } => Boolean(item));
}
