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
import { inspectSlides, type PptxSlide } from "./ooxml/pptx.js";
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
  slides?: string;
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
  styleInventory?: Record<string, unknown>;
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
  const scopedSlides = scopePptxSlides(slides, options.slides);
  const scopedObjectMap = scopePptxObjectMap(objectMap, options.slides);
  const summaryDepth = options.depth === "summary";
  const themePaths = paths.filter((path) => /^ppt\/theme\/theme\d+\.xml$/i.test(path));
  const masterPaths = paths.filter((path) => /^ppt\/slideMasters\/slideMaster\d+\.xml$/i.test(path));
  const layoutPaths = paths.filter((path) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/i.test(path));
  const chartPaths = paths.filter((path) => /^ppt\/charts\/chart\d+\.xml$/i.test(path));
  const styleInventory = pptxStyleInventory(scopedObjectMap);
  const layoutIssues = pptxLayoutIssues(scopedObjectMap, scopedSlides.length || slides.length);
  const slidePayload = summaryDepth
    ? scopedSlides.map((slide) => ({
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
    : scopedSlides;

  const macros = paths.filter((path) => /vbaProject\.bin$/i.test(path));
  return withObjectGraph({
    schema: "officegen.inspect.result@1.2",
    trusted: trustedMeta(
      "officegen.inspect.result@1.2",
      input,
      {
        slides: slides.length,
        scopedSlides: scopedSlides.length,
        textObjects: scopedObjectMap.length,
        semanticTextObjects: scopedObjectMap.filter((entry) => (entry as { semantic?: unknown }).semantic).length,
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
        styleInventory,
        layoutIssues: summaryDepth ? layoutIssues.slice(0, 40) : layoutIssues,
        placeholders: objectMap
          .filter((entry) => entry.selectorHints?.placeholder)
          .filter((entry) => scopedObjectMap.includes(entry))
          .slice(0, summaryDepth ? 40 : undefined)
          .map((entry) => ({
            stableObjectId: entry.stableObjectId,
            slide: entry.selectorHints?.slide,
            placeholder: entry.selectorHints?.placeholder,
            label: entry.label,
            untrusted: true
          }))
      },
      filters: {
        slides: options.slides,
        scoped: Boolean(options.slides)
      },
      ...(options.depth === "full" || options.include?.includes("rawPaths") ? { rawPaths: paths } : {})
    },
    objectMap: summaryDepth ? compactObjectMap(scopedObjectMap, 25) : scopedObjectMap,
    styleInventory,
    agentInstruction: AGENT_UNTRUSTED_INSTRUCTION
  }, scopedObjectMap, input, options, riskFlagsFromMacros(macros, "PPTX_MACROS_PRESENT"));
}

function pptxStyleInventory(objectMap: InspectResult["objectMap"]): Record<string, unknown> {
  const fontsLatin = new Set<string>();
  const fontsEastAsia = new Set<string>();
  const fontsComplexScript = new Set<string>();
  const fontSizes = new Map<number, number>();
  const languages = new Set<string>();
  let semanticObjects = 0;
  let runCount = 0;
  let boldRuns = 0;
  let italicRuns = 0;
  let noProofRuns = 0;
  let bulletObjects = 0;
  let numberingObjects = 0;
  const mixedFontObjects: Array<Record<string, unknown>> = [];

  for (const entry of objectMap) {
    const semantic = asPlainRecord((entry as { semantic?: unknown }).semantic);
    const paragraphs = semanticParagraphs(semantic);
    if (!paragraphs.length) continue;
    semanticObjects += 1;
    const objectFonts = new Set<string>();
    if (paragraphs.some((paragraph) => asPlainRecord(paragraph).bullet)) bulletObjects += 1;
    if (paragraphs.some((paragraph) => asPlainRecord(paragraph).numbering)) numberingObjects += 1;
    for (const paragraph of paragraphs) {
      for (const runValue of semanticRuns(asPlainRecord(paragraph))) {
        const run = asPlainRecord(runValue);
        runCount += 1;
        if (run.bold === true) boldRuns += 1;
        if (run.italic === true) italicRuns += 1;
        if (run.noProof === true) noProofRuns += 1;
        addString(fontsLatin, run.fontFamilyLatin);
        addString(fontsEastAsia, run.fontFamilyEastAsia);
        addString(fontsComplexScript, run.fontFamilyComplexScript);
        addString(languages, run.lang);
        addString(objectFonts, run.fontFamilyLatin);
        addString(objectFonts, run.fontFamilyEastAsia);
        addString(objectFonts, run.fontFamilyComplexScript);
        const fontSize = Number(run.fontSizePt);
        if (Number.isFinite(fontSize) && fontSize > 0) fontSizes.set(fontSize, (fontSizes.get(fontSize) ?? 0) + 1);
      }
    }
    if (objectFonts.size > 2) {
      mixedFontObjects.push({
        stableObjectId: entry.stableObjectId,
        slide: entry.selectorHints?.slide,
        fonts: [...objectFonts].sort().slice(0, 8),
        untrusted: true
      });
    }
  }

  return {
    schema: "officegen.styleInventory@1.0",
    source: "pptx.objectMap.semantic",
    objectCount: objectMap.length,
    semanticObjects,
    runCount,
    fonts: {
      latin: [...fontsLatin].sort(),
      eastAsia: [...fontsEastAsia].sort(),
      complexScript: [...fontsComplexScript].sort()
    },
    fontSizesPt: [...fontSizes.entries()].sort((left, right) => left[0] - right[0]).map(([sizePt, count]) => ({ sizePt, count })),
    languages: [...languages].sort(),
    boldRuns,
    italicRuns,
    noProofRuns,
    bulletObjects,
    numberingObjects,
    mixedFontObjects: mixedFontObjects.slice(0, 20),
    untrusted: true
  };
}

function pptxLayoutIssues(objectMap: InspectResult["objectMap"], slideCount: number): Array<Record<string, unknown>> {
  const issues: Array<Record<string, unknown>> = [];
  const perSlide = new Map<number, { textObjects: number; characters: number }>();
  for (const entry of objectMap) {
    const slide = Number(entry.selectorHints?.slide ?? slideIndexFromPath(entry.sourcePath ?? entry.xmlPath) ?? 1);
    const text = String(entry.text ?? entry.textPreview ?? "");
    if (text.trim()) {
      const current = perSlide.get(slide) ?? { textObjects: 0, characters: 0 };
      current.textObjects += 1;
      current.characters += text.length;
      perSlide.set(slide, current);
    }
    if (entry.bounds && (entry.bounds.x < -1 || entry.bounds.y < -1 || entry.bounds.x + entry.bounds.width > 961 || entry.bounds.y + entry.bounds.height > 541)) {
      issues.push(pptxLayoutIssue("PPTX_OBJECT_OFF_SLIDE", "warning", entry, "Object bounds extend beyond the nominal 16:9 slide area.", { bounds: entry.bounds }));
    }
    const semantic = asPlainRecord((entry as { semantic?: unknown }).semantic);
    const paragraphs = semanticParagraphs(semantic);
    if (!paragraphs.length || !text.trim()) continue;
    const runFormats = paragraphs.flatMap((paragraph) => semanticRuns(asPlainRecord(paragraph)).map(asPlainRecord));
    const fontSizes = runFormats.map((run) => Number(run.fontSizePt)).filter((size) => Number.isFinite(size) && size > 0);
    const medianFontSize = median(fontSizes) ?? (entry.kind === "shape" ? 18 : 12);
    if (fontSizes.some((size) => size < 9)) {
      issues.push(pptxLayoutIssue("PPTX_TINY_TEXT", "warning", entry, "Text contains font sizes below 9pt, which is risky for human presentation review.", { fontSizesPt: fontSizes.slice(0, 12) }));
    }
    if (runFormats.some((run) => run.italic === true) && text.length > 24) {
      issues.push(pptxLayoutIssue("PPTX_ITALIC_BODY_TEXT", "info", entry, "Long body text has italic formatting; this often indicates stray run formatting after replacement.", {}));
    }
    const fontFamilies = new Set<string>();
    for (const run of runFormats) {
      addString(fontFamilies, run.fontFamilyLatin);
      addString(fontFamilies, run.fontFamilyEastAsia);
      addString(fontFamilies, run.fontFamilyComplexScript);
    }
    if (fontFamilies.size > 2) {
      issues.push(pptxLayoutIssue("PPTX_MIXED_FONT_FAMILIES", "warning", entry, "A text object mixes several explicit font families; Japanese/Latin fallback may look inconsistent.", { fonts: [...fontFamilies].sort().slice(0, 8) }));
    }
    if (entry.bounds && estimatedTextHeight(text, entry.bounds.width, medianFontSize, paragraphs.length) > entry.bounds.height * 1.15) {
      issues.push(pptxLayoutIssue("PPTX_TEXT_OVERFLOW_ESTIMATE", "warning", entry, "Estimated wrapped text height exceeds the shape bounds; verify in PowerPoint or run fit/repair before final use.", {
        bounds: entry.bounds,
        fontSizePt: medianFontSize,
        textLength: text.length
      }));
    }
  }
  for (const [slide, density] of perSlide.entries()) {
    if (density.textObjects > 18 || density.characters > 950) {
      issues.push({
        code: "PPTX_SLIDE_DENSITY_HIGH",
        severity: "warning",
        slide,
        message: "Slide has high text density and may be hard to scan.",
        metrics: density,
        repairCommand: "officegen improve <input.pptx> --profile business --agent --json",
        untrusted: true
      });
    }
  }
  if (slideCount > 1 && perSlide.size < slideCount) {
    issues.push({
      code: "PPTX_BLANK_SLIDE_RISK",
      severity: "warning",
      message: "One or more slides have no detected text objects.",
      metrics: { slideCount, slidesWithText: perSlide.size },
      repairCommand: "officegen inspect <input.pptx> --depth full --agent --json",
      untrusted: true
    });
  }
  return issues.slice(0, 80);
}

function pptxLayoutIssue(code: string, severity: "info" | "warning", entry: InspectResult["objectMap"][number], message: string, metrics: Record<string, unknown>): Record<string, unknown> {
  return {
    code,
    severity,
    stableObjectId: entry.stableObjectId,
    slide: entry.selectorHints?.slide,
    message,
    metrics,
    repairCommand: "officegen improve <input.pptx> --profile business --agent --json",
    untrusted: true
  };
}

function asPlainRecord(value: unknown): Record<string, unknown> {
  return value && typeof value === "object" && !Array.isArray(value) ? value as Record<string, unknown> : {};
}

function semanticParagraphs(semantic: Record<string, unknown>): Array<Record<string, unknown>> {
  return Array.isArray(semantic.paragraphs) ? semantic.paragraphs.map(asPlainRecord) : [];
}

function semanticRuns(paragraph: Record<string, unknown>): Array<Record<string, unknown>> {
  return Array.isArray(paragraph.runs) ? paragraph.runs.map(asPlainRecord) : [];
}

function addString(target: Set<string>, value: unknown): void {
  if (typeof value === "string" && value.trim()) target.add(value.trim());
}

function median(values: number[]): number | undefined {
  const sorted = values.filter((value) => Number.isFinite(value)).sort((left, right) => left - right);
  if (!sorted.length) return undefined;
  return sorted[Math.floor(sorted.length / 2)];
}

function estimatedTextHeight(text: string, widthPx: number, fontSizePt: number, paragraphCount: number): number {
  const fontPx = Math.max(8, fontSizePt * 1.333);
  const charsPerLine = Math.max(8, Math.floor(Math.max(48, widthPx) / Math.max(4.5, fontPx * 0.52)));
  const explicitLines = text.split(/\r?\n/).reduce((sum, line) => sum + Math.max(1, Math.ceil(line.length / charsPerLine)), 0);
  const lines = Math.max(explicitLines, paragraphCount);
  return lines * fontPx * 1.22 + Math.max(0, paragraphCount - 1) * fontPx * 0.28 + 8;
}

function scopePptxSlides(slides: PptxSlide[], slideRange: string | undefined): PptxSlide[] {
  const selected = parseNumberSelection(slideRange);
  if (!selected) return slides;
  return slides.filter((slide) => selected.has(slide.index));
}

function scopePptxObjectMap(objectMap: InspectResult["objectMap"], slideRange: string | undefined): InspectResult["objectMap"] {
  const selected = parseNumberSelection(slideRange);
  if (!selected) return objectMap;
  return objectMap.filter((entry) => {
    const slide = typeof entry.selectorHints?.slide === "number"
      ? entry.selectorHints.slide
      : slideIndexFromPath(entry.sourcePath ?? entry.xmlPath);
    return slide !== undefined && selected.has(slide);
  });
}

function parseNumberSelection(value: string | undefined): Set<number> | undefined {
  if (!value) return undefined;
  const selected = new Set<number>();
  for (const part of value.split(",")) {
    const trimmed = part.trim();
    if (!trimmed) continue;
    const range = /^(\d+)\s*-\s*(\d+)$/.exec(trimmed);
    if (range) {
      const start = Number(range[1]);
      const end = Number(range[2]);
      for (let current = Math.min(start, end); current <= Math.max(start, end); current += 1) selected.add(current);
      continue;
    }
    const single = Number(trimmed);
    if (Number.isInteger(single) && single > 0) selected.add(single);
  }
  return selected.size ? selected : undefined;
}

function slideIndexFromPath(path: string | undefined): number | undefined {
  const match = /^ppt\/slides\/slide(\d+)\.xml$/i.exec(path ?? "");
  return match ? Number(match[1]) : undefined;
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
  const styleInventory = await inspectDocxStyleInventory(zip);

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
    ...(styleInventory ? { styleInventory } : {}),
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
      samples: extractWorksheetFormulaSamples(xml)
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

function extractWorksheetFormulaSamples(xml: string): Array<{ ref: string; formula: string; untrusted: true }> {
  const samples: Array<{ ref: string; formula: string; untrusted: true }> = [];
  for (const match of xml.matchAll(/<c\b([^>]*)>([\s\S]*?)<\/c>/g)) {
    const attrs = match[1] ?? "";
    const inner = match[2] ?? "";
    const ref = /\br="([^"]+)"/.exec(attrs)?.[1];
    const formulaMatch = /<f\b[^>]*>([\s\S]*?)<\/f>/.exec(inner);
    if (!ref || !formulaMatch) continue;
    samples.push({ ref, formula: decodeXmlEntities(formulaMatch[1] ?? ""), untrusted: true });
    if (samples.length >= 12) break;
  }
  return samples;
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

async function inspectDocxStyleInventory(zip: Awaited<ReturnType<typeof loadZip>>): Promise<{ sourcePath: string; styleCount: number; styles: string[] } | undefined> {
  const stylesXml = await readZipText(zip, "word/styles.xml");
  if (!stylesXml) return undefined;
  const styles = [...stylesXml.matchAll(/<w:style\b([^>]*)>([\s\S]*?)<\/w:style>/g)]
    .map((match) => {
      const attrs = match[1] ?? "";
      const body = match[2] ?? "";
      const styleId = /\bw:styleId="([^"]+)"/.exec(attrs)?.[1];
      const styleName = /<w:name\b[^>]*\bw:val="([^"]+)"/.exec(body)?.[1];
      return decodeXmlAttr(styleName ?? styleId ?? "");
    })
    .filter(Boolean);
  return { sourcePath: "word/styles.xml", styleCount: styles.length, styles: styles.slice(0, 80) };
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
