import {
  type InputLike,
  type OfficegenConfig,
  type ObjectMapEntry,
  getLoadedZipSafetyReport,
  isOfficeFormat,
  loadZip,
  normalizeInput,
  readZipBytes,
  readZipText,
  writeOutput,
  zipSafetyCaveats,
  zipToBytes
} from "./shared.js";
import { inspect } from "./inspect.js";
import { commentXml, insertParagraphAfter, insertedParagraphXml, replaceOrCreateHeaderFooter, setParagraphText } from "./ooxml/docx.js";
import { embedPdfFonts } from "./pdfFonts.js";
import { duplicateSlide, extractShapes, reorderSlides, replaceShapeBulletItems } from "./ooxml/pptx.js";
import { appendRows, insertRows, setCell, sheetPath } from "./ooxml/xlsx.js";
import { escapeXmlText, pxToEmu, replaceAllXmlText, setFirstTextInBlock } from "./ooxml/xml.js";
import { nextRelationshipId } from "./ooxml/relationships.js";
import { PDFDocument, rgb } from "pdf-lib";
import JSZip from "jszip";

export type EditSelector = {
  stableObjectId?: string;
  contains?: string;
  placeholderKey?: string;
  shapeName?: string;
  contentControlTag?: string;
  namedRange?: string;
  textMatch?: { text: string; exact?: boolean };
};

export type EditOperation =
  | { type: "replaceText"; from: string; to: string; selector?: EditSelector }
  | { type: "setText"; text: string; selector: EditSelector }
  | { type: "pdf.textOverlay"; page: number; text: string; x: number; y: number; size?: number; color?: string }
  | { type: "pdf.annotation"; page: number; text: string; x: number; y: number; width?: number; height?: number }
  | { op: "replaceText"; from: string; to: string; selector?: EditSelector }
  | { op: "setText"; text: string; selector: EditSelector }
  | { op: "pptx.duplicateSlide"; slide?: number; after?: number; selector?: EditSelector }
  | { op: "pptx.reorderSlides"; order: number[]; selector?: EditSelector }
  | { op: "pptx.insertBulletItems"; items: string[]; selector: EditSelector }
  | { op: "pptx.replaceBulletItems"; items: string[]; selector: EditSelector }
  | { op: "pptx.replaceImageByShape"; selector: EditSelector; replacementBase64: string; replacementPath?: string; fit?: "contain" | "cover" | "stretch"; crop?: CropRect }
  | { op: "pptx.updateChartData"; selector: EditSelector; categories: string[]; values: number[]; seriesName?: string }
  | { op: "pptx.setBounds"; selector: EditSelector; bounds: { x: number; y: number; width: number; height: number } }
  | { op: "docx.insertParagraphAfter"; text: string; selector: EditSelector }
  | { op: "docx.setHeader"; text: string }
  | { op: "docx.setFooter"; text: string }
  | { op: "docx.setStyle"; styleId: string; font?: string; size?: number; bold?: boolean }
  | { op: "docx.addComment"; text: string; selector: EditSelector; author?: string }
  | { op: "docx.addRedline"; text: string; selector: EditSelector; author?: string }
  | { op: "xlsx.insertRows"; sheet?: number; rowIndex: number; rows: unknown[][]; selector?: EditSelector }
  | { op: "xlsx.appendRows"; sheet?: number; rows: unknown[][]; selector?: EditSelector }
  | { op: "xlsx.setCell"; sheet?: number; cell: string; value: unknown; selector?: EditSelector }
  | { op: "xlsx.setFormula"; sheet?: number; cell: string; formula: string; selector?: EditSelector }
  | { op: "xlsx.updateTable"; sheet?: number; startCell: string; rows: unknown[][]; selector?: EditSelector }
  | { op: "xlsx.writeTable"; sheet?: number; startCell: string; rows: unknown[][]; tableName?: string; selector?: EditSelector }
  | { op: "xlsx.table.resize"; selector: EditSelector; ref: string }
  | { op: "xlsx.chart.setData"; selector: EditSelector; categories: string[]; values: number[]; seriesName?: string }
  | { op: "xlsx.pivot.refreshDefinition"; selector: EditSelector };

interface CropRect {
  left?: number;
  right?: number;
  top?: number;
  bottom?: number;
}

export interface EditOptions {
  out?: string;
  dryRun?: boolean;
  resolveSelectors?: boolean;
  format?: "pptx" | "docx" | "xlsx" | "pdf" | "unknown";
  atomic?: boolean;
  validateFirst?: boolean;
  idempotencyKey?: string;
  continueOnError?: boolean;
  config?: OfficegenConfig;
}

export interface EditSelectorResolution {
  operationIndex: number;
  selector: EditSelector;
  stableObjectId?: string;
  matched: boolean;
  matchCount: number;
  matches: Array<{
    stableObjectId: string;
    kind: string;
    label?: string;
    text?: string;
    sourcePath?: string;
    xmlPath?: string;
  }>;
  reason?: "not-found" | "ambiguous" | "unsupported-selector";
}

export interface ResolveEditSelectorsResult {
  schema: "officegen.edit.selectors@1.2";
  format: string;
  resolutions: EditSelectorResolution[];
  objectMap: ObjectMapEntry[];
  caveats: string[];
}

export interface EditOperationResult {
  operationIndex: number;
  op: string;
  applied: boolean;
  reason?: "not-found" | "ambiguous" | "unsupported" | "validation-failed" | "idempotency-replay" | "skipped-after-error";
  message?: string;
}

export interface EditResult {
  schema: "officegen.edit.result@1.2";
  format: string;
  changed: boolean;
  applied: number;
  skipped: number;
  out?: string;
  bytes?: Uint8Array;
  resolvedSelectors?: EditSelectorResolution[];
  opResults?: EditOperationResult[];
  errors?: EditOperationResult[];
  caveats: string[];
}

export async function edit(input: InputLike, operations: EditOperation[], options: EditOptions = {}): Promise<EditResult> {
  const normalized = await normalizeInput(input, options.format ?? "unknown");
  const selectorResult = options.resolveSelectors || options.validateFirst !== false
    ? await resolveEditSelectorsForNormalized(normalized, operations, options.config)
    : undefined;
  const result = isOfficeFormat(normalized.format)
    ? await editOfficeXml(normalized, operations, options, selectorResult)
    : normalized.format === "pdf"
      ? await editPdf(normalized, operations, options)
      : undefined;
  if (result) {
    if (selectorResult) result.resolvedSelectors = selectorResult.resolutions;
    return result;
  }
  throw new Error(`Unsupported edit format: ${normalized.format}`);
}

export const editDocument = edit;

export async function resolveEditSelectors(
  input: InputLike,
  operations: EditOperation[],
  options: Pick<EditOptions, "format" | "config"> = {}
): Promise<ResolveEditSelectorsResult> {
  const normalized = await normalizeInput(input, options.format ?? "unknown");
  return resolveEditSelectorsForNormalized(normalized, operations, options.config);
}

async function resolveEditSelectorsForNormalized(
  normalized: Awaited<ReturnType<typeof normalizeInput>>,
  operations: EditOperation[],
  config?: OfficegenConfig
): Promise<ResolveEditSelectorsResult> {
  const inspected = await inspect({ data: normalized.bytes, format: normalized.format }, { config });
  const resolutions = operations.flatMap((operation, index) => {
    const selector = selectorForOperation(operation);
    if (!selector) return [];
    const matches = resolveMatches(inspected.objectMap, selector);
    return [
      {
        operationIndex: index,
        selector,
        stableObjectId: selector.stableObjectId,
        matched: matches.length > 0,
        matchCount: matches.length,
        matches: matches.map(selectorMatch),
        reason: matches.length === 0 ? "not-found" : matches.length > 1 ? "ambiguous" : undefined
      } satisfies EditSelectorResolution
    ];
  });
  return {
    schema: "officegen.edit.selectors@1.2",
    format: inspected.trusted.format,
    resolutions,
    objectMap: inspected.objectMap,
    caveats: ["Selector resolution is based on the current inspect objectMap stableObjectId values."]
  };
}

async function editOfficeXml(
  input: Awaited<ReturnType<typeof normalizeInput>>,
  operations: EditOperation[],
  options: EditOptions,
  selectorResult: ResolveEditSelectorsResult | undefined
): Promise<EditResult> {
  const zip = await loadZip(input, { zipSafety: { config: options.config } });
  const atomic = options.atomic ?? true;
  const continueOnError = options.continueOnError ?? false;
  const opResults: EditOperationResult[] = [];
  let applied = 0;
  let skipped = 0;

  if (options.idempotencyKey) {
    const markerPath = idempotencyMarkerPath(options.idempotencyKey);
    if (zip.file(markerPath)) {
      return {
        schema: "officegen.edit.result@1.2",
        format: input.format,
        changed: false,
        applied: 0,
        skipped: operations.length,
        opResults: operations.map((operation, index) => ({
          operationIndex: index,
          op: operationName(operation),
          applied: false,
          reason: "idempotency-replay",
          message: `idempotencyKey already applied: ${options.idempotencyKey}`
        })),
        caveats: ["IDEMPOTENCY_REPLAY: idempotencyKey marker already exists.", ...zipSafetyCaveats(getLoadedZipSafetyReport(zip))]
      };
    }
  }

  const validationErrors = options.validateFirst === false ? [] : validationFailures(selectorResult);
  if (validationErrors.length && atomic) {
    return editAbortResult(input.format, operations.length, validationErrors, [
      "Atomic edit aborted before writing because selector validation failed.",
      ...zipSafetyCaveats(getLoadedZipSafetyReport(zip))
    ]);
  }

  for (const [index, operation] of operations.entries()) {
    if (opResults.some((result) => result.applied === false && result.reason && result.reason !== "unsupported") && !continueOnError) {
      skipped += 1;
      opResults.push({ operationIndex: index, op: operationName(operation), applied: false, reason: "skipped-after-error" });
      continue;
    }
    const validation = validationErrors.find((failure) => failure.operationIndex === index);
    if (validation) {
      skipped += 1;
      opResults.push(validation);
      continue;
    }
    try {
      const changed = await applyOfficeOperation(zip, input.format as "pptx" | "docx" | "xlsx", operation, selectorResult?.objectMap ?? [], index);
      if (changed) {
        applied += 1;
        opResults.push({ operationIndex: index, op: operationName(operation), applied: true });
      } else {
        skipped += 1;
        opResults.push({ operationIndex: index, op: operationName(operation), applied: false, reason: "not-found" });
      }
    } catch (error) {
      skipped += 1;
      opResults.push({
        operationIndex: index,
        op: operationName(operation),
        applied: false,
        reason: classifyEditError(error),
        message: error instanceof Error ? error.message : String(error)
      });
      if (!continueOnError && atomic) break;
    }
  }

  const errors = opResults.filter((result) => !result.applied && result.reason && result.reason !== "unsupported");
  if (errors.length && atomic) {
    return editAbortResult(input.format, skipped, opResults, [
      "Atomic edit aborted; no output bytes were written.",
      ...zipSafetyCaveats(getLoadedZipSafetyReport(zip))
    ]);
  }

  if (options.idempotencyKey && applied > 0) zip.file(idempotencyMarkerPath(options.idempotencyKey), new Date().toISOString());
  const bytes = options.dryRun ? undefined : await zipToBytes(zip);
  if (!options.dryRun) await writeOutput(options.out, bytes as Uint8Array);
  return {
    schema: "officegen.edit.result@1.2",
    format: input.format,
    changed: applied > 0,
    applied,
    skipped,
    out: options.dryRun ? undefined : options.out,
    bytes: options.dryRun || options.out ? undefined : bytes,
    opResults,
    errors: errors.length ? errors : undefined,
    caveats: [
      "Office XML edits preserve unknown parts but do not recalculate native layout, formulas, or theme-derived rendering.",
      ...zipSafetyCaveats(getLoadedZipSafetyReport(zip))
    ]
  };
}

async function applyOfficeOperation(
  zip: Awaited<ReturnType<typeof loadZip>>,
  format: "pptx" | "docx" | "xlsx",
  operation: EditOperation,
  objectMap: ObjectMapEntry[],
  index: number
): Promise<boolean> {
  const op = operationName(operation);
  if (op === "replaceText") return replaceTextInEditableParts(zip, format, (operation as { from: string; to: string }).from, (operation as { from: string; to: string }).to);
  if (op === "setText") return setSelectedText(zip, format, operation as { selector: EditSelector; text: string }, objectMap);
  if (format === "pptx" && op === "pptx.duplicateSlide") {
    const duplicate = operation as { slide?: number; after?: number; selector?: EditSelector };
    await duplicateSlide(zip, duplicate.slide ?? slideNumberFromSelector(duplicate.selector, objectMap) ?? 1, duplicate.after);
    return true;
  }
  if (format === "pptx" && op === "pptx.reorderSlides") {
    await reorderSlides(zip, (operation as { order: number[] }).order);
    return true;
  }
  if (format === "pptx" && (op === "pptx.insertBulletItems" || op === "pptx.replaceBulletItems")) {
    return editPptxBullets(zip, operation as { selector: EditSelector; items: string[] }, objectMap, op === "pptx.insertBulletItems" ? "insert" : "replace");
  }
  if (format === "pptx" && op === "pptx.replaceImageByShape") {
    return editPptxReplaceImageByShape(zip, operation as { selector: EditSelector; replacementBase64: string; replacementPath?: string; fit?: "contain" | "cover" | "stretch"; crop?: CropRect }, objectMap);
  }
  if (format === "pptx" && op === "pptx.updateChartData") {
    return editPptxUpdateChartData(zip, operation as { selector: EditSelector; categories: string[]; values: number[]; seriesName?: string }, objectMap);
  }
  if (format === "pptx" && op === "pptx.setBounds") {
    return editPptxSetBounds(zip, operation as { selector: EditSelector; bounds: { x: number; y: number; width: number; height: number } }, objectMap);
  }
  if (format === "docx" && op === "docx.insertParagraphAfter") {
    return editDocxInsertParagraph(zip, operation as { selector: EditSelector; text: string }, objectMap);
  }
  if (format === "docx" && (op === "docx.setHeader" || op === "docx.setFooter")) {
    return editDocxHeaderFooter(zip, op === "docx.setHeader" ? "header" : "footer", (operation as { text: string }).text);
  }
  if (format === "docx" && op === "docx.setStyle") {
    return editDocxStyle(zip, operation as { styleId: string; font?: string; size?: number; bold?: boolean });
  }
  if (format === "docx" && op === "docx.addComment") {
    return editDocxAddComment(zip, operation as { selector: EditSelector; text: string; author?: string }, objectMap);
  }
  if (format === "docx" && op === "docx.addRedline") {
    return editDocxAddRedline(zip, operation as { selector: EditSelector; text: string; author?: string }, objectMap);
  }
  if (format === "xlsx" && op === "xlsx.insertRows") {
    const rowOp = operation as { sheet?: number; rowIndex: number; rows: unknown[][] };
    const path = sheetPath(rowOp.sheet);
    const xml = (await readZipText(zip, path)) ?? "";
    const next = insertRows(xml, rowOp.rowIndex, rowOp.rows);
    if (next.changed) zip.file(path, next.xml);
    return next.changed;
  }
  if (format === "xlsx" && op === "xlsx.appendRows") {
    const rowOp = operation as { sheet?: number; rows: unknown[][] };
    const path = sheetPath(rowOp.sheet);
    const xml = (await readZipText(zip, path)) ?? "";
    const next = appendRows(xml, rowOp.rows);
    if (next.changed) zip.file(path, next.xml);
    return next.changed;
  }
  if (format === "xlsx" && op === "xlsx.setCell") {
    const cellOp = operation as { sheet?: number; cell: string; value: unknown };
    return editXlsxSetCell(zip, cellOp.sheet, cellOp.cell, cellOp.value);
  }
  if (format === "xlsx" && op === "xlsx.setFormula") {
    const formulaOp = operation as { sheet?: number; cell: string; formula: string };
    return editXlsxSetFormula(zip, formulaOp.sheet, formulaOp.cell, formulaOp.formula);
  }
  if (format === "xlsx" && (op === "xlsx.updateTable" || op === "xlsx.writeTable")) {
    const tableOp = operation as { sheet?: number; startCell: string; rows: unknown[][]; tableName?: string };
    let changed = false;
    const start = /^([A-Z]+)(\d+)$/i.exec(tableOp.startCell);
    if (!start) throw new Error(`SELECTOR_NOT_FOUND: invalid startCell for operation ${index}.`);
    const startCol = columnIndex(start[1] ?? "A");
    const startRow = Number(start[2]);
    for (const [r, row] of tableOp.rows.entries()) {
      for (const [c, value] of row.entries()) {
        changed = (await editXlsxSetCell(zip, tableOp.sheet, `${columnName(startCol + c)}${startRow + r}`, value)) || changed;
      }
    }
    changed = (await ensureXlsxTable(zip, tableOp.sheet, tableOp.startCell, tableOp.rows, tableOp.tableName)) || changed;
    return changed;
  }
  if (format === "xlsx" && op === "xlsx.table.resize") {
    const tableOp = operation as { selector: EditSelector; ref: string };
    const target = singleMatch(objectMap, tableOp.selector);
    if (target.kind !== "table" || !target.xmlPath) throw new Error("SELECTOR_NOT_FOUND: selected object is not an XLSX table.");
    const xml = (await readZipText(zip, target.xmlPath)) ?? "";
    const ref = escapeXmlText(tableOp.ref);
    let next = xml.replace(/(<table\b[^>]*\bref=")[^"]*(")/, `$1${ref}$2`);
    next = next.replace(/(<autoFilter\b[^>]*\bref=")[^"]*(")/, `$1${ref}$2`);
    if (next !== xml) zip.file(target.xmlPath, next);
    return next !== xml;
  }
  if (format === "xlsx" && op === "xlsx.chart.setData") {
    const chartOp = operation as { selector: EditSelector; categories: string[]; values: number[]; seriesName?: string };
    const target = singleMatch(objectMap, chartOp.selector);
    if (target.kind !== "chart" || !target.xmlPath) throw new Error("SELECTOR_NOT_FOUND: selected object is not an XLSX chart.");
    const points = chartOp.categories.map((category, pointIndex) => ({ category, value: Number(chartOp.values[pointIndex] ?? 0) }));
    const xml = (await readZipText(zip, target.xmlPath)) ?? "";
    assertSingleSeriesChart(xml, "xlsx.chart.setData");
    const next = replaceChartCaches(xml, chartOp.seriesName ?? target.label ?? "Series 1", points);
    if (next !== xml) zip.file(target.xmlPath, next);
    const workbookChanged = await updateXlsxChartBackingRanges(zip, xml, chartOp.seriesName ?? target.label ?? "Series 1", points);
    return next !== xml || workbookChanged;
  }
  if (format === "xlsx" && op === "xlsx.pivot.refreshDefinition") {
    const pivotOp = operation as { selector: EditSelector };
    const target = singleMatch(objectMap, pivotOp.selector);
    if (target.kind !== "pivotTable" || !target.xmlPath) throw new Error("SELECTOR_NOT_FOUND: selected object is not an XLSX pivotTable.");
    const xml = (await readZipText(zip, target.xmlPath)) ?? "";
    const next = xml.replace(/<pivotTableDefinition\b([^>]*)>/, (match, attrs: string) =>
      /\brefreshOnLoad=/.test(attrs) ? match.replace(/\brefreshOnLoad="[^"]*"/, 'refreshOnLoad="1"') : `<pivotTableDefinition${attrs} refreshOnLoad="1">`
    );
    if (next !== xml) zip.file(target.xmlPath, next);
    return next !== xml;
  }
  return false;
}

async function replaceTextInEditableParts(zip: Awaited<ReturnType<typeof loadZip>>, format: "pptx" | "docx" | "xlsx", from: string, to: string): Promise<boolean> {
  if (!from) return false;
  const paths = Object.keys(zip.files)
    .filter((path) => !zip.files[path]?.dir)
    .filter((path) =>
      format === "pptx"
        ? /^ppt\/slides\/slide\d+\.xml$/i.test(path)
        : format === "docx"
          ? /^word\/(document|header\d+|footer\d+)\.xml$/i.test(path)
          : /^xl\/(worksheets\/sheet\d+|sharedStrings)\.xml$/i.test(path)
    );
  let changed = false;
  for (const path of paths) {
    const xml = (await readZipText(zip, path)) ?? "";
    const next = replaceAllXmlText(xml, from, to);
    if (next !== xml) {
      zip.file(path, next);
      changed = true;
    }
  }
  return changed;
}

async function setSelectedText(zip: Awaited<ReturnType<typeof loadZip>>, format: "pptx" | "docx" | "xlsx", operation: { selector: EditSelector; text: string }, objectMap: ObjectMapEntry[]): Promise<boolean> {
  const target = singleMatch(objectMap, operation.selector);
  if (!target?.sourcePath) throw new Error("SELECTOR_NOT_FOUND: selected object has no sourcePath.");
  const xml = (await readZipText(zip, target.sourcePath)) ?? "";
  if (format === "pptx") {
    const shapes = extractShapes(xml, Number(target.selectorHints?.slide ?? 1), "", target.sourcePath);
    const ordinal = shapes.findIndex((shape) => shape.stableObjectId === target.stableObjectId) + 1;
    if (!ordinal) throw new Error(`SELECTOR_NOT_FOUND: ${target.stableObjectId}`);
    const next = replaceShapeText(xml, ordinal, operation.text);
    if (next.changed) zip.file(target.sourcePath, next.xml);
    return next.changed;
  }
  if (format === "docx") {
    const ordinal = Number(target.selectorHints?.paragraph ?? stableOrdinal(target.stableObjectId));
    const next = setParagraphText(xml, ordinal, operation.text);
    if (next.changed) zip.file(target.sourcePath, next.xml);
    return next.changed;
  }
  if (format === "xlsx") {
    return editXlsxSetCell(zip, Number(target.selectorHints?.sheet ?? 1), String(target.label ?? ""), operation.text);
  }
  return false;
}

async function editPptxBullets(zip: Awaited<ReturnType<typeof loadZip>>, operation: { selector: EditSelector; items: string[] }, objectMap: ObjectMapEntry[], mode: "insert" | "replace"): Promise<boolean> {
  const target = singleMatch(objectMap, operation.selector);
  if (!target?.sourcePath) throw new Error("SELECTOR_NOT_FOUND: selected PPTX shape has no sourcePath.");
  const xml = (await readZipText(zip, target.sourcePath)) ?? "";
  const shapes = extractShapes(xml, Number(target.selectorHints?.slide ?? 1), "", target.sourcePath);
  const ordinal = shapes.findIndex((shape) => shape.stableObjectId === target.stableObjectId) + 1;
  if (!ordinal) throw new Error(`SELECTOR_NOT_FOUND: ${target.stableObjectId}`);
  const next = replaceShapeBulletItems(xml, ordinal, operation.items, mode);
  if (next.changed) zip.file(target.sourcePath, next.xml);
  return next.changed;
}

async function editPptxReplaceImageByShape(
  zip: Awaited<ReturnType<typeof loadZip>>,
  operation: { selector: EditSelector; replacementBase64: string; replacementPath?: string; fit?: "contain" | "cover" | "stretch"; crop?: CropRect },
  objectMap: ObjectMapEntry[]
): Promise<boolean> {
  const target = singleMatch(objectMap, operation.selector);
  const assetPath = String(target.media?.assetPath ?? target.selectorHints?.assetPath ?? "");
  if (target.kind !== "picture" || !target.sourcePath || !assetPath) {
    throw new Error("SELECTOR_NOT_FOUND: selected object is not a PPTX picture with an asset relationship.");
  }
  const replacement = Buffer.from(operation.replacementBase64, "base64");
  const existing = (await readZipBytes(zip, assetPath)) ?? new Uint8Array();
  const currentMediaType = detectMediaType(existing, assetPath);
  const replacementMediaType = detectMediaType(replacement, operation.replacementPath ?? assetPath);
  const expectedMediaType = mediaTypeFromExtension(assetPath);
  if (expectedMediaType && replacementMediaType !== expectedMediaType) {
    throw new Error(`ASSET_UNSUPPORTED_FORMAT: replacement media type ${replacementMediaType} does not match ${assetPath} (${expectedMediaType}).`);
  }
  if (currentMediaType !== "application/octet-stream" && replacementMediaType !== currentMediaType) {
    throw new Error(`ASSET_UNSUPPORTED_FORMAT: replacement media type ${replacementMediaType} does not match existing asset type ${currentMediaType}.`);
  }
  zip.file(assetPath, replacement);
  const slideXml = (await readZipText(zip, target.sourcePath)) ?? "";
  const pictureIndex = Number(target.selectorHints?.pictureIndex ?? stableOrdinal(target.stableObjectId));
  const nextXml = updatePictureCrop(slideXml, target, pictureIndex, operation.fit, operation.crop, replacement, replacementMediaType);
  if (nextXml !== slideXml) zip.file(target.sourcePath, nextXml);
  return true;
}

async function editPptxUpdateChartData(
  zip: Awaited<ReturnType<typeof loadZip>>,
  operation: { selector: EditSelector; categories: string[]; values: number[]; seriesName?: string },
  objectMap: ObjectMapEntry[]
): Promise<boolean> {
  const target = singleMatch(objectMap, operation.selector);
  const chartPath = String(target.media?.chartPath ?? target.selectorHints?.chartPath ?? target.xmlPath ?? "");
  if (target.kind !== "chart" || !chartPath) throw new Error("SELECTOR_NOT_FOUND: selected object is not a PPTX chart.");
  if (!operation.categories.length || operation.categories.length !== operation.values.length) {
    throw new Error("SCHEMA_INVALID: pptx.updateChartData requires categories and values arrays with equal length.");
  }
  const xml = (await readZipText(zip, chartPath)) ?? "";
  assertSingleSeriesChart(xml, "pptx.updateChartData");
  const points = operation.categories.map((category, index) => ({
    category,
    value: Number(operation.values[index] ?? 0)
  }));
  const next = replaceChartCaches(xml, operation.seriesName ?? "Series 1", points);
  if (next !== xml) zip.file(chartPath, next);
  const workbookChanged = await updateEmbeddedChartWorkbook(zip, chartPath, operation.seriesName ?? "Series 1", points);
  return next !== xml || workbookChanged;
}

async function editPptxSetBounds(
  zip: Awaited<ReturnType<typeof loadZip>>,
  operation: { selector: EditSelector; bounds: { x: number; y: number; width: number; height: number } },
  objectMap: ObjectMapEntry[]
): Promise<boolean> {
  const target = singleMatch(objectMap, operation.selector);
  if (!target.sourcePath || !["shape", "picture", "chart"].includes(target.kind)) {
    throw new Error("SELECTOR_NOT_FOUND: selected object is not a PPTX shape, picture, or chart.");
  }
  const xml = (await readZipText(zip, target.sourcePath)) ?? "";
  const shapeId = String(target.selectorHints?.shapeId ?? "");
  const next = replacePptxObjectBlock(xml, target.kind, shapeId, (block) => setBlockBounds(block, operation.bounds));
  if (next !== xml) zip.file(target.sourcePath, next);
  return next !== xml;
}

async function editDocxInsertParagraph(zip: Awaited<ReturnType<typeof loadZip>>, operation: { selector: EditSelector; text: string }, objectMap: ObjectMapEntry[]): Promise<boolean> {
  const target = singleMatch(objectMap, operation.selector);
  if (!target?.sourcePath) throw new Error("SELECTOR_NOT_FOUND: selected DOCX paragraph has no sourcePath.");
  const xml = (await readZipText(zip, target.sourcePath)) ?? "";
  const ordinal = Number(target.selectorHints?.paragraph ?? stableOrdinal(target.stableObjectId));
  const next = insertParagraphAfter(xml, ordinal, operation.text);
  if (next.changed) zip.file(target.sourcePath, next.xml);
  return next.changed;
}

async function editDocxHeaderFooter(zip: Awaited<ReturnType<typeof loadZip>>, kind: "header" | "footer", text: string): Promise<boolean> {
  const partPath = `word/${kind}1.xml`;
  const relType = `http://schemas.openxmlformats.org/officeDocument/2006/relationships/${kind}`;
  const contentType = kind === "header"
    ? "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
    : "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";
  const xml = await readZipText(zip, partPath);
  zip.file(partPath, replaceOrCreateHeaderFooter(xml, kind, text));
  await ensureContentTypeOverride(zip, `/${partPath}`, contentType);
  const relId = await ensureDocumentRelationship(zip, relType, `${kind}1.xml`);
  const documentXml = (await readZipText(zip, "word/document.xml")) ?? "";
  const nextDocumentXml = ensureSectionReference(documentXml, kind, relId);
  if (nextDocumentXml !== documentXml) zip.file("word/document.xml", nextDocumentXml);
  return true;
}

async function editDocxAddComment(zip: Awaited<ReturnType<typeof loadZip>>, operation: { selector: EditSelector; text: string; author?: string }, objectMap: ObjectMapEntry[]): Promise<boolean> {
  const target = singleMatch(objectMap, operation.selector);
  if (!target.sourcePath) throw new Error("SELECTOR_NOT_FOUND: selected DOCX paragraph has no sourcePath.");
  await ensureContentTypeOverride(zip, "/word/comments.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml");
  await ensureDocumentRelationship(zip, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments", "comments.xml");
  const commentsXml = (await readZipText(zip, "word/comments.xml")) ?? '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:comments>';
  const nextId = Math.max(-1, ...[...commentsXml.matchAll(/\bw:id="(\d+)"/g)].map((match) => Number(match[1])).filter(Number.isFinite)) + 1;
  const comment = commentXml(nextId, operation.author ?? "officegen", operation.text);
  const nextCommentsXml = /<\/w:comments>\s*$/.test(commentsXml)
    ? commentsXml.replace(/<\/w:comments>\s*$/, `${comment}</w:comments>`)
    : `<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${comment}</w:comments>`;
  zip.file("word/comments.xml", nextCommentsXml);
  const xml = (await readZipText(zip, target.sourcePath)) ?? "";
  const ordinal = Number(target.selectorHints?.paragraph ?? stableOrdinal(target.stableObjectId));
  const next = replaceNthParagraph(xml, ordinal, (paragraph) => {
    const withStart = paragraph.replace(/(<w:p\b[^>]*>)/, `$1<w:commentRangeStart w:id="${nextId}"/>`);
    return withStart.replace(/<\/w:p>$/, `<w:commentRangeEnd w:id="${nextId}"/><w:r><w:commentReference w:id="${nextId}"/></w:r></w:p>`);
  });
  if (next.changed) zip.file(target.sourcePath, next.xml);
  return next.changed;
}

async function editDocxAddRedline(zip: Awaited<ReturnType<typeof loadZip>>, operation: { selector: EditSelector; text: string; author?: string }, objectMap: ObjectMapEntry[]): Promise<boolean> {
  const target = singleMatch(objectMap, operation.selector);
  if (!target.sourcePath) throw new Error("SELECTOR_NOT_FOUND: selected DOCX paragraph has no sourcePath.");
  const xml = (await readZipText(zip, target.sourcePath)) ?? "";
  const ordinal = Number(target.selectorHints?.paragraph ?? stableOrdinal(target.stableObjectId));
  const nextId = nextDocxRevisionId(xml);
  const next = replaceNthParagraph(xml, ordinal, (paragraph) => `${paragraph}${insertedParagraphXml(operation.text, operation.author ?? "officegen", new Date(), nextId)}`);
  if (next.changed) zip.file(target.sourcePath, next.xml);
  return next.changed;
}

async function editDocxStyle(zip: Awaited<ReturnType<typeof loadZip>>, operation: { styleId: string; font?: string; size?: number; bold?: boolean }): Promise<boolean> {
  const path = "word/styles.xml";
  const styles = (await readZipText(zip, path)) ?? '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:styles>';
  const styleXml = buildDocxStyleXml(operation);
  const re = new RegExp(`<w:style\\b[^>]*\\bw:styleId="${escapeRegExp(operation.styleId)}"[\\s\\S]*?<\\/w:style>`);
  const next = re.test(styles)
    ? styles.replace(re, styleXml)
    : styles.replace(/<\/w:styles>\s*$/, `${styleXml}</w:styles>`);
  if (next !== styles) zip.file(path, next);
  await ensureContentTypeOverride(zip, "/word/styles.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml");
  return next !== styles;
}

function buildDocxStyleXml(operation: { styleId: string; font?: string; size?: number; bold?: boolean }): string {
  const styleId = escapeXmlText(operation.styleId);
  const runProps = [
    operation.font ? `<w:rFonts w:ascii="${escapeXmlText(operation.font)}" w:hAnsi="${escapeXmlText(operation.font)}" w:eastAsia="${escapeXmlText(operation.font)}"/>` : "",
    typeof operation.size === "number" ? `<w:sz w:val="${Math.round(operation.size * 2)}"/>` : "",
    operation.bold ? "<w:b/>" : ""
  ].join("");
  return `<w:style w:type="paragraph" w:styleId="${styleId}"><w:name w:val="${styleId}"/>${runProps ? `<w:rPr>${runProps}</w:rPr>` : ""}</w:style>`;
}

async function editXlsxSetCell(zip: Awaited<ReturnType<typeof loadZip>>, sheet: number | undefined, ref: string, value: unknown): Promise<boolean> {
  if (!ref) throw new Error("SELECTOR_NOT_FOUND: xlsx cell ref is required.");
  const path = sheetPath(sheet);
  const xml = (await readZipText(zip, path)) ?? "";
  const next = setCell(xml, ref, value);
  if (next.changed) zip.file(path, next.xml);
  return next.changed;
}

async function editXlsxSetFormula(zip: Awaited<ReturnType<typeof loadZip>>, sheet: number | undefined, ref: string, formula: string): Promise<boolean> {
  if (!ref) throw new Error("SELECTOR_NOT_FOUND: xlsx cell ref is required.");
  const path = sheetPath(sheet);
  const xml = (await readZipText(zip, path)) ?? "";
  const cellXml = `<c r="${escapeXmlText(ref)}"><f>${escapeXmlText(formula.replace(/^=/, ""))}</f></c>`;
  const pattern = new RegExp(`<c\\b[^>]*\\br=["']${escapeRegExp(ref)}["'][^>]*?(?:\\/>|>[\\s\\S]*?<\\/c>)`);
  const rowNo = rowFromRef(ref);
  const rowPattern = new RegExp(`<row\\b([^>]*)\\br=["']${rowNo}["'][^>]*>[\\s\\S]*?<\\/row>`);
  const next = pattern.test(xml)
    ? xml.replace(pattern, cellXml)
    : rowPattern.test(xml)
      ? xml.replace(rowPattern, (row) => row.replace(/<\/row>$/, `${cellXml}</row>`))
      : xml.replace(/<\/sheetData>/, `<row r="${rowNo}">${cellXml}</row></sheetData>`);
  if (next !== xml) zip.file(path, next);
  return next !== xml;
}

function updatePictureCrop(
  slideXml: string,
  target: ObjectMapEntry,
  pictureIndex: number,
  fit: "contain" | "cover" | "stretch" | undefined,
  crop: CropRect | undefined,
  replacement: Uint8Array,
  mediaType: string
): string {
  const shapeId = String(target.selectorHints?.shapeId ?? "");
  let index = 0;
  return slideXml.replace(/<p:pic\b[\s\S]*?<\/p:pic>/g, (picture) => {
    index += 1;
    const cNvPr = /<p:cNvPr\b([^>]*)\/>/.exec(picture)?.[1] ?? "";
    const candidateId = /\bid="([^"]+)"/.exec(cNvPr)?.[1] ?? "";
    if ((shapeId && candidateId !== shapeId) || (!shapeId && index !== pictureIndex)) return picture;
    const rect = crop ?? cropForFit(target, replacement, mediaType, fit);
    if (!rect || fit === "contain" || fit === "stretch") return picture.replace(/<a:srcRect\b[^>]*\/>/g, "");
    const srcRect = `<a:srcRect${cropAttr("l", rect.left)}${cropAttr("r", rect.right)}${cropAttr("t", rect.top)}${cropAttr("b", rect.bottom)}/>`;
    if (/<a:srcRect\b[^>]*\/>/.test(picture)) return picture.replace(/<a:srcRect\b[^>]*\/>/, srcRect);
    return picture.replace(/(<a:blip\b[^>]*\/>|<a:blip\b[\s\S]*?<\/a:blip>)/, `$1${srcRect}`);
  });
}

function replacePptxObjectBlock(xml: string, kind: string, shapeId: string, updater: (block: string) => string): string {
  const pattern = kind === "picture"
    ? /<p:pic\b[\s\S]*?<\/p:pic>/g
    : kind === "chart"
      ? /<p:graphicFrame\b[\s\S]*?<\/p:graphicFrame>/g
      : /<p:sp\b[\s\S]*?<\/p:sp>/g;
  let index = 0;
  return xml.replace(pattern, (block) => {
    index += 1;
    const cNvPr = /<p:cNvPr\b([^>]*)\/>/.exec(block)?.[1] ?? "";
    const candidateId = /\bid="([^"]+)"/.exec(cNvPr)?.[1];
    if ((shapeId && candidateId !== shapeId) || (!shapeId && index !== 1)) return block;
    return updater(block);
  });
}

function setBlockBounds(block: string, bounds: { x: number; y: number; width: number; height: number }): string {
  const off = `<a:off x="${pxToEmu(bounds.x)}" y="${pxToEmu(bounds.y)}"/>`;
  const ext = `<a:ext cx="${pxToEmu(bounds.width)}" cy="${pxToEmu(bounds.height)}"/>`;
  if (/<p:xfrm\b[\s\S]*?<\/p:xfrm>/.test(block)) {
    return block.replace(/<p:xfrm\b([^>]*)>[\s\S]*?<\/p:xfrm>/, `<p:xfrm$1>${off}${ext}</p:xfrm>`);
  }
  if (/<a:xfrm\b[\s\S]*?<\/a:xfrm>/.test(block)) {
    return block.replace(/<a:xfrm\b([^>]*)>[\s\S]*?<\/a:xfrm>/, `<a:xfrm$1>${off}${ext}</a:xfrm>`);
  }
  if (/<p:graphicFrame\b/.test(block)) {
    return block.replace(/(<p:nvGraphicFramePr\b[\s\S]*?<\/p:nvGraphicFramePr>)/, `$1<p:xfrm>${off}${ext}</p:xfrm>`);
  }
  return block.replace(/(<p:spPr\b[^>]*>)/, `$1<a:xfrm>${off}${ext}</a:xfrm>`);
}

function cropForFit(target: ObjectMapEntry, replacement: Uint8Array, mediaType: string, fit: "contain" | "cover" | "stretch" | undefined): CropRect | undefined {
  if (fit !== "cover") return undefined;
  const dimensions = detectDimensions(replacement, mediaType);
  const bounds = target.bounds;
  if (!dimensions.width || !dimensions.height || !bounds?.width || !bounds.height) return undefined;
  const imageAspect = dimensions.width / dimensions.height;
  const boxAspect = bounds.width / bounds.height;
  if (!Number.isFinite(imageAspect) || !Number.isFinite(boxAspect) || imageAspect <= 0 || boxAspect <= 0) return undefined;
  if (imageAspect > boxAspect) {
    const keep = boxAspect / imageAspect;
    const side = (1 - keep) / 2;
    return { left: side, right: side };
  }
  const keep = imageAspect / boxAspect;
  const side = (1 - keep) / 2;
  return { top: side, bottom: side };
}

function cropAttr(name: string, value: number | undefined): string {
  if (value === undefined) return "";
  const normalized = Math.max(0, Math.min(1, value));
  return ` ${name}="${Math.round(normalized * 100000)}"`;
}

function replaceChartCaches(xml: string, seriesName: string, points: Array<{ category: string; value: number }>): string {
  const pointCount = points.length;
  const multiLevelStringCache = `<c:multiLvlStrCache><c:ptCount val="${pointCount}"/><c:lvl>${points.map((point, index) => `<c:pt idx="${index}"><c:v>${escapeXmlText(point.category)}</c:v></c:pt>`).join("")}</c:lvl></c:multiLvlStrCache>`;
  const stringCache = `<c:strCache><c:ptCount val="${pointCount}"/>${points.map((point, index) => `<c:pt idx="${index}"><c:v>${escapeXmlText(point.category)}</c:v></c:pt>`).join("")}</c:strCache>`;
  const numberCache = `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${pointCount}"/>${points.map((point, index) => `<c:pt idx="${index}"><c:v>${Number.isFinite(point.value) ? point.value : 0}</c:v></c:pt>`).join("")}</c:numCache>`;
  let next = xml.replace(/<c:tx>\s*<c:strRef>[\s\S]*?<\/c:strRef>\s*<\/c:tx>|<c:tx>\s*<c:v>[\s\S]*?<\/c:v>\s*<\/c:tx>/, `<c:tx><c:v>${escapeXmlText(seriesName)}</c:v></c:tx>`);
  next = next.replace(/<c:cat>[\s\S]*?<\/c:cat>/, (cat) => {
    let updated = cat.replace(/<c:f>([^<]*)<\/c:f>/, (_match, formula: string) => `<c:f>${escapeXmlText(updateChartRangeFormula(formula, pointCount, "A"))}</c:f>`);
    if (/<c:multiLvlStrCache>[\s\S]*?<\/c:multiLvlStrCache>/.test(updated)) {
      updated = updated.replace(/<c:multiLvlStrCache>[\s\S]*?<\/c:multiLvlStrCache>/, multiLevelStringCache);
    } else if (/<c:strCache>[\s\S]*?<\/c:strCache>/.test(updated)) {
      updated = updated.replace(/<c:strCache>[\s\S]*?<\/c:strCache>/, stringCache);
    }
    return updated;
  });
  next = next.replace(/<c:val>[\s\S]*?<\/c:val>/, (val) => val
    .replace(/<c:f>([^<]*)<\/c:f>/, (_match, formula: string) => `<c:f>${escapeXmlText(updateChartRangeFormula(formula, pointCount, "B"))}</c:f>`)
    .replace(/<c:numCache>[\s\S]*?<\/c:numCache>/, numberCache));
  return next;
}

function assertSingleSeriesChart(xml: string, operation: string): void {
  const seriesCount = (xml.match(/<c:ser\b/g) ?? []).length;
  if (seriesCount > 1) {
    throw new Error(`SCHEMA_INVALID: ${operation} currently supports single-series charts only; refusing partial multi-series update.`);
  }
}

async function updateXlsxChartBackingRanges(
  zip: Awaited<ReturnType<typeof loadZip>>,
  chartXml: string,
  seriesName: string,
  points: Array<{ category: string; value: number }>
): Promise<boolean> {
  const catFormula = /<c:cat>[\s\S]*?<c:f>([^<]+)<\/c:f>[\s\S]*?<\/c:cat>/.exec(chartXml)?.[1];
  const valFormula = /<c:val>[\s\S]*?<c:f>([^<]+)<\/c:f>[\s\S]*?<\/c:val>/.exec(chartXml)?.[1];
  const categoryRange = catFormula ? parseA1RangeFormula(catFormula) : undefined;
  const valueRange = valFormula ? parseA1RangeFormula(valFormula) : undefined;
  if (!categoryRange || !valueRange || categoryRange.sheet !== valueRange.sheet) {
    throw new Error("SCHEMA_INVALID: xlsx.chart.setData cannot resolve chart backing worksheet ranges.");
  }
  const sheetNumber = worksheetNumberFromName(categoryRange.sheet);
  const path = sheetPath(sheetNumber);
  const xml = await readZipText(zip, path);
  if (!xml) throw new Error(`SELECTOR_NOT_FOUND: chart backing worksheet ${categoryRange.sheet} was not found.`);
  let next = xml;
  const valueHeaderRow = Math.max(1, valueRange.startRow - 1);
  next = setCell(next, `${valueRange.startCol}${valueHeaderRow}`, seriesName).xml;
  for (const [index, point] of points.entries()) {
    const categoryRow = categoryRange.startRow + index;
    const valueRow = valueRange.startRow + index;
    next = setCell(next, `${categoryRange.startCol}${categoryRow}`, point.category).xml;
    next = setCell(next, `${valueRange.startCol}${valueRow}`, Number.isFinite(point.value) ? point.value : 0).xml;
  }
  if (next !== xml) zip.file(path, next);
  return next !== xml;
}

function parseA1RangeFormula(formula: string): { sheet: string; startCol: string; startRow: number; endCol: string; endRow: number } | undefined {
  const normalized = formula.replace(/&apos;/g, "'").trim();
  const match = /^(?:'([^']+)'|([^!]+))!\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$/i.exec(normalized);
  if (!match) return undefined;
  return {
    sheet: match[1] ?? match[2] ?? "Sheet1",
    startCol: (match[3] ?? "A").toUpperCase(),
    startRow: Number(match[4] ?? 1),
    endCol: (match[5] ?? "A").toUpperCase(),
    endRow: Number(match[6] ?? 1)
  };
}

function updateChartRangeFormula(formula: string, pointCount: number, fallbackColumn: string): string {
  const range = parseA1RangeFormula(formula);
  if (!range) return `Sheet1!$${fallbackColumn}$2:$${fallbackColumn}$${pointCount + 1}`;
  return `${formatSheetName(range.sheet)}!$${range.startCol}$${range.startRow}:$${range.endCol}$${range.startRow + pointCount - 1}`;
}

function formatSheetName(name: string): string {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`;
}

function worksheetNumberFromName(name: string): number {
  const match = /(\d+)$/.exec(name.trim());
  return match ? Number(match[1]) : 1;
}

function nextDocxRevisionId(xml: string): number {
  const ids = [...xml.matchAll(/<w:(?:ins|del)\b[^>]*\bw:id="(\d+)"/g)]
    .map((match) => Number(match[1]))
    .filter(Number.isFinite);
  return Math.max(0, ...ids) + 1;
}

async function updateEmbeddedChartWorkbook(
  zip: Awaited<ReturnType<typeof loadZip>>,
  chartPath: string,
  seriesName: string,
  points: Array<{ category: string; value: number }>
): Promise<boolean> {
  const relsPath = chartPath.replace(/^ppt\/charts\//, "ppt/charts/_rels/") + ".rels";
  const relsXml = (await readZipText(zip, relsPath)) ?? "";
  const packageRel = /<Relationship\b[^>]*\bType="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/package"[^>]*\bTarget="([^"]+)"/.exec(relsXml)?.[1];
  if (!packageRel) return false;
  const workbookPath = normalizePackageTarget("ppt/charts", packageRel);
  const workbookBytes = await readZipBytes(zip, workbookPath);
  if (!workbookBytes) return false;
  const workbookZip = await JSZip.loadAsync(workbookBytes);
  const sheet = (await readZipText(workbookZip, "xl/worksheets/sheet1.xml")) ?? "";
  let nextSheet = setCell(sheet, "B1", seriesName).xml;
  for (const [index, point] of points.entries()) {
    const row = index + 2;
    nextSheet = setCell(nextSheet, `A${row}`, point.category).xml;
    nextSheet = setCell(nextSheet, `B${row}`, Number.isFinite(point.value) ? point.value : 0).xml;
  }
  workbookZip.file("xl/worksheets/sheet1.xml", nextSheet);
  zip.file(workbookPath, await workbookZip.generateAsync({ type: "uint8array", compression: "DEFLATE", compressionOptions: { level: 6 } }));
  return true;
}

async function ensureXlsxTable(
  zip: Awaited<ReturnType<typeof loadZip>>,
  sheet: number | undefined,
  startCell: string,
  rows: unknown[][],
  tableName?: string
): Promise<boolean> {
  if (!rows.length || !rows[0]?.length) return false;
  const sheetNo = sheet && sheet > 0 ? sheet : 1;
  const worksheetPath = sheetPath(sheetNo);
  const worksheetXml = (await readZipText(zip, worksheetPath)) ?? "";
  const start = /^([A-Z]+)(\d+)$/i.exec(startCell);
  if (!start) return false;
  const startCol = columnIndex(start[1] ?? "A");
  const startRow = Number(start[2]);
  const endCol = columnName(startCol + (rows[0]?.length ?? 1) - 1);
  const endRow = startRow + rows.length - 1;
  const ref = `${columnName(startCol)}${startRow}:${endCol}${endRow}`;
  const headers = (rows[0] ?? []).map((value, index) => String(value ?? `Column${index + 1}`));
  const existing = await findXlsxTableForStart(zip, startCell, tableName);
  const tablePath = existing?.path ?? `xl/tables/table${nextTableNumber(zip)}.xml`;
  const displayName = sanitizeTableName(tableName ?? existing?.name ?? `Table${tablePath.match(/table(\d+)\.xml$/)?.[1] ?? "1"}`);
  const tableId = Number(tablePath.match(/table(\d+)\.xml$/)?.[1] ?? 1);
  zip.file(tablePath, tableXml(tableId, displayName, ref, headers));
  await ensureContentTypeOverride(zip, `/${tablePath}`, "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml");
  const relId = await ensureWorksheetRelationship(zip, sheetNo, `../tables/${tablePath.split("/").pop()}`);
  const nextWorksheet = ensureWorksheetTablePart(worksheetXml, relId);
  if (nextWorksheet !== worksheetXml) zip.file(worksheetPath, nextWorksheet);
  return true;
}

async function findXlsxTableForStart(zip: Awaited<ReturnType<typeof loadZip>>, startCell: string, tableName?: string): Promise<{ path: string; name?: string } | undefined> {
  for (const path of Object.keys(zip.files).filter((item) => /^xl\/tables\/table\d+\.xml$/i.test(item)).sort()) {
    const xml = (await readZipText(zip, path)) ?? "";
    const attrs = /<table\b([^>]*)/.exec(xml)?.[1] ?? "";
    const ref = /\bref="([^"]+)"/.exec(attrs)?.[1] ?? "";
    const name = /\bdisplayName="([^"]+)"/.exec(attrs)?.[1] ?? /\bname="([^"]+)"/.exec(attrs)?.[1];
    if ((tableName && name === tableName) || ref.toUpperCase().startsWith(`${startCell.toUpperCase()}:`)) return { path, name };
  }
  return undefined;
}

async function ensureWorksheetRelationship(zip: Awaited<ReturnType<typeof loadZip>>, sheetNo: number, target: string): Promise<string> {
  const relsPath = `xl/worksheets/_rels/sheet${sheetNo}.xml.rels`;
  const xml = (await readZipText(zip, relsPath)) ?? '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
  const existing = new RegExp(`<Relationship\\b[^>]*\\bTarget="${escapeRegExp(target)}"[^>]*/>`).exec(xml)?.[0];
  const existingId = existing ? /\bId="([^"]+)"/.exec(existing)?.[1] : undefined;
  if (existingId) return existingId;
  const id = nextRelationshipId(xml);
  zip.file(relsPath, xml.replace(/<\/Relationships>\s*$/, `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="${target}"/></Relationships>`));
  return id;
}

function ensureWorksheetTablePart(xml: string, relId: string): string {
  if (new RegExp(`<tablePart\\b[^>]*r:id="${escapeRegExp(relId)}"[^>]*/>`).test(xml)) return xml;
  if (/<tableParts\b[\s\S]*?<\/tableParts>/.test(xml)) {
    return xml.replace(/<tableParts\b([^>]*)>([\s\S]*?)<\/tableParts>/, (_match, attrs: string, body: string) => {
      const count = (body.match(/<tablePart\b/g) ?? []).length + 1;
      const nextAttrs = /\bcount="/.test(attrs) ? attrs.replace(/\bcount="[^"]*"/, `count="${count}"`) : `${attrs} count="${count}"`;
      return `<tableParts${nextAttrs}>${body}<tablePart r:id="${relId}"/></tableParts>`;
    });
  }
  return xml.replace(/<\/worksheet>\s*$/, `<tableParts count="1"><tablePart r:id="${relId}"/></tableParts></worksheet>`);
}

function tableXml(id: number, name: string, ref: string, headers: string[]): string {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="${id}" name="${escapeXmlText(name)}" displayName="${escapeXmlText(name)}" ref="${ref}" totalsRowShown="0">`,
    `<autoFilter ref="${ref}"/>`,
    `<tableColumns count="${headers.length}">`,
    headers.map((header, index) => `<tableColumn id="${index + 1}" name="${escapeXmlText(header)}"/>`).join(""),
    "</tableColumns>",
    '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>',
    "</table>"
  ].join("");
}

function nextTableNumber(zip: Awaited<ReturnType<typeof loadZip>>): number {
  const numbers = Object.keys(zip.files)
    .map((path) => /^xl\/tables\/table(\d+)\.xml$/i.exec(path)?.[1])
    .filter(Boolean)
    .map(Number);
  return Math.max(0, ...numbers) + 1;
}

function sanitizeTableName(value: string): string {
  return value.replace(/[^A-Za-z0-9_]/g, "_").replace(/^[^A-Za-z_]+/, "T_").slice(0, 120) || "OfficegenTable";
}

function normalizePackageTarget(base: string, target: string): string {
  const normalizedTarget = target.replace(/\\/g, "/");
  const packageAbsolute = normalizedTarget.startsWith("/");
  const parts = `${packageAbsolute || !base ? "" : `${base}/`}${packageAbsolute ? normalizedTarget.slice(1) : normalizedTarget}`.split("/");
  const normalized: string[] = [];
  for (const part of parts) {
    if (!part || part === ".") continue;
    if (part === "..") normalized.pop();
    else normalized.push(part);
  }
  return normalized.join("/");
}

async function ensureContentTypeOverride(zip: Awaited<ReturnType<typeof loadZip>>, partName: string, contentType: string): Promise<void> {
  const xml = (await readZipText(zip, "[Content_Types].xml")) ?? '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>';
  if (xml.includes(`PartName="${partName}"`)) return;
  zip.file("[Content_Types].xml", xml.replace(/<\/Types>\s*$/, `<Override PartName="${partName}" ContentType="${contentType}"/></Types>`));
}

async function ensureDocumentRelationship(zip: Awaited<ReturnType<typeof loadZip>>, type: string, target: string): Promise<string> {
  const relsPath = "word/_rels/document.xml.rels";
  const xml = (await readZipText(zip, relsPath)) ?? '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
  const existing = new RegExp(`<Relationship\\b[^>]*\\bType="${escapeRegExp(type)}"[^>]*\\bTarget="${escapeRegExp(target)}"[^>]*/>`).exec(xml)?.[0];
  const existingId = existing ? /\bId="([^"]+)"/.exec(existing)?.[1] : undefined;
  if (existingId) return existingId;
  const id = nextRelationshipId(xml);
  zip.file(relsPath, xml.replace(/<\/Relationships>\s*$/, `<Relationship Id="${id}" Type="${type}" Target="${target}"/></Relationships>`));
  return id;
}

function ensureSectionReference(documentXml: string, kind: "header" | "footer", relId: string): string {
  const tag = kind === "header" ? "w:headerReference" : "w:footerReference";
  const reference = `<${tag} w:type="default" r:id="${relId}"/>`;
  if (new RegExp(`<${tag}\\b[^>]*w:type="default"[^>]*/>`).test(documentXml)) {
    return documentXml.replace(new RegExp(`<${tag}\\b[^>]*w:type="default"[^>]*/>`), reference);
  }
  if (/<w:sectPr\b[\s\S]*?<\/w:sectPr>/.test(documentXml)) {
    return documentXml.replace(/<w:sectPr\b([^>]*)>/, `<w:sectPr$1>${reference}`);
  }
  return documentXml.replace(/<\/w:body>/, `<w:sectPr>${reference}</w:sectPr></w:body>`);
}

function replaceNthParagraph(xml: string, ordinal: number, replacer: (paragraph: string) => string): { changed: boolean; xml: string } {
  let index = 0;
  let changed = false;
  const next = xml.replace(/<w:p\b[\s\S]*?<\/w:p>/g, (paragraph) => {
    index += 1;
    if (index !== ordinal) return paragraph;
    const replaced = replacer(paragraph);
    changed = replaced !== paragraph;
    return replaced;
  });
  return { changed, xml: next };
}

function detectMediaType(bytes: Uint8Array, path?: string): string {
  const ext = path?.split(".").pop()?.toLowerCase();
  if (bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47) return "image/png";
  if (bytes[0] === 0xff && bytes[1] === 0xd8) return "image/jpeg";
  if (String.fromCharCode(...bytes.slice(0, 120)).includes("<svg")) return "image/svg+xml";
  if (bytes.length >= 6 && (Buffer.from(bytes.slice(0, 6)).toString("ascii") === "GIF87a" || Buffer.from(bytes.slice(0, 6)).toString("ascii") === "GIF89a")) return "image/gif";
  if (ext === "emf" || ext === "wmf" || ext === "gif") return "application/octet-stream";
  return "application/octet-stream";
}

function detectDimensions(bytes: Uint8Array, mediaType: string): { width?: number; height?: number } {
  if (mediaType === "image/png" && bytes.length >= 24) {
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    return { width: view.getUint32(16), height: view.getUint32(20) };
  }
  if (mediaType === "image/jpeg") {
    let offset = 2;
    while (offset + 9 < bytes.length) {
      if (bytes[offset] !== 0xff) break;
      const marker = bytes[offset + 1];
      const length = (bytes[offset + 2] << 8) + bytes[offset + 3];
      if (marker >= 0xc0 && marker <= 0xc3) {
        return { height: (bytes[offset + 5] << 8) + bytes[offset + 6], width: (bytes[offset + 7] << 8) + bytes[offset + 8] };
      }
      offset += 2 + length;
    }
  }
  return {};
}

function mediaTypeFromExtension(path: string): string | undefined {
  const ext = path.split(".").pop()?.toLowerCase();
  if (ext === "png") return "image/png";
  if (ext === "jpg" || ext === "jpeg") return "image/jpeg";
  if (ext === "svg") return "image/svg+xml";
  if (ext === "gif") return "image/gif";
  return undefined;
}

function replaceShapeText(xml: string, ordinal: number, text: string): { changed: boolean; xml: string } {
  let index = 0;
  let changed = false;
  const next = xml.replace(/<p:sp\b[\s\S]*?<\/p:sp>/g, (shape) => {
    index += 1;
    if (index !== ordinal) return shape;
    const replaced = setFirstTextInBlock(shape, "a:t", text);
    changed = replaced !== shape;
    return replaced;
  });
  return { changed, xml: next };
}

function validationFailures(selectorResult: ResolveEditSelectorsResult | undefined): EditOperationResult[] {
  return (selectorResult?.resolutions ?? [])
    .filter((resolution) => resolution.reason === "not-found" || resolution.reason === "ambiguous")
    .map((resolution) => ({
      operationIndex: resolution.operationIndex,
      op: "selector",
      applied: false,
      reason: resolution.reason === "ambiguous" ? "ambiguous" : "not-found",
      message: resolution.reason === "ambiguous"
        ? `SELECTOR_AMBIGUOUS: selector matched ${resolution.matchCount} objects.`
        : "SELECTOR_NOT_FOUND: selector matched no objects."
    }));
}

function editAbortResult(format: string, skipped: number, opResults: EditOperationResult[], caveats: string[]): EditResult {
  const errors = opResults.filter((result) => result.reason && result.reason !== "unsupported");
  return {
    schema: "officegen.edit.result@1.2",
    format,
    changed: false,
    applied: 0,
    skipped,
    opResults,
    errors: errors.length ? errors : undefined,
    caveats
  };
}

function selectorForOperation(operation: EditOperation): EditSelector | undefined {
  if ("selector" in operation) return operation.selector;
  return undefined;
}

function resolveMatches(objectMap: ObjectMapEntry[], selector: EditSelector): ObjectMapEntry[] {
  if (selector.stableObjectId) return objectMap.filter((entry) => entry.stableObjectId === selector.stableObjectId);
  if (selector.shapeName) return objectMap.filter((entry) => entry.label === selector.shapeName || entry.selectorHints?.shapeName === selector.shapeName || entry.selectorHints?.name === selector.shapeName);
  if (selector.placeholderKey) return objectMap.filter((entry) => entry.selectorHints?.placeholderKey === selector.placeholderKey || entry.selectorHints?.placeholder === selector.placeholderKey);
  if (selector.contentControlTag) return objectMap.filter((entry) => entry.selectorHints?.contentControlTag === selector.contentControlTag || entry.selectorHints?.tag === selector.contentControlTag);
  if (selector.namedRange) return objectMap.filter((entry) => entry.selectorHints?.namedRange === selector.namedRange || entry.label === selector.namedRange);
  const text = selector.textMatch?.text ?? selector.contains;
  if (!text) return [];
  return objectMap.filter((entry) => selector.textMatch?.exact ? entry.text === text : entry.text?.includes(text));
}

function singleMatch(objectMap: ObjectMapEntry[], selector: EditSelector): ObjectMapEntry {
  const matches = resolveMatches(objectMap, selector);
  if (!matches.length) throw new Error("SELECTOR_NOT_FOUND: selector matched no objects.");
  if (matches.length > 1) throw new Error(`SELECTOR_AMBIGUOUS: selector matched ${matches.length} objects.`);
  return matches[0] as ObjectMapEntry;
}

function selectorMatch(entry: ObjectMapEntry): EditSelectorResolution["matches"][number] {
  return {
    stableObjectId: entry.stableObjectId,
    kind: entry.kind,
    label: entry.label,
    text: entry.text,
    sourcePath: entry.sourcePath,
    xmlPath: entry.xmlPath
  };
}

function operationName(operation: EditOperation): string {
  return "op" in operation ? operation.op : operation.type;
}

function stableOrdinal(stableObjectId: string): number {
  return Number(stableObjectId.split(":").at(-1) ?? 0);
}

function slideNumberFromSelector(selector: EditSelector | undefined, objectMap: ObjectMapEntry[]): number | undefined {
  if (!selector) return undefined;
  const match = singleMatch(objectMap, selector);
  return Number(match.selectorHints?.slide);
}

function idempotencyMarkerPath(key: string): string {
  return `officegen/idempotency/${simpleHash(key)}.txt`;
}

function simpleHash(value: string): string {
  let hash = 2166136261;
  for (const char of value) {
    hash ^= char.charCodeAt(0);
    hash = Math.imul(hash, 16777619);
  }
  return (hash >>> 0).toString(16);
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function classifyEditError(error: unknown): EditOperationResult["reason"] {
  const message = error instanceof Error ? error.message : String(error);
  if (message.includes("SELECTOR_AMBIGUOUS")) return "ambiguous";
  if (message.includes("SELECTOR_NOT_FOUND")) return "not-found";
  return "validation-failed";
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

function rowFromRef(ref: string): number {
  return Number(/\d+/.exec(ref)?.[0] ?? 1);
}

async function editPdf(
  input: Awaited<ReturnType<typeof normalizeInput>>,
  operations: EditOperation[],
  options: EditOptions
): Promise<EditResult> {
  const pdf = await PDFDocument.load(input.bytes, { ignoreEncryption: true });
  const fontSet = await embedPdfFonts(pdf, operations.map((op) => "text" in op ? String(op.text) : ""));
  const font = fontSet.font;
  let applied = 0;
  let skipped = 0;

  for (const op of operations) {
    const name = operationName(op);
    if (name === "pdf.textOverlay") {
      const textOp = op as { page: number; text: string; x: number; y: number; size?: number; color?: string };
      if (!isValidPage(pdf, textOp.page)) {
        skipped += 1;
        continue;
      }
      const page = pdf.getPage(textOp.page - 1);
      page.drawText(textOp.text, {
        x: textOp.x,
        y: textOp.y,
        size: textOp.size ?? 12,
        font,
        color: parseRgb(textOp.color)
      });
      applied += 1;
    } else if (name === "pdf.annotation") {
      const annotation = op as { page: number; text: string; x: number; y: number; width?: number; height?: number };
      if (!isValidPage(pdf, annotation.page)) {
        skipped += 1;
        continue;
      }
      const page = pdf.getPage(annotation.page - 1);
      page.drawRectangle({
        x: annotation.x,
        y: annotation.y,
        width: annotation.width ?? 160,
        height: annotation.height ?? 48,
        borderColor: rgb(0.91, 0.59, 0.12),
        borderWidth: 1,
        color: rgb(1, 0.96, 0.82),
        opacity: 0.9
      });
      page.drawText(annotation.text, { x: annotation.x + 6, y: annotation.y + (annotation.height ?? 48) - 18, size: 10, font, color: rgb(0.2, 0.2, 0.2) });
      applied += 1;
    } else {
      skipped += 1;
    }
  }

  const bytes = options.dryRun ? undefined : await pdf.save({ useObjectStreams: false });
  if (!options.dryRun) await writeOutput(options.out, bytes as Uint8Array);
  return {
    schema: "officegen.edit.result@1.2",
    format: "pdf",
    changed: applied > 0,
    applied,
    skipped,
    out: options.dryRun ? undefined : options.out,
    bytes: options.dryRun || options.out ? undefined : bytes,
    caveats: ["PDF edit is additive; existing text/content is not removed in the MVP."]
  };
}

function isValidPage(pdf: PDFDocument, page: number): boolean {
  return Number.isInteger(page) && page >= 1 && page <= pdf.getPageCount();
}

function parseRgb(hex?: string): ReturnType<typeof rgb> {
  if (!hex || !/^#[0-9a-f]{6}$/i.test(hex)) return rgb(0, 0, 0);
  const r = Number.parseInt(hex.slice(1, 3), 16) / 255;
  const g = Number.parseInt(hex.slice(3, 5), 16) / 255;
  const b = Number.parseInt(hex.slice(5, 7), 16) / 255;
  return rgb(r, g, b);
}

void escapeXmlText;
