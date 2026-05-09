import {
  type InputLike,
  type ObjectMapEntry,
  getLoadedZipSafetyReport,
  isOfficeFormat,
  loadZip,
  normalizeInput,
  readZipText,
  writeOutput,
  zipSafetyCaveats,
  zipToBytes
} from "./shared.js";
import { inspect } from "./inspect.js";
import { insertParagraphAfter, setParagraphText } from "./ooxml/docx.js";
import { duplicateSlide, extractShapes, reorderSlides, replaceShapeBulletItems } from "./ooxml/pptx.js";
import { insertRows, setCell, sheetPath } from "./ooxml/xlsx.js";
import { escapeXmlText, replaceAllXmlText, setFirstTextInBlock } from "./ooxml/xml.js";
import { PDFDocument, rgb, StandardFonts } from "pdf-lib";

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
  | { op: "docx.insertParagraphAfter"; text: string; selector: EditSelector }
  | { op: "xlsx.insertRows"; sheet?: number; rowIndex: number; rows: unknown[][]; selector?: EditSelector }
  | { op: "xlsx.setCell"; sheet?: number; cell: string; value: string; selector?: EditSelector }
  | { op: "xlsx.updateTable"; sheet?: number; startCell: string; rows: unknown[][]; selector?: EditSelector };

export interface EditOptions {
  out?: string;
  dryRun?: boolean;
  resolveSelectors?: boolean;
  format?: "pptx" | "docx" | "xlsx" | "pdf" | "unknown";
  atomic?: boolean;
  validateFirst?: boolean;
  idempotencyKey?: string;
  continueOnError?: boolean;
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
    ? await resolveEditSelectorsForNormalized(normalized, operations)
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
  options: Pick<EditOptions, "format"> = {}
): Promise<ResolveEditSelectorsResult> {
  const normalized = await normalizeInput(input, options.format ?? "unknown");
  return resolveEditSelectorsForNormalized(normalized, operations);
}

async function resolveEditSelectorsForNormalized(
  normalized: Awaited<ReturnType<typeof normalizeInput>>,
  operations: EditOperation[]
): Promise<ResolveEditSelectorsResult> {
  const inspected = await inspect({ data: normalized.bytes, format: normalized.format });
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
  const zip = await loadZip(input);
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
  if (format === "docx" && op === "docx.insertParagraphAfter") {
    return editDocxInsertParagraph(zip, operation as { selector: EditSelector; text: string }, objectMap);
  }
  if (format === "xlsx" && op === "xlsx.insertRows") {
    const rowOp = operation as { sheet?: number; rowIndex: number; rows: unknown[][] };
    const path = sheetPath(rowOp.sheet);
    const xml = (await readZipText(zip, path)) ?? "";
    const next = insertRows(xml, rowOp.rowIndex, rowOp.rows);
    if (next.changed) zip.file(path, next.xml);
    return next.changed;
  }
  if (format === "xlsx" && op === "xlsx.setCell") {
    const cellOp = operation as { sheet?: number; cell: string; value: string };
    return editXlsxSetCell(zip, cellOp.sheet, cellOp.cell, cellOp.value);
  }
  if (format === "xlsx" && op === "xlsx.updateTable") {
    const tableOp = operation as { sheet?: number; startCell: string; rows: unknown[][] };
    let changed = false;
    const start = /^([A-Z]+)(\d+)$/i.exec(tableOp.startCell);
    if (!start) throw new Error(`SELECTOR_NOT_FOUND: invalid startCell for operation ${index}.`);
    const startCol = columnIndex(start[1] ?? "A");
    const startRow = Number(start[2]);
    for (const [r, row] of tableOp.rows.entries()) {
      for (const [c, value] of row.entries()) {
        changed = (await editXlsxSetCell(zip, tableOp.sheet, `${columnName(startCol + c)}${startRow + r}`, String(value ?? ""))) || changed;
      }
    }
    return changed;
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

async function editDocxInsertParagraph(zip: Awaited<ReturnType<typeof loadZip>>, operation: { selector: EditSelector; text: string }, objectMap: ObjectMapEntry[]): Promise<boolean> {
  const target = singleMatch(objectMap, operation.selector);
  if (!target?.sourcePath) throw new Error("SELECTOR_NOT_FOUND: selected DOCX paragraph has no sourcePath.");
  const xml = (await readZipText(zip, target.sourcePath)) ?? "";
  const ordinal = Number(target.selectorHints?.paragraph ?? stableOrdinal(target.stableObjectId));
  const next = insertParagraphAfter(xml, ordinal, operation.text);
  if (next.changed) zip.file(target.sourcePath, next.xml);
  return next.changed;
}

async function editXlsxSetCell(zip: Awaited<ReturnType<typeof loadZip>>, sheet: number | undefined, ref: string, value: string): Promise<boolean> {
  if (!ref) throw new Error("SELECTOR_NOT_FOUND: xlsx cell ref is required.");
  const path = sheetPath(sheet);
  const xml = (await readZipText(zip, path)) ?? "";
  const next = setCell(xml, ref, value);
  if (next.changed) zip.file(path, next.xml);
  return next.changed;
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
