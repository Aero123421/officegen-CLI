import {
  type InputLike,
  type OfficegenConfig,
  type ObjectMapEntry,
  decodeXmlEntities,
  getLoadedZipSafetyReport,
  isOfficeFormat,
  loadZip,
  normalizeInput,
  readZipBytes,
  readZipText,
  stripXmlTags,
  writeOutput,
  zipSafetyCaveats,
  zipToBytes
} from "./shared.js";
import { inspect } from "./inspect.js";
import { commentXml, insertParagraphAfter, insertedParagraphXml, replaceOrCreateHeaderFooter, setParagraphText } from "./ooxml/docx.js";
import { embedPdfFonts } from "./pdfFonts.js";
import { addBlankSlide, addTextBox, applyBoundsToPptxBlock, duplicateSlide, extractShapes, getSlidePaths, reorderSlides, replaceShapeBulletItems } from "./ooxml/pptx.js";
import { appendRows, insertRows, setCell, sheetPath } from "./ooxml/xlsx.js";
import { escapeXmlText, pxToEmu, replaceAllXmlText, setFirstTextInBlock, xmlAttr } from "./ooxml/xml.js";
import { nextRelationshipId } from "./ooxml/relationships.js";
import { PackageGraph } from "./ooxml/packageGraph.js";
import { applyXmlPatches, type XmlPatch } from "./ooxml/patchEngine.js";
import { createSourceFingerprint } from "./ooxml/sourceSpan.js";
import { buildTokenIndex } from "./ooxml/tokenIndex.js";
import { createEditTransaction, type EditPartStore, type EditTransaction } from "./ooxml/transaction.js";
import { createOperationRegistry, type OperationRegistry } from "./ooxml/operations/registry.js";
import type { OperationResult } from "./ooxml/operations/types.js";
import { buildObjectGraph, type ObjectGraph, type ObjectGraphNode } from "./graphs/objectGraph.js";
import {
  SELECTOR_GRAPH_LOW_CONFIDENCE_THRESHOLD,
  objectGraphHash as hashObjectGraph,
  selectionLockForNode,
  selectorResolutionNextActions,
  sourceFingerprintForNode,
  type SelectorResolutionV2,
  type SelectorResolutionV2Status,
  type SelectorSelectionLock
} from "./graphs/selectorGraph.js";
import { PDFDocument, rgb } from "pdf-lib";
import JSZip from "jszip";
import { createHash } from "node:crypto";

export type EditSelector = {
  stableObjectId?: string;
  slide?: number;
  shapeId?: string;
  contains?: string;
  placeholderKey?: string;
  placeholder?: string;
  shapeName?: string;
  contentControlTag?: string;
  namedRange?: string;
  sheetName?: string;
  cell?: string;
  tableName?: string;
  chartPath?: string;
  textMatch?: { text: string; exact?: boolean };
  textHash?: string;
  positionHash?: string;
  sourcePath?: string;
  xmlPath?: string;
  page?: number;
  story?: string;
  paragraph?: number;
  table?: number;
  row?: number;
  column?: number;
  range?: string;
  relationshipId?: string;
  assetPath?: string;
  commentId?: string;
  revisionId?: string;
  nearestTo?: { slide?: number; x: number; y: number };
  rightOf?: string | { text: string; slide?: number };
  largestTextOnSlide?: number | boolean;
  nthBodyShape?: { slide: number; n: number };
};

export type PptxBulletListItem = string | { text: string; level?: number; bold?: boolean; numbering?: boolean };
type XlsxSheetRef = number | string;

export type EditOperation =
  | { type: "replaceText"; from: string; to: string; selector?: EditSelector }
  | { type: "setText"; text: string; selector: EditSelector }
  | { type: "pdf.textOverlay"; page: number; text: string; x: number; y: number; size?: number; color?: string }
  | { type: "pdf.annotation"; page: number; text: string; x: number; y: number; width?: number; height?: number }
  | { type: "pdf.redact"; page?: number; text?: string; selector?: EditSelector }
  | { op: "replaceText"; from: string; to: string; selector?: EditSelector }
  | { op: "setText"; text: string; selector: EditSelector }
  | { op: "pdf.redact"; page?: number; text?: string; selector?: EditSelector }
  | { op: "pptx.duplicateSlide"; slide?: number; after?: number; selector?: EditSelector }
  | { op: "pptx.addSlide"; after?: number }
  | { op: "pptx.addSlideFromLayout"; after?: number; layout?: string | number }
  | { op: "pptx.reorderSlides"; order: number[]; selector?: EditSelector }
  | { op: "pptx.addTextbox"; slide: number; text: string; bounds: { x: number; y: number; width: number; height: number }; name?: string; fontSize?: number; bold?: boolean }
  | { op: "pptx.formatTitle"; selector: EditSelector; fontSize?: number; bold?: boolean; textCase?: "upper" | "lower" | "title" | "sentence" }
  | { op: "pptx.formatAllTitles"; fontSize?: number; bold?: boolean; textCase?: "upper" | "lower" | "title" | "sentence" }
  | { op: "pptx.replaceBodyBullets"; slide: number; items: PptxBulletListItem[]; spaceBeforeForLevel1ExceptFirst?: number }
  | { op: "pptx.fitContentToPlaceholder"; selector: EditSelector; minFontSize?: number }
  | { op: "pptx.alignObjects"; selectors: EditSelector[]; mode: "left" | "right" | "center" | "top" | "bottom" | "middle" }
  | { op: "pptx.distributeObjects"; selectors: EditSelector[]; axis: "x" | "y" }
  | { op: "pptx.setAltText"; selector: EditSelector; title?: string; description?: string; decorative?: boolean }
  | { op: "pptx.setSpeakerNotes"; slide: number; text: string; mode?: "replace" | "append" }
  | { op: "pptx.replaceWithBulletList"; items: PptxBulletListItem[]; selector: EditSelector; spaceBeforeForLevel1ExceptFirst?: number }
  | { op: "pptx.insertBulletItems"; items: string[]; selector: EditSelector }
  | { op: "pptx.replaceBulletItems"; items: string[]; selector: EditSelector }
  | { op: "pptx.setFontSize"; selector: EditSelector; fontSize: number }
  | { op: "pptx.setBold"; selector: EditSelector; bold: boolean }
  | { op: "pptx.setBulletLevel"; selector: EditSelector; level: number }
  | { op: "pptx.setNumbering"; selector: EditSelector; level?: number; startAt?: number }
  | { op: "pptx.setLineSpacing"; selector: EditSelector; lineSpacing: number }
  | { op: "pptx.setSpaceBefore"; selector: EditSelector; spaceBefore: number }
  | { op: "pptx.setTextCase"; selector: EditSelector; textCase: "upper" | "lower" | "title" | "sentence" }
  | { op: "pptx.setTableCellText"; selector: EditSelector; text: string }
  | { op: "pptx.replaceImageByShape"; selector: EditSelector; replacementBase64: string; replacementPath?: string; fit?: "contain" | "cover" | "stretch"; crop?: CropRect }
  | { op: "pptx.updateChartData"; selector: EditSelector; categories: string[]; values: number[]; seriesName?: string }
  | { op: "pptx.setBounds"; selector: EditSelector; bounds: { x: number; y: number; width: number; height: number } }
  | { op: "docx.insertParagraphAfter"; text: string; selector: EditSelector }
  | { op: "docx.replaceTextSmart"; from: string; to: string; selector?: EditSelector }
  | { op: "docx.setTableCellText"; text: string; selector: EditSelector }
  | { op: "docx.setHeader"; text: string }
  | { op: "docx.setFooter"; text: string }
  | { op: "docx.setStyle"; styleId: string; font?: string; size?: number; bold?: boolean }
  | { op: "docx.addComment"; text: string; selector: EditSelector; author?: string }
  | { op: "docx.addRedline"; text: string; selector: EditSelector; author?: string }
  | { op: "docx.redline.insert"; text: string; selector: EditSelector; author?: string }
  | { op: "docx.redline.delete"; selector: EditSelector; author?: string }
  | { op: "docx.redline.replace"; text: string; selector: EditSelector; author?: string }
  | { op: "docx.applyStyle"; styleId: string; selector: EditSelector }
  | { op: "docx.headerFooter.setText"; kind: "header" | "footer"; text: string }
  | { op: "xlsx.insertRows"; sheet?: XlsxSheetRef; sheetName?: string; rowIndex: number; rows: unknown[][]; selector?: EditSelector }
  | { op: "xlsx.appendRows"; sheet?: XlsxSheetRef; sheetName?: string; rows: unknown[][]; selector?: EditSelector }
  | { op: "xlsx.setCell"; sheet?: XlsxSheetRef; sheetName?: string; cell: string; value: unknown; selector?: EditSelector }
  | { op: "xlsx.setFormula"; sheet?: XlsxSheetRef; sheetName?: string; cell: string; formula: string; selector?: EditSelector }
  | { op: "xlsx.definedName.set"; name: string; ref: string }
  | { op: "xlsx.definedName.delete"; name: string }
  | { op: "xlsx.setRange"; sheet?: XlsxSheetRef; sheetName?: string; startCell: string; values: unknown[][]; selector?: EditSelector }
  | { op: "xlsx.updateTable"; sheet?: XlsxSheetRef; sheetName?: string; startCell: string; rows: unknown[][]; selector?: EditSelector }
  | { op: "xlsx.writeTable"; sheet?: XlsxSheetRef; sheetName?: string; startCell: string; rows: unknown[][]; tableName?: string; selector?: EditSelector }
  | { op: "xlsx.table.resize"; selector: EditSelector; ref: string }
  | { op: "xlsx.chart.setData"; selector: EditSelector; categories: string[]; values: number[]; seriesName?: string }
  | { op: "xlsx.pivot.refreshDefinition"; selector: EditSelector }
  | { op: "xlsx.pivot.refreshAll" }
  | { op: "xlsx.slicer.setSelection"; selector: EditSelector; selected: string[] };

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
  expectedInputSha256?: string;
  expectedObjectMapHash?: string;
  expectedObjectGraphHash?: string;
  selectionLock?: SelectorSelectionLock;
  minSelectorConfidence?: number;
  continueOnError?: boolean;
  allowPartial?: boolean;
  config?: OfficegenConfig;
}

export interface EditSelectorResolution {
  operationIndex: number;
  selector: EditSelector;
  stableObjectId?: string;
  matched: boolean;
  matchCount: number;
  status?: SelectorResolutionV2Status;
  confidence?: number;
  matches: Array<{
    nodeId?: string;
    stableObjectId: string;
    kind: string;
    confidence?: number;
    label?: string;
    text?: string;
    sourcePath?: string;
    xmlPath?: string;
    selectorHints?: Record<string, unknown>;
  }>;
  evidence?: SelectorResolutionV2["evidence"];
  ambiguityReason?: string;
  nextActions?: string[];
  selectionLock?: SelectorSelectionLock;
  selectorResolution?: SelectorResolutionV2;
  reason?: "not-found" | "ambiguous" | "low-confidence" | "unsupported-selector";
  diagnostics?: EditSelectorDiagnostic[];
  suggestions?: string[];
}

export interface EditSelectorNearCandidate {
  stableObjectId: string;
  kind: string;
  label?: string;
  text?: string;
  textPreview?: string;
  sourcePath?: string;
  xmlPath?: string;
  selectorHints?: Record<string, unknown>;
  suggestedSelector: EditSelector;
}

export interface EditSelectorDiagnostic {
  code: "SELECTOR_NEAR_WHITESPACE_INSENSITIVE_MATCH";
  severity: "info";
  message: string;
  selectorField: "contains" | "textMatch";
  requestedText: string;
  normalizedRequestedText: string;
  candidates: EditSelectorNearCandidate[];
}

export interface ResolveEditSelectorsResult {
  schema: "officegen.edit.selectors@1.2";
  format: string;
  inputSha256: string;
  objectMapHash: string;
  objectGraphHash: string;
  resolutions: EditSelectorResolution[];
  objectMap: ObjectMapEntry[];
  caveats: string[];
}

export interface EditSourceFingerprint {
  algorithm: "sha256";
  hash: string;
  byteLength: number;
  path?: string;
}

export interface EditBlockedEvidence {
  code: string;
  field?: string;
  expected?: string;
  current?: string;
  expectedHash?: string;
  currentHash?: string;
  operationCount?: number;
  wouldWrite?: false;
}

export interface EditOperationResult {
  operationIndex: number;
  op: string;
  applied: boolean;
  reason?: "not-found" | "ambiguous" | "low-confidence" | "unsupported" | "validation-failed" | "idempotency-replay" | "skipped-after-error" | "stale-plan";
  message?: string;
  evidence?: EditBlockedEvidence;
  diagnostics?: EditSelectorDiagnostic[];
}

export interface PatchPlanTouchedPart {
  path: string;
  change: "modified" | "created" | "deleted";
  beforeSha256?: string;
  afterSha256?: string;
  sourceFingerprint?: EditSourceFingerprint;
}

export interface PatchPlanOperation {
  operationIndex: number;
  op: string;
  wouldApply: boolean;
  reason?: EditOperationResult["reason"];
  message?: string;
  selector?: EditSelector;
}

export interface PatchPlan {
  schema: "officegen.patchPlan@2";
  format: string;
  wouldWrite: false;
  inputSha256: string;
  objectMapHash?: string;
  objectGraphHash?: string;
  sourceFingerprint: EditSourceFingerprint;
  operations: PatchPlanOperation[];
  touchedParts: PatchPlanTouchedPart[];
  expectedChangedParts: string[];
  sourceFingerprints: EditSourceFingerprint[];
  blocked: EditOperationResult[];
}

export interface EditResult {
  schema: "officegen.edit.result@1.2";
  format: string;
  dryRun?: boolean;
  inputSha256?: string;
  objectMapHash?: string;
  objectGraphHash?: string;
  sourceFingerprint?: EditSourceFingerprint;
  rolledBack?: boolean;
  changed: boolean;
  applied: number;
  skipped: number;
  out?: string;
  bytes?: Uint8Array;
  resolvedSelectors?: EditSelectorResolution[];
  opResults?: EditOperationResult[];
  errors?: EditOperationResult[];
  partial?: boolean;
  allowPartial?: boolean;
  patchPlan?: PatchPlan;
  caveats: string[];
}

type OfficeXmlFormat = "pptx" | "docx" | "xlsx";
type OfficeZip = Awaited<ReturnType<typeof loadZip>>;
type OfficePartValue = string | Uint8Array;

interface OfficeOperationShared {
  zip: OfficeZip;
  graph: PackageGraph;
  objectMap: ObjectMapEntry[];
}

interface OfficeEditContext {
  zip: OfficeZip;
  graph: PackageGraph;
  store: EditPartStore<OfficePartValue>;
  transaction: EditTransaction<OfficePartValue>;
  registry: OperationRegistry<EditOperation, OperationResult, OfficePartValue, OfficeOperationShared>;
  initialPartPaths: Set<string>;
  rollback?: Awaited<ReturnType<typeof rollbackOfficeTransaction>>;
}

export async function edit(input: InputLike, operations: EditOperation[], options: EditOptions = {}): Promise<EditResult> {
  const normalized = await normalizeInput(input, options.format ?? "unknown");
  const selectorResult = options.resolveSelectors || options.validateFirst !== false
    ? await resolveEditSelectorsForNormalized(normalized, operations, options.config)
    : undefined;
  const result = isOfficeFormat(normalized.format)
    ? await editOfficeXml(normalized, operations, options, selectorResult)
    : normalized.format === "pdf"
      ? await editPdf(normalized, operations, options, selectorResult)
      : undefined;
  if (result) {
    if (selectorResult && !(normalized.format === "pptx" && hasSelectorAfterPptxCreator(operations))) result.resolvedSelectors = selectorResult.resolutions;
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
  const objectGraph = buildObjectGraph(inspected.objectMap);
  const currentObjectGraphHash = hashObjectGraph(objectGraph);
  const resolutions = operations.flatMap((operation, index) => {
    const selector = selectorForOperation(operation);
    if (!selector) return [];
    return [selectorResolutionForObjectMap(index, selector, inspected.objectMap, objectGraph, currentObjectGraphHash, operationName(operation))];
  });
  return {
    schema: "officegen.edit.selectors@1.2",
    format: inspected.trusted.format,
    inputSha256: sha256Bytes(normalized.bytes),
    objectMapHash: objectMapHash(inspected.objectMap),
    objectGraphHash: currentObjectGraphHash,
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
  const officeFormat = input.format as OfficeXmlFormat;
  const graph = await PackageGraph.fromZip(zip, { format: officeFormat });
  const atomic = options.atomic ?? true;
  const continueOnError = options.continueOnError ?? false;
  const opResults: EditOperationResult[] = [];
  const inputSha256 = selectorResult?.inputSha256 ?? sha256Bytes(input.bytes);
  const sourceFingerprint = inputSourceFingerprint(input.bytes);
  const store = new ZipPartStore(zip);
  const officeContext: OfficeEditContext = {
    zip,
    graph,
    store,
    transaction: createEditTransaction<OfficePartValue>(store, { atomic, continueOnError: true }),
    registry: createOfficeOperationRegistry(),
    initialPartPaths: zipPartPathSet(zip)
  };
  let applied = 0;
  let skipped = 0;

  if (options.idempotencyKey) {
    const markerPath = idempotencyMarkerPath(options.idempotencyKey);
    if (zip.file(markerPath)) {
      const idempotencyResults = operations.map((operation, index) => ({
        operationIndex: index,
        op: operationName(operation),
        applied: false,
        reason: "idempotency-replay" as const,
        message: `idempotencyKey already applied: ${options.idempotencyKey}`
      }));
      return {
        schema: "officegen.edit.result@1.2",
        format: input.format,
        dryRun: options.dryRun,
        inputSha256,
        objectMapHash: selectorResult?.objectMapHash,
        objectGraphHash: selectorResult?.objectGraphHash,
        sourceFingerprint,
        changed: false,
        applied: 0,
        skipped: operations.length,
        opResults: idempotencyResults,
        patchPlan: options.dryRun ? await buildPatchPlan(input.format, input.bytes, operations, idempotencyResults, selectorResult) : undefined,
        caveats: ["IDEMPOTENCY_REPLAY: idempotencyKey marker already exists.", ...zipSafetyCaveats(getLoadedZipSafetyReport(zip))]
      };
    }
  }

  const staleErrors = stalePlanFailures(selectorResult, options, operations);
  if (staleErrors.length) {
    const patchPlan = options.dryRun
      ? await buildPatchPlan(input.format, input.bytes, operations, staleErrors, selectorResult)
      : undefined;
    return {
      schema: "officegen.edit.result@1.2",
      format: input.format,
      dryRun: options.dryRun,
      inputSha256,
      objectMapHash: selectorResult?.objectMapHash,
      objectGraphHash: selectorResult?.objectGraphHash,
      sourceFingerprint,
      changed: false,
      applied: 0,
      skipped: operations.length,
      opResults: staleErrors,
      errors: staleErrors,
      patchPlan,
      caveats: [
        "EDIT_STALE_PLAN: expected input or object map hash does not match the current file.",
        ...zipSafetyCaveats(getLoadedZipSafetyReport(zip))
      ]
    };
  }

  const dynamicPptxSelectors = input.format === "pptx" && hasSelectorAfterPptxCreator(operations);
  const runtimeResolutions = dynamicPptxSelectors
    ? (selectorResult?.resolutions.filter((resolution) => !hasPriorPptxCreator(operations, resolution.operationIndex)) ?? [])
    : undefined;
  const validationErrors = options.validateFirst === false
    ? []
    : validationFailures(selectorResult, options.minSelectorConfidence).filter((failure) => !dynamicPptxSelectors || !hasPriorPptxCreator(operations, failure.operationIndex));
  if (validationErrors.length && atomic) {
    const opResults = [...validationErrors];
    appendSkippedAfterErrorResults(opResults, operations, "Skipped because atomic selector validation failed.");
    return editAbortResult(input.format, operations.length, opResults, operations, [
      "Atomic edit aborted before writing because selector validation failed.",
      ...zipSafetyCaveats(getLoadedZipSafetyReport(zip))
    ], selectorResult, options.dryRun, input.bytes);
  }

  await snapshotExistingOfficeParts(officeContext.transaction, graph);

  for (const [index, operation] of operations.entries()) {
    if (opResults.some(isRequiredEditFailure) && !continueOnError) {
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
      const objectMap = dynamicPptxSelectors
        ? await inspectCurrentObjectMap(zip, input.format as "pptx" | "docx" | "xlsx", options.config)
        : selectorResult?.objectMap ?? [];
      const selector = selectorForOperation(operation);
      if (dynamicPptxSelectors && selector && hasPriorPptxCreator(operations, index)) {
        const currentGraph = buildObjectGraph(objectMap);
        runtimeResolutions?.push(selectorResolutionForObjectMap(index, selector, objectMap, currentGraph, hashObjectGraph(currentGraph), operationName(operation)));
      }
      const txResult = await officeContext.transaction.run(index, () =>
        applyOfficeOperation(officeContext, input.format as OfficeXmlFormat, operation, objectMap, index)
      );
      if (!txResult.applied) {
        throw txResult.error;
      }
      const changed = txResult.value === true;
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
      if (!continueOnError && atomic) {
        skipped += appendSkippedAfterErrorResults(opResults, operations, "Skipped because atomic edit aborted after an earlier error.");
        break;
      }
    }
  }

  const errors = opResults.filter(isRequiredEditFailure);
  if (errors.length && (atomic || !options.allowPartial || applied <= 0)) {
    officeContext.rollback ??= await rollbackOfficeTransaction(zip, officeContext.transaction, officeContext.initialPartPaths);
    return editAbortResult(input.format, skipped, opResults, operations, [
      atomic
        ? "Atomic edit aborted; no output bytes were written."
        : "Edit aborted before writing because not all required operations succeeded. Pass allowPartial to permit best-effort output.",
      rollbackCaveat(officeContext.rollback),
      ...zipSafetyCaveats(getLoadedZipSafetyReport(zip))
    ], selectorResult, options.dryRun, input.bytes);
  }

  if (options.idempotencyKey && applied > 0) await officeContext.transaction.writePart(idempotencyMarkerPath(options.idempotencyKey), new Date().toISOString());
  const patchPlan = options.dryRun
    ? await buildPatchPlan(input.format, input.bytes, operations, opResults, selectorResult, officeContext.transaction, officeContext.store, officeContext.initialPartPaths, zipPartPathSet(zip))
    : undefined;
  const commit = officeContext.transaction.closedForWrites ? undefined : await officeContext.transaction.commit();
  const bytes = options.dryRun ? undefined : await zipToBytes(zip);
  if (!options.dryRun) await writeOutput(options.out, bytes as Uint8Array);
  return {
    schema: "officegen.edit.result@1.2",
    format: input.format,
    dryRun: options.dryRun,
    inputSha256,
    objectMapHash: selectorResult?.objectMapHash,
    objectGraphHash: selectorResult?.objectGraphHash,
    sourceFingerprint,
    changed: applied > 0,
    applied,
    skipped,
    out: options.dryRun ? undefined : options.out,
    bytes: options.dryRun || options.out ? undefined : bytes,
    opResults,
    resolvedSelectors: runtimeResolutions,
    errors: errors.length ? errors : undefined,
    partial: errors.length ? true : undefined,
    allowPartial: options.allowPartial || undefined,
    patchPlan,
    caveats: [
      "Office XML edits preserve unknown parts but do not recalculate native layout, formulas, or theme-derived rendering.",
      ...(commit ? [`EDIT_TRANSACTION: journaled ${commit.journaledParts} package part snapshot(s).`] : []),
      ...zipSafetyCaveats(getLoadedZipSafetyReport(zip))
    ]
  };
}

async function applyOfficeOperation(
  context: OfficeEditContext,
  format: OfficeXmlFormat,
  operation: EditOperation,
  objectMap: ObjectMapEntry[],
  index: number
): Promise<boolean> {
  const op = operationName(operation);
  const registryResult = await applyRegisteredOfficeOperation(context, format, operation, objectMap);
  if (registryResult !== undefined) return registryResult;
  const zip = context.zip;
  if (op === "replaceText") return replaceTextInEditableParts(zip, format, (operation as { from: string; to: string }).from, (operation as { from: string; to: string }).to);
  if (op === "setText") return setSelectedText(zip, format, operation as { selector: EditSelector; text: string }, objectMap);
  if (format === "pptx" && op === "pptx.duplicateSlide") {
    const duplicate = operation as { slide?: number; after?: number; selector?: EditSelector };
    await duplicateSlide(zip, duplicate.slide ?? slideNumberFromSelector(duplicate.selector, objectMap) ?? 1, duplicate.after);
    return true;
  }
  if (format === "pptx" && (op === "pptx.addSlide" || op === "pptx.addSlideFromLayout")) {
    await addBlankSlide(zip, (operation as { after?: number }).after);
    return true;
  }
  if (format === "pptx" && op === "pptx.reorderSlides") {
    await reorderSlides(zip, (operation as { order: number[] }).order);
    return true;
  }
  if (format === "pptx" && op === "pptx.addTextbox") {
    const add = operation as { slide: number; text: string; bounds: { x: number; y: number; width: number; height: number }; name?: string; fontSize?: number; bold?: boolean };
    await addTextBox(zip, add.slide, add);
    return true;
  }
  if (format === "pptx" && op === "pptx.formatTitle") {
    return editPptxTextStyle(zip, operation as { selector: EditSelector; fontSize?: number; bold?: boolean; textCase?: "upper" | "lower" | "title" | "sentence" }, objectMap);
  }
  if (format === "pptx" && op === "pptx.formatAllTitles") {
    return editPptxFormatAllTitles(zip, operation as { fontSize?: number; bold?: boolean; textCase?: "upper" | "lower" | "title" | "sentence" }, objectMap);
  }
  if (format === "pptx" && op === "pptx.replaceBodyBullets") {
    return editPptxReplaceBodyBullets(zip, operation as { slide: number; items: PptxBulletListItem[]; spaceBeforeForLevel1ExceptFirst?: number }, objectMap);
  }
  if (format === "pptx" && op === "pptx.fitContentToPlaceholder") {
    return editPptxFitContent(zip, operation as { selector: EditSelector; minFontSize?: number }, objectMap);
  }
  if (format === "pptx" && op === "pptx.alignObjects") {
    return editPptxAlignObjects(zip, operation as { selectors: EditSelector[]; mode: "left" | "right" | "center" | "top" | "bottom" | "middle" }, objectMap);
  }
  if (format === "pptx" && op === "pptx.distributeObjects") {
    return editPptxDistributeObjects(zip, operation as { selectors: EditSelector[]; axis: "x" | "y" }, objectMap);
  }
  if (format === "pptx" && op === "pptx.setAltText") {
    return editPptxSetAltText(zip, operation as { selector: EditSelector; title?: string; description?: string; decorative?: boolean }, objectMap);
  }
  if (format === "pptx" && op === "pptx.setSpeakerNotes") {
    return editPptxSetSpeakerNotes(zip, operation as { slide: number; text: string; mode?: "replace" | "append" });
  }
  if (format === "pptx" && op === "pptx.replaceWithBulletList") {
    return editPptxRichBullets(zip, operation as { selector: EditSelector; items: PptxBulletListItem[]; spaceBeforeForLevel1ExceptFirst?: number }, objectMap);
  }
  if (format === "pptx" && (op === "pptx.insertBulletItems" || op === "pptx.replaceBulletItems")) {
    return editPptxBullets(zip, operation as { selector: EditSelector; items: string[] }, objectMap, op === "pptx.insertBulletItems" ? "insert" : "replace");
  }
  if (format === "pptx" && op === "pptx.setFontSize") {
    return editPptxTextStyle(zip, operation as { selector: EditSelector; fontSize: number }, objectMap);
  }
  if (format === "pptx" && op === "pptx.setBold") {
    return editPptxTextStyle(zip, operation as { selector: EditSelector; bold: boolean }, objectMap);
  }
  if (format === "pptx" && op === "pptx.setBulletLevel") {
    return editPptxParagraphStyle(zip, operation as { selector: EditSelector; level: number }, objectMap);
  }
  if (format === "pptx" && op === "pptx.setNumbering") {
    return editPptxParagraphStyle(zip, { ...(operation as { selector: EditSelector; level?: number; startAt?: number }), numbering: true }, objectMap);
  }
  if (format === "pptx" && op === "pptx.setLineSpacing") {
    return editPptxParagraphStyle(zip, operation as { selector: EditSelector; lineSpacing: number }, objectMap);
  }
  if (format === "pptx" && op === "pptx.setSpaceBefore") {
    return editPptxParagraphStyle(zip, operation as { selector: EditSelector; spaceBefore: number }, objectMap);
  }
  if (format === "pptx" && op === "pptx.setTextCase") {
    return editPptxTextStyle(zip, operation as { selector: EditSelector; textCase: "upper" | "lower" | "title" | "sentence" }, objectMap);
  }
  if (format === "pptx" && op === "pptx.setTableCellText") {
    return editPptxSetTableCellText(zip, operation as { selector: EditSelector; text: string }, objectMap);
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
  if (format === "docx" && op === "docx.replaceTextSmart") {
    const smart = operation as { from: string; to: string; selector?: EditSelector };
    return smart.selector ? editDocxReplaceTextSmartSelected(zip, smart as { from: string; to: string; selector: EditSelector }, objectMap) : replaceTextInEditableParts(zip, format, smart.from, smart.to);
  }
  if (format === "docx" && op === "docx.setTableCellText") {
    return editDocxSetTableCellText(zip, operation as { selector: EditSelector; text: string }, objectMap);
  }
  if (format === "docx" && (op === "docx.setHeader" || op === "docx.setFooter" || op === "docx.headerFooter.setText")) {
    const headerFooter = operation as { kind?: "header" | "footer"; text: string };
    return editDocxHeaderFooter(zip, headerFooter.kind ?? (op === "docx.setHeader" ? "header" : "footer"), headerFooter.text);
  }
  if (format === "docx" && op === "docx.setStyle") {
    return editDocxStyle(zip, operation as { styleId: string; font?: string; size?: number; bold?: boolean });
  }
  if (format === "docx" && op === "docx.addComment") {
    return editDocxAddComment(zip, operation as { selector: EditSelector; text: string; author?: string }, objectMap);
  }
  if (format === "docx" && (op === "docx.addRedline" || op === "docx.redline.insert")) {
    return editDocxAddRedline(zip, operation as { selector: EditSelector; text: string; author?: string }, objectMap);
  }
  if (format === "docx" && op === "docx.redline.delete") {
    return editDocxDeleteRedline(zip, operation as { selector: EditSelector; author?: string }, objectMap);
  }
  if (format === "docx" && op === "docx.redline.replace") {
    const replace = operation as { selector: EditSelector; text: string; author?: string };
    const deleted = await editDocxDeleteRedline(zip, replace, objectMap);
    const inserted = await editDocxAddRedline(zip, replace, objectMap);
    return deleted || inserted;
  }
  if (format === "docx" && op === "docx.applyStyle") {
    return editDocxApplyStyle(zip, operation as { selector: EditSelector; styleId: string }, objectMap);
  }
  if (format === "xlsx" && op === "xlsx.insertRows") {
    const rowOp = operation as { sheet?: XlsxSheetRef; sheetName?: string; rowIndex: number; rows: unknown[][] };
    const sheetNo = await resolveXlsxSheet(zip, rowOp);
    const path = sheetPath(sheetNo);
    const xml = (await readZipText(zip, path)) ?? "";
    const next = insertRows(xml, rowOp.rowIndex, rowOp.rows);
    if (next.changed) zip.file(path, next.xml);
    return next.changed;
  }
  if (format === "xlsx" && op === "xlsx.appendRows") {
    const rowOp = operation as { sheet?: XlsxSheetRef; sheetName?: string; rows: unknown[][] };
    const sheetNo = await resolveXlsxSheet(zip, rowOp);
    const path = sheetPath(sheetNo);
    const xml = (await readZipText(zip, path)) ?? "";
    const next = appendRows(xml, rowOp.rows);
    if (next.changed) zip.file(path, next.xml);
    return next.changed;
  }
  if (format === "xlsx" && op === "xlsx.setCell") {
    const cellOp = operation as { sheet?: XlsxSheetRef; sheetName?: string; cell: string; value: unknown };
    return editXlsxSetCell(zip, await resolveXlsxSheet(zip, cellOp), cellOp.cell, cellOp.value);
  }
  if (format === "xlsx" && op === "xlsx.setFormula") {
    const formulaOp = operation as { sheet?: XlsxSheetRef; sheetName?: string; cell: string; formula: string };
    return editXlsxSetFormula(zip, await resolveXlsxSheet(zip, formulaOp), formulaOp.cell, formulaOp.formula);
  }
  if (format === "xlsx" && op === "xlsx.definedName.set") {
    return editXlsxDefinedName(zip, operation as { name: string; ref: string }, "set");
  }
  if (format === "xlsx" && op === "xlsx.definedName.delete") {
    return editXlsxDefinedName(zip, operation as { name: string }, "delete");
  }
  if (format === "xlsx" && op === "xlsx.setRange") {
    const rangeOp = operation as { sheet?: XlsxSheetRef; sheetName?: string; startCell: string; values: unknown[][] };
    return editXlsxSetRange(zip, await resolveXlsxSheet(zip, rangeOp), rangeOp.startCell, rangeOp.values);
  }
  if (format === "xlsx" && (op === "xlsx.updateTable" || op === "xlsx.writeTable")) {
    const tableOp = operation as { sheet?: XlsxSheetRef; sheetName?: string; startCell: string; rows: unknown[][]; tableName?: string };
    const sheetNo = await resolveXlsxSheet(zip, tableOp);
    let changed = false;
    const start = /^([A-Z]+)(\d+)$/i.exec(tableOp.startCell);
    if (!start) throw new Error(`SELECTOR_NOT_FOUND: invalid startCell for operation ${index}.`);
    const startCol = columnIndex(start[1] ?? "A");
    const startRow = Number(start[2]);
    for (const [r, row] of tableOp.rows.entries()) {
      for (const [c, value] of row.entries()) {
        changed = (await editXlsxSetCell(zip, sheetNo, `${columnName(startCol + c)}${startRow + r}`, value)) || changed;
      }
    }
    changed = (await ensureXlsxTable(zip, sheetNo, tableOp.startCell, tableOp.rows, tableOp.tableName)) || changed;
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
  if (format === "xlsx" && op === "xlsx.pivot.refreshAll") {
    return refreshAllPivotDefinitions(zip);
  }
  if (format === "xlsx" && op === "xlsx.slicer.setSelection") {
    return editXlsxSlicerSelection(zip, operation as { selector: EditSelector; selected: string[] }, objectMap);
  }
  return false;
}

function createOfficeOperationRegistry(): OperationRegistry<EditOperation, OperationResult, OfficePartValue, OfficeOperationShared> {
  const registry = createOperationRegistry<EditOperation, OperationResult, OfficePartValue, OfficeOperationShared>();
  for (const format of ["pptx", "docx", "xlsx"] as const) {
    registry.register({
      format,
      opName: "replaceText",
      handler: async (operation, { transaction, shared }) => {
        if (!shared) return { applied: false, reason: "unsupported" };
        const changed = await replaceTextInEditablePartsWithPatches(
          shared.zip,
          shared.graph,
          transaction,
          format,
          (operation as { from: string }).from,
          (operation as { to: string }).to
        );
        return { applied: changed, changed, reason: changed ? undefined : "not-found" };
      }
    });
    registry.register({
      format,
      opName: "setText",
      handler: async (operation, { transaction, shared }) => {
        if (!shared) return { applied: false, reason: "unsupported" };
        const safe = await setSelectedTextWithPatchEngine(
          shared.objectMap,
          transaction,
          operation as { selector: EditSelector; text: string }
        );
        if (!safe.handled) return { applied: false, reason: "legacy-fallback" };
        return { applied: safe.changed, changed: safe.changed, reason: safe.changed ? undefined : "not-found" };
      }
    });
  }
  return registry;
}

async function applyRegisteredOfficeOperation(
  context: OfficeEditContext,
  format: OfficeXmlFormat,
  operation: EditOperation,
  objectMap: ObjectMapEntry[]
): Promise<boolean | undefined> {
  const lookup = context.registry.lookup(format, operation);
  if (!lookup.supported) return undefined;
  const result = await lookup.registration.handler(operation, {
    format,
    transaction: context.transaction,
    store: context.store,
    options: { atomic: context.transaction.atomic, continueOnError: context.transaction.continueOnError },
    shared: { zip: context.zip, graph: context.graph, objectMap }
  });
  if (result.reason === "legacy-fallback") return undefined;
  return result.changed ?? result.applied;
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

async function replaceTextInEditablePartsWithPatches(
  zip: OfficeZip,
  graph: PackageGraph,
  transaction: EditTransaction<OfficePartValue>,
  format: OfficeXmlFormat,
  from: string,
  to: string
): Promise<boolean> {
  if (!from) return false;
  const escapedFrom = escapeXmlText(from);
  const escapedTo = escapeXmlText(to);
  let changed = false;
  for (const path of editableXmlPartPaths(zip, graph, format)) {
    const xml = await transaction.readPart(path);
    if (typeof xml !== "string" || !xml.includes(escapedFrom)) continue;
    const index = buildTokenIndex(xml);
    const patches: XmlPatch[] = index.textRuns
      .filter((run) => run.text.includes(escapedFrom))
      .map((run) => ({
        type: "replace",
        span: run.valueSpan,
        value: run.text.split(escapedFrom).join(escapedTo),
        fingerprint: createSourceFingerprint(xml, run.valueSpan)
      }));
    if (!patches.length) continue;
    const next = applyXmlPatches(xml, patches);
    if (next !== xml) {
      await transaction.writePart(path, next);
      changed = true;
    }
  }
  return changed;
}

async function setSelectedTextWithPatchEngine(
  objectMap: ObjectMapEntry[],
  transaction: EditTransaction<OfficePartValue>,
  operation: { selector: EditSelector; text: string }
): Promise<{ handled: boolean; changed: boolean }> {
  const target = singleMatch(objectMap, operation.selector);
  if (!target?.sourcePath || target.text === undefined) return { handled: false, changed: false };
  const xml = await transaction.readPart(target.sourcePath);
  if (typeof xml !== "string") return { handled: false, changed: false };
  const escapedCurrent = escapeXmlText(target.text);
  const index = buildTokenIndex(xml);
  const candidates = index.textRuns.filter((run) => run.text === escapedCurrent);
  if (candidates.length !== 1) return { handled: false, changed: false };
  const run = candidates[0];
  if (!run) return { handled: false, changed: false };
  const escapedNext = escapeXmlText(operation.text);
  if (run.text === escapedNext) return { handled: true, changed: false };
  const next = applyXmlPatches(xml, [{
    type: "replace",
    span: run.valueSpan,
    value: escapedNext,
    fingerprint: createSourceFingerprint(xml, run.valueSpan)
  }]);
  await transaction.writePart(target.sourcePath, next);
  return { handled: true, changed: next !== xml };
}

function editableXmlPartPaths(zip: OfficeZip, graph: PackageGraph, format: OfficeXmlFormat): string[] {
  const graphPaths = graph.listParts().map((part) => part.path);
  const zipPaths = Object.keys(zip.files).filter((path) => !zip.files[path]?.dir);
  return [...new Set([...graphPaths, ...zipPaths])]
    .filter((path) =>
      format === "pptx"
        ? /^ppt\/slides\/slide\d+\.xml$/i.test(path)
        : format === "docx"
          ? /^word\/(document|header\d+|footer\d+)\.xml$/i.test(path)
          : /^xl\/(worksheets\/sheet\d+|sharedStrings)\.xml$/i.test(path)
    )
    .sort((left, right) => left.localeCompare(right));
}

async function snapshotExistingOfficeParts(transaction: EditTransaction<OfficePartValue>, graph: PackageGraph): Promise<void> {
  for (const part of graph.listParts()) {
    await transaction.snapshotPart(part.path);
  }
}

async function rollbackOfficeTransaction(
  zip: OfficeZip,
  transaction: EditTransaction<OfficePartValue>,
  initialPartPaths: Set<string>
): Promise<{ restoredParts: number; removedCreatedParts: number; errors: Array<{ path: string; error: unknown }> }> {
  const rollback = await transaction.rollback();
  let removedCreatedParts = 0;
  for (const path of zipPartPathSet(zip)) {
    if (!initialPartPaths.has(path)) {
      zip.remove(path);
      removedCreatedParts += 1;
    }
  }
  return {
    restoredParts: rollback.restoredParts,
    removedCreatedParts,
    errors: rollback.errors
  };
}

function rollbackCaveat(rollback: Awaited<ReturnType<typeof rollbackOfficeTransaction>> | undefined): string {
  if (!rollback) return "EDIT_TRANSACTION_ROLLBACK: no transaction metadata was available.";
  const failed = rollback.errors.length ? `, rollback errors: ${rollback.errors.length}` : "";
  return `EDIT_TRANSACTION_ROLLBACK: restored ${rollback.restoredParts} package part snapshot(s), removed ${rollback.removedCreatedParts} created part(s)${failed}.`;
}

function zipPartPathSet(zip: OfficeZip): Set<string> {
  return new Set(Object.keys(zip.files).filter((path) => !zip.files[path]?.dir).map((path) => path.replace(/\\/g, "/").replace(/^\/+/, "")));
}

class ZipPartStore implements EditPartStore<OfficePartValue> {
  constructor(private readonly zip: OfficeZip) {}

  async readPart(path: string): Promise<OfficePartValue | undefined> {
    const normalized = normalizeZipPartPath(path);
    const file = this.zip.file(normalized);
    if (!file) return undefined;
    return isXmlPart(normalized) || isTextPart(normalized) ? file.async("string") : file.async("uint8array");
  }

  writePart(path: string, value: OfficePartValue): void {
    this.zip.file(normalizeZipPartPath(path), value);
  }

  deletePart(path: string): void {
    this.zip.remove(normalizeZipPartPath(path));
  }
}

function normalizeZipPartPath(path: string): string {
  return path.replace(/\\/g, "/").replace(/^\/+/, "");
}

function isXmlPart(path: string): boolean {
  return /\.xml$/i.test(path) || /\.rels$/i.test(path) || path === "[Content_Types].xml";
}

function isTextPart(path: string): boolean {
  return /\.(txt|csv|json)$/i.test(path);
}

async function setSelectedText(zip: Awaited<ReturnType<typeof loadZip>>, format: "pptx" | "docx" | "xlsx", operation: { selector: EditSelector; text: string }, objectMap: ObjectMapEntry[]): Promise<boolean> {
  const target = singleMatch(objectMap, operation.selector);
  if (!target?.sourcePath) throw new Error("SELECTOR_NOT_FOUND: selected object has no sourcePath.");
  const xml = (await readZipText(zip, target.sourcePath)) ?? "";
  if (format === "pptx") {
    if (target.kind === "tableCell") {
      return editPptxSetTableCellText(zip, operation, objectMap);
    }
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

async function editPptxRichBullets(
  zip: Awaited<ReturnType<typeof loadZip>>,
  operation: { selector: EditSelector; items: PptxBulletListItem[]; spaceBeforeForLevel1ExceptFirst?: number },
  objectMap: ObjectMapEntry[]
): Promise<boolean> {
  const target = singleMatch(objectMap, operation.selector);
  if (!target?.sourcePath || target.kind !== "shape") throw new Error("SELECTOR_NOT_FOUND: selected object is not a PPTX text shape.");
  const xml = (await readZipText(zip, target.sourcePath)) ?? "";
  const next = updateSelectedPptxShape(xml, target, (shape) => replaceTextBodyParagraphs(shape, operation.items, operation.spaceBeforeForLevel1ExceptFirst));
  if (next !== xml) zip.file(target.sourcePath, next);
  return next !== xml;
}

async function editPptxTextStyle(
  zip: Awaited<ReturnType<typeof loadZip>>,
  operation: { selector: EditSelector; fontSize?: number; bold?: boolean; textCase?: "upper" | "lower" | "title" | "sentence" },
  objectMap: ObjectMapEntry[]
): Promise<boolean> {
  const target = singleMatch(objectMap, operation.selector, { kinds: ["shape"] });
  if (!target?.sourcePath || target.kind !== "shape") throw new Error("SELECTOR_NOT_FOUND: selected object is not a PPTX text shape.");
  const xml = (await readZipText(zip, target.sourcePath)) ?? "";
  let styleTargetSeen = operation.fontSize === undefined && operation.bold === undefined;
  const result = updateSelectedPptxShapeWithResult(xml, target, (shape) => {
    let updated = shape;
    if (operation.textCase) updated = applyPptxTextCase(updated, operation.textCase);
    if (operation.fontSize !== undefined || operation.bold !== undefined) {
      styleTargetSeen = /<a:(r|fld)\b[\s\S]*?<\/a:\1>/.test(updated);
      updated = updateShapeRunProperties(updated, { fontSize: operation.fontSize, bold: operation.bold });
    }
    return updated;
  });
  if (result.changed) zip.file(target.sourcePath, result.xml);
  return result.changed || (result.matched && styleTargetSeen);
}

async function editPptxFormatAllTitles(
  zip: Awaited<ReturnType<typeof loadZip>>,
  operation: { fontSize?: number; bold?: boolean; textCase?: "upper" | "lower" | "title" | "sentence" },
  objectMap: ObjectMapEntry[]
): Promise<boolean> {
  let changed = false;
  const slideNumbers = [...new Set(objectMap.map((entry) => Number(entry.selectorHints?.slide)).filter(Number.isFinite))].sort((a, b) => a - b);
  for (const slide of slideNumbers) {
    const candidates = objectMap.filter((entry) => entry.kind === "shape" && Number(entry.selectorHints?.slide) === slide && entry.sourcePath);
    const title = candidates.find((entry) => ["title", "ctrTitle"].includes(String(entry.selectorHints?.placeholder ?? entry.selectorHints?.placeholderKey ?? "")))
      ?? candidates.sort((left, right) => textProminenceScore(right) - textProminenceScore(left))[0];
    if (!title?.sourcePath) continue;
    const xml = (await readZipText(zip, title.sourcePath)) ?? "";
    const next = updateSelectedPptxShape(xml, title, (shape) => {
      let updated = shape;
      if (operation.textCase) updated = applyPptxTextCase(updated, operation.textCase);
      return updateShapeRunProperties(updated, { fontSize: operation.fontSize, bold: operation.bold });
    });
    if (next !== xml) {
      zip.file(title.sourcePath, next);
      changed = true;
    }
  }
  return changed;
}

async function editPptxReplaceBodyBullets(
  zip: Awaited<ReturnType<typeof loadZip>>,
  operation: { slide: number; items: PptxBulletListItem[]; spaceBeforeForLevel1ExceptFirst?: number },
  objectMap: ObjectMapEntry[]
): Promise<boolean> {
  const bodies = objectMap.filter((entry) => entry.kind === "shape" && Number(entry.selectorHints?.slide) === operation.slide && entry.sourcePath);
  const target = bodies.find((entry) => {
    const placeholder = String(entry.selectorHints?.placeholder ?? entry.selectorHints?.placeholderKey ?? "");
    return placeholder === "body" || placeholder === "subTitle" || placeholder === "obj";
  }) ?? bodies.find((entry) => !["title", "ctrTitle"].includes(String(entry.selectorHints?.placeholder ?? entry.selectorHints?.placeholderKey ?? "")));
  if (!target?.sourcePath) throw new Error(`SELECTOR_NOT_FOUND: no body shape found on slide ${operation.slide}.`);
  const xml = (await readZipText(zip, target.sourcePath)) ?? "";
  const next = updateSelectedPptxShape(xml, target, (shape) => replaceTextBodyParagraphs(shape, operation.items, operation.spaceBeforeForLevel1ExceptFirst));
  if (next !== xml) zip.file(target.sourcePath, next);
  return next !== xml;
}

async function editPptxFitContent(
  zip: Awaited<ReturnType<typeof loadZip>>,
  operation: { selector: EditSelector; minFontSize?: number },
  objectMap: ObjectMapEntry[]
): Promise<boolean> {
  const target = singleMatch(objectMap, operation.selector);
  if (!target.sourcePath || target.kind !== "shape") throw new Error("SELECTOR_NOT_FOUND: selected object is not a PPTX text shape.");
  const xml = (await readZipText(zip, target.sourcePath)) ?? "";
  const next = updateSelectedPptxShape(xml, target, (shape) => {
    return ensurePptxNormAutofit(shape, operation.minFontSize);
  });
  if (next !== xml) zip.file(target.sourcePath, next);
  return next !== xml;
}

function ensurePptxNormAutofit(shapeXml: string, minFontSize?: number): string {
  const autofit = `<a:normAutofit${minFontSize ? ` lnSpcReduction="20000" fontScale="${Math.max(1000, Math.round(minFontSize * 2500))}"` : ""}/>`;
  const cleaned = shapeXml.replace(/<a:(?:spAutoFit|normAutofit|noAutofit)\b[^>]*\/>/g, "");
  if (/<a:bodyPr\b[^>]*\/>/.test(cleaned)) {
    return cleaned.replace(/<a:bodyPr\b([^>]*)\/>/, (_match, attrs: string) => `<a:bodyPr${attrs}>${autofit}</a:bodyPr>`);
  }
  return cleaned.replace(/(<a:bodyPr\b[^>]*>)([\s\S]*?)(<\/a:bodyPr>)/, (_match, open: string, body: string, close: string) => `${open}${body}${autofit}${close}`);
}

async function editPptxAlignObjects(
  zip: Awaited<ReturnType<typeof loadZip>>,
  operation: { selectors: EditSelector[]; mode: "left" | "right" | "center" | "top" | "bottom" | "middle" },
  objectMap: ObjectMapEntry[]
): Promise<boolean> {
  const targets = operation.selectors.map((selector) => singleMatch(objectMap, selector)).filter((entry) => entry.bounds);
  if (targets.length < 2) throw new Error("SELECTOR_NOT_FOUND: pptx.alignObjects requires at least two bounded objects.");
  const anchor = targets[0]?.bounds;
  if (!anchor) return false;
  let changed = false;
  for (const target of targets.slice(1)) {
    const bounds = target.bounds!;
    const nextBounds = { ...bounds };
    if (operation.mode === "left") nextBounds.x = anchor.x;
    if (operation.mode === "right") nextBounds.x = anchor.x + anchor.width - bounds.width;
    if (operation.mode === "center") nextBounds.x = anchor.x + (anchor.width - bounds.width) / 2;
    if (operation.mode === "top") nextBounds.y = anchor.y;
    if (operation.mode === "bottom") nextBounds.y = anchor.y + anchor.height - bounds.height;
    if (operation.mode === "middle") nextBounds.y = anchor.y + (anchor.height - bounds.height) / 2;
    changed = (await setPptxTargetBounds(zip, target, nextBounds)) || changed;
  }
  return changed;
}

async function editPptxDistributeObjects(
  zip: Awaited<ReturnType<typeof loadZip>>,
  operation: { selectors: EditSelector[]; axis: "x" | "y" },
  objectMap: ObjectMapEntry[]
): Promise<boolean> {
  const targets = operation.selectors.map((selector) => singleMatch(objectMap, selector)).filter((entry) => entry.bounds)
    .sort((left, right) => operation.axis === "x" ? left.bounds!.x - right.bounds!.x : left.bounds!.y - right.bounds!.y);
  if (targets.length < 3) throw new Error("SELECTOR_NOT_FOUND: pptx.distributeObjects requires at least three bounded objects.");
  const first = targets[0]!.bounds!;
  const last = targets[targets.length - 1]!.bounds!;
  const start = operation.axis === "x" ? first.x : first.y;
  const end = operation.axis === "x" ? last.x : last.y;
  const gap = (end - start) / (targets.length - 1);
  let changed = false;
  for (const [index, target] of targets.entries()) {
    if (index === 0 || index === targets.length - 1) continue;
    const bounds = { ...target.bounds! };
    if (operation.axis === "x") bounds.x = start + gap * index;
    else bounds.y = start + gap * index;
    changed = (await setPptxTargetBounds(zip, target, bounds)) || changed;
  }
  return changed;
}

async function editPptxSetAltText(
  zip: Awaited<ReturnType<typeof loadZip>>,
  operation: { selector: EditSelector; title?: string; description?: string; decorative?: boolean },
  objectMap: ObjectMapEntry[]
): Promise<boolean> {
  const target = singleMatch(objectMap, operation.selector);
  if (!target.sourcePath || !["shape", "picture", "chart"].includes(target.kind)) throw new Error("SELECTOR_NOT_FOUND: selected object cannot receive alt text.");
  const xml = (await readZipText(zip, target.sourcePath)) ?? "";
  const next = replacePptxObjectBlock(xml, target.kind, String(target.selectorHints?.shapeId ?? ""), Number(target.selectorHints?.shapeIndex ?? stableOrdinal(target.stableObjectId)), (block) =>
    block.replace(/<p:cNvPr\b([^>]*)\/>/, (_match, attrs: string) => {
      let nextAttrs = attrs;
      if (operation.title !== undefined) nextAttrs = upsertXmlAttr(nextAttrs, "title", operation.title);
      nextAttrs = upsertXmlAttr(nextAttrs, "descr", operation.decorative ? "" : operation.description ?? operation.title ?? "");
      return `<p:cNvPr ${nextAttrs}/>`;
    })
  );
  if (next !== xml) zip.file(target.sourcePath, next);
  return next !== xml;
}

async function editPptxParagraphStyle(
  zip: Awaited<ReturnType<typeof loadZip>>,
  operation: { selector: EditSelector; level?: number; numbering?: boolean; startAt?: number; lineSpacing?: number; spaceBefore?: number },
  objectMap: ObjectMapEntry[]
): Promise<boolean> {
  const target = singleMatch(objectMap, operation.selector);
  if (!target?.sourcePath || target.kind !== "shape") throw new Error("SELECTOR_NOT_FOUND: selected object is not a PPTX text shape.");
  const xml = (await readZipText(zip, target.sourcePath)) ?? "";
  const next = updateSelectedPptxShape(xml, target, (shape) => updateShapeParagraphProperties(shape, operation));
  if (next !== xml) zip.file(target.sourcePath, next);
  return next !== xml;
}

async function editPptxSetTableCellText(
  zip: Awaited<ReturnType<typeof loadZip>>,
  operation: { selector: EditSelector; text: string },
  objectMap: ObjectMapEntry[]
): Promise<boolean> {
  const target = singleMatch(objectMap, operation.selector);
  if (!target?.sourcePath || target.kind !== "tableCell") throw new Error("SELECTOR_NOT_FOUND: selected object is not a PPTX table cell.");
  const xml = (await readZipText(zip, target.sourcePath)) ?? "";
  const ordinal = Number(target.selectorHints?.tableCell ?? stableOrdinal(target.stableObjectId));
  let index = 0;
  let changed = false;
  const next = xml.replace(/<a:tc\b[\s\S]*?<\/a:tc>/g, (cell) => {
    index += 1;
    if (index !== ordinal) return cell;
    const replaced = setFirstTextInBlock(cell, "a:t", operation.text);
    changed = replaced !== cell;
    return replaced;
  });
  if (changed) zip.file(target.sourcePath, next);
  return changed;
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
  return setPptxTargetBounds(zip, target, operation.bounds);
}

async function setPptxTargetBounds(
  zip: Awaited<ReturnType<typeof loadZip>>,
  target: ObjectMapEntry,
  bounds: { x: number; y: number; width: number; height: number }
): Promise<boolean> {
  if (!target.sourcePath || !["shape", "picture", "chart"].includes(target.kind)) {
    throw new Error("SELECTOR_NOT_FOUND: selected object is not a PPTX shape, picture, or chart.");
  }
  const xml = (await readZipText(zip, target.sourcePath)) ?? "";
  const shapeId = String(target.selectorHints?.shapeId ?? "");
  const ordinal = Number(
    target.kind === "picture"
      ? target.selectorHints?.pictureIndex
      : target.kind === "chart"
        ? target.selectorHints?.chartIndex
        : target.selectorHints?.shapeIndex ?? stableOrdinal(target.stableObjectId)
  );
  const next = replacePptxObjectBlock(xml, target.kind, shapeId, ordinal, (block) => applyBoundsToPptxBlock(block, bounds));
  if (next !== xml) zip.file(target.sourcePath, next);
  return next !== xml;
}

async function editPptxSetSpeakerNotes(
  zip: Awaited<ReturnType<typeof loadZip>>,
  operation: { slide: number; text: string; mode?: "replace" | "append" }
): Promise<boolean> {
  const slidePaths = await getSlidePaths(zip);
  const slidePath = slidePaths[operation.slide - 1];
  if (!slidePath) throw new Error(`SELECTOR_NOT_FOUND: pptx slide ${operation.slide} not found.`);
  const notesPath = `ppt/notesSlides/notesSlide${operation.slide}.xml`;
  const existing = await readZipText(zip, notesPath);
  const text = operation.mode === "append" && existing
    ? `${extractNotesText(existing)}\n${operation.text}`.trim()
    : operation.text;
  zip.file(notesPath, notesSlideXml(text));
  await ensureContentTypeOverride(zip, `/${notesPath}`, "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml");
  const relsPath = slidePath.replace(/^ppt\/slides\//, "ppt/slides/_rels/") + ".rels";
  const relsXml = (await readZipText(zip, relsPath)) ?? '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
  if (!/\/notesSlide"/.test(relsXml) && !relsXml.includes(`../notesSlides/notesSlide${operation.slide}.xml`)) {
    const rId = nextRelationshipId(relsXml);
    zip.file(relsPath, relsXml.replace(/<\/Relationships>\s*$/, `<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide${operation.slide}.xml"/></Relationships>`));
  }
  return true;
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

async function editDocxDeleteRedline(zip: Awaited<ReturnType<typeof loadZip>>, operation: { selector: EditSelector; author?: string }, objectMap: ObjectMapEntry[]): Promise<boolean> {
  const target = singleMatch(objectMap, operation.selector);
  if (!target.sourcePath) throw new Error("SELECTOR_NOT_FOUND: selected DOCX paragraph has no sourcePath.");
  const xml = (await readZipText(zip, target.sourcePath)) ?? "";
  const ordinal = Number(target.selectorHints?.paragraph ?? stableOrdinal(target.stableObjectId));
  const nextId = nextDocxRevisionId(xml);
  const next = replaceNthParagraph(xml, ordinal, (paragraph) => {
    const text = paragraph.replace(/<[^>]+>/g, "");
    return `<w:p><w:del w:id="${nextId}" w:author="${escapeXmlText(operation.author ?? "officegen")}" w:date="${new Date().toISOString()}"><w:r><w:delText>${escapeXmlText(text)}</w:delText></w:r></w:del></w:p>`;
  });
  if (next.changed) zip.file(target.sourcePath, next.xml);
  return next.changed;
}

async function editDocxApplyStyle(zip: Awaited<ReturnType<typeof loadZip>>, operation: { selector: EditSelector; styleId: string }, objectMap: ObjectMapEntry[]): Promise<boolean> {
  const target = singleMatch(objectMap, operation.selector);
  if (!target.sourcePath) throw new Error("SELECTOR_NOT_FOUND: selected DOCX paragraph has no sourcePath.");
  const xml = (await readZipText(zip, target.sourcePath)) ?? "";
  const ordinal = Number(target.selectorHints?.paragraph ?? stableOrdinal(target.stableObjectId));
  const style = escapeXmlText(operation.styleId);
  const next = replaceNthParagraph(xml, ordinal, (paragraph) => {
    if (/<w:pPr\b[\s\S]*?<\/w:pPr>/.test(paragraph)) {
      if (/<w:pStyle\b[^>]*\/>/.test(paragraph)) return paragraph.replace(/<w:pStyle\b[^>]*\/>/, `<w:pStyle w:val="${style}"/>`);
      return paragraph.replace(/<w:pPr\b([^>]*)>/, `<w:pPr$1><w:pStyle w:val="${style}"/>`);
    }
    return paragraph.replace(/(<w:p\b[^>]*>)/, `$1<w:pPr><w:pStyle w:val="${style}"/></w:pPr>`);
  });
  if (next.changed) zip.file(target.sourcePath, next.xml);
  return next.changed;
}

async function editDocxReplaceTextSmartSelected(
  zip: Awaited<ReturnType<typeof loadZip>>,
  operation: { selector: EditSelector; from: string; to: string },
  objectMap: ObjectMapEntry[]
): Promise<boolean> {
  const target = singleMatch(objectMap, operation.selector);
  if (!target.sourcePath) throw new Error("SELECTOR_NOT_FOUND: selected DOCX object has no sourcePath.");
  const xml = (await readZipText(zip, target.sourcePath)) ?? "";
  const ordinal = Number(target.selectorHints?.paragraph ?? stableOrdinal(target.stableObjectId));
  const next = replaceNthParagraph(xml, ordinal, (paragraph) => smartReplaceDocxParagraph(paragraph, operation.from, operation.to));
  if (next.changed) zip.file(target.sourcePath, next.xml);
  return next.changed;
}

async function editDocxSetTableCellText(zip: Awaited<ReturnType<typeof loadZip>>, operation: { selector: EditSelector; text: string }, objectMap: ObjectMapEntry[]): Promise<boolean> {
  const target = singleMatch(objectMap, operation.selector);
  if (!target.sourcePath || target.kind !== "tableCell") throw new Error("SELECTOR_NOT_FOUND: selected object is not a DOCX table cell.");
  const xml = (await readZipText(zip, target.sourcePath)) ?? "";
  const cellOrdinal = Number(target.selectorHints?.cell ?? stableOrdinal(target.stableObjectId));
  let index = 0;
  let changed = false;
  const next = xml.replace(/<w:tc\b[\s\S]*?<\/w:tc>/g, (cell) => {
    index += 1;
    if (index !== cellOrdinal) return cell;
    const replaced = setFirstTextInBlock(cell, "w:t", operation.text);
    changed = replaced !== cell;
    return replaced;
  });
  if (changed) zip.file(target.sourcePath, next);
  return changed;
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

function smartReplaceDocxParagraph(paragraph: string, from: string, to: string): string {
  if (!from) return paragraph;
  const textRuns = [...paragraph.matchAll(/<w:t\b[^>]*>[\s\S]*?<\/w:t>/g)];
  const joined = textRuns.map((match) => stripXmlTags(match[0])).join("");
  if (!joined.includes(from)) return paragraph;
  const replaced = joined.split(from).join(to);
  let used = false;
  return paragraph.replace(/<w:t\b([^>]*)>[\s\S]*?<\/w:t>/g, (_match, attrs: string) => {
    if (used) return `<w:t${attrs}></w:t>`;
    used = true;
    return `<w:t${attrs}>${escapeXmlText(replaced)}</w:t>`;
  });
}

async function resolveXlsxSheet(zip: Awaited<ReturnType<typeof loadZip>>, operation: { sheet?: XlsxSheetRef; sheetName?: string }): Promise<number | undefined> {
  if (typeof operation.sheet === "number") return operation.sheet;
  const requested = typeof operation.sheetName === "string" && operation.sheetName.trim()
    ? operation.sheetName
    : typeof operation.sheet === "string"
      ? operation.sheet
      : undefined;
  if (!requested) return undefined;
  const workbook = (await readZipText(zip, "xl/workbook.xml")) ?? "";
  const names = readXlsxWorkbookSheetNames(workbook);
  const requestedKey = requested.trim().toLowerCase();
  const index = names.findIndex((name) => name.toLowerCase() === requestedKey);
  if (index >= 0) return index + 1;
  const numeric = Number(requested);
  if (Number.isInteger(numeric) && numeric > 0) return numeric;
  throw new Error(`SELECTOR_NOT_FOUND: XLSX sheet '${requested}' was not found.`);
}

function readXlsxWorkbookSheetNames(workbookXml: string): string[] {
  return [...workbookXml.matchAll(/<sheet\b[^>]*\bname="([^"]*)"/g)]
    .map((match) => decodeXmlEntities(match[1] ?? ""))
    .filter(Boolean);
}

async function editXlsxSetCell(zip: Awaited<ReturnType<typeof loadZip>>, sheet: number | undefined, ref: string, value: unknown): Promise<boolean> {
  if (!ref) throw new Error("SELECTOR_NOT_FOUND: xlsx cell ref is required.");
  const path = sheetPath(sheet);
  const xml = (await readZipText(zip, path)) ?? "";
  const next = setCell(xml, ref, value);
  if (next.changed) zip.file(path, next.xml);
  return next.changed;
}

async function editXlsxSetRange(zip: Awaited<ReturnType<typeof loadZip>>, sheet: number | undefined, startCell: string, values: unknown[][]): Promise<boolean> {
  const start = /^([A-Z]+)(\d+)$/i.exec(startCell);
  if (!start || !values.length) throw new Error("SELECTOR_NOT_FOUND: xlsx.setRange requires a valid startCell and non-empty values.");
  const startCol = columnIndex(start[1] ?? "A");
  const startRow = Number(start[2]);
  let changed = false;
  for (const [rowOffset, row] of values.entries()) {
    for (const [colOffset, value] of row.entries()) {
      changed = (await editXlsxSetCell(zip, sheet, `${columnName(startCol + colOffset)}${startRow + rowOffset}`, value)) || changed;
    }
  }
  return changed;
}

async function editXlsxSetFormula(zip: Awaited<ReturnType<typeof loadZip>>, sheet: number | undefined, ref: string, formula: string): Promise<boolean> {
  if (!ref) throw new Error("SELECTOR_NOT_FOUND: xlsx cell ref is required.");
  assertSafeXlsxFormula(formula);
  const path = sheetPath(sheet);
  const xml = (await readZipText(zip, path)) ?? "";
  const pattern = new RegExp(`<c\\b[^>]*\\br=["']${escapeRegExp(ref)}["'][^>]*?(?:\\/>|>[\\s\\S]*?<\\/c>)`);
  const rowNo = rowFromRef(ref);
  const rowPattern = new RegExp(`<row\\b([^>]*)\\br=["']${rowNo}["'][^>]*>[\\s\\S]*?<\\/row>`);
  const next = pattern.test(xml)
    ? xml.replace(pattern, (cell) => formulaCellXml(ref, formula, /^<c\b([^>]*)/.exec(cell)?.[1] ?? ""))
    : rowPattern.test(xml)
      ? xml.replace(rowPattern, (row) => row.replace(/<\/row>$/, `${formulaCellXml(ref, formula)}</row>`))
      : xml.replace(/<\/sheetData>/, `<row r="${rowNo}">${formulaCellXml(ref, formula)}</row></sheetData>`);
  if (next !== xml) {
    zip.file(path, next);
    await markXlsxRecalcNeeded(zip);
  }
  return next !== xml;
}

function formulaCellXml(ref: string, formula: string, existingAttrs = ""): string {
  const preservedAttrs = preserveXlsxCellAttrs(existingAttrs);
  return `<c r="${escapeXmlText(ref)}"${preservedAttrs}><f>${escapeXmlText(formula.replace(/^=/, ""))}</f></c>`;
}

function preserveXlsxCellAttrs(attrs: string): string {
  const preserved = [];
  for (const name of ["s", "cm", "vm", "ph"]) {
    const value = xmlAttr(attrs, name);
    if (value !== undefined) preserved.push(`${name}="${escapeXmlText(value)}"`);
  }
  return preserved.length ? ` ${preserved.join(" ")}` : "";
}

async function editXlsxDefinedName(zip: Awaited<ReturnType<typeof loadZip>>, operation: { name: string; ref?: string }, mode: "set" | "delete"): Promise<boolean> {
  const path = "xl/workbook.xml";
  const xml = (await readZipText(zip, path)) ?? "";
  if (!operation.name) throw new Error("SCHEMA_INVALID: defined name requires a non-empty name.");
  const name = escapeXmlText(operation.name);
  const tagPattern = new RegExp(`<definedName\\b[^>]*\\bname="${escapeRegExp(operation.name)}"[^>]*>[\\s\\S]*?<\\/definedName>`, "g");
  let next = xml.replace(tagPattern, "");
  if (mode === "set") {
    const ref = escapeXmlText(operation.ref ?? "");
    if (!ref) throw new Error("SCHEMA_INVALID: xlsx.definedName.set requires ref.");
    const definedName = `<definedName name="${name}">${ref}</definedName>`;
    if (/<definedNames\b[\s\S]*?<\/definedNames>/.test(next)) next = next.replace(/<\/definedNames>/, `${definedName}</definedNames>`);
    else next = next.replace(/<\/workbook>/, `<definedNames>${definedName}</definedNames></workbook>`);
  }
  if (next !== xml) zip.file(path, next);
  return next !== xml;
}

function assertSafeXlsxFormula(formula: string): void {
  const normalized = formula.replace(/^=/, "").toUpperCase();
  const forbidden = [
    /\[[^\]]+\]/,
    /\bWEBSERVICE\s*\(/,
    /\bHYPERLINK\s*\(/,
    /\bRTD\s*\(/,
    /\bDDE\s*\(/,
    /\bINDIRECT\s*\(/,
    /\bNOW\s*\(/,
    /\bRAND(?:BETWEEN)?\s*\(/
  ];
  if (forbidden.some((pattern) => pattern.test(normalized))) {
    throw new Error("SCHEMA_INVALID: xlsx.setFormula strict safety blocks external, volatile, indirection, and link-capable formulas.");
  }
}

async function markXlsxRecalcNeeded(zip: Awaited<ReturnType<typeof loadZip>>): Promise<void> {
  const workbookPath = "xl/workbook.xml";
  const workbook = (await readZipText(zip, workbookPath)) ?? "";
  if (workbook) {
    const next = updateXlsxWorkbookCalcPr(workbook);
    if (next !== workbook) zip.file(workbookPath, next);
  }
  zip.remove("xl/calcChain.xml");
}

function updateXlsxWorkbookCalcPr(workbook: string): string {
  const workbookPrefix = /^<\?xml\b[\s\S]*?\?>\s*<([A-Za-z_][\w.-]*:)?workbook\b|<([A-Za-z_][\w.-]*:)?workbook\b/.exec(workbook)?.[1]
    ?? /^<\?xml\b[\s\S]*?\?>\s*<([A-Za-z_][\w.-]*:)?workbook\b|<([A-Za-z_][\w.-]*:)?workbook\b/.exec(workbook)?.[2]
    ?? "";
  const calcName = `${workbookPrefix}calcPr`;
  const calcPattern = new RegExp(`<${escapeRegExp(calcName)}\\b([^>]*?)(?:\\s*\\/\\s*>|>[\\s\\S]*?<\\/${escapeRegExp(calcName)}>)`, "i");
  if (calcPattern.test(workbook)) {
    return workbook.replace(calcPattern, (_match, attrs: string) => calcPrXml(calcName, attrs));
  }
  const genericCalcPattern = /<([A-Za-z_][\w.-]*:)?calcPr\b([^>]*?)(?:\s*\/\s*>|>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?calcPr>)/i;
  if (genericCalcPattern.test(workbook)) {
    return workbook.replace(genericCalcPattern, (_match, prefix: string | undefined, attrs: string) => calcPrXml(`${prefix ?? ""}calcPr`, attrs));
  }
  return insertXlsxCalcPr(workbook, calcName);
}

function calcPrXml(tagName: string, attrs: string): string {
  const nextAttrs = upsertXmlAttr(upsertXmlAttr(attrs, "fullCalcOnLoad", "1"), "forceFullCalc", "1");
  return `<${tagName}${nextAttrs ? ` ${nextAttrs}` : ""}/>`;
}

function insertXlsxCalcPr(workbook: string, tagName: string): string {
  const calcPr = calcPrXml(tagName, "");
  if (/<(?:[A-Za-z_][\w.-]*:)?extLst\b/i.test(workbook)) {
    return workbook.replace(/<([A-Za-z_][\w.-]*:)?extLst\b/i, `${calcPr}<$1extLst`);
  }
  const workbookPrefix = tagName.includes(":") ? `${tagName.split(":")[0]}:` : "";
  const closePattern = new RegExp(`</${escapeRegExp(workbookPrefix)}workbook>`, "i");
  return workbook.replace(closePattern, `${calcPr}</${workbookPrefix}workbook>`);
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

function replacePptxObjectBlock(xml: string, kind: string, shapeId: string, ordinal: number, updater: (block: string) => string): string {
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
    if (shapeId && candidateId !== shapeId) return block;
    if (!shapeId && Number.isFinite(ordinal) && index !== ordinal) return block;
    if (!shapeId && !Number.isFinite(ordinal) && index !== 1) return block;
    return updater(block);
  });
}

function updateSelectedPptxShape(xml: string, target: ObjectMapEntry, updater: (shape: string) => string): string {
  return updateSelectedPptxShapeWithResult(xml, target, updater).xml;
}

function updateSelectedPptxShapeWithResult(xml: string, target: ObjectMapEntry, updater: (shape: string) => string): { xml: string; matched: boolean; changed: boolean } {
  const resolved = resolveCurrentPptxShape(xml, target);
  const shapeId = resolved?.shapeId ?? String(target.selectorHints?.shapeId ?? "");
  const ordinal = Number(resolved?.shapeIndex ?? target.selectorHints?.shapeIndex ?? stableOrdinal(target.stableObjectId));
  let index = 0;
  let matched = false;
  const next = xml.replace(/<p:sp\b[\s\S]*?<\/p:sp>/g, (shape) => {
    index += 1;
    const cNvPr = /<p:cNvPr\b([^>]*)\/>/.exec(shape)?.[1] ?? "";
    const candidateId = /\bid="([^"]+)"/.exec(cNvPr)?.[1];
    if (shapeId && candidateId !== shapeId) return shape;
    if (Number.isFinite(ordinal) && index !== ordinal) return shape;
    matched = true;
    return updater(shape);
  });
  return { xml: next, matched, changed: next !== xml };
}

function resolveCurrentPptxShape(xml: string, target: ObjectMapEntry): { shapeId?: string; shapeIndex: number } | undefined {
  if (!target.sourcePath) return undefined;
  const slide = Number(target.selectorHints?.slide ?? 1);
  const shapes = extractShapes(xml, slide, "", target.sourcePath);
  const byStableIdCandidates = shapes.filter((shape) => shape.stableObjectId === target.stableObjectId);
  const byStableId = preferPptxShapeCandidate(byStableIdCandidates, target);
  if (byStableId) return { shapeId: byStableId.shapeId, shapeIndex: byStableId.shapeIndex };
  const targetShapeId = String(target.selectorHints?.shapeId ?? "");
  if (targetShapeId) {
    const byShapeId = preferPptxShapeCandidate(shapes.filter((shape) => shape.shapeId === targetShapeId), target);
    if (byShapeId) return { shapeId: byShapeId.shapeId, shapeIndex: byShapeId.shapeIndex };
  }
  const targetShapeIndex = Number(target.selectorHints?.shapeIndex);
  if (Number.isFinite(targetShapeIndex)) {
    const byShapeIndex = shapes.find((shape) => shape.shapeIndex === targetShapeIndex);
    if (byShapeIndex) return { shapeId: byShapeIndex.shapeId, shapeIndex: byShapeIndex.shapeIndex };
  }
  return undefined;
}

function preferPptxShapeCandidate(
  candidates: ReturnType<typeof extractShapes>,
  target: ObjectMapEntry
): ReturnType<typeof extractShapes>[number] | undefined {
  if (candidates.length <= 1) return candidates[0];
  const targetText = target.text ?? target.textPreview;
  if (targetText) {
    const byExactText = candidates.find((shape) => shape.text === targetText);
    if (byExactText) return byExactText;
    const byContainedText = candidates.find((shape) => shape.text.includes(targetText) || targetText.includes(shape.text));
    if (byContainedText) return byContainedText;
  }
  const targetName = String(target.selectorHints?.name ?? target.selectorHints?.shapeName ?? target.label ?? "");
  if (targetName) {
    const byName = candidates.find((shape) => shape.name === targetName);
    if (byName) return byName;
  }
  return candidates[0];
}

function replaceTextBodyParagraphs(shape: string, items: PptxBulletListItem[], spaceBeforeForLevel1ExceptFirst?: number): string {
  const paragraphs = items.map((item, index) => richBulletParagraphXml(item, index, spaceBeforeForLevel1ExceptFirst)).join("");
  return shape.replace(/(<p:txBody\b[^>]*>)([\s\S]*?)(<\/p:txBody>)/, (_match, open: string, body: string, close: string) => {
    const bodyPr = firstXmlElement(body, "a:bodyPr") ?? "<a:bodyPr/>";
    const lstStyle = firstXmlElement(body, "a:lstStyle") ?? "<a:lstStyle/>";
    return `${open}${bodyPr}${lstStyle}${paragraphs}${close}`;
  });
}

function richBulletParagraphXml(item: PptxBulletListItem, index: number, spaceBeforeForLevel1ExceptFirst?: number): string {
  const record = typeof item === "string" ? { text: item, level: 0, bold: false, numbering: false } : item;
  const level = Math.max(0, Math.min(8, Math.round(Number(record.level ?? 0))));
  const bold = record.bold ? ' b="1"' : "";
  const bullet = record.numbering ? '<a:buAutoNum type="arabicPeriod"/>' : '<a:buChar char="&#8226;"/>';
  const spaceBefore = level === 0 && index > 0 && spaceBeforeForLevel1ExceptFirst
    ? `<a:spcBef><a:spcPts val="${Math.round(spaceBeforeForLevel1ExceptFirst * 100)}"/></a:spcBef>`
    : "";
  return `<a:p><a:pPr lvl="${level}">${spaceBefore}${bullet}</a:pPr><a:r><a:rPr${bold}/><a:t>${escapeXmlText(record.text)}</a:t></a:r></a:p>`;
}

function updateShapeRunProperties(shape: string, options: { fontSize?: number; bold?: boolean }): string {
  return shape
    .replace(/<a:r\b[^>]*>[\s\S]*?<\/a:r>/g, (run) => updateTextRunPropertyBlock(run, "a:r", options))
    .replace(/<a:fld\b[^>]*>[\s\S]*?<\/a:fld>/g, (field) => updateTextRunPropertyBlock(field, "a:fld", options));
}

function updateTextRunPropertyBlock(block: string, tag: "a:r" | "a:fld", options: { fontSize?: number; bold?: boolean }): string {
  if (/<a:rPr\b[^>]*\/>/.test(block)) {
    return block.replace(/<a:rPr\b([^>]*)\/>/, (_match, attrs: string) => `<a:rPr${runPropertyAttrs(attrs, options)}/>`);
  }
  if (/<a:rPr\b/.test(block)) {
    return block.replace(/<a:rPr\b([^>]*)>/, (_match, attrs: string) => `<a:rPr${runPropertyAttrs(attrs, options)}>`);
  }
  const escapedTag = escapeRegExp(tag);
  return block.replace(new RegExp(`(<${escapedTag}\\b[^>]*>)`), `$1<a:rPr${runPropertyAttrs("", options)}/>`);
}

function runPropertyAttrs(attrs: string, options: { fontSize?: number; bold?: boolean }): string {
  let next = attrs;
  if (options.fontSize !== undefined) next = upsertXmlAttr(next, "sz", String(Math.round(options.fontSize * 100)));
  if (options.bold !== undefined) next = upsertXmlAttr(next, "b", options.bold ? "1" : "0");
  return next ? ` ${next.trim()}` : "";
}

function updateShapeParagraphProperties(
  shape: string,
  options: { level?: number; numbering?: boolean; startAt?: number; lineSpacing?: number; spaceBefore?: number }
): string {
  return shape.replace(/<a:p\b[^>]*>[\s\S]*?<\/a:p>/g, (paragraph) => {
    if (/<a:pPr\b/.test(paragraph)) {
      return paragraph.replace(/<a:pPr\b([^>]*)\/>/, (_match, attrs: string) => paragraphPropertiesXml(attrs, "", options))
        .replace(/<a:pPr\b([^>]*)>([\s\S]*?)<\/a:pPr>/, (_match, attrs: string, body: string) => paragraphPropertiesXml(attrs, body, options));
    }
    return paragraph.replace(/(<a:p\b[^>]*>)/, `$1${paragraphPropertiesXml("", "", options)}`);
  });
}

function paragraphPropertiesXml(
  attrs: string,
  body: string,
  options: { level?: number; numbering?: boolean; startAt?: number; lineSpacing?: number; spaceBefore?: number }
): string {
  let nextAttrs = attrs;
  if (options.level !== undefined) nextAttrs = upsertXmlAttr(nextAttrs, "lvl", String(Math.max(0, Math.min(8, Math.round(options.level)))));
  let nextBody = body;
  const existingLineSpacing = extractFirstParagraphChild(nextBody, "a:lnSpc");
  nextBody = removeFirstParagraphChild(nextBody, "a:lnSpc");
  const existingSpaceBefore = extractFirstParagraphChild(nextBody, "a:spcBef");
  nextBody = removeFirstParagraphChild(nextBody, "a:spcBef");
  const existingBullet = extractBulletParagraphChild(nextBody);
  nextBody = removeBulletParagraphChildren(nextBody);
  let lineSpacing = existingLineSpacing;
  if (options.lineSpacing !== undefined) {
    const val = Math.round((options.lineSpacing <= 10 ? options.lineSpacing * 100000 : options.lineSpacing * 1000));
    lineSpacing = `<a:lnSpc><a:spcPct val="${val}"/></a:lnSpc>`;
  }
  let spaceBefore = existingSpaceBefore;
  if (options.spaceBefore !== undefined) {
    spaceBefore = `<a:spcBef><a:spcPts val="${Math.round(options.spaceBefore * 100)}"/></a:spcBef>`;
  }
  let bullet = existingBullet;
  if (options.numbering) {
    bullet = `<a:buAutoNum type="arabicPeriod"${options.startAt ? ` startAt="${Math.max(1, Math.round(options.startAt))}"` : ""}/>`;
  }
  nextBody = `${lineSpacing ?? ""}${spaceBefore ?? ""}${bullet ?? ""}${nextBody}`;
  return `<a:pPr${nextAttrs ? ` ${nextAttrs.trim()}` : ""}>${nextBody}</a:pPr>`;
}

function firstXmlElement(body: string, tag: string): string | undefined {
  const escaped = escapeRegExp(tag);
  const selfClosing = new RegExp(`<${escaped}\\b[^>]*/>`);
  const selfClosingMatch = selfClosing.exec(body)?.[0];
  if (selfClosingMatch) return selfClosingMatch;
  return new RegExp(`<${escaped}\\b[^>]*>[\\s\\S]*?<\\/${escaped}>`).exec(body)?.[0];
}

function extractFirstParagraphChild(body: string, tag: string): string | undefined {
  return firstXmlElement(body, tag);
}

function removeFirstParagraphChild(body: string, tag: string): string {
  const child = firstXmlElement(body, tag);
  return child ? body.replace(child, "") : body;
}

function extractBulletParagraphChild(body: string): string | undefined {
  return /<a:bu(?:Char|AutoNum|None)\b[\s\S]*?\/>/.exec(body)?.[0] ?? /<a:buBlip\b[\s\S]*?<\/a:buBlip>/.exec(body)?.[0];
}

function removeBulletParagraphChildren(body: string): string {
  return body.replace(/<a:bu(?:Char|AutoNum|None)\b[\s\S]*?\/>/g, "").replace(/<a:buBlip\b[\s\S]*?<\/a:buBlip>/g, "");
}

function applyPptxTextCase(shape: string, textCase: "upper" | "lower" | "title" | "sentence"): string {
  return shape.replace(/<a:t(?:\s[^>]*)?>([\s\S]*?)<\/a:t>/g, (match, text: string) => match.replace(text, escapeXmlText(convertTextCase(decodeXmlText(text), textCase))));
}

function convertTextCase(value: string, textCase: "upper" | "lower" | "title" | "sentence"): string {
  if (textCase === "upper") return value.toLocaleUpperCase();
  if (textCase === "lower") return value.toLocaleLowerCase();
  if (textCase === "title") return value.replace(/\p{L}[\p{L}\p{N}'-]*/gu, (word) => word.charAt(0).toLocaleUpperCase() + word.slice(1).toLocaleLowerCase());
  const lower = value.toLocaleLowerCase();
  return lower.replace(/(^\s*\p{L}|[.!?]\s+\p{L})/gu, (letter) => letter.toLocaleUpperCase());
}

function decodeXmlText(value: string): string {
  return value.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, "\"").replace(/&apos;/g, "'").replace(/&amp;/g, "&");
}

function upsertXmlAttr(attrs: string, name: string, value: string): string {
  const escaped = escapeRegExp(name);
  const pattern = new RegExp(`(?:^|\\s)${escaped}="[^"]*"`);
  const attr = `${name}="${escapeXmlText(value)}"`;
  return pattern.test(attrs) ? attrs.replace(pattern, ` ${attr}`).trim() : `${attrs.trim()} ${attr}`.trim();
}

function notesSlideXml(text: string): string {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">',
    '<p:cSld><p:spTree>',
    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>',
    '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>',
    `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes Placeholder"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:txBody><a:bodyPr/><a:lstStyle/>${text.split(/\r?\n/).map((line) => `<a:p><a:r><a:t>${escapeXmlText(line)}</a:t></a:r></a:p>`).join("")}</p:txBody></p:sp>`,
    '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>'
  ].join("");
}

function extractNotesText(xml: string): string {
  return [...xml.matchAll(/<a:t\b[^>]*>([\s\S]*?)<\/a:t>/g)].map((match) => match[1] ?? "").join("\n");
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

async function refreshAllPivotDefinitions(zip: Awaited<ReturnType<typeof loadZip>>): Promise<boolean> {
  let changed = false;
  for (const path of Object.keys(zip.files).filter((item) => /^xl\/pivotTables\/pivotTable\d+\.xml$/i.test(item))) {
    const xml = (await readZipText(zip, path)) ?? "";
    const next = xml.replace(/<pivotTableDefinition\b([^>]*)>/, (match, attrs: string) =>
      /\brefreshOnLoad=/.test(attrs) ? match.replace(/\brefreshOnLoad="[^"]*"/, 'refreshOnLoad="1"') : `<pivotTableDefinition${attrs} refreshOnLoad="1">`
    );
    if (next !== xml) {
      zip.file(path, next);
      changed = true;
    }
  }
  for (const path of Object.keys(zip.files).filter((item) => /^xl\/pivotCache\/pivotCacheDefinition\d+\.xml$/i.test(item))) {
    const xml = (await readZipText(zip, path)) ?? "";
    const next = xml.replace(/<pivotCacheDefinition\b([^>]*)>/, (match, attrs: string) =>
      /\brefreshOnLoad=/.test(attrs) ? match.replace(/\brefreshOnLoad="[^"]*"/, 'refreshOnLoad="1"') : `<pivotCacheDefinition${attrs} refreshOnLoad="1">`
    );
    if (next !== xml) {
      zip.file(path, next);
      changed = true;
    }
  }
  return changed;
}

async function editXlsxSlicerSelection(
  zip: Awaited<ReturnType<typeof loadZip>>,
  operation: { selector: EditSelector; selected: string[] },
  objectMap: ObjectMapEntry[]
): Promise<boolean> {
  const target = singleMatch(objectMap, operation.selector);
  if (target.kind !== "slicer" || !target.xmlPath) throw new Error("SELECTOR_NOT_FOUND: selected object is not an XLSX slicer.");
  const xml = (await readZipText(zip, target.xmlPath)) ?? "";
  const selected = new Set(operation.selected.map(String));
  let touched = false;
  const next = xml.replace(/<[^:>]*:?(?:slicerItem|item)\b([^>]*)>/g, (match, attrs: string) => {
    const value = /\b(?:n|x|name)="([^"]+)"/.exec(attrs)?.[1];
    if (!value) return match;
    touched = true;
    const hidden = selected.has(value) ? "0" : "1";
    return /\bh="/.test(match) ? match.replace(/\bh="[^"]*"/, `h="${hidden}"`) : match.replace(/\/?>$/, ` h="${hidden}"/>`);
  });
  if (touched && next !== xml) zip.file(target.xmlPath, next);
  return touched && next !== xml;
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

function validationFailures(selectorResult: ResolveEditSelectorsResult | undefined, minSelectorConfidence = 0): EditOperationResult[] {
  return (selectorResult?.resolutions ?? [])
    .filter((resolution) => resolution.reason === "not-found" || resolution.reason === "ambiguous" || resolution.reason === "low-confidence" || ((resolution.confidence ?? 1) < minSelectorConfidence))
    .map((resolution) => ({
      operationIndex: resolution.operationIndex,
      op: "selector",
      applied: false,
      reason: resolution.reason === "ambiguous" ? "ambiguous" : resolution.reason === "not-found" ? "not-found" : "low-confidence",
      message: resolution.reason === "ambiguous"
        ? `SELECTOR_AMBIGUOUS: selector matched ${resolution.matchCount} objects.`
        : resolution.reason === "not-found"
          ? selectorNotFoundMessage(resolution)
          : `SELECTOR_LOW_CONFIDENCE: selector confidence ${resolution.confidence ?? 0} is below required ${minSelectorConfidence}.`,
      diagnostics: resolution.diagnostics
    }));
}

function selectorNotFoundMessage(resolution: EditSelectorResolution): string {
  const near = resolution.diagnostics?.find((diagnostic) => diagnostic.code === "SELECTOR_NEAR_WHITESPACE_INSENSITIVE_MATCH");
  const candidate = near?.candidates[0];
  if (!near || !candidate) return "SELECTOR_NOT_FOUND: selector matched no objects.";
  return [
    "SELECTOR_NOT_FOUND: selector matched no objects.",
    "Near whitespace/newline-insensitive text candidate found, but atomic validation did not auto-select it.",
    `Suggested stableObjectId selector: ${candidate.stableObjectId}.`
  ].join(" ");
}

function stalePlanFailures(selectorResult: ResolveEditSelectorsResult | undefined, options: EditOptions, operations: EditOperation[]): EditOperationResult[] {
  if (!selectorResult) return [];
  const failures: EditOperationResult[] = [];
  if (options.expectedInputSha256 && options.expectedInputSha256 !== selectorResult.inputSha256) {
    failures.push({
      operationIndex: -1,
      op: "stalePlan",
      applied: false,
      reason: "stale-plan",
      message: `EDIT_STALE_PLAN: blocked before write because expectedInputSha256 ${options.expectedInputSha256} does not match current ${selectorResult.inputSha256}.`,
      evidence: {
        code: "EDIT_STALE_PLAN",
        field: "expectedInputSha256",
        expected: options.expectedInputSha256,
        current: selectorResult.inputSha256,
        expectedHash: options.expectedInputSha256,
        currentHash: selectorResult.inputSha256,
        operationCount: operations.length,
        wouldWrite: false
      }
    });
  }
  if (options.expectedObjectMapHash && options.expectedObjectMapHash !== selectorResult.objectMapHash) {
    failures.push({
      operationIndex: -1,
      op: "stalePlan",
      applied: false,
      reason: "stale-plan",
      message: `EDIT_STALE_PLAN: blocked before write because objectMapHash stale mismatch: expected ${options.expectedObjectMapHash} does not match current ${selectorResult.objectMapHash}.`,
      evidence: {
        code: "EDIT_STALE_PLAN",
        field: "expectedObjectMapHash",
        expected: options.expectedObjectMapHash,
        current: selectorResult.objectMapHash,
        expectedHash: options.expectedObjectMapHash,
        currentHash: selectorResult.objectMapHash,
        operationCount: operations.length,
        wouldWrite: false
      }
    });
  }
  if (options.expectedObjectGraphHash && options.expectedObjectGraphHash !== selectorResult.objectGraphHash) {
    failures.push({
      operationIndex: -1,
      op: "stalePlan",
      applied: false,
      reason: "stale-plan",
      message: `EDIT_STALE_PLAN: blocked before write because objectGraphHash stale mismatch: expected ${options.expectedObjectGraphHash} does not match current ${selectorResult.objectGraphHash}.`,
      evidence: {
        code: "EDIT_STALE_PLAN",
        field: "expectedObjectGraphHash",
        expected: options.expectedObjectGraphHash,
        current: selectorResult.objectGraphHash,
        expectedHash: options.expectedObjectGraphHash,
        currentHash: selectorResult.objectGraphHash,
        operationCount: operations.length,
        wouldWrite: false
      }
    });
  }
  if (options.selectionLock) {
    const lock = options.selectionLock;
    if (lock.objectGraphHash !== selectorResult.objectGraphHash) {
      failures.push({
        operationIndex: -1,
        op: "selectionLock",
        applied: false,
        reason: "stale-plan",
        message: `EDIT_STALE_SELECTION_LOCK: blocked before write because selectionLock.objectGraphHash ${lock.objectGraphHash} does not match current ${selectorResult.objectGraphHash}.`,
        evidence: {
          code: "EDIT_STALE_SELECTION_LOCK",
          field: "selectionLock.objectGraphHash",
          expected: lock.objectGraphHash,
          current: selectorResult.objectGraphHash,
          expectedHash: lock.objectGraphHash,
          currentHash: selectorResult.objectGraphHash,
          operationCount: operations.length,
          wouldWrite: false
        }
      });
    } else if (lock.nodeId || lock.sourceFingerprint) {
      const currentLock = selectorResult.resolutions.find((resolution) =>
        (!lock.nodeId || resolution.selectionLock?.nodeId === lock.nodeId) &&
        (!lock.sourceFingerprint || resolution.selectionLock?.sourceFingerprint === lock.sourceFingerprint)
      )?.selectionLock;
      if (!currentLock) {
        failures.push({
          operationIndex: -1,
          op: "selectionLock",
          applied: false,
          reason: "stale-plan",
          message: "EDIT_STALE_SELECTION_LOCK: blocked before write because selectionLock nodeId/sourceFingerprint no longer matches any resolved selector.",
          evidence: {
            code: "EDIT_STALE_SELECTION_LOCK",
            field: "selectionLock",
            expected: JSON.stringify(lock),
            current: JSON.stringify(selectorResult.resolutions.map((resolution) => resolution.selectionLock).filter(Boolean)),
            expectedHash: lock.sourceFingerprint,
            operationCount: operations.length,
            wouldWrite: false
          }
        });
      }
    }
  }
  return failures.length ? failures.map((failure) => ({ ...failure, message: `${failure.message} ${operations.length} operations were not applied.` })) : [];
}

async function inspectCurrentObjectMap(
  zip: Awaited<ReturnType<typeof loadZip>>,
  format: "pptx" | "docx" | "xlsx",
  config?: OfficegenConfig
): Promise<ObjectMapEntry[]> {
  const bytes = await zipToBytes(zip);
  const inspected = await inspect({ data: bytes, format }, { config });
  return inspected.objectMap;
}

function sha256Bytes(bytes: Uint8Array): string {
  return `sha256:${createHash("sha256").update(bytes).digest("hex")}`;
}

function objectMapHash(objectMap: ObjectMapEntry[]): string {
  return `sha256:${createHash("sha256").update(stableStringify(objectMap.map((entry) => ({
    stableObjectId: entry.stableObjectId,
    kind: entry.kind,
    label: entry.label,
    text: entry.text ?? entry.textPreview,
    bbox: entry.bbox,
    selectorHints: entry.selectorHints
  })))).digest("hex")}`;
}

async function buildPatchPlan(
  format: string,
  inputBytes: Uint8Array,
  operations: EditOperation[],
  opResults: EditOperationResult[],
  selectorResult?: ResolveEditSelectorsResult,
  transaction?: EditTransaction<OfficePartValue>,
  store?: EditPartStore<OfficePartValue>,
  initialPartPaths = new Set<string>(),
  currentPartPaths = new Set<string>()
): Promise<PatchPlan> {
  const touchedParts = transaction && store
    ? await patchPlanTouchedParts(transaction, store, initialPartPaths, currentPartPaths)
    : [];
  const sourceFingerprints = touchedParts
    .map((part) => part.sourceFingerprint)
    .filter((fingerprint): fingerprint is EditSourceFingerprint => Boolean(fingerprint));
  return {
    schema: "officegen.patchPlan@2",
    format,
    wouldWrite: false,
    inputSha256: selectorResult?.inputSha256 ?? sha256Bytes(inputBytes),
    objectMapHash: selectorResult?.objectMapHash,
    objectGraphHash: selectorResult?.objectGraphHash,
    sourceFingerprint: inputSourceFingerprint(inputBytes),
    operations: operations.map((operation, index) => {
      const result = opResults.find((item) => item.operationIndex === index);
      return {
        operationIndex: index,
        op: operationName(operation),
        wouldApply: result?.applied === true,
        reason: result?.reason,
        message: result?.message,
        selector: selectorForOperation(operation)
      };
    }),
    touchedParts,
    expectedChangedParts: touchedParts.map((part) => part.path),
    sourceFingerprints,
    blocked: opResults.filter((result) => result.applied === false && result.reason !== "idempotency-replay")
  };
}

async function patchPlanTouchedParts(
  transaction: EditTransaction<OfficePartValue>,
  store: EditPartStore<OfficePartValue>,
  initialPartPaths: Set<string>,
  currentPartPaths: Set<string>
): Promise<PatchPlanTouchedPart[]> {
  const byPath = new Map<string, PatchPlanTouchedPart>();
  for (const entry of transaction.snapshot()) {
    const current = await store.readPart(entry.path);
    if (entry.existed && current === undefined) {
      byPath.set(entry.path, {
        path: entry.path,
        change: "deleted",
        beforeSha256: sha256Part(entry.value as OfficePartValue),
        sourceFingerprint: partSourceFingerprint(entry.path, entry.value as OfficePartValue)
      });
      continue;
    }
    if (!entry.existed && current !== undefined) {
      byPath.set(entry.path, {
        path: entry.path,
        change: "created",
        afterSha256: sha256Part(current)
      });
      continue;
    }
    if (entry.existed && current !== undefined) {
      const beforeSha256 = sha256Part(entry.value as OfficePartValue);
      const afterSha256 = sha256Part(current);
      if (beforeSha256 !== afterSha256) {
        byPath.set(entry.path, {
          path: entry.path,
          change: "modified",
          beforeSha256,
          afterSha256,
          sourceFingerprint: partSourceFingerprint(entry.path, entry.value as OfficePartValue)
        });
      }
    }
  }
  for (const path of currentPartPaths) {
    if (initialPartPaths.has(path) || byPath.has(path)) continue;
    const current = await store.readPart(path);
    if (current === undefined) continue;
    byPath.set(path, {
      path,
      change: "created",
      afterSha256: sha256Part(current)
    });
  }
  return [...byPath.values()].sort((left, right) => left.path.localeCompare(right.path));
}

function inputSourceFingerprint(bytes: Uint8Array): EditSourceFingerprint {
  return {
    algorithm: "sha256",
    hash: sha256Digest(bytes),
    byteLength: bytes.byteLength
  };
}

function partSourceFingerprint(path: string, value: OfficePartValue): EditSourceFingerprint {
  const bytes = partBytes(value);
  return {
    algorithm: "sha256",
    hash: sha256Digest(bytes),
    byteLength: bytes.byteLength,
    path
  };
}

function sha256Part(value: OfficePartValue): string {
  return `sha256:${sha256Digest(partBytes(value))}`;
}

function sha256Digest(bytes: Uint8Array): string {
  return createHash("sha256").update(bytes).digest("hex");
}

function partBytes(value: OfficePartValue): Uint8Array {
  return typeof value === "string" ? Buffer.from(value, "utf8") : value;
}

function stableStringify(value: unknown): string {
  if (value === null || typeof value !== "object") return JSON.stringify(value);
  if (Array.isArray(value)) return `[${value.map(stableStringify).join(",")}]`;
  const record = value as Record<string, unknown>;
  return `{${Object.keys(record).sort().map((key) => `${JSON.stringify(key)}:${stableStringify(record[key])}`).join(",")}}`;
}

function hasSelectorAfterPptxCreator(operations: EditOperation[]): boolean {
  return operations.some((operation, index) => Boolean(selectorForOperation(operation)) && hasPriorPptxCreator(operations, index));
}

function hasPriorPptxCreator(operations: EditOperation[], index: number): boolean {
  const selector = selectorForOperation(operations[index] as EditOperation);
  return operations.slice(0, index).some((operation) => {
    const op = operationName(operation);
    if (op === "pptx.addSlide" || op === "pptx.addTextbox") return true;
    return op === "pptx.duplicateSlide" && selectorTargetsExplicitSlide(selector);
  });
}

function selectorTargetsExplicitSlide(selector: EditSelector | undefined): boolean {
  if (!selector) return false;
  if (selector.slide !== undefined || selector.nearestTo?.slide !== undefined || selector.nthBodyShape?.slide !== undefined) return true;
  if (typeof selector.rightOf === "object" && selector.rightOf.slide !== undefined) return true;
  return typeof selector.largestTextOnSlide === "number";
}

function editAbortResult(
  format: string,
  skipped: number,
  opResults: EditOperationResult[],
  operations: EditOperation[],
  caveats: string[],
  selectorResult?: ResolveEditSelectorsResult,
  dryRun?: boolean,
  inputBytes?: Uint8Array
): EditResult {
  const errors = opResults.filter(isRequiredEditFailure);
  const rolledBack = opResults.some((result) => result.applied);
  const inputSha256 = inputBytes ? (selectorResult?.inputSha256 ?? sha256Bytes(inputBytes)) : selectorResult?.inputSha256;
  const sourceFingerprint = inputBytes ? inputSourceFingerprint(inputBytes) : undefined;
  return {
    schema: "officegen.edit.result@1.2",
    format,
    dryRun,
    inputSha256,
    objectMapHash: selectorResult?.objectMapHash,
    objectGraphHash: selectorResult?.objectGraphHash,
    sourceFingerprint,
    rolledBack: rolledBack || undefined,
    changed: false,
    applied: 0,
    skipped,
    opResults,
    errors: errors.length ? errors : undefined,
    patchPlan: dryRun && inputBytes ? {
      schema: "officegen.patchPlan@2",
      format,
      wouldWrite: false,
      inputSha256: inputSha256 ?? sha256Bytes(inputBytes),
      objectMapHash: selectorResult?.objectMapHash,
      objectGraphHash: selectorResult?.objectGraphHash,
      sourceFingerprint: inputSourceFingerprint(inputBytes),
      operations: operations.map((operation, index) => {
        const result = opResults.find((item) => item.operationIndex === index);
        return {
          operationIndex: index,
          op: operationName(operation),
          wouldApply: result?.applied === true,
          reason: result?.reason,
          message: result?.message,
          selector: selectorForOperation(operation)
        };
      }),
      touchedParts: [],
      expectedChangedParts: [],
      sourceFingerprints: [],
      blocked: opResults.filter((result) => result.applied === false)
    } : undefined,
    caveats
  };
}

function isRequiredEditFailure(result: EditOperationResult): boolean {
  if (result.applied !== false) return false;
  return result.reason === "not-found"
    || result.reason === "ambiguous"
    || result.reason === "low-confidence"
    || result.reason === "unsupported"
    || result.reason === "validation-failed"
    || result.reason === "skipped-after-error"
    || result.reason === "stale-plan";
}

function appendSkippedAfterErrorResults(opResults: EditOperationResult[], operations: EditOperation[], message: string): number {
  const existing = new Set(opResults.map((result) => result.operationIndex));
  let added = 0;
  for (const [index, operation] of operations.entries()) {
    if (existing.has(index)) continue;
    opResults.push({
      operationIndex: index,
      op: operationName(operation),
      applied: false,
      reason: "skipped-after-error",
      message
    });
    added += 1;
  }
  opResults.sort((left, right) => left.operationIndex - right.operationIndex);
  return added;
}

function selectorForOperation(operation: EditOperation): EditSelector | undefined {
  if ("selector" in operation) return operation.selector;
  return undefined;
}

function selectorResolutionForObjectMap(
  operationIndex: number,
  selector: EditSelector,
  objectMap: ObjectMapEntry[],
  graph: ObjectGraph,
  currentObjectGraphHash: string,
  operation?: string
): EditSelectorResolution {
  const matches = resolveMatches(objectMap, selector, matchOptionsForOperation(operation));
  const diagnostics = matches.length ? [] : selectorNearCandidateDiagnostics(objectMap, selector, matchOptionsForOperation(operation));
  const graphNodesByStableId = new Map(graph.nodes.map((node) => [node.stableId, node]));
  const confidence = matches.length === 1 ? selectorConfidence(matches[0] as ObjectMapEntry, selector, matches) : undefined;
  const status = selectorStatus(matches.length, confidence);
  const reason = status === "not_found"
    ? "not-found"
    : status === "ambiguous"
      ? "ambiguous"
      : status === "low_confidence"
        ? "low-confidence"
        : undefined;
  const matchedNode = matches.length === 1 ? graphNodesByStableId.get(matches[0]?.stableObjectId ?? "") : undefined;
  const selectorResolution = selectorResolutionV2ForObjectMap(selector, matches, graphNodesByStableId, graph, currentObjectGraphHash, status, confidence);
  const suggestions = selectorDiagnosticSuggestions(diagnostics);
  if (suggestions.length) selectorResolution.nextActions = [...selectorResolution.nextActions, ...suggestions];
  return {
    operationIndex,
    selector,
    stableObjectId: selector.stableObjectId,
    matched: status === "matched",
    matchCount: matches.length,
    status,
    confidence,
    matches: matches.map((match) => selectorMatch(match, selector, matches, graphNodesByStableId)),
    evidence: selectorResolution.evidence,
    ambiguityReason: status === "ambiguous" ? "multiple-matches" : undefined,
    nextActions: selectorResolution.nextActions,
    selectionLock: selectionLockForNode(graph, matchedNode),
    selectorResolution,
    reason,
    diagnostics: diagnostics.length ? diagnostics : undefined,
    suggestions: suggestions.length ? suggestions : undefined
  };
}

function selectorStatus(matchCount: number, confidence: number | undefined): SelectorResolutionV2Status {
  if (matchCount === 0) return "not_found";
  if (matchCount > 1) return "ambiguous";
  if (confidence !== undefined && confidence < SELECTOR_GRAPH_LOW_CONFIDENCE_THRESHOLD) return "low_confidence";
  return "matched";
}

function selectorResolutionV2ForObjectMap(
  selector: EditSelector,
  matches: ObjectMapEntry[],
  graphNodesByStableId: Map<string, ObjectGraphNode>,
  graph: ObjectGraph,
  currentObjectGraphHash: string,
  status: SelectorResolutionV2Status,
  confidence: number | undefined
): SelectorResolutionV2 {
  const matchedNode = matches.length === 1 ? graphNodesByStableId.get(matches[0]?.stableObjectId ?? "") : undefined;
  const evidence = matches.flatMap((match) => graphNodesByStableId.get(match.stableObjectId)?.evidence ?? []);
  return {
    schema: "officegen.selectorResolution@2",
    status,
    confidence,
    candidates: matches.map((match) => {
      const node = graphNodesByStableId.get(match.stableObjectId);
      return {
        nodeId: node?.nodeId,
        stableObjectId: match.stableObjectId,
        type: match.kind,
        label: match.label,
        text: match.text,
        confidence: selectorConfidence(match, selector, matches),
        source: node?.source
      };
    }),
    evidence,
    ambiguityReason: status === "ambiguous" ? "multiple-matches" : undefined,
    nextActions: selectorResolutionNextActions(status),
    selectionLock: matchedNode
      ? {
          objectGraphHash: currentObjectGraphHash,
          nodeId: matchedNode.nodeId,
          sourceFingerprint: sourceFingerprintForNode(matchedNode)
        }
      : selectionLockForNode(graph, undefined)
  };
}

interface ResolveMatchOptions {
  kinds?: string[];
}

function matchOptionsForOperation(operation: string | undefined): ResolveMatchOptions | undefined {
  if (operation === "pptx.setBold" || operation === "pptx.setFontSize" || operation === "pptx.setTextCase") return { kinds: ["shape"] };
  if (operation === "pptx.setBulletLevel" || operation === "pptx.setNumbering" || operation === "pptx.setLineSpacing" || operation === "pptx.setSpaceBefore") return { kinds: ["shape"] };
  return undefined;
}

function resolveMatches(objectMap: ObjectMapEntry[], selector: EditSelector, options: ResolveMatchOptions = {}): ObjectMapEntry[] {
  if (selector.stableObjectId) return objectMap.filter((entry) => entry.stableObjectId === selector.stableObjectId && matchesResolveOptions(entry, options));
  let candidates = objectMap;
  if (options.kinds?.length) candidates = candidates.filter((entry) => matchesResolveOptions(entry, options));
  if (selector.slide !== undefined) candidates = candidates.filter((entry) => Number(entry.selectorHints?.slide) === selector.slide);
  if (selector.shapeId) candidates = candidates.filter((entry) => String(entry.selectorHints?.shapeId ?? "") === selector.shapeId);
  if (selector.shapeName) candidates = candidates.filter((entry) => entry.label === selector.shapeName || entry.selectorHints?.shapeName === selector.shapeName || entry.selectorHints?.name === selector.shapeName);
  if (selector.sourcePath) candidates = candidates.filter((entry) => entry.sourcePath === selector.sourcePath || entry.selectorHints?.sourcePath === selector.sourcePath);
  if (selector.xmlPath) candidates = candidates.filter((entry) => entry.xmlPath === selector.xmlPath || entry.selectorHints?.xmlPath === selector.xmlPath);
  if (selector.page !== undefined) candidates = candidates.filter((entry) => Number(entry.selectorHints?.page ?? entry.selectorHints?.slide ?? entry.selectorHints?.sheet) === selector.page);
  if (selector.story) candidates = candidates.filter((entry) => entry.selectorHints?.story === selector.story || entry.selectorHints?.partKind === selector.story);
  if (selector.paragraph !== undefined) candidates = candidates.filter((entry) => Number(entry.selectorHints?.paragraph) === selector.paragraph);
  if (selector.table !== undefined) candidates = candidates.filter((entry) => Number(entry.selectorHints?.table ?? entry.selectorHints?.tableIndex) === selector.table);
  if (selector.row !== undefined) candidates = candidates.filter((entry) => Number(entry.selectorHints?.row) === selector.row);
  if (selector.column !== undefined) candidates = candidates.filter((entry) => Number(entry.selectorHints?.column) === selector.column);
  if (selector.range) candidates = candidates.filter((entry) => String(entry.selectorHints?.range ?? entry.selectorHints?.ref ?? "") === selector.range);
  if (selector.relationshipId) candidates = candidates.filter((entry) => entry.selectorHints?.relationshipId === selector.relationshipId || entry.media?.relationshipId === selector.relationshipId);
  if (selector.assetPath) candidates = candidates.filter((entry) => entry.selectorHints?.assetPath === selector.assetPath || entry.media?.assetPath === selector.assetPath);
  if (selector.commentId) candidates = candidates.filter((entry) => String(entry.selectorHints?.commentId ?? "") === selector.commentId);
  if (selector.revisionId) candidates = candidates.filter((entry) => String(entry.selectorHints?.revisionId ?? "") === selector.revisionId);
  const placeholderKey = selector.placeholderKey ?? selector.placeholder;
  if (placeholderKey) candidates = candidates.filter((entry) => entry.selectorHints?.placeholderKey === placeholderKey || entry.selectorHints?.placeholder === placeholderKey);
  if (selector.textHash) candidates = candidates.filter((entry) => entry.selectorHints?.textHash === selector.textHash);
  if (selector.positionHash) candidates = candidates.filter((entry) => entry.selectorHints?.positionHash === selector.positionHash);
  if (selector.contentControlTag) candidates = candidates.filter((entry) => entry.selectorHints?.contentControlTag === selector.contentControlTag || entry.selectorHints?.tag === selector.contentControlTag);
  if (selector.namedRange) candidates = candidates.filter((entry) => entry.selectorHints?.namedRange === selector.namedRange || entry.label === selector.namedRange);
  if (selector.sheetName) candidates = candidates.filter((entry) => {
    const expected = selector.sheetName?.toLowerCase();
    const hinted = String(entry.selectorHints?.sheetName ?? "").toLowerCase();
    const fallback = entry.selectorHints?.sheet ? `sheet${entry.selectorHints.sheet}`.toLowerCase() : "";
    return hinted === expected || fallback === expected;
  });
  if (selector.cell) candidates = candidates.filter((entry) => String(entry.selectorHints?.cell ?? entry.label ?? "").toUpperCase() === selector.cell?.toUpperCase());
  if (selector.tableName) candidates = candidates.filter((entry) => String(entry.selectorHints?.tableName ?? entry.label ?? "") === selector.tableName);
  if (selector.chartPath) candidates = candidates.filter((entry) => String(entry.selectorHints?.chartPath ?? entry.xmlPath ?? "") === selector.chartPath);
  const text = selector.textMatch?.text ?? selector.contains;
  if (text) candidates = candidates.filter((entry) => selector.textMatch?.exact ? entry.text === text : entry.text?.includes(text));
  if (selector.nearestTo) return nearestMatches(candidates, selector);
  if (selector.rightOf) return rightOfMatches(objectMap, candidates, selector);
  if (selector.largestTextOnSlide) return largestTextMatches(candidates, selector);
  if (selector.nthBodyShape) return nthBodyShapeMatches(candidates, selector.nthBodyShape);
  return candidates === objectMap ? [] : candidates;
}

function singleMatch(objectMap: ObjectMapEntry[], selector: EditSelector, options?: ResolveMatchOptions): ObjectMapEntry {
  const matches = resolveMatches(objectMap, selector, options);
  if (!matches.length) throw new Error("SELECTOR_NOT_FOUND: selector matched no objects.");
  if (matches.length > 1) throw new Error(`SELECTOR_AMBIGUOUS: selector matched ${matches.length} objects.`);
  return matches[0] as ObjectMapEntry;
}

function matchesResolveOptions(entry: ObjectMapEntry, options: ResolveMatchOptions): boolean {
  return !options.kinds?.length || options.kinds.includes(entry.kind);
}

function selectorNearCandidateDiagnostics(objectMap: ObjectMapEntry[], selector: EditSelector, options: ResolveMatchOptions = {}): EditSelectorDiagnostic[] {
  const requestedText = selector.textMatch?.text ?? selector.contains;
  if (!requestedText || compactSelectorText(requestedText).length === 0) return [];
  const selectorField: EditSelectorDiagnostic["selectorField"] = selector.textMatch ? "textMatch" : "contains";
  const scoped = resolveMatchesWithoutText(objectMap, selector, options);
  const candidates = scoped
    .filter((entry) => whitespaceInsensitiveTextMatches(entry.text, requestedText, selector.textMatch?.exact === true))
    .slice(0, 5)
    .map((entry) => nearCandidate(entry));
  if (!candidates.length) return [];
  return [{
    code: "SELECTOR_NEAR_WHITESPACE_INSENSITIVE_MATCH",
    severity: "info",
    message: "A candidate would match if spaces/newlines were ignored. It was not selected automatically; use stableObjectId or a literal selector that matches the rendered text.",
    selectorField,
    requestedText,
    normalizedRequestedText: compactSelectorText(requestedText),
    candidates
  }];
}

function resolveMatchesWithoutText(objectMap: ObjectMapEntry[], selector: EditSelector, options: ResolveMatchOptions = {}): ObjectMapEntry[] {
  const matches = resolveMatches(objectMap, { ...selector, contains: undefined, textMatch: undefined }, options);
  if (matches.length || selectorHasNonTextCriterion(selector)) return matches;
  return objectMap.filter((entry) => matchesResolveOptions(entry, options));
}

function selectorHasNonTextCriterion(selector: EditSelector): boolean {
  return selector.stableObjectId !== undefined ||
    selector.slide !== undefined ||
    selector.shapeId !== undefined ||
    selector.placeholderKey !== undefined ||
    selector.placeholder !== undefined ||
    selector.shapeName !== undefined ||
    selector.contentControlTag !== undefined ||
    selector.namedRange !== undefined ||
    selector.sheetName !== undefined ||
    selector.cell !== undefined ||
    selector.tableName !== undefined ||
    selector.chartPath !== undefined ||
    selector.textHash !== undefined ||
    selector.positionHash !== undefined ||
    selector.sourcePath !== undefined ||
    selector.xmlPath !== undefined ||
    selector.page !== undefined ||
    selector.story !== undefined ||
    selector.paragraph !== undefined ||
    selector.table !== undefined ||
    selector.row !== undefined ||
    selector.column !== undefined ||
    selector.range !== undefined ||
    selector.relationshipId !== undefined ||
    selector.assetPath !== undefined ||
    selector.commentId !== undefined ||
    selector.revisionId !== undefined ||
    selector.nearestTo !== undefined ||
    selector.rightOf !== undefined ||
    selector.largestTextOnSlide !== undefined ||
    selector.nthBodyShape !== undefined;
}

function whitespaceInsensitiveTextMatches(actual: string | undefined, expected: string, exact: boolean): boolean {
  const compactActual = compactSelectorText(actual);
  const compactExpected = compactSelectorText(expected);
  if (!compactActual || !compactExpected) return false;
  return exact ? compactActual === compactExpected : compactActual.includes(compactExpected);
}

function nearCandidate(entry: ObjectMapEntry): EditSelectorNearCandidate {
  return {
    stableObjectId: entry.stableObjectId,
    kind: entry.kind,
    label: entry.label,
    text: entry.text,
    textPreview: entry.textPreview,
    sourcePath: entry.sourcePath,
    xmlPath: entry.xmlPath,
    selectorHints: entry.selectorHints,
    suggestedSelector: { stableObjectId: entry.stableObjectId }
  };
}

function selectorDiagnosticSuggestions(diagnostics: EditSelectorDiagnostic[]): string[] {
  const near = diagnostics.find((diagnostic) => diagnostic.code === "SELECTOR_NEAR_WHITESPACE_INSENSITIVE_MATCH");
  const candidate = near?.candidates[0];
  if (!candidate) return [];
  return [
    `Whitespace/newline-insensitive near match found. Validation was not relaxed; retry with selector stableObjectId '${candidate.stableObjectId}' or a literal text selector that includes the newline/space characters.`
  ];
}

function compactSelectorText(value: string | undefined): string {
  return String(value ?? "").replace(/\s+/g, "").toLowerCase();
}

function selectorMatch(
  entry: ObjectMapEntry,
  selector: EditSelector,
  matches: ObjectMapEntry[],
  graphNodesByStableId: Map<string, ObjectGraphNode>
): EditSelectorResolution["matches"][number] {
  const node = graphNodesByStableId.get(entry.stableObjectId);
  return {
    nodeId: node?.nodeId,
    stableObjectId: entry.stableObjectId,
    kind: entry.kind,
    confidence: selectorConfidence(entry, selector, matches),
    label: entry.label,
    text: entry.text,
    sourcePath: entry.sourcePath,
    xmlPath: entry.xmlPath,
    selectorHints: entry.selectorHints
  };
}

function nearestMatches(objectMap: ObjectMapEntry[], selector: EditSelector): ObjectMapEntry[] {
  const point = selector.nearestTo;
  if (!point) return [];
  const ranked = objectMap
    .filter((entry) => entry.bounds && (point.slide === undefined || Number(entry.selectorHints?.slide) === point.slide))
    .map((entry) => ({ entry, distance: centerDistance(entry, point.x, point.y) }))
    .sort((left, right) => left.distance - right.distance);
  const best = ranked[0];
  if (!best || best.distance > 1000) return [];
  const second = ranked[1];
  if (second && (second.distance - best.distance <= 24 || second.distance <= best.distance * 1.15)) {
    return ranked.filter((item) => item.distance - best.distance <= 24 || item.distance <= best.distance * 1.15).map((item) => item.entry);
  }
  return [best.entry];
}

function rightOfMatches(objectMap: ObjectMapEntry[], candidates: ObjectMapEntry[], selector: EditSelector): ObjectMapEntry[] {
  const spec = selector.rightOf;
  const text = typeof spec === "string" ? spec : spec?.text;
  const slide = typeof spec === "object" ? spec.slide : selector.slide;
  if (!text) return [];
  const anchors = objectMap.filter((entry) => entry.bounds && entry.text?.includes(text) && (slide === undefined || Number(entry.selectorHints?.slide) === slide));
  if (!anchors.length) return [];
  if (anchors.length > 1) return anchors;
  const anchor = anchors[0] as ObjectMapEntry;
  const anchorRight = (anchor.bounds?.x ?? 0) + (anchor.bounds?.width ?? 0);
  const ranked = candidates
    .filter((entry) => entry.bounds && entry.stableObjectId !== anchor.stableObjectId && Number(entry.selectorHints?.slide) === Number(anchor.selectorHints?.slide) && (entry.bounds?.x ?? 0) >= anchorRight)
    .map((entry) => ({ entry, delta: (entry.bounds?.x ?? 0) - anchorRight }))
    .sort((left, right) => left.delta - right.delta);
  const best = ranked[0];
  if (!best) return [];
  const close = ranked.filter((item) => item.delta - best.delta <= 24);
  return close.map((item) => item.entry);
}

function largestTextMatches(objectMap: ObjectMapEntry[], selector: EditSelector): ObjectMapEntry[] {
  const slide = typeof selector.largestTextOnSlide === "number" ? selector.largestTextOnSlide : selector.slide;
  if (slide === undefined) return [];
  const ranked = objectMap
    .filter((entry) => entry.kind === "shape" && entry.text && Number(entry.selectorHints?.slide) === slide)
    .map((entry) => ({ entry, score: textProminenceScore(entry) }))
    .sort((left, right) => right.score - left.score || String(right.entry.text ?? "").length - String(left.entry.text ?? "").length);
  const best = ranked[0];
  if (!best) return [];
  const close = ranked.filter((item) => best.score > 0 && item.score >= best.score * 0.95);
  return close.map((item) => item.entry);
}

function nthBodyShapeMatches(objectMap: ObjectMapEntry[], selector: { slide: number; n: number }): ObjectMapEntry[] {
  const bodies = objectMap.filter((entry) => {
    if (entry.kind !== "shape" || Number(entry.selectorHints?.slide) !== selector.slide) return false;
    const placeholder = String(entry.selectorHints?.placeholder ?? entry.selectorHints?.placeholderKey ?? "");
    return placeholder !== "title" && placeholder !== "ctrTitle";
  });
  return bodies.slice(Math.max(0, selector.n - 1), Math.max(0, selector.n));
}

function centerDistance(entry: ObjectMapEntry, x: number, y: number): number {
  const bounds = entry.bounds;
  if (!bounds) return Number.MAX_SAFE_INTEGER;
  const cx = bounds.x + bounds.width / 2;
  const cy = bounds.y + bounds.height / 2;
  return Math.hypot(cx - x, cy - y);
}

function textVisualArea(entry: ObjectMapEntry): number {
  const bounds = entry.bounds;
  return bounds ? bounds.width * bounds.height : 0;
}

function textProminenceScore(entry: ObjectMapEntry): number {
  const bounds = entry.bounds;
  if (!bounds) return 0;
  const placeholder = String(entry.selectorHints?.placeholder ?? entry.selectorHints?.placeholderKey ?? "");
  const titleBoost = placeholder === "title" || placeholder === "ctrTitle" ? 2.25 : 1;
  const textLength = Math.max(1, String(entry.text ?? "").trim().length);
  return bounds.height * Math.sqrt(textLength) * titleBoost;
}

function selectorConfidence(entry: ObjectMapEntry, selector: EditSelector, matches: ObjectMapEntry[]): number {
  if (selector.stableObjectId || selector.shapeId || selector.textHash || selector.positionHash) return 1;
  if (selector.nearestTo) {
    const distance = centerDistance(entry, selector.nearestTo.x, selector.nearestTo.y);
    return Number(Math.max(0.55, Math.min(0.98, 1 - distance / 1000)).toFixed(2));
  }
  if (selector.rightOf || selector.largestTextOnSlide || selector.nthBodyShape) return 0.82;
  if (matches.length === 1) return 0.9;
  return Number(Math.max(0.35, 0.8 / Math.max(1, matches.length)).toFixed(2));
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
  options: EditOptions,
  selectorResult?: ResolveEditSelectorsResult
): Promise<EditResult> {
  const inputSha256 = selectorResult?.inputSha256 ?? sha256Bytes(input.bytes);
  const sourceFingerprint = inputSourceFingerprint(input.bytes);
  if (hasPdfEncryptEntry(input.bytes)) {
    const opResults = operations.map((operation, index) => ({
      operationIndex: index,
      op: operationName(operation),
      applied: false,
      reason: "validation-failed" as const,
      message: "PDF_ENCRYPTED_BLOCKED: encrypted PDFs may be inspected for risk reporting, but mutation is blocked by default."
    }));
    return {
      schema: "officegen.edit.result@1.2",
      format: "pdf",
      dryRun: options.dryRun,
      inputSha256,
      objectMapHash: selectorResult?.objectMapHash,
      objectGraphHash: selectorResult?.objectGraphHash,
      sourceFingerprint,
      changed: false,
      applied: 0,
      skipped: operations.length,
      opResults,
      errors: opResults,
      patchPlan: options.dryRun ? await buildPatchPlan("pdf", input.bytes, operations, opResults, selectorResult) : undefined,
      caveats: [
        "PDF_ENCRYPTED_BLOCKED: edit does not use ignoreEncryption; inspect the PDF and provide an unencrypted copy before mutation.",
        "No output bytes were written."
      ]
    };
  }
  const staleErrors = stalePlanFailures(selectorResult, options, operations);
  if (staleErrors.length) {
    return {
      schema: "officegen.edit.result@1.2",
      format: "pdf",
      dryRun: options.dryRun,
      inputSha256,
      objectMapHash: selectorResult?.objectMapHash,
      objectGraphHash: selectorResult?.objectGraphHash,
      sourceFingerprint,
      changed: false,
      applied: 0,
      skipped: operations.length,
      opResults: staleErrors,
      errors: staleErrors,
      patchPlan: options.dryRun ? await buildPatchPlan("pdf", input.bytes, operations, staleErrors, selectorResult) : undefined,
      caveats: ["EDIT_STALE_PLAN: expected input or object map hash does not match the current PDF."]
    };
  }
  const redactionBlocks = pdfRedactionBlocks(operations);
  if (redactionBlocks.length) {
    const redactionIndexes = new Set(redactionBlocks.map((result) => result.operationIndex));
    const opResults = operations.map((operation, index) => redactionIndexes.has(index)
      ? redactionBlocks.find((result) => result.operationIndex === index) as EditOperationResult
      : {
          operationIndex: index,
          op: operationName(operation),
          applied: false,
          reason: "skipped-after-error" as const,
          message: "Skipped because PDF redaction operations are blocked atomically."
        });
    return {
      schema: "officegen.edit.result@1.2",
      format: "pdf",
      dryRun: options.dryRun,
      inputSha256,
      objectMapHash: selectorResult?.objectMapHash,
      objectGraphHash: selectorResult?.objectGraphHash,
      sourceFingerprint,
      changed: false,
      applied: 0,
      skipped: operations.length,
      opResults,
      errors: redactionBlocks,
      patchPlan: options.dryRun ? await buildPatchPlan("pdf", input.bytes, operations, opResults, selectorResult) : undefined,
      caveats: [
        "PDF_REDACTION_BLOCKED: true PDF redaction is not implemented; overlay operations must not be treated as redaction.",
        "No output bytes were written because redaction-like PDF operations were present."
      ]
    };
  }
  const pdf = await PDFDocument.load(input.bytes);
  const fontSet = await embedPdfFonts(pdf, operations.map((op) => "text" in op ? String(op.text) : ""));
  const font = fontSet.font;
  let applied = 0;
  let skipped = 0;
  const opResults: EditOperationResult[] = [];

  for (const [index, op] of operations.entries()) {
    const name = operationName(op);
    if (name === "pdf.textOverlay") {
      const textOp = op as { page: number; text: string; x: number; y: number; size?: number; color?: string };
      if (!isValidPage(pdf, textOp.page)) {
        skipped += 1;
        opResults.push({ operationIndex: index, op: name, applied: false, reason: "not-found", message: `PDF page ${textOp.page} was not found.` });
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
      opResults.push({ operationIndex: index, op: name, applied: true });
    } else if (name === "pdf.annotation") {
      const annotation = op as { page: number; text: string; x: number; y: number; width?: number; height?: number };
      if (!isValidPage(pdf, annotation.page)) {
        skipped += 1;
        opResults.push({ operationIndex: index, op: name, applied: false, reason: "not-found", message: `PDF page ${annotation.page} was not found.` });
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
      opResults.push({ operationIndex: index, op: name, applied: true });
    } else {
      skipped += 1;
      opResults.push({ operationIndex: index, op: name, applied: false, reason: "unsupported", message: `Unsupported PDF edit operation: ${name}` });
    }
  }

  const errors = opResults.filter(isRequiredEditFailure);
  if (errors.length && (!options.allowPartial || applied <= 0)) {
    return editAbortResult("pdf", operations.length, opResults, operations, [
      "PDF edit aborted before writing because not all required operations succeeded. Pass allowPartial to permit best-effort output.",
      "No output bytes were written because the edit result was incomplete."
    ], selectorResult, options.dryRun, input.bytes);
  }

  const bytes = options.dryRun ? undefined : await pdf.save({ useObjectStreams: false });
  if (!options.dryRun) await writeOutput(options.out, bytes as Uint8Array);
  return {
    schema: "officegen.edit.result@1.2",
    format: "pdf",
    dryRun: options.dryRun,
    inputSha256,
    objectMapHash: selectorResult?.objectMapHash,
    objectGraphHash: selectorResult?.objectGraphHash,
    sourceFingerprint,
    changed: applied > 0,
    applied,
    skipped,
    out: options.dryRun ? undefined : options.out,
    bytes: options.dryRun || options.out ? undefined : bytes,
    opResults,
    errors: errors.length ? errors : undefined,
    partial: errors.length ? true : undefined,
    allowPartial: options.allowPartial || undefined,
    patchPlan: options.dryRun ? await buildPatchPlan("pdf", input.bytes, operations, opResults, selectorResult) : undefined,
    caveats: ["PDF edit is additive; existing text/content is not removed in the MVP."]
  };
}

function pdfRedactionBlocks(operations: EditOperation[]): EditOperationResult[] {
  return operations.flatMap((operation, index) => {
    const op = operationName(operation);
    if (!isPdfRedactionOperation(operation, op)) return [];
    return [{
      operationIndex: index,
      op,
      applied: false,
      reason: "unsupported" as const,
      message: "PDF_REDACTION_UNSUPPORTED: true PDF redaction is not implemented. The operation was blocked so overlay output cannot be mistaken for removed content."
    }];
  });
}

function isPdfRedactionOperation(operation: EditOperation, name = operationName(operation)): boolean {
  if (/^pdf\..*redact/i.test(name) || /redaction/i.test(name)) return true;
  const record = operation as Record<string, unknown>;
  return Boolean(record.redact || record.redaction || record.redactions);
}

function hasPdfEncryptEntry(bytes: Uint8Array): boolean {
  return /\/Encrypt\b/.test(Buffer.from(bytes).toString("latin1"));
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
