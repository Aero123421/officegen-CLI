import { inspect, type InspectResult } from "./inspect.js";
import { view } from "./view.js";
import { exportDocument } from "./export.js";
import { type InputLike, type ObjectMapEntry, type OfficegenConfig, loadZip, normalizeInput, sortedZipFiles } from "./shared.js";
import { PDFDocument } from "pdf-lib";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { createHash } from "node:crypto";

export interface DiffOptions {
  config?: OfficegenConfig;
  visual?: boolean;
  native?: boolean;
  maxPages?: number;
}

export interface DiffResult {
  schema: "officegen.diff.result@1.2";
  formatBefore: string;
  formatAfter: string;
  changed: boolean;
  summary: {
    addedObjects: number;
    removedObjects: number;
    changedTextObjects: number;
    changedGeometryObjects: number;
    beforePages: number;
    afterPages: number;
    pageCountChanged: boolean;
    changedParts?: number;
    visualRegressionScore?: number;
  };
  semantic: {
    added: ObjectMapEntry[];
    removed: ObjectMapEntry[];
    changedText: Array<{
      stableObjectId: string;
      kind: string;
      before?: string;
      after?: string;
    }>;
    changedGeometry: Array<{
      stableObjectId: string;
      kind: string;
      beforeBbox?: [number, number, number, number];
      afterBbox?: [number, number, number, number];
      delta: { x: number; y: number; width: number; height: number };
    }>;
    partChanges?: Array<{ path: string; kind: string; beforeHash?: string; afterHash?: string; status: "added" | "removed" | "changed" }>;
  };
  visual?: {
    fidelity: "approximate" | "native";
    pagesCompared: number;
    beforePages: number;
    afterPages: number;
    pageCountChanged: boolean;
    pageScores: Array<{ page: number; score: number; beforeHash: string; afterHash: string }>;
    renderer?: string;
  };
  caveats: string[];
}

export async function diffDocuments(before: InputLike, after: InputLike, options: DiffOptions = {}): Promise<DiffResult> {
  const beforeInspect = await inspect(before, { depth: "shallow", config: options.config });
  const afterInspect = await inspect(after, { depth: "shallow", config: options.config });
  const semantic = semanticDiff(beforeInspect, afterInspect);
  semantic.partChanges = await semanticPartDiff(before, after, beforeInspect.trusted.format, afterInspect.trusted.format);
  const visual = options.visual ? await visualDiff(before, after, beforeInspect, afterInspect, options) : undefined;
  const visualRegressionScore = visual?.pageScores.length
    ? Number((visual.pageScores.reduce((sum, page) => sum + page.score, 0) / visual.pageScores.length).toFixed(4))
    : undefined;
  const beforePages = pageLikeCount(beforeInspect);
  const afterPages = pageLikeCount(afterInspect);
  const pageCountChanged = beforePages !== afterPages;
  const changed = semantic.added.length > 0 || semantic.removed.length > 0 || semantic.changedText.length > 0 || semantic.changedGeometry.length > 0 || (semantic.partChanges?.length ?? 0) > 0 || pageCountChanged || (visualRegressionScore ?? 0) > 0;
  return {
    schema: "officegen.diff.result@1.2",
    formatBefore: beforeInspect.trusted.format,
    formatAfter: afterInspect.trusted.format,
    changed,
    summary: {
      addedObjects: semantic.added.length,
      removedObjects: semantic.removed.length,
      changedTextObjects: semantic.changedText.length,
      changedGeometryObjects: semantic.changedGeometry.length,
      beforePages,
      afterPages,
      pageCountChanged,
      changedParts: semantic.partChanges?.length ?? 0,
      visualRegressionScore
    },
    semantic,
    visual,
    caveats: [
      visual?.fidelity === "native"
        ? "Native visual regression compares trusted renderer PDF outputs; fidelity depends on installed renderer filters and fonts."
        : visual?.renderer === "pdf-bytes"
          ? "PDF visual diff without native renderer compares page-aware PDF byte windows; it avoids zero-content false negatives but is not raster fidelity."
          : "Visual diff is based on officegen's approximate SVG/HTML view, not a native Office rasterization.",
      "StableObjectId matching is best-effort across generated files and preserves strongest value for edits within the same document lineage."
    ]
  };
}

function semanticDiff(before: InspectResult, after: InspectResult): DiffResult["semantic"] {
  const beforeMap = new Map(before.objectMap.map((entry) => [entry.stableObjectId, entry]));
  const afterMap = new Map(after.objectMap.map((entry) => [entry.stableObjectId, entry]));
  const added = [...afterMap.values()].filter((entry) => !beforeMap.has(entry.stableObjectId));
  const removed = [...beforeMap.values()].filter((entry) => !afterMap.has(entry.stableObjectId));
  const changedText = [...beforeMap.entries()]
    .map(([stableObjectId, beforeEntry]) => {
      const afterEntry = afterMap.get(stableObjectId);
      if (!afterEntry || (beforeEntry.text ?? beforeEntry.textPreview) === (afterEntry.text ?? afterEntry.textPreview)) return undefined;
      return {
        stableObjectId,
        kind: beforeEntry.kind,
        before: beforeEntry.text ?? beforeEntry.textPreview,
        after: afterEntry.text ?? afterEntry.textPreview
      };
    })
    .filter((entry): entry is NonNullable<typeof entry> => Boolean(entry));
  const changedGeometry = [...beforeMap.entries()]
    .map(([stableObjectId, beforeEntry]) => {
      const afterEntry = afterMap.get(stableObjectId);
      const beforeBbox = normalizedBbox(beforeEntry);
      const afterBbox = afterEntry ? normalizedBbox(afterEntry) : undefined;
      if (!afterEntry || !beforeBbox || !afterBbox || bboxEqual(beforeBbox, afterBbox)) return undefined;
      return {
        stableObjectId,
        kind: beforeEntry.kind,
        beforeBbox,
        afterBbox,
        delta: {
          x: Number((afterBbox[0] - beforeBbox[0]).toFixed(2)),
          y: Number((afterBbox[1] - beforeBbox[1]).toFixed(2)),
          width: Number((afterBbox[2] - beforeBbox[2]).toFixed(2)),
          height: Number((afterBbox[3] - beforeBbox[3]).toFixed(2))
        }
      };
    })
    .filter((entry): entry is NonNullable<typeof entry> => Boolean(entry));
  return { added, removed, changedText, changedGeometry };
}

async function semanticPartDiff(before: InputLike, after: InputLike, formatBefore: string, formatAfter: string): Promise<NonNullable<DiffResult["semantic"]["partChanges"]>> {
  if (formatBefore !== formatAfter || !["pptx", "docx", "xlsx"].includes(formatBefore)) return [];
  const beforeNormalized = await normalizeInput(before);
  const afterNormalized = await normalizeInput(after);
  const beforeZip = await loadZip(beforeNormalized);
  const afterZip = await loadZip(afterNormalized);
  const beforeParts = await packagePartHashes(beforeZip);
  const afterParts = await packagePartHashes(afterZip);
  const keys = new Set([...beforeParts.keys(), ...afterParts.keys()]);
  const changes: NonNullable<DiffResult["semantic"]["partChanges"]> = [];
  for (const key of [...keys].sort()) {
    const beforeHash = beforeParts.get(key);
    const afterHash = afterParts.get(key);
    if (beforeHash === afterHash) continue;
    changes.push({
      path: key,
      kind: classifyPackagePart(key),
      beforeHash,
      afterHash,
      status: beforeHash ? afterHash ? "changed" : "removed" : "added"
    });
  }
  return changes;
}

function normalizedBbox(entry: ObjectMapEntry): [number, number, number, number] | undefined {
  if (entry.bbox && entry.bbox.length === 4) return entry.bbox as [number, number, number, number];
  if (!entry.bounds) return undefined;
  return [entry.bounds.x, entry.bounds.y, entry.bounds.width, entry.bounds.height];
}

function bboxEqual(left: [number, number, number, number], right: [number, number, number, number]): boolean {
  return left.every((value, index) => Math.abs(value - (right[index] ?? 0)) < 0.01);
}

function pageLikeCount(inspected: InspectResult): number {
  const summary = inspected.trusted.summary as Record<string, unknown>;
  return Number(summary.pages ?? summary.slides ?? summary.sheets ?? (inspected.trusted.format === "docx" ? 1 : 0));
}

async function packagePartHashes(zip: Awaited<ReturnType<typeof loadZip>>): Promise<Map<string, string>> {
  const map = new Map<string, string>();
  const interesting = sortedZipFiles(zip).filter((file) =>
    /\/charts\/chart\d+\.xml$/i.test(file) ||
    /\/embeddings\//i.test(file) ||
    /\/theme\/theme\d+\.xml$/i.test(file) ||
    /\/tables\/table\d+\.xml$/i.test(file) ||
    /\/media\//i.test(file) ||
    /\/comments.*\.xml$/i.test(file) ||
    /\/styles\.xml$/i.test(file)
  );
  for (const file of interesting) {
    const entry = zip.file(file);
    if (!entry) continue;
    const bytes = await entry.async("uint8array");
    map.set(file, bytesHash(bytes));
  }
  return map;
}

function classifyPackagePart(file: string): string {
  if (/\/charts\//i.test(file)) return "chartXml";
  if (/\/embeddings\//i.test(file)) return "embeddedWorkbook";
  if (/\/theme\//i.test(file)) return "theme";
  if (/\/tables\//i.test(file)) return "table";
  if (/\/media\//i.test(file)) return "imageOrMedia";
  if (/\/comments/i.test(file)) return "comments";
  if (/\/styles\.xml$/i.test(file)) return "styles";
  return "packagePart";
}

function bytesHash(bytes: Uint8Array): string {
  return `sha256:${createHash("sha256").update(bytes).digest("hex")}`;
}

async function visualDiff(
  beforeInput: InputLike,
  afterInput: InputLike,
  before: InspectResult,
  after: InspectResult,
  options: DiffOptions
): Promise<NonNullable<DiffResult["visual"]>> {
  if (options.native) return nativeVisualDiff(beforeInput, afterInput, options);
  if (before.trusted.format === "pdf" && after.trusted.format === "pdf") return pdfByteVisualDiff(beforeInput, afterInput, options);

  const beforeView = await view(before, { format: "svg", maxPages: options.maxPages, config: options.config });
  const afterView = await view(after, { format: "svg", maxPages: options.maxPages, config: options.config });
  const pagesCompared = Math.min(beforeView.pages.length, afterView.pages.length);
  const pageScores = [];
  for (let index = 0; index < pagesCompared; index += 1) {
    const beforeHash = textHash(beforeView.pages[index]?.content ?? "");
    const afterHash = textHash(afterView.pages[index]?.content ?? "");
    pageScores.push({
      page: index + 1,
      beforeHash,
      afterHash,
      score: beforeHash === afterHash ? 0 : normalizedStringDistance(beforeView.pages[index]?.content ?? "", afterView.pages[index]?.content ?? "")
    });
  }
  return {
    fidelity: "approximate",
    pagesCompared,
    beforePages: beforeView.pages.length,
    afterPages: afterView.pages.length,
    pageCountChanged: beforeView.pages.length !== afterView.pages.length,
    pageScores
  };
}

async function pdfByteVisualDiff(beforeInput: InputLike, afterInput: InputLike, options: DiffOptions): Promise<NonNullable<DiffResult["visual"]>> {
  const beforeNormalized = await normalizeInput(beforeInput);
  const afterNormalized = await normalizeInput(afterInput);
  const beforeDoc = await PDFDocument.load(beforeNormalized.bytes, { ignoreEncryption: true });
  const afterDoc = await PDFDocument.load(afterNormalized.bytes, { ignoreEncryption: true });
  const pagesCompared = Math.min(beforeDoc.getPageCount(), afterDoc.getPageCount(), options.maxPages ?? Number.MAX_SAFE_INTEGER);
  const beforeWindows = byteWindows(beforeNormalized.bytes, pagesCompared);
  const afterWindows = byteWindows(afterNormalized.bytes, pagesCompared);
  const pageScores = [];
  for (let index = 0; index < pagesCompared; index += 1) {
    pageScores.push({
      page: index + 1,
      beforeHash: textHash(beforeWindows[index] ?? ""),
      afterHash: textHash(afterWindows[index] ?? ""),
      score: normalizedStringDistance(beforeWindows[index] ?? "", afterWindows[index] ?? "")
    });
  }
  return {
    fidelity: "approximate",
    pagesCompared,
    beforePages: beforeDoc.getPageCount(),
    afterPages: afterDoc.getPageCount(),
    pageCountChanged: beforeDoc.getPageCount() !== afterDoc.getPageCount(),
    pageScores,
    renderer: "pdf-bytes"
  };
}

async function nativeVisualDiff(beforeInput: InputLike, afterInput: InputLike, options: DiffOptions): Promise<NonNullable<DiffResult["visual"]>> {
  const dir = await mkdtemp(path.join(os.tmpdir(), "officegen-diff-native-"));
  try {
    const beforePdf = path.join(dir, "before.pdf");
    const afterPdf = path.join(dir, "after.pdf");
    const beforeExport = await exportDocument(beforeInput, { to: "pdf", mode: "native", out: beforePdf, config: options.config });
    await exportDocument(afterInput, { to: "pdf", mode: "native", out: afterPdf, config: options.config });
    const beforeBytes = await readFile(beforePdf);
    const afterBytes = await readFile(afterPdf);
    const beforeDoc = await PDFDocument.load(beforeBytes, { ignoreEncryption: true });
    const afterDoc = await PDFDocument.load(afterBytes, { ignoreEncryption: true });
    const pagesCompared = Math.min(beforeDoc.getPageCount(), afterDoc.getPageCount(), options.maxPages ?? Number.MAX_SAFE_INTEGER);
    const beforeText = byteWindows(beforeBytes, pagesCompared);
    const afterText = byteWindows(afterBytes, pagesCompared);
    const pageScores = [];
    for (let index = 0; index < pagesCompared; index += 1) {
      pageScores.push({
        page: index + 1,
        beforeHash: textHash(beforeText[index] ?? ""),
        afterHash: textHash(afterText[index] ?? ""),
        score: normalizedStringDistance(beforeText[index] ?? "", afterText[index] ?? "")
      });
    }
    return {
      fidelity: "native",
      pagesCompared,
      beforePages: beforeDoc.getPageCount(),
      afterPages: afterDoc.getPageCount(),
      pageCountChanged: beforeDoc.getPageCount() !== afterDoc.getPageCount(),
      pageScores,
      renderer: beforeExport.renderer?.id ?? "native"
    };
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
}

function byteWindows(bytes: Uint8Array, windows: number): string[] {
  if (windows <= 0) return [];
  const chunkSize = Math.max(1, Math.ceil(bytes.length / windows));
  const chunks = [];
  for (let index = 0; index < windows; index += 1) {
    const chunk = bytes.subarray(index * chunkSize, Math.min(bytes.length, (index + 1) * chunkSize));
    chunks.push(Buffer.from(chunk).toString("latin1"));
  }
  return chunks;
}

function normalizedStringDistance(before: string, after: string): number {
  if (!before && !after) return 0;
  const max = Math.max(before.length, after.length, 1);
  let changed = Math.abs(before.length - after.length);
  const limit = Math.min(before.length, after.length);
  for (let index = 0; index < limit; index += 1) {
    if (before.charCodeAt(index) !== after.charCodeAt(index)) changed += 1;
  }
  return Number(Math.min(1, changed / max).toFixed(4));
}

function textHash(value: string): string {
  return `sha256:${createHash("sha256").update(value).digest("hex")}`;
}

export const diff = diffDocuments;
