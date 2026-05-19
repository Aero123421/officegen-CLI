import * as pdfjs from "pdfjs-dist/legacy/build/pdf.mjs";
import { inflateSync } from "node:zlib";

export interface PdfTextBlock {
  page: number;
  index: number;
  text: string;
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  source: "pdfjs" | "content-stream";
  untrusted: true;
}

export interface PdfAnnotationSummary {
  page: number;
  index: number;
  subtype?: string;
  contents?: string;
  rect?: number[];
  hasAppearance?: boolean;
  untrusted: true;
}

export interface PdfRiskFlag {
  code: string;
  severity: "info" | "warning" | "error";
  message: string;
}

export interface PdfObjectGraph {
  pageCount: number;
  textBlocks: PdfTextBlock[];
  annotations: PdfAnnotationSummary[];
  metadata: Record<string, unknown>;
  scan: {
    objects: number;
    streams: number;
    imageObjects: number;
    annotationObjects: number;
    metadataObjects: number;
    embeddedFiles: number;
    objectStreams: number;
    xrefStreams: number;
    encrypted: boolean;
    incrementalUpdates: number;
    filters: Array<{ name: string; count: number; supported: boolean }>;
    unsupportedFilters: string[];
    hasAcroForm: boolean;
    hasJavascript: boolean;
    hasRedactionAnnotations: boolean;
  };
  riskFlags: PdfRiskFlag[];
  caveats: string[];
}

const TEXT_EXTRACTION_SUPPORTED_FILTERS = new Set([
  "ASCII85Decode",
  "ASCIIHexDecode",
  "FlateDecode",
  "RunLengthDecode"
]);

export async function inspectPdfObjectGraph(bytes: Uint8Array, pageSizes: Array<{ width: number; height: number }>): Promise<PdfObjectGraph> {
  const raw = Buffer.from(bytes).toString("latin1");
  const scan = scanPdfSyntax(raw);
  const pdfjsResult = await inspectWithPdfjs(bytes, pageSizes).catch((error: unknown) => ({
    textBlocks: [] as PdfTextBlock[],
    annotations: [] as PdfAnnotationSummary[],
    metadata: {},
    caveats: [`PDFJS_INSPECT_FAILED: ${error instanceof Error ? error.message : String(error)}`]
  }));
  const looseTextBlocks = extractLooseTextBlocks(raw);
  const streamTextBlocks = pdfjsResult.textBlocks.length ? [] : extractStreamTextBlocks(raw);
  const textBlocks = mergePdfTextBlocks(looseTextBlocks, pdfjsResult.textBlocks.length ? pdfjsResult.textBlocks : streamTextBlocks);
  const riskFlags = pdfRiskFlags(scan, textBlocks, pdfjsResult.annotations);
  return {
    pageCount: pageSizes.length,
    textBlocks,
    annotations: pdfjsResult.annotations,
    metadata: pdfjsResult.metadata,
    scan,
    riskFlags,
    caveats: [
      ...pdfjsResult.caveats,
      ...(scan.unsupportedFilters.length
        ? [`PDF_UNSUPPORTED_FILTERS: ${scan.unsupportedFilters.join(", ")} stream filter(s) are not decoded by the lightweight inspector.`]
        : []),
      ...(textBlocks.length ? [] : ["PDF_QUALITY_TEXT_BLOCKS_ZERO: no extractable text blocks were found; page preview artifacts or native PDF tooling are recommended."])
    ]
  };
}

export function scanPdfForForbiddenText(bytes: Uint8Array, forbidden: string[] | RegExp[]): {
  found: Array<{ pattern: string; source: "raw" | "content-stream"; sample: string }>;
  checkedSources: string[];
} {
  const raw = Buffer.from(bytes).toString("latin1");
  const decodedStreams = decodeContentStreams(raw);
  const sources = [
    { kind: "raw" as const, text: raw },
    ...decodedStreams.flatMap((text) => [
      { kind: "content-stream" as const, text },
      { kind: "content-stream" as const, text: extractPdfTextFromContent(text) }
    ])
  ];
  const found: Array<{ pattern: string; source: "raw" | "content-stream"; sample: string }> = [];
  for (const pattern of forbidden) {
    for (const source of sources) {
      const match = typeof pattern === "string"
        ? indexOfCaseInsensitive(source.text, pattern)
        : regexIndex(source.text, pattern);
      if (match.index >= 0) {
        found.push({
          pattern: String(pattern),
          source: source.kind,
          sample: source.text.slice(Math.max(0, match.index - 24), match.index + match.length + 24).replace(/\s+/g, " ")
        });
        break;
      }
    }
  }
  return { found, checkedSources: [...new Set(sources.map((source) => source.kind))] };
}

async function inspectWithPdfjs(bytes: Uint8Array, pageSizes: Array<{ width: number; height: number }>): Promise<{
  textBlocks: PdfTextBlock[];
  annotations: PdfAnnotationSummary[];
  metadata: Record<string, unknown>;
  caveats: string[];
}> {
  const loadingTask = pdfjs.getDocument({
    data: new Uint8Array(bytes),
    disableFontFace: true,
    isEvalSupported: false,
    useSystemFonts: true
  } as Parameters<typeof pdfjs.getDocument>[0]);
  const document = await loadingTask.promise;
  const textBlocks: PdfTextBlock[] = [];
  const annotations: PdfAnnotationSummary[] = [];
  const metadata: Record<string, unknown> = {};
  const caveats: string[] = [];

  const meta = await document.getMetadata().catch(() => undefined as unknown);
  if (meta && typeof meta === "object") {
    const record = meta as { info?: Record<string, unknown>; metadata?: { getAll?: () => Record<string, unknown> } };
    if (record.info) metadata.info = record.info;
    const xmp = record.metadata?.getAll?.();
    if (xmp && Object.keys(xmp).length) metadata.xmp = xmp;
  }

  const pages = Math.min(document.numPages, pageSizes.length || document.numPages);
  for (let pageNumber = 1; pageNumber <= pages; pageNumber += 1) {
    const page = await document.getPage(pageNumber);
    const size = pageSizes[pageNumber - 1] ?? { width: 612, height: 792 };
    const textContent = await page.getTextContent().catch(() => ({ items: [] as unknown[] }));
    for (const [index, item] of ((textContent as { items?: unknown[] }).items ?? []).entries()) {
      const textItem = item as { str?: string; transform?: number[]; width?: number; height?: number };
      const text = String(textItem.str ?? "").trim();
      if (!text) continue;
      const transform = textItem.transform ?? [];
      const x = Number(transform[4] ?? 0);
      const baselineY = Number(transform[5] ?? size.height);
      const inferredHeight = textItem.height ?? Math.abs(Number(transform[3] ?? 10));
      const height = Number(inferredHeight || 10);
      textBlocks.push({
        page: pageNumber,
        index: textBlocks.filter((block) => block.page === pageNumber).length + 1,
        text,
        x,
        y: Math.max(0, size.height - baselineY),
        width: Number(textItem.width ?? Math.max(24, text.length * 6)),
        height,
        source: "pdfjs",
        untrusted: true
      });
    }

    const pageAnnotations = await page.getAnnotations({ intent: "display" }).catch(() => [] as unknown[]);
    for (const [index, annotation] of pageAnnotations.entries()) {
      const item = annotation as { subtype?: string; contents?: string; rect?: number[]; hasAppearance?: boolean };
      annotations.push({
        page: pageNumber,
        index: index + 1,
        subtype: item.subtype,
        contents: item.contents,
        rect: Array.isArray(item.rect) ? item.rect.map(Number) : undefined,
        hasAppearance: item.hasAppearance,
        untrusted: true
      });
    }
    page.cleanup();
  }
  if (document.numPages !== pageSizes.length && pageSizes.length) {
    caveats.push(`PDF_PAGE_COUNT_MISMATCH: pdfjs reported ${document.numPages} page(s), pdf-lib reported ${pageSizes.length}.`);
  }
  await document.destroy();
  return { textBlocks, annotations, metadata, caveats };
}

function scanPdfSyntax(raw: string): PdfObjectGraph["scan"] {
  const filterCounts = countFilters(raw);
  const filters = [...filterCounts.entries()].sort(([a], [b]) => a.localeCompare(b)).map(([name, count]) => ({
    name,
    count,
    supported: TEXT_EXTRACTION_SUPPORTED_FILTERS.has(name)
  }));
  const unsupportedFilters = filters.filter((filter) => !filter.supported).map((filter) => filter.name);
  const incrementalUpdates = Math.max(0, (raw.match(/%%EOF/g) ?? []).length - 1);
  return {
    objects: (raw.match(/\b\d+\s+\d+\s+obj\b/g) ?? []).length,
    streams: (raw.match(/\bstream\r?\n/g) ?? []).length,
    imageObjects: (raw.match(/\/Subtype\s*\/Image\b/g) ?? []).length,
    annotationObjects: (raw.match(/\/Annots\b|\/Type\s*\/Annot\b/g) ?? []).length,
    metadataObjects: (raw.match(/\/Metadata\b|\/Type\s*\/Metadata\b/g) ?? []).length,
    embeddedFiles: (raw.match(/\/EmbeddedFiles\b|\/EmbeddedFile\b|\/Filespec\b/g) ?? []).length,
    objectStreams: (raw.match(/\/Type\s*\/ObjStm\b/g) ?? []).length,
    xrefStreams: (raw.match(/\/Type\s*\/XRef\b/g) ?? []).length,
    encrypted: /\/Encrypt\b/.test(raw),
    incrementalUpdates,
    filters,
    unsupportedFilters,
    hasAcroForm: /\/AcroForm\b/.test(raw),
    hasJavascript: /\/JavaScript\b|\/OpenAction\b|\/AA\b|\/JS\b/.test(raw),
    hasRedactionAnnotations: /\/Subtype\s*\/Redact\b/.test(raw)
  };
}

function pdfRiskFlags(scan: PdfObjectGraph["scan"], textBlocks: PdfTextBlock[], annotations: PdfAnnotationSummary[]): PdfRiskFlag[] {
  const flags: PdfRiskFlag[] = [];
  if (!textBlocks.length) flags.push({ code: "PDF_TEXT_BLOCKS_ZERO", severity: "warning", message: "No extractable PDF text was found." });
  if (scan.encrypted) flags.push({ code: "PDF_ENCRYPTED", severity: "warning", message: "The PDF has an Encrypt entry; inspection may be incomplete." });
  if (scan.objectStreams || scan.xrefStreams) flags.push({ code: "PDF_COMPRESSED_OBJECT_GRAPH", severity: "info", message: "The PDF uses object or xref streams; byte-level graph inspection is approximate." });
  if (scan.unsupportedFilters.length) flags.push({ code: "PDF_UNSUPPORTED_FILTERS", severity: "warning", message: `Unsupported stream filters detected: ${scan.unsupportedFilters.join(", ")}.` });
  if (scan.embeddedFiles) flags.push({ code: "PDF_EMBEDDED_FILES", severity: "warning", message: "Embedded file references were detected." });
  if (scan.hasAcroForm) flags.push({ code: "PDF_ACROFORM", severity: "info", message: "Interactive form structures were detected." });
  if (scan.hasJavascript) flags.push({ code: "PDF_ACTIVE_CONTENT", severity: "warning", message: "PDF JavaScript or automatic actions were detected." });
  if (scan.incrementalUpdates) flags.push({ code: "PDF_INCREMENTAL_UPDATES", severity: "info", message: "Multiple EOF markers suggest incremental updates; stale content may remain in previous revisions." });
  if (scan.hasRedactionAnnotations || annotations.some((annotation) => annotation.subtype?.toLowerCase() === "redact")) {
    flags.push({ code: "PDF_REDACTION_ANNOTATIONS", severity: "warning", message: "Redaction annotations were detected; verify that they have been applied by a PDF processor." });
  }
  if (textBlocks.length) flags.push({ code: "PDF_EXTRACTABLE_TEXT_PRESENT", severity: "info", message: "Extractable text remains present in the PDF content graph." });
  return flags;
}

function countFilters(raw: string): Map<string, number> {
  const counts = new Map<string, number>();
  for (const match of raw.matchAll(/\/Filter\s*(\[[^\]]+\]|\/[A-Za-z0-9]+)/g)) {
    const value = match[1] ?? "";
    for (const name of value.matchAll(/\/([A-Za-z0-9]+)/g)) {
      counts.set(name[1] as string, (counts.get(name[1] as string) ?? 0) + 1);
    }
  }
  return counts;
}

function extractStreamTextBlocks(raw: string): PdfTextBlock[] {
  return decodeContentStreams(raw)
    .map((streamText) => extractPdfTextFromContent(streamText))
    .filter(Boolean)
    .map((text, index) => ({
      page: 1,
      index: index + 1,
      text,
      source: "content-stream" as const,
      untrusted: true as const
    }));
}

function extractLooseTextBlocks(raw: string): PdfTextBlock[] {
  const text = extractPdfTextFromContent(raw);
  if (!text) return [];
  return [{
    page: 1,
    index: 1,
    text,
    source: "content-stream",
    untrusted: true
  }];
}

function mergePdfTextBlocks(priorityBlocks: PdfTextBlock[], blocks: PdfTextBlock[]): PdfTextBlock[] {
  const merged: PdfTextBlock[] = [];
  const seen = new Set<string>();
  for (const block of [...priorityBlocks, ...blocks]) {
    const key = `${block.page}:${block.text}`;
    if (seen.has(key)) continue;
    seen.add(key);
    const pageIndex = merged.filter((item) => item.page === block.page).length + 1;
    merged.push({ ...block, index: pageIndex });
  }
  return merged;
}

function decodeContentStreams(raw: string): string[] {
  const streams: string[] = [];
  for (const object of raw.matchAll(/\b\d+\s+\d+\s+obj([\s\S]*?)endobj/g)) {
    const body = object[1] ?? "";
    const streamMatch = /stream\r?\n([\s\S]*?)\r?\nendstream/.exec(body);
    if (!streamMatch) continue;
    const encoded = Buffer.from(streamMatch[1] ?? "", "latin1");
    const filterMatch = /\/Filter\s*(\[[^\]]+\]|\/[A-Za-z0-9]+)/.exec(body);
    const filters = filterMatch ? [...(filterMatch[1] ?? "").matchAll(/\/([A-Za-z0-9]+)/g)].map((match) => match[1] as string) : [];
    if (!filters.length) {
      streams.push(encoded.toString("latin1"));
      continue;
    }
    if (filters.length === 1 && filters[0] === "FlateDecode") {
      try {
        streams.push(inflateSync(encoded).toString("latin1"));
      } catch {
        // Keep scanning other streams.
      }
    }
  }
  return streams;
}

function extractPdfTextFromContent(content: string): string {
  const strings: string[] = [];
  for (const match of content.matchAll(/\((?:\\.|[^\\)])*\)\s*(?:Tj|'|")/g)) {
    strings.push(decodePdfLiteral(match[0].replace(/\)\s*(?:Tj|'|")$/, "").slice(1)));
  }
  for (const match of content.matchAll(/\[((?:\s*(?:\((?:\\.|[^\\)])*\)|<[\da-fA-F\s]+>)\s*-?\d*)+)\]\s*TJ/g)) {
    const array = match[1] ?? "";
    for (const item of array.matchAll(/\((?:\\.|[^\\)])*\)|<[\da-fA-F\s]+>/g)) strings.push(decodePdfStringToken(item[0]));
  }
  for (const match of content.matchAll(/<([\da-fA-F\s]+)>\s*Tj/g)) {
    strings.push(decodePdfHex(match[1] ?? ""));
  }
  return strings.join(" ").replace(/\s+/g, " ").trim();
}

function decodePdfStringToken(token: string): string {
  return token.startsWith("<") ? decodePdfHex(token.slice(1, -1)) : decodePdfLiteral(token.slice(1, -1));
}

function decodePdfLiteral(value: string): string {
  return value
    .replace(/\\([nrtbf()\\])/g, (_match, code: string) => ({ n: "\n", r: "\r", t: "\t", b: "\b", f: "\f", "(": "(", ")": ")", "\\": "\\" })[code] ?? code)
    .replace(/\\([0-7]{1,3})/g, (_match, octal: string) => String.fromCharCode(Number.parseInt(octal, 8)))
    .replace(/\)$/g, "");
}

function decodePdfHex(hex: string): string {
  const clean = hex.replace(/\s+/g, "");
  if (!clean) return "";
  const bytes = Buffer.from(clean.length % 2 ? `${clean}0` : clean, "hex");
  if (bytes[0] === 0xfe && bytes[1] === 0xff) return Buffer.from(bytes.subarray(2)).swap16().toString("utf16le");
  return bytes.toString("latin1");
}

function indexOfCaseInsensitive(text: string, pattern: string): { index: number; length: number } {
  if (!pattern) return { index: -1, length: 0 };
  return { index: text.toLowerCase().indexOf(pattern.toLowerCase()), length: pattern.length };
}

function regexIndex(text: string, pattern: RegExp): { index: number; length: number } {
  const flags = pattern.flags.includes("g") ? pattern.flags : `${pattern.flags}g`;
  const match = new RegExp(pattern.source, flags).exec(text);
  return { index: match?.index ?? -1, length: match?.[0]?.length ?? 0 };
}
