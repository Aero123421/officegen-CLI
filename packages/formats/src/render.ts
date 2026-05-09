import { extname } from "node:path";
import { OfficegenError, type OfficegenConfig } from "@officegen/core";
import { assertPdfStandardFontText, writeOutput } from "./shared.js";
import { PDFDocument, StandardFonts, rgb } from "pdf-lib";

export type RenderTarget = "pptx" | "docx" | "xlsx" | "pdf";

export interface DocumentIR {
  title?: string;
  kind?: string;
  targets?: string[];
  sections?: Array<{
    id?: string;
    title?: string;
    body?: string | string[];
    blocks?: Array<{ type?: string; text?: string; rows?: Array<Record<string, unknown> | unknown[]> }>;
    rows?: Array<Record<string, unknown> | unknown[]>;
    items?: string[];
  }>;
  slides?: Array<{ title?: string; body?: string | string[] }>;
  sheets?: Array<{ name?: string; rows?: Array<Record<string, unknown> | unknown[]> }>;
}

export interface RenderOptions {
  out?: string;
  target?: string;
  config?: OfficegenConfig;
}

export interface RenderResult {
  schema: "officegen.render.result@1.2";
  target: RenderTarget;
  out?: string;
  bytes?: Uint8Array | Buffer;
  caveats: string[];
}

export async function render(ir: DocumentIR, options: RenderOptions = {}): Promise<RenderResult> {
  const target = resolveRenderTarget(ir, options);
  if (target === "pptx") return renderPptx(ir, options);
  if (target === "docx") return renderDocx(ir, options);
  if (target === "xlsx") return renderXlsx(ir, options);
  return renderPdf(ir, options);
}

export const renderDocument = render;

async function renderPptx(ir: DocumentIR, options: RenderOptions): Promise<RenderResult> {
  const mod = (await import("pptxgenjs")) as unknown as { default: new () => any };
  const pptx = new mod.default();
  pptx.layout = "LAYOUT_WIDE";
  const slides = ir.slides?.length ? ir.slides : normalizedSections(ir);
  for (const section of slides) {
    const slide = pptx.addSlide();
    slide.background = { color: "FFFFFF" };
    slide.addText(section.title ?? ir.title ?? "Untitled", { x: 0.55, y: 0.35, w: 12.2, h: 0.5, fontSize: 28, bold: true, color: "111111" });
    const body = Array.isArray(section.body) ? section.body.join("\n") : section.body ?? "";
    slide.addText(body, { x: 0.65, y: 1.1, w: 12, h: 5.7, fontSize: 16, color: "333333", breakLine: false, fit: "shrink" });
  }
  const bytes = await pptx.write({ outputType: "nodebuffer" });
  await writeOutput(options.out, bytes);
  return {
    schema: "officegen.render.result@1.2",
    target: "pptx",
    out: options.out,
    bytes: options.out ? undefined : bytes,
    caveats: ["Basic PPTX generation uses plain text boxes; advanced theme/layout support is not included."]
  };
}

async function renderDocx(ir: DocumentIR, options: RenderOptions): Promise<RenderResult> {
  const docx = await import("docx");
  const sections = normalizedSections(ir);
  const children = [
    new docx.Paragraph({ text: ir.title ?? "Untitled", heading: docx.HeadingLevel.TITLE }),
    ...sections.flatMap((section) => [
      ...(section.title ? [new docx.Paragraph({ text: section.title, heading: docx.HeadingLevel.HEADING_1 })] : []),
      ...toLines(section.body).map((line) => new docx.Paragraph({ children: [new docx.TextRun(line)] }))
    ])
  ];
  const document = new docx.Document({ sections: [{ properties: {}, children }] });
  const bytes = await docx.Packer.toBuffer(document);
  await writeOutput(options.out, bytes);
  return {
    schema: "officegen.render.result@1.2",
    target: "docx",
    out: options.out,
    bytes: options.out ? undefined : bytes,
    caveats: ["Basic DOCX generation supports headings and paragraphs only."]
  };
}

async function renderXlsx(ir: DocumentIR, options: RenderOptions): Promise<RenderResult> {
  const ExcelJS = (await import("exceljs")) as unknown as { default: any };
  const workbook = new ExcelJS.default.Workbook();
  const sheets = ir.sheets?.length ? ir.sheets : [{ name: ir.title ?? "Sheet1", rows: rowsFromSections(ir) }];
  for (const sheetSpec of sheets) {
    const sheet = workbook.addWorksheet(sanitizeSheetName(sheetSpec.name ?? "Sheet"));
    const rows = sheetSpec.rows ?? [];
    if (rows.length && !Array.isArray(rows[0])) {
      const keys = Object.keys(rows[0] as Record<string, unknown>);
      sheet.addRow(keys);
      for (const row of rows as Array<Record<string, unknown>>) sheet.addRow(keys.map((key) => row[key]));
    } else {
      for (const row of rows as unknown[][]) sheet.addRow(row);
    }
    sheet.columns?.forEach((column: { width?: number }) => {
      column.width = Math.max(column.width ?? 12, 12);
    });
  }
  const bytes = Buffer.from(await workbook.xlsx.writeBuffer());
  await writeOutput(options.out, bytes);
  return {
    schema: "officegen.render.result@1.2",
    target: "xlsx",
    out: options.out,
    bytes: options.out ? undefined : bytes,
    caveats: ["Basic XLSX generation supports worksheets and tabular rows only."]
  };
}

async function renderPdf(ir: DocumentIR, options: RenderOptions): Promise<RenderResult> {
  const pdf = await PDFDocument.create();
  const font = await pdf.embedFont(StandardFonts.Helvetica);
  const bold = await pdf.embedFont(StandardFonts.HelveticaBold);
  const sections = normalizedSections(ir);
  for (const section of sections) {
    const page = pdf.addPage([612, 792]);
    const { height } = page.getSize();
    const title = assertPdfStandardFontText(section.title ?? ir.title ?? "Untitled", bold, "render.pdf.title");
    page.drawText(title, { x: 54, y: height - 72, size: 22, font: bold, color: rgb(0.07, 0.07, 0.07) });
    let y = height - 112;
    for (const line of toLines(section.body)) {
      page.drawText(assertPdfStandardFontText(line.slice(0, 95), font, "render.pdf.body"), { x: 54, y, size: 11, font, color: rgb(0.16, 0.16, 0.16) });
      y -= 18;
      if (y < 54) break;
    }
  }
  const bytes = await pdf.save({ useObjectStreams: false });
  await writeOutput(options.out, bytes);
  return {
    schema: "officegen.render.result@1.2",
    target: "pdf",
    out: options.out,
    bytes: options.out ? undefined : bytes,
    caveats: ["PDF direct render is fixed-layout and is not a native Office conversion path."]
  };
}

function resolveRenderTarget(ir: DocumentIR, options: RenderOptions): RenderTarget {
  const explicit = options.target ?? ir.kind;
  const outputTarget = inferTargetFromOutput(options.out);
  if (explicit !== undefined) {
    const target = parseRenderTarget(explicit, "render target");
    assertOutputTargetMatches(target, outputTarget, options.out);
    return target;
  }
  if (outputTarget !== undefined) return outputTarget;
  if (ir.targets !== undefined && ir.targets.length > 0) {
    return parseRenderTarget(ir.targets[0], "IR targets[0]");
  }
  return "pdf";
}

function inferTargetFromOutput(out?: string): RenderTarget | undefined {
  if (!out) return undefined;
  const ext = extname(out).slice(1).toLowerCase();
  if (!ext) return undefined;
  return parseRenderTarget(ext, "output extension");
}

function isRenderTarget(value: unknown): value is RenderTarget {
  return value === "pptx" || value === "docx" || value === "xlsx" || value === "pdf";
}

function parseRenderTarget(value: unknown, source: string): RenderTarget {
  if (isRenderTarget(value)) return value;
  throw new OfficegenError(
    "EXPORT_UNSUPPORTED",
    `Unsupported ${source}: ${String(value)}. Supported render targets are pptx, docx, xlsx, and pdf.`,
    { source, value: String(value), supported: ["pptx", "docx", "xlsx", "pdf"] }
  );
}

function assertOutputTargetMatches(target: RenderTarget, outputTarget: RenderTarget | undefined, out?: string): void {
  if (outputTarget === undefined || outputTarget === target) return;
  throw new OfficegenError(
    "TARGET_EXTENSION_MISMATCH",
    `Render target ${target} does not match output extension .${outputTarget}${out ? ` for ${out}` : ""}.`,
    { target, outputTarget, ...(out ? { out } : {}) }
  );
}

function toLines(value?: string | string[]): string[] {
  if (Array.isArray(value)) return value.flatMap((item) => String(item).split(/\r?\n/));
  return String(value ?? "").split(/\r?\n/).filter(Boolean);
}

function normalizedSections(ir: DocumentIR): Array<{ title?: string; body?: string | string[]; rows?: Array<Record<string, unknown> | unknown[]> }> {
  const sections = ir.sections?.length ? ir.sections : [{ title: ir.title ?? "Untitled", body: "" }];
  return sections.map((section) => ({
    ...section,
    title: section.title ?? ir.title ?? "Untitled",
    body: section.body ?? bodyFromBlocks(section.blocks)
  }));
}

function bodyFromBlocks(blocks?: Array<{ type?: string; text?: string }>): string {
  return (blocks ?? [])
    .map((block) => block.text ?? "")
    .filter(Boolean)
    .join("\n");
}

function rowsFromSections(ir: DocumentIR): Array<Record<string, unknown> | unknown[]> {
  const rows = (ir.sections ?? []).flatMap((section) => [
    ...(section.rows ?? []),
    ...section.blocks?.flatMap((block) => block.rows ?? []) ?? []
  ]);
  return rows.length ? rows : [["title", "body"], [ir.title ?? "Untitled", bodyFromBlocks(ir.sections?.[0]?.blocks)]];
}

function sanitizeSheetName(name: string): string {
  return name.replace(/[\[\]*?/\\:]/g, " ").slice(0, 31) || "Sheet";
}
