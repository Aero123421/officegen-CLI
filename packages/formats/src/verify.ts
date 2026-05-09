import { mkdtemp, rm, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { diagnose } from "./diagnose.js";
import { exportDocument } from "./export.js";
import { inspect } from "./inspect.js";
import { type InputLike, type OfficegenConfig, normalizeInput } from "./shared.js";
import { view } from "./view.js";
import { PDFDocument } from "pdf-lib";

export interface VerifyOptions {
  native?: boolean;
  visual?: boolean;
  out?: string;
  formulas?: boolean;
  namedRanges?: boolean;
  externalLinks?: boolean;
  protectedSheets?: boolean;
  config?: OfficegenConfig;
}

export interface VerifyResult {
  schema: "officegen.verify.result@1.2";
  readiness: "pass" | "warning" | "blocked";
  score: number;
  format: string;
  openable: boolean;
  noRepairDialogExpected: boolean;
  nativeRenderer?: { attempted: boolean; ok: boolean; message?: string; artifact?: string };
  visual?: { fidelity: "approximate" | "native"; pagesChecked: number; blankPages: number };
  blockingIssues: string[];
  warnings: string[];
  warningSummary: Array<{ code: string; count: number; severity: "warning" | "error"; examples: string[] }>;
  topRisks: Array<{ code: string; severity: "warning" | "error"; count: number; message: string; repair?: string }>;
  scoreBreakdown: Record<string, unknown>;
  recommendedRepairs: Array<{ code: string; command?: string; reason: string }>;
  artifacts: Record<string, unknown>;
}

export async function verify(input: InputLike, options: VerifyOptions = {}): Promise<VerifyResult> {
  const normalized = await normalizeInput(input, "unknown");
  const artifacts: Record<string, unknown> = {};
  let openable = true;
  const warnings: string[] = [];
  const blockingIssues: string[] = [];
  const inspected = await inspect({ data: normalized.bytes, path: normalized.path, format: normalized.format }, { depth: "summary", config: options.config }).catch((error) => {
    openable = false;
    blockingIssues.push(error instanceof Error ? error.message : String(error));
    return undefined;
  });
  const diagnosed = inspected ? await diagnose({ data: normalized.bytes, path: normalized.path, format: normalized.format }, { config: options.config }) : undefined;
  const overflowIssues: Array<{ code: string; message: string; severity: "warning" | "error"; slide?: number; page?: number; repair?: string }> = [];
  for (const issue of diagnosed?.issues ?? []) {
    if (issue.severity === "error") blockingIssues.push(`${issue.code}: ${issue.message}`);
    if (issue.severity === "warning") {
      warnings.push(`${issue.code}: ${issue.message}`);
      if (issue.code === "TEXT_OVERFLOW_RISK") {
        const record = issue as typeof issue & { location?: { slide?: number; page?: number } };
        overflowIssues.push({
          code: issue.code,
          message: issue.message,
          severity: "warning",
          slide: record.location?.slide,
          page: record.location?.page,
          repair: "Run layout repair or shorten/split the object; verify will report the worst five overflow candidates."
        });
      }
    }
  }
  const noRepairDialogExpected = ![...(diagnosed?.issues ?? [])].some((issue) => issue.code.startsWith("OFFICE_REPAIR_RISK"));

  const visual = options.visual && inspected
    ? await verifyVisual({ data: normalized.bytes, format: normalized.format }, options.config)
    : undefined;
  if (visual?.blankPages) warnings.push(`VISUAL_BLANK_PAGE: ${visual.blankPages} blank preview pages detected.`);

  const nativeRenderer = options.native
    ? await verifyNative(normalized, options, artifacts)
    : undefined;
  if (nativeRenderer && !nativeRenderer.ok) warnings.push(nativeRenderer.message ?? "Native renderer verification did not complete.");
  if (!options.native && ["pptx", "docx", "xlsx"].includes(normalized.format)) {
    warnings.push("NATIVE_RENDERER_NOT_RUN: native repair-dialog/openability verification is optional-gated; use --native under an enabled renderer policy.");
  }
  if (normalized.format === "pdf" && inspected?.trusted.summary && (inspected.trusted.summary as Record<string, unknown>).textBlocks === 0) {
    warnings.push("PDF_TEXT_BLOCKS_ZERO: no extractable text blocks; page preview artifacts or native PDF tooling recommended.");
  }
  if (normalized.format === "xlsx") {
    const workbookMap = (inspected?.untrusted as Record<string, any> | undefined)?.workbookMap;
    if (options.formulas && !workbookMap?.formulas?.some((entry: any) => entry.count > 0)) warnings.push("XLSX_FORMULAS_NONE: no formulas detected.");
    if (options.namedRanges && !(workbookMap?.namedRanges?.length > 0)) warnings.push("XLSX_NAMED_RANGES_NONE: no named ranges detected.");
    if (options.externalLinks && workbookMap?.externalLinks?.length > 0) blockingIssues.push("XLSX_EXTERNAL_LINKS_PRESENT");
    if (options.protectedSheets && workbookMap?.protectedSheets?.length > 0) warnings.push("XLSX_PROTECTED_SHEETS_PRESENT: protected sheets may require manual review.");
  }

  if (!openable) blockingIssues.push("INPUT_NOT_OPENABLE");
  const warningSummary = aggregateWarnings(warnings, blockingIssues);
  const topRisks = warningSummary.slice(0, 8).map((item) => ({
    code: item.code,
    severity: item.severity,
    count: item.count,
    message: item.examples[0] ?? item.code,
    repair: repairForCode(item.code)
  }));
  for (const issue of overflowIssues.slice(0, 5)) {
    if (!topRisks.some((risk) => risk.code === issue.code)) {
      topRisks.push({ code: issue.code, severity: "warning", count: overflowIssues.length, message: issue.message, repair: issue.repair });
    }
  }
  const readiness = blockingIssues.length ? "blocked" : warnings.length ? "warning" : "pass";
  const warningPenalty = warningSummary.reduce((sum, item) => sum + Math.min(0.16, item.count * 0.04), 0);
  const blockingPenalty = Math.min(0.85, blockingIssues.length * 0.35);
  const score = Number(Math.max(0, 1 - blockingPenalty - warningPenalty).toFixed(2));
  const scoreBreakdown = {
    base: 1,
    blockingPenalty,
    warningPenalty,
    cappedWarningKinds: warningSummary.length,
    repeatedWarningsCapped: true
  };
  const recommendedRepairs = topRisks
    .filter((risk) => risk.repair)
    .map((risk) => ({ code: risk.code, reason: risk.repair ?? "", command: commandForRisk(risk.code, normalized.format) }));
  const result: VerifyResult = {
    schema: "officegen.verify.result@1.2",
    readiness,
    score,
    format: normalized.format,
    openable,
    noRepairDialogExpected,
    nativeRenderer,
    visual,
    blockingIssues,
    warnings,
    warningSummary,
    topRisks,
    scoreBreakdown,
    recommendedRepairs,
    artifacts
  };
  if (options.out) await writeFile(options.out, `${JSON.stringify(result, null, 2)}\n`, "utf8");
  return result;
}

function aggregateWarnings(warnings: string[], blockingIssues: string[]): VerifyResult["warningSummary"] {
  const map = new Map<string, { code: string; count: number; severity: "warning" | "error"; examples: string[] }>();
  const entries: Array<{ message: string; severity: "warning" | "error" }> = [
    ...warnings.map((message) => ({ message, severity: "warning" as const })),
    ...blockingIssues.map((message) => ({ message, severity: "error" as const }))
  ];
  for (const entry of entries) {
    const code = entry.message.split(":")[0]?.trim() || entry.message;
    const current = map.get(code) ?? { code, count: 0, severity: entry.severity, examples: [] };
    current.count += 1;
    current.severity = current.severity === "error" || entry.severity === "error" ? "error" : "warning";
    if (current.examples.length < 3) current.examples.push(entry.message);
    map.set(code, current);
  }
  return [...map.values()].sort((left, right) => severityRank(right.severity) - severityRank(left.severity) || right.count - left.count || left.code.localeCompare(right.code));
}

function severityRank(severity: "warning" | "error"): number {
  return severity === "error" ? 2 : 1;
}

function repairForCode(code: string): string | undefined {
  if (code === "TEXT_OVERFLOW_RISK") return "Shorten text, enlarge the text box, reduce font size, or split the slide/page.";
  if (code === "PDF_TEXT_BLOCKS_ZERO") return "Create page previews and inspect them with the AI vision layer, or use native PDF tooling.";
  if (code === "NATIVE_RENDERER_NOT_RUN") return "Run verify --native with a trusted renderer profile when repair-dialog evidence is required.";
  if (code === "XLSX_EXTERNAL_LINKS_PRESENT") return "Review and sanitize external workbook links before autonomous use.";
  return undefined;
}

function commandForRisk(code: string, format: string): string | undefined {
  if (code === "TEXT_OVERFLOW_RISK") return `officegen diagnose <input.${format}> --json`;
  if (code === "PDF_TEXT_BLOCKS_ZERO") return "officegen view input.pdf --out .officegen/runs/pdf-view --json";
  if (code === "NATIVE_RENDERER_NOT_RUN") return `OFFICEGEN_PROFILE=enterprise officegen verify input.${format} --native --visual --json`;
  return undefined;
}

async function verifyVisual(input: InputLike, config?: OfficegenConfig): Promise<NonNullable<VerifyResult["visual"]>> {
  const preview = await view(input, { format: "svg", maxPages: 10, config });
  const blankPages = preview.pages.filter((page) => !/<text\b|<rect\b|data-kind=/.test(page.content)).length;
  return { fidelity: "approximate", pagesChecked: preview.pages.length, blankPages };
}

async function verifyNative(input: Awaited<ReturnType<typeof normalizeInput>>, options: VerifyOptions, artifacts: Record<string, unknown>): Promise<NonNullable<VerifyResult["nativeRenderer"]>> {
  if (!["pptx", "docx", "xlsx"].includes(input.format)) return { attempted: false, ok: false, message: "Native renderer verification is only available for Office inputs." };
  if (!input.path) return { attempted: false, ok: false, message: "Native renderer verification requires an input file path." };
  const dir = await mkdtemp(path.join(os.tmpdir(), "officegen-verify-"));
  const pdfPath = path.join(dir, "native.pdf");
  try {
    const exported = await exportDocument(input.path, { to: "pdf", mode: "native", out: pdfPath, config: options.config });
    const pdf = await PDFDocument.load(await import("node:fs/promises").then((fs) => fs.readFile(pdfPath)), { ignoreEncryption: true });
    artifacts.nativePdf = pdfPath;
    return { attempted: true, ok: true, artifact: pdfPath, message: `Native renderer produced ${pdf.getPageCount()} PDF page(s) with ${exported.renderer?.id ?? "renderer"}.` };
  } catch (error) {
    await rm(dir, { recursive: true, force: true });
    return { attempted: true, ok: false, message: error instanceof Error ? error.message : String(error) };
  }
}

export const verifyDocument = verify;
