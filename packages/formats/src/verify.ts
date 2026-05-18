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
  gates?: VerifyGates;
  formulas?: boolean;
  namedRanges?: boolean;
  externalLinks?: boolean;
  protectedSheets?: boolean;
  timeoutMs?: number;
  config?: OfficegenConfig;
}

export interface VerifyGates {
  expectedSlides?: number;
  expectedPages?: number;
  requiredText?: string[];
  forbiddenText?: string[];
  maxWarnings?: number;
  requireNoRepairDialog?: boolean;
  maxBlankPages?: number;
}

export interface VerifyResult {
  schema: "officegen.verify.result@1.2";
  readiness: "pass" | "pass_with_environment_gap" | "warning" | "blocked";
  partial?: boolean;
  phaseTimings?: Array<{ phase: string; durationMs: number; timeout?: boolean }>;
  score: number;
  format: string;
  openable: boolean;
  noRepairDialogExpected: boolean;
  nativeRenderer?: { attempted: boolean; ok: boolean; message?: string; artifact?: string; repairDialogExpected?: boolean };
  visual?: { fidelity: "approximate" | "native"; pagesChecked: number; blankPages: number };
  blockingIssues: string[];
  warnings: string[];
  warningSummary: Array<{ code: string; count: number; severity: "warning" | "error"; category: WarningCategory; examples: string[] }>;
  topRisks: Array<{ code: string; severity: "warning" | "error"; category: WarningCategory; count: number; message: string; slide?: number; page?: number; stableObjectId?: string; repair?: string }>;
  scoreBreakdown: Record<string, unknown>;
  recommendedRepairs: Array<{ code: string; command?: string; reason: string }>;
  artifacts: Record<string, unknown>;
  gates?: { passed: boolean; failed: string[]; warnings: string[] };
}

export async function verify(input: InputLike, options: VerifyOptions = {}): Promise<VerifyResult> {
  const normalized = await normalizeInput(input, "unknown");
  const artifacts: Record<string, unknown> = {};
  const phaseTimings: NonNullable<VerifyResult["phaseTimings"]> = [];
  let openable = true;
  let partial = false;
  const warnings: string[] = [];
  const blockingIssues: string[] = [];
  const inspectDepth = gatesNeedFullText(options.gates) ? "full" : "summary";
  const inspected = await timedPhase("inspect", phaseTimings, options.timeoutMs, () => inspect({ data: normalized.bytes, path: normalized.path, format: normalized.format }, { depth: inspectDepth, config: options.config })).catch((error) => {
    if (isTimeout(error)) {
      partial = true;
      warnings.push(`VERIFY_TIMEOUT: inspect exceeded ${options.timeoutMs}ms.`);
      return undefined;
    }
    openable = false;
    blockingIssues.push(error instanceof Error ? error.message : String(error));
    return undefined;
  });
  const diagnosed = inspected ? await timedPhase("diagnose", phaseTimings, options.timeoutMs, () => diagnose({ data: normalized.bytes, path: normalized.path, format: normalized.format }, { config: options.config })).catch((error) => {
    if (isTimeout(error)) {
      partial = true;
      warnings.push(`VERIFY_TIMEOUT: diagnose exceeded ${options.timeoutMs}ms.`);
      return undefined;
    }
    warnings.push(`DIAGNOSE_FAILED: ${error instanceof Error ? error.message : String(error)}`);
    return undefined;
  }) : undefined;
  const overflowIssues: Array<{ code: string; message: string; severity: "warning" | "error"; slide?: number; page?: number; stableObjectId?: string; repair?: string }> = [];
  for (const issue of diagnosed?.issues ?? []) {
    if (issue.severity === "error") blockingIssues.push(`${issue.code}: ${issue.message}`);
    if (issue.severity === "warning") {
      warnings.push(`${issue.code}: ${issue.message}`);
      if (issue.code === "TEXT_OVERFLOW_RISK") {
        const record = issue as typeof issue & { location?: { slide?: number; page?: number; stableObjectId?: string } };
        overflowIssues.push({
          code: issue.code,
          message: issue.message,
          severity: "warning",
          slide: record.location?.slide,
          page: record.location?.page,
          stableObjectId: record.location?.stableObjectId,
          repair: "Run layout repair or shorten/split the object; verify will report the worst five overflow candidates."
        });
      }
    }
  }
  let noRepairDialogExpected = ![...(diagnosed?.issues ?? [])].some((issue) => issue.code.startsWith("OFFICE_REPAIR_RISK"));

  const visual = options.visual && inspected
    ? await timedPhase("visual", phaseTimings, options.timeoutMs, () => verifyVisual({ data: normalized.bytes, format: normalized.format }, options.config)).catch((error) => {
        if (isTimeout(error)) {
          partial = true;
          warnings.push(`VERIFY_TIMEOUT: visual preview exceeded ${options.timeoutMs}ms.`);
          return undefined;
        }
        warnings.push(`VISUAL_VERIFY_FAILED: ${error instanceof Error ? error.message : String(error)}`);
        return undefined;
      })
    : undefined;
  if (visual?.blankPages) warnings.push(`VISUAL_BLANK_PAGE: ${visual.blankPages} blank preview pages detected.`);

  const nativeRenderer = options.native
    ? await timedPhase("native", phaseTimings, options.timeoutMs, () => verifyNative(normalized, options, artifacts)).catch((error): NonNullable<VerifyResult["nativeRenderer"]> => {
        if (isTimeout(error)) {
          partial = true;
          warnings.push(`VERIFY_TIMEOUT: native verification exceeded ${options.timeoutMs}ms.`);
          return { attempted: true, ok: false, message: `Native verification exceeded ${options.timeoutMs}ms.` };
        }
        return { attempted: true, ok: false, message: error instanceof Error ? error.message : String(error) };
      })
    : undefined;
  if (nativeRenderer && !nativeRenderer.ok) warnings.push(nativeRenderer.message ?? "Native renderer verification did not complete.");
  if (nativeRenderer?.repairDialogExpected === true) {
    noRepairDialogExpected = false;
    blockingIssues.push("OFFICE_REPAIR_DIALOG_EXPECTED_NATIVE");
  }
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

  const gateResult = inspected ? evaluateGates(options.gates, inspected, visual, warnings.length, noRepairDialogExpected, nativeRenderer) : undefined;
  for (const issue of gateResult?.failed ?? []) blockingIssues.push(issue);
  for (const issue of gateResult?.warnings ?? []) warnings.push(issue);

  if (!openable) blockingIssues.push("INPUT_NOT_OPENABLE");
  const warningSummary = aggregateWarnings(warnings, blockingIssues);
  const topRisks: VerifyResult["topRisks"] = warningSummary.slice(0, 8).map((item) => ({
    code: item.code,
    severity: item.severity,
    category: item.category,
    count: item.count,
    message: item.examples[0] ?? item.code,
    repair: repairForCode(item.code)
  }));
  for (const issue of worstOverflowIssues(overflowIssues).slice(0, 5)) {
    topRisks.push({
      code: issue.code,
      severity: "warning",
      category: "quality",
      count: overflowIssues.length,
      message: issue.message,
      slide: issue.slide,
      page: issue.page,
      stableObjectId: issue.stableObjectId,
      repair: issue.repair
    });
  }
  const hasNonEnvironmentWarnings = warningSummary.some((item) => item.severity === "warning" && item.category !== "environment");
  const readiness = blockingIssues.length ? "blocked" : hasNonEnvironmentWarnings ? "warning" : warnings.length ? "pass_with_environment_gap" : "pass";
  const warningPenalty = warningSummary.reduce((sum, item) => sum + (item.category === "environment" ? 0.01 : Math.min(0.16, item.count * 0.04)), 0);
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
    partial,
    phaseTimings,
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
    artifacts,
    gates: gateResult
  };
  if (options.out) await writeFile(options.out, `${JSON.stringify(result, null, 2)}\n`, "utf8");
  return result;
}

type WarningCategory = "quality" | "compatibility" | "security" | "environment";

function aggregateWarnings(warnings: string[], blockingIssues: string[]): VerifyResult["warningSummary"] {
  const map = new Map<string, { code: string; count: number; severity: "warning" | "error"; category: WarningCategory; examples: string[] }>();
  const entries: Array<{ message: string; severity: "warning" | "error" }> = [
    ...warnings.map((message) => ({ message, severity: "warning" as const })),
    ...blockingIssues.map((message) => ({ message, severity: "error" as const }))
  ];
  for (const entry of entries) {
    const code = entry.message.split(":")[0]?.trim() || entry.message;
    const current = map.get(code) ?? { code, count: 0, severity: entry.severity, category: warningCategory(code), examples: [] };
    current.count += 1;
    current.severity = current.severity === "error" || entry.severity === "error" ? "error" : "warning";
    if (current.examples.length < 3) current.examples.push(entry.message);
    map.set(code, current);
  }
  return [...map.values()].sort((left, right) => severityRank(right.severity) - severityRank(left.severity) || right.count - left.count || left.code.localeCompare(right.code));
}

function warningCategory(code: string): WarningCategory {
  if (code.startsWith("SECURITY_") || code.includes("MACRO") || code.includes("EXTERNAL_LINK")) return "security";
  if (code === "NATIVE_RENDERER_NOT_RUN" || code === "VERIFY_TIMEOUT" || code.includes("RENDERER")) return "environment";
  if (code.includes("REPAIR") || code.includes("OPENABLE") || code.includes("UNSUPPORTED")) return "compatibility";
  return "quality";
}

function worstOverflowIssues(issues: Array<{ message: string; code: string; severity: "warning" | "error"; slide?: number; page?: number; stableObjectId?: string; repair?: string }>): typeof issues {
  return [...issues].sort((left, right) => right.message.length - left.message.length);
}

async function timedPhase<T>(phase: string, timings: NonNullable<VerifyResult["phaseTimings"]>, timeoutMs: number | undefined, task: () => Promise<T>): Promise<T> {
  const started = Date.now();
  try {
    const result = timeoutMs ? await withTimeout(task(), timeoutMs, phase) : await task();
    timings.push({ phase, durationMs: Date.now() - started });
    return result;
  } catch (error) {
    timings.push({ phase, durationMs: Date.now() - started, timeout: isTimeout(error) });
    throw error;
  }
}

async function withTimeout<T>(promise: Promise<T>, timeoutMs: number, phase: string): Promise<T> {
  let timer: NodeJS.Timeout | undefined;
  try {
    return await Promise.race([
      promise,
      new Promise<T>((_resolve, reject) => {
        timer = setTimeout(() => reject(new Error(`VERIFY_TIMEOUT:${phase}`)), timeoutMs);
      })
    ]);
  } finally {
    if (timer) clearTimeout(timer);
  }
}

function isTimeout(error: unknown): boolean {
  return error instanceof Error && error.message.startsWith("VERIFY_TIMEOUT:");
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

function evaluateGates(
  gates: VerifyGates | undefined,
  inspected: Awaited<ReturnType<typeof inspect>>,
  visual: VerifyResult["visual"] | undefined,
  warningCount: number,
  noRepairDialogExpected: boolean,
  nativeRenderer: VerifyResult["nativeRenderer"] | undefined
): VerifyResult["gates"] | undefined {
  if (!gates) return undefined;
  const failed: string[] = [];
  const warnings: string[] = [];
  const summary = inspected.trusted.summary as Record<string, unknown>;
  if (gates.expectedSlides !== undefined && Number(summary.slides ?? 0) !== gates.expectedSlides) {
    failed.push(`GATE_EXPECTED_SLIDES: expected ${gates.expectedSlides}, got ${Number(summary.slides ?? 0)}.`);
  }
  if (gates.expectedPages !== undefined) {
    const pages = Number(summary.pages ?? summary.slides ?? summary.sheets ?? 0);
    if (pages !== gates.expectedPages) failed.push(`GATE_EXPECTED_PAGES: expected ${gates.expectedPages}, got ${pages}.`);
  }
  const searchableText = inspected.objectMap.map((entry) => `${entry.text ?? ""}\n${entry.textPreview ?? ""}`).join("\n");
  for (const text of gates.requiredText ?? []) {
    if (!searchableText.includes(text)) failed.push(`GATE_REQUIRED_TEXT_MISSING: ${text}`);
  }
  for (const text of gates.forbiddenText ?? []) {
    if (searchableText.includes(text)) failed.push(`GATE_FORBIDDEN_TEXT_PRESENT: ${text}`);
  }
  if (gates.maxBlankPages !== undefined) {
    if (!visual) failed.push("GATE_MAX_BLANK_PAGES_UNEVALUATED: run verify with visual enabled to evaluate maxBlankPages.");
    else if (visual.blankPages > gates.maxBlankPages) failed.push(`GATE_MAX_BLANK_PAGES: expected <= ${gates.maxBlankPages}, got ${visual.blankPages}.`);
  }
  if (gates.maxWarnings !== undefined && warningCount > gates.maxWarnings) {
    failed.push(`GATE_MAX_WARNINGS: expected <= ${gates.maxWarnings}, got ${warningCount}.`);
  }
  if (gates.requireNoRepairDialog) {
    if (!noRepairDialogExpected) failed.push("GATE_REPAIR_DIALOG_EXPECTED: repair dialog risk was detected.");
    if (!nativeRenderer?.ok) failed.push("GATE_REPAIR_DIALOG_NATIVE_UNEVALUATED: run verify with native enabled to prove repair-dialog behavior.");
    if (inspected.trusted.caveats.some((caveat) => /repair/i.test(caveat))) {
      warnings.push("GATE_REPAIR_DIALOG_EVIDENCE_LIMITED: inspect caveats mention repair risk.");
    }
  }
  return { passed: failed.length === 0, failed, warnings };
}

function gatesNeedFullText(gates: VerifyGates | undefined): boolean {
  return Boolean(gates?.requiredText?.length || gates?.forbiddenText?.length);
}

async function verifyVisual(input: InputLike, config?: OfficegenConfig): Promise<NonNullable<VerifyResult["visual"]>> {
  const preview = await view(input, { format: "svg", maxPages: 10, config });
  const blankPages = preview.pages.filter((page) => !page.objectMap.some(hasVisiblePreviewObject)).length;
  return { fidelity: "approximate", pagesChecked: preview.pages.length, blankPages };
}

function hasVisiblePreviewObject(entry: { kind: string; text?: string; textPreview?: string; label?: string }): boolean {
  const text = `${entry.text ?? ""}${entry.textPreview ?? ""}`.trim();
  if (text) return true;
  return ["picture", "image", "chart"].includes(entry.kind);
}

async function verifyNative(input: Awaited<ReturnType<typeof normalizeInput>>, options: VerifyOptions, artifacts: Record<string, unknown>): Promise<NonNullable<VerifyResult["nativeRenderer"]>> {
  if (!["pptx", "docx", "xlsx"].includes(input.format)) return { attempted: false, ok: false, message: "Native renderer verification is only available for Office inputs." };
  if (!input.path) return { attempted: false, ok: false, message: "Native renderer verification requires an input file path." };
  const dir = await mkdtemp(path.join(os.tmpdir(), "officegen-verify-"));
  const pdfPath = path.join(dir, "native.pdf");
  try {
    const exported = await exportDocument(input.path, { to: "pdf", mode: "native", out: pdfPath, config: options.config, timeoutMs: options.timeoutMs });
    const pdf = await PDFDocument.load(await import("node:fs/promises").then((fs) => fs.readFile(pdfPath)), { ignoreEncryption: true });
    artifacts.nativePdf = pdfPath;
    return {
      attempted: true,
      ok: true,
      artifact: pdfPath,
      repairDialogExpected: exported.renderer?.repairDialogExpected,
      message: `Native renderer produced ${pdf.getPageCount()} PDF page(s) with ${exported.renderer?.id ?? "renderer"}.`
    };
  } catch (error) {
    await rm(dir, { recursive: true, force: true });
    return { attempted: true, ok: false, message: error instanceof Error ? error.message : String(error) };
  }
}

export const verifyDocument = verify;
