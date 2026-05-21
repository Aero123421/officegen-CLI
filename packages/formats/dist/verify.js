import { mkdir, rm, writeFile } from "node:fs/promises";
import path from "node:path";
import { diagnose } from "./diagnose.js";
import { exportDocument } from "./export.js";
import { inspect } from "./inspect.js";
import { normalizeInput } from "./shared.js";
import { view } from "./view.js";
import { PDFDocument } from "pdf-lib";
export async function verify(input, options = {}) {
    const normalized = await normalizeInput(input, "unknown");
    const artifacts = {};
    const phaseTimings = [];
    let openable = true;
    let partial = false;
    const warnings = [];
    const blockingIssues = [];
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
    const overflowIssues = [];
    for (const issue of diagnosed?.issues ?? []) {
        if (issue.severity === "error")
            blockingIssues.push(`${issue.code}: ${issue.message}`);
        if (issue.severity === "warning") {
            warnings.push(`${issue.code}: ${issue.message}`);
            if (issue.code === "TEXT_OVERFLOW_RISK") {
                const record = issue;
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
    const nativeRequested = options.native === true || options.mode === "proof";
    const visual = options.visual && inspected
        ? await timedPhase("visual", phaseTimings, options.timeoutMs, () => verifyVisual({ data: normalized.bytes, path: normalized.path, format: normalized.format }, options.config, options.mode)).catch((error) => {
            if (isTimeout(error)) {
                partial = true;
                warnings.push(`VERIFY_TIMEOUT: visual preview exceeded ${options.timeoutMs}ms.`);
                return undefined;
            }
            warnings.push(`VISUAL_VERIFY_FAILED: ${error instanceof Error ? error.message : String(error)}`);
            return undefined;
        })
        : undefined;
    if (visual?.blankPages)
        warnings.push(`VISUAL_BLANK_PAGE: ${visual.blankPages} blank preview pages detected.`);
    if (visual?.identicalPages.length)
        warnings.push(`VISUAL_IDENTICAL_PAGES: raster preview pages ${visual.identicalPages.join(", ")} share page hashes.`);
    for (const warning of visual?.pixelDensityWarnings ?? [])
        warnings.push(`VISUAL_PIXEL_DENSITY: ${warning}`);
    for (const issue of visualBlockingIssues(visual))
        blockingIssues.push(issue);
    const visualDiff = options.visual
        ? {
            status: "skipped",
            expectedDiffOnly: false,
            fidelity: visual?.fidelity,
            pagesCompared: 0,
            message: "Visual preview verification ran without an expected/baseline document, so pixel diff was skipped."
        }
        : undefined;
    const nativeRenderer = nativeRequested
        ? await timedPhase("native", phaseTimings, options.timeoutMs, () => verifyNative(normalized, options, artifacts)).catch((error) => {
            if (isTimeout(error)) {
                partial = true;
                warnings.push(`VERIFY_TIMEOUT: native verification exceeded ${options.timeoutMs}ms.`);
                return { attempted: true, ok: false, message: `Native verification exceeded ${options.timeoutMs}ms.` };
            }
            return { attempted: true, ok: false, message: error instanceof Error ? error.message : String(error) };
        })
        : undefined;
    if (nativeRenderer && !nativeRenderer.ok)
        warnings.push(nativeRenderer.message ?? "Native renderer verification did not complete.");
    if (nativeRequested && nativeRenderer && !nativeRenderer.ok) {
        blockingIssues.push(`NATIVE_RENDERER_BLOCKED: ${nativeRenderer.message ?? "Native renderer verification did not complete."}`);
    }
    if (nativeRenderer?.repairDialogExpected === true) {
        noRepairDialogExpected = false;
        blockingIssues.push("OFFICE_REPAIR_DIALOG_EXPECTED_NATIVE");
    }
    if (!nativeRequested && ["pptx", "docx", "xlsx"].includes(normalized.format)) {
        warnings.push("NATIVE_RENDERER_NOT_RUN: native repair-dialog/openability verification is optional-gated; use --native under an enabled renderer policy.");
    }
    if (normalized.format === "pdf" && inspected?.trusted.summary && inspected.trusted.summary.textBlocks === 0) {
        warnings.push("PDF_TEXT_BLOCKS_ZERO: no extractable text blocks; page preview artifacts or native PDF tooling recommended.");
    }
    for (const warning of packageRiskWarningsFromCaveats(inspected?.trusted.caveats ?? []))
        warnings.push(warning);
    if (normalized.format === "xlsx") {
        const workbookMap = inspected?.untrusted?.workbookMap;
        if (options.formulas && !workbookMap?.formulas?.some((entry) => entry.count > 0))
            warnings.push("XLSX_FORMULAS_NONE: no formulas detected.");
        if (options.namedRanges && !(workbookMap?.namedRanges?.length > 0))
            warnings.push("XLSX_NAMED_RANGES_NONE: no named ranges detected.");
        if (options.externalLinks && workbookMap?.externalLinks?.length > 0)
            blockingIssues.push("XLSX_EXTERNAL_LINKS_PRESENT");
        if (options.protectedSheets && workbookMap?.protectedSheets?.length > 0)
            warnings.push("XLSX_PROTECTED_SHEETS_PRESENT: protected sheets may require manual review.");
    }
    const gateResult = inspected ? evaluateGates(options.gates, inspected, visual, warnings.length, noRepairDialogExpected, nativeRenderer) : undefined;
    for (const issue of gateResult?.failed ?? [])
        blockingIssues.push(issue);
    for (const issue of gateResult?.warnings ?? [])
        warnings.push(issue);
    if (!openable)
        blockingIssues.push("INPUT_NOT_OPENABLE");
    const warningSummary = aggregateWarnings(warnings, blockingIssues);
    const topRisks = warningSummary
        .filter((item) => !(item.code === "TEXT_OVERFLOW_RISK" && overflowIssues.length > 0))
        .slice(0, 8)
        .map((item) => ({
        code: item.code,
        severity: item.severity,
        category: item.category,
        count: item.count,
        message: item.examples[0] ?? item.code,
        repair: repairForCode(item.code)
    }));
    const worstOverflows = worstOverflowIssues(overflowIssues).slice(0, 5);
    if (worstOverflows.length) {
        const first = worstOverflows[0];
        topRisks.push({
            code: "TEXT_OVERFLOW_RISK",
            severity: "warning",
            category: "quality",
            count: overflowIssues.length,
            message: `${overflowIssues.length} text overflow risk(s) detected; worst ${worstOverflows.length} candidate(s) are listed in examples.`,
            slide: first.slide,
            page: first.page,
            stableObjectId: first.stableObjectId,
            repair: first.repair,
            examples: worstOverflows.map((issue) => ({
                slide: issue.slide,
                page: issue.page,
                stableObjectId: issue.stableObjectId,
                message: issue.message
            }))
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
    const recommendedRepairs = uniqueRepairs(topRisks
        .filter((risk) => risk.repair)
        .map((risk) => ({ code: risk.code, reason: risk.repair ?? "", command: commandForRisk(risk.code, normalized.format) })));
    const nativeProof = nativeProofFromRenderer(nativeRenderer, nativeRequested, normalized.format);
    const result = {
        schema: "officegen.verify.result@1.2",
        verificationReport: buildVerificationReport({
            format: normalized.format,
            readiness,
            partial,
            score,
            openable,
            noRepairDialogExpected,
            inspected,
            diagnosed,
            visual,
            visualDiff,
            nativeRenderer,
            nativeProof,
            gateResult,
            warnings,
            blockingIssues,
            warningSummary,
            recommendedRepairs,
            artifacts
        }),
        readiness,
        partial,
        phaseTimings,
        score,
        format: normalized.format,
        openable,
        noRepairDialogExpected,
        nativeRenderer,
        nativeProof,
        visual,
        visualDiff,
        expectedDiffOnly: visualDiff?.expectedDiffOnly ?? false,
        blockingIssues,
        warnings,
        warningSummary,
        topRisks,
        scoreBreakdown,
        recommendedRepairs,
        artifacts,
        gates: gateResult
    };
    if (options.out)
        await writeFile(options.out, `${JSON.stringify(result, null, 2)}\n`, "utf8");
    return result;
}
function aggregateWarnings(warnings, blockingIssues) {
    const map = new Map();
    const entries = [
        ...warnings.map((message) => ({ message, severity: "warning" })),
        ...blockingIssues.map((message) => ({ message, severity: "error" }))
    ];
    for (const entry of entries) {
        const code = warningCodeFromMessage(entry.message);
        const current = map.get(code) ?? { code, count: 0, severity: entry.severity, category: warningCategory(code), examples: [] };
        current.count += 1;
        current.severity = current.severity === "error" || entry.severity === "error" ? "error" : "warning";
        if (current.examples.length < 3)
            current.examples.push(entry.message);
        map.set(code, current);
    }
    return [...map.values()].sort((left, right) => severityRank(right.severity) - severityRank(left.severity) || right.count - left.count || left.code.localeCompare(right.code));
}
function packageRiskWarningsFromCaveats(caveats) {
    return caveats.filter((caveat) => /^Zip safety (warning|error|critical): /i.test(caveat));
}
function warningCodeFromMessage(message) {
    const zipSafety = /^Zip safety (?:warning|error|critical): ([A-Z0-9_]+)/i.exec(message);
    if (zipSafety?.[1])
        return zipSafety[1];
    return message.split(":")[0]?.trim() || message;
}
function warningCategory(code) {
    if (code.startsWith("SECURITY_") || code.includes("MACRO") || code.includes("EXTERNAL_LINK"))
        return "security";
    if (code === "ZIP_EMBEDDED_OBJECT" || code === "ZIP_EXTERNAL_RELATIONSHIP")
        return "security";
    if (code.startsWith("ZIP_"))
        return "compatibility";
    if (code === "NATIVE_RENDERER_NOT_RUN" || code === "VERIFY_TIMEOUT" || code.includes("RENDERER"))
        return "environment";
    if (code.includes("REPAIR") || code.includes("OPENABLE") || code.includes("UNSUPPORTED"))
        return "compatibility";
    return "quality";
}
function worstOverflowIssues(issues) {
    return [...issues].sort((left, right) => right.message.length - left.message.length);
}
async function timedPhase(phase, timings, timeoutMs, task) {
    const started = Date.now();
    try {
        const result = timeoutMs ? await withTimeout(task(), timeoutMs, phase) : await task();
        timings.push({ phase, durationMs: Date.now() - started });
        return result;
    }
    catch (error) {
        timings.push({ phase, durationMs: Date.now() - started, timeout: isTimeout(error) });
        throw error;
    }
}
async function withTimeout(promise, timeoutMs, phase) {
    let timer;
    try {
        return await Promise.race([
            promise,
            new Promise((_resolve, reject) => {
                timer = setTimeout(() => reject(new Error(`VERIFY_TIMEOUT:${phase}`)), timeoutMs);
            })
        ]);
    }
    finally {
        if (timer)
            clearTimeout(timer);
    }
}
function isTimeout(error) {
    return error instanceof Error && error.message.startsWith("VERIFY_TIMEOUT:");
}
function severityRank(severity) {
    return severity === "error" ? 2 : 1;
}
function repairForCode(code) {
    if (code === "TEXT_OVERFLOW_RISK")
        return "Shorten text, enlarge the text box, reduce font size, or split the slide/page.";
    if (code === "PDF_TEXT_BLOCKS_ZERO")
        return "Create page previews and inspect them with the AI vision layer, or use native PDF tooling.";
    if (code === "NATIVE_RENDERER_NOT_RUN")
        return "Run verify --native with a trusted renderer profile when repair-dialog evidence is required.";
    if (code === "XLSX_EXTERNAL_LINKS_PRESENT")
        return "Review and sanitize external workbook links before autonomous use.";
    return undefined;
}
function commandForRisk(code, format) {
    if (code === "TEXT_OVERFLOW_RISK")
        return `officegen diagnose <input.${format}> --json`;
    if (code === "PDF_TEXT_BLOCKS_ZERO")
        return "officegen view input.pdf --out .officegen/runs/pdf-view --json";
    if (code === "NATIVE_RENDERER_NOT_RUN")
        return profileCommand("enterprise", `officegen verify input.${format} --native --visual --json`);
    return undefined;
}
function uniqueRepairs(repairs) {
    const seen = new Set();
    return repairs.filter((repair) => {
        const key = `${repair.code}:${repair.command ?? ""}:${repair.reason}`;
        if (seen.has(key))
            return false;
        seen.add(key);
        return true;
    });
}
function profileCommand(profile, command) {
    if (process.platform === "win32")
        return `$env:OFFICEGEN_PROFILE='${profile}'; ${command}`;
    return `OFFICEGEN_PROFILE=${profile} ${command}`;
}
function evaluateGates(gates, inspected, visual, warningCount, noRepairDialogExpected, nativeRenderer) {
    if (!gates)
        return undefined;
    const failed = [];
    const warnings = [];
    const summary = inspected.trusted.summary;
    if (gates.expectedSlides !== undefined && Number(summary.slides ?? 0) !== gates.expectedSlides) {
        failed.push(`GATE_EXPECTED_SLIDES: expected ${gates.expectedSlides}, got ${Number(summary.slides ?? 0)}.`);
    }
    if (gates.expectedPages !== undefined) {
        const pages = Number(summary.pages ?? summary.slides ?? summary.sheets ?? 0);
        if (pages !== gates.expectedPages)
            failed.push(`GATE_EXPECTED_PAGES: expected ${gates.expectedPages}, got ${pages}.`);
    }
    const searchableText = inspected.objectMap.map((entry) => `${entry.text ?? ""}\n${entry.textPreview ?? ""}`).join("\n");
    for (const text of gates.requiredText ?? []) {
        if (!searchableText.includes(text))
            failed.push(`GATE_REQUIRED_TEXT_MISSING: ${text}`);
    }
    for (const text of gates.forbiddenText ?? []) {
        if (searchableText.includes(text))
            failed.push(`GATE_FORBIDDEN_TEXT_PRESENT: ${text}`);
    }
    if (gates.maxBlankPages !== undefined) {
        if (!visual)
            failed.push("GATE_MAX_BLANK_PAGES_UNEVALUATED: run verify with visual enabled to evaluate maxBlankPages.");
        else if (visual.blankPages > gates.maxBlankPages)
            failed.push(`GATE_MAX_BLANK_PAGES: expected <= ${gates.maxBlankPages}, got ${visual.blankPages}.`);
    }
    if (gates.maxWarnings !== undefined && warningCount > gates.maxWarnings) {
        failed.push(`GATE_MAX_WARNINGS: expected <= ${gates.maxWarnings}, got ${warningCount}.`);
    }
    if (gates.requireNoRepairDialog) {
        if (!noRepairDialogExpected)
            failed.push("GATE_REPAIR_DIALOG_EXPECTED: repair dialog risk was detected.");
        if (!nativeRenderer?.ok)
            failed.push("GATE_REPAIR_DIALOG_NATIVE_UNEVALUATED: run verify with native enabled to prove repair-dialog behavior.");
        if (inspected.trusted.caveats.some((caveat) => /repair/i.test(caveat))) {
            warnings.push("GATE_REPAIR_DIALOG_EVIDENCE_LIMITED: inspect caveats mention repair risk.");
        }
    }
    return { passed: failed.length === 0, failed, warnings };
}
function gatesNeedFullText(gates) {
    return Boolean(gates?.requiredText?.length || gates?.forbiddenText?.length);
}
async function verifyVisual(input, config, mode) {
    try {
        const preview = await view(input, { format: "png", maxPages: 10, config, mode });
        const raster = preview.rasterDiagnostics;
        if (raster) {
            return {
                fidelity: preview.fidelity === "native" ? "native" : "approximate",
                pagesChecked: preview.pages.length,
                blankPages: raster.blankPages.length,
                identicalPages: raster.identicalPages,
                pixelDensityWarnings: raster.pixelDensityWarnings,
                allPagesIdentical: raster.allPagesIdentical
            };
        }
    }
    catch (error) {
        if (mode === "proof") {
            throw error;
        }
        const preview = await view(input, { format: "svg", maxPages: 10, config });
        const blankPages = preview.pages.filter((page) => !page.objectMap.some(hasVisiblePreviewObject)).length;
        return {
            fidelity: "approximate",
            pagesChecked: preview.pages.length,
            blankPages,
            identicalPages: [],
            pixelDensityWarnings: [`VISUAL_RASTER_UNAVAILABLE: raster pixel diagnostics could not run (${error instanceof Error ? error.message : String(error)}).`],
            rasterDiagnosticsUnavailable: true
        };
    }
    const preview = await view(input, { format: "svg", maxPages: 10, config });
    const blankPages = preview.pages.filter((page) => !page.objectMap.some(hasVisiblePreviewObject)).length;
    return {
        fidelity: "approximate",
        pagesChecked: preview.pages.length,
        blankPages,
        identicalPages: [],
        pixelDensityWarnings: ["VISUAL_RASTER_UNAVAILABLE: raster pixel diagnostics did not produce page density metadata."],
        rasterDiagnosticsUnavailable: true
    };
}
function hasVisiblePreviewObject(entry) {
    const text = `${entry.text ?? ""}${entry.textPreview ?? ""}`.trim();
    if (text)
        return true;
    return ["picture", "image", "chart"].includes(entry.kind);
}
async function verifyNative(input, options, artifacts) {
    if (!["pptx", "docx", "xlsx"].includes(input.format))
        return { attempted: false, ok: false, message: "Native renderer verification is only available for Office inputs." };
    if (!input.path)
        return { attempted: false, ok: false, message: "Native renderer verification requires an input file path." };
    const artifactDir = managedVerifyArtifactDir(input.path, options.out);
    await mkdir(artifactDir, { recursive: true });
    const pdfPath = path.join(artifactDir, `${path.basename(input.path, path.extname(input.path))}.native.pdf`);
    try {
        const exported = await exportDocument(input.path, { to: "pdf", mode: options.mode === "proof" ? "proof" : "native", out: pdfPath, config: options.config, timeoutMs: options.timeoutMs });
        const pdf = await PDFDocument.load(await import("node:fs/promises").then((fs) => fs.readFile(pdfPath)));
        artifacts.nativePdf = {
            artifactId: "verify-native-pdf",
            role: "native-render",
            path: pdfPath,
            format: "pdf",
            managed: true,
            exists: true,
            sourceCommand: "verify --native"
        };
        return {
            attempted: true,
            ok: true,
            artifact: pdfPath,
            repairDialogExpected: exported.renderer?.repairDialogExpected,
            renderer: exported.nativeProof?.renderer ?? nativeProofRenderer(exported.renderer?.id),
            message: `Native renderer produced ${pdf.getPageCount()} PDF page(s) with ${exported.renderer?.id ?? "renderer"}.`
        };
    }
    catch (error) {
        await rm(pdfPath, { force: true });
        return { attempted: true, ok: false, message: error instanceof Error ? error.message : String(error) };
    }
}
function managedVerifyArtifactDir(inputPath, reportOut) {
    if (reportOut) {
        return path.join(path.dirname(reportOut), `${path.basename(reportOut, path.extname(reportOut))}-artifacts`);
    }
    return path.join(path.dirname(inputPath), ".officegen", "verify-artifacts");
}
function nativeProofFromRenderer(nativeRenderer, requested, format) {
    if (!requested) {
        return ["pptx", "docx", "xlsx"].includes(format)
            ? { status: "not_run", reason: "Native proof was not requested; use --native or --mode proof for final layout evidence." }
            : { status: "not_run", reason: "Native proof is only relevant for Office inputs." };
    }
    if (!nativeRenderer)
        return { status: "failed", reason: "Native proof was requested but did not produce a renderer result." };
    if (nativeRenderer.ok) {
        return {
            status: "passed",
            renderer: nativeRenderer.renderer ?? nativeProofRenderer(nativeRenderer.message),
            artifact: nativeRenderer.artifact,
            reason: nativeRenderer.message
        };
    }
    const reason = nativeRenderer.message ?? "Native renderer verification did not complete.";
    const policyHint = /external process|externalProcess|denied|disabled/i.test(reason)
        ? " Run officegen renderer doctor --json and officegen config show --json to confirm native renderer policy before retrying."
        : "";
    return {
        status: /not found|requires|disabled|denied|unavailable/i.test(reason) ? "unavailable" : "failed",
        reason: `${reason}${policyHint}`
    };
}
function nativeProofRenderer(message) {
    if (/powerpoint-com/i.test(message ?? ""))
        return "powerpoint";
    if (/libreoffice/i.test(message ?? ""))
        return "libreoffice";
    if (/word-com|excel-com/i.test(message ?? ""))
        return "office-com";
    return undefined;
}
function buildVerificationReport(context) {
    const summary = (context.inspected?.trusted.summary ?? {});
    const caveats = context.inspected?.trusted.caveats ?? [];
    const riskFlags = asArrayRecord(context.inspected?.untrusted?.pdfGraph)
        .riskFlags;
    const securityIssues = [
        ...context.warningSummary.filter((item) => item.category === "security").flatMap((item) => item.examples),
        ...(riskFlags ?? []).filter((flag) => flag.severity === "warning" || flag.severity === "error").map((flag) => `${flag.code}: ${flag.message}`)
    ];
    const goalIssues = [...(context.gateResult?.failed ?? []), ...(context.gateResult?.warnings ?? [])];
    const gates = {
        schema: gate(context.openable && !context.blockingIssues.includes("INPUT_NOT_OPENABLE"), {
            format: context.format,
            openable: context.openable
        }, context.openable ? [] : ["INPUT_NOT_OPENABLE"]),
        package: gate(context.openable && context.noRepairDialogExpected, {
            noRepairDialogExpected: context.noRepairDialogExpected,
            caveatCount: caveats.length,
            summary
        }, [
            ...context.blockingIssues.filter((issue) => /REPAIR|OPENABLE|PACKAGE|ZIP/i.test(issue)),
            ...caveats.filter((caveat) => /repair|zip|package/i.test(caveat))
        ]),
        semantic: gate(Boolean(context.inspected), {
            objectCount: context.inspected?.objectMap.length ?? 0,
            textBlocks: summary.textBlocks,
            textObjects: summary.textObjects
        }, context.blockingIssues.filter((issue) => /TEXT|SEMANTIC|REQUIRED|FORBIDDEN/i.test(issue))),
        visual: context.visual || context.visualDiff
            ? gate(!(context.visualDiff?.status === "blocked") && !visualHasQualityFailures(context.visual), {
                fidelity: context.visual?.fidelity ?? context.visualDiff?.fidelity,
                pagesChecked: context.visual?.pagesChecked,
                blankPages: context.visual?.blankPages ?? 0,
                identicalPages: context.visual?.identicalPages ?? [],
                allPagesIdentical: context.visual?.allPagesIdentical ?? false,
                pixelDensityWarnings: context.visual?.pixelDensityWarnings?.length ?? 0,
                rasterDiagnosticsUnavailable: context.visual?.rasterDiagnosticsUnavailable ?? false,
                diffStatus: context.visualDiff?.status
            }, context.warnings.filter((issue) => /^VISUAL_|GATE_MAX_BLANK_PAGES/.test(issue)))
            : skippedGate("Visual verification was not requested."),
        native: context.nativeRenderer
            ? gate(context.nativeRenderer.ok, {
                attempted: context.nativeRenderer.attempted,
                ok: context.nativeRenderer.ok,
                repairDialogExpected: context.nativeRenderer.repairDialogExpected,
                artifact: context.nativeRenderer.artifact,
                nativeProof: context.nativeProof
            }, context.blockingIssues.filter((issue) => /NATIVE|REPAIR_DIALOG/i.test(issue)))
            : skippedGate(context.nativeProof.reason ?? "Native renderer verification was not requested."),
        security: securityIssues.length
            ? { status: "warning", summary: { riskFlagCount: riskFlags?.length ?? 0 }, issues: securityIssues }
            : { status: "pass", summary: { riskFlagCount: riskFlags?.length ?? 0 }, issues: [] },
        accessibility: skippedGate("Accessibility checks are not implemented for this format yet."),
        goal: context.gateResult
            ? gate(context.gateResult.passed, {
                passed: context.gateResult.passed,
                failed: context.gateResult.failed.length,
                warnings: context.gateResult.warnings.length
            }, goalIssues)
            : skippedGate("No explicit verification gates were supplied.")
    };
    return {
        schema: "officegen.verify@2",
        version: 2,
        format: context.format,
        readiness: context.readiness,
        score: context.score,
        partial: context.partial,
        gates,
        issues: reportIssues(context.warningSummary, context.blockingIssues, gates),
        artifacts: reportArtifacts(context.artifacts),
        recommendedRepairs: context.recommendedRepairs
    };
}
function visualHasQualityFailures(visual) {
    return Boolean(visual
        && (visual.blankPages > 0 || visualPixelQualityFailures(visual).length > 0));
}
function visualBlockingIssues(visual) {
    if (!visualHasQualityFailures(visual) || !visual)
        return [];
    const pixelWarnings = visualPixelQualityFailures(visual);
    const reasons = [
        visual.blankPages > 0 ? `${visual.blankPages} blank preview page(s)` : undefined,
        pixelWarnings.length > 0 ? `${pixelWarnings.length} raster pixel-density warning(s)` : undefined
    ].filter((reason) => Boolean(reason));
    return [`VISUAL_GATE_FAILED: visual verification failed (${reasons.join(", ")}).`];
}
function visualPixelQualityFailures(visual) {
    return visual.pixelDensityWarnings.filter((warning) => !warning.startsWith("VISUAL_RASTER_UNAVAILABLE"));
}
function gate(ok, summary, issues) {
    return {
        status: ok ? issues.length ? "warning" : "pass" : "fail",
        summary,
        issues
    };
}
function skippedGate(reason) {
    return { status: "skipped", issues: [reason] };
}
function reportIssues(warningSummary, blockingIssues, gates) {
    const gateByIssue = new Map();
    for (const [gateName, projection] of Object.entries(gates)) {
        for (const issue of projection.issues) {
            if (!gateByIssue.has(issue))
                gateByIssue.set(issue, gateName);
        }
    }
    const warnings = warningSummary.flatMap((item) => item.examples.map((message) => ({
        code: item.code,
        severity: item.severity,
        category: item.category,
        message,
        gate: gateByIssue.get(message)
    })));
    const blocking = blockingIssues.map((message) => ({
        code: message.split(":")[0]?.trim() || message,
        severity: "error",
        category: warningCategory(message.split(":")[0]?.trim() || message),
        message,
        gate: gateByIssue.get(message)
    }));
    return [...warnings, ...blocking];
}
function reportArtifacts(artifacts) {
    return Object.entries(artifacts).flatMap(([key, value]) => {
        const record = value && typeof value === "object" && !Array.isArray(value) ? value : {};
        const pathValue = typeof record.path === "string" ? record.path : typeof value === "string" ? value : undefined;
        if (!pathValue && !Object.keys(record).length)
            return [];
        return [{
                artifactId: String(record.artifactId ?? key),
                role: String(record.role ?? key),
                path: pathValue,
                format: typeof record.format === "string" ? record.format : undefined,
                managed: record.managed !== false,
                exists: typeof record.exists === "boolean" ? record.exists : undefined,
                sourceCommand: typeof record.sourceCommand === "string" ? record.sourceCommand : undefined
            }];
    });
}
function asArrayRecord(value) {
    return value && typeof value === "object" && !Array.isArray(value) ? value : {};
}
export const verifyDocument = verify;
//# sourceMappingURL=verify.js.map