import { extname } from "node:path";
import { OfficegenError } from "../../core/dist/index.js";
import { assertPdfStandardFontText, writeOutput } from "./shared.js";
import { PDFDocument, StandardFonts, rgb } from "pdf-lib";
export async function render(ir, options = {}) {
    const target = resolveRenderTarget(ir, options);
    if (target === "pptx")
        return renderPptx(ir, options);
    if (target === "docx")
        return renderDocx(ir, options);
    if (target === "xlsx")
        return renderXlsx(ir, options);
    return renderPdf(ir, options);
}
export const renderDocument = render;
async function renderPptx(ir, options) {
    const mod = (await import("pptxgenjs"));
    const pptx = new mod.default();
    pptx.layout = "LAYOUT_WIDE";
    const theme = ir.design?.theme;
    if (theme) {
        pptx.author = "officegen";
        pptx.theme = {
            headFontFace: String(theme.headFontFace ?? "Aptos Display"),
            bodyFontFace: String(theme.bodyFontFace ?? "Aptos"),
            lang: String(theme.lang ?? "en-US")
        };
    }
    const slides = ir.slides?.length ? ir.slides : normalizedSections(ir);
    const diagnostics = [];
    const caveats = ["PPTX generation supports native text boxes, lists, images, callouts, tables, and basic editable Office charts."];
    for (const [sectionIndex, section] of slides.entries()) {
        const slide = pptx.addSlide();
        const palette = pptxPalette(ir);
        slide.background = { color: palette.background };
        slide.addText(section.title ?? ir.title ?? "Untitled", { x: 0.55, y: 0.32, w: 12.2, h: 0.45, fontSize: 26, bold: true, color: palette.text });
        let y = 1.0;
        const blocks = section.blocks?.length ? section.blocks : blocksFromBody(section.body);
        for (const block of blocks) {
            const type = block.type ?? "paragraph";
            if (type === "heading") {
                slide.addText(block.text ?? block.title ?? "", { x: 0.65, y, w: 12, h: 0.35, fontSize: 18, bold: true, color: palette.accent, fit: "shrink" });
                y += 0.48;
                continue;
            }
            if (type === "list" || block.items?.length) {
                const items = block.items?.length ? block.items : toLines(block.text);
                slide.addText(items.map((item) => ({ text: item, options: { bullet: { type: "ul" } } })), { x: 0.78, y, w: 11.6, h: Math.min(2.8, items.length * 0.28 + 0.15), fontSize: 13.5, color: palette.text, fit: "shrink" });
                y += Math.min(2.9, items.length * 0.3 + 0.25);
                continue;
            }
            if (type === "table" && block.rows?.length) {
                const allRows = normalizeTableRows(block.rows);
                const rows = allRows.slice(0, 18);
                if (allRows.length > rows.length) {
                    diagnostics.push({
                        code: "RENDER_TABLE_TRUNCATED",
                        severity: "warning",
                        slide: sectionIndex + 1,
                        rows: allRows.length,
                        renderedRows: rows.length,
                        message: "Table block exceeded the per-slide row limit and was truncated."
                    });
                }
                slide.addTable(rows, {
                    x: 0.65,
                    y,
                    w: 12,
                    h: Math.min(3.6, rows.length * 0.32 + 0.15),
                    fontSize: 10.5,
                    color: palette.text,
                    border: { type: "solid", color: "D1D5DB", pt: 0.7 },
                    fill: { color: "FFFFFF" },
                    margin: 0.05
                });
                y += Math.min(3.75, rows.length * 0.34 + 0.3);
                if (y > 7.0)
                    diagnostics.push(renderOverflowDiagnostic(sectionIndex + 1, type, y));
                continue;
            }
            if (type === "callout") {
                slide.addText(block.text ?? "", { x: 0.65, y, w: 12, h: 0.7, fontSize: 15, bold: true, color: palette.text, fill: { color: palette.callout }, fit: "shrink", margin: 0.12 });
                y += 0.88;
                if (y > 7.0)
                    diagnostics.push(renderOverflowDiagnostic(sectionIndex + 1, type, y));
                continue;
            }
            if (type === "image" && block.path) {
                slide.addImage({ path: block.path, x: 0.65, y, w: 4.2, h: 2.4 });
                y += 2.6;
                if (y > 7.0)
                    diagnostics.push(renderOverflowDiagnostic(sectionIndex + 1, type, y));
                continue;
            }
            if (type === "chart") {
                const chart = nativeChartData(block);
                if (chart) {
                    const chartType = pptxChartType(pptx, block.chartType);
                    slide.addChart(chartType, chart.data, {
                        x: 0.75,
                        y,
                        w: 11.6,
                        h: 3.0,
                        showTitle: Boolean(block.title),
                        title: block.title,
                        showLegend: chart.data.length > 1,
                        showValue: false,
                        catAxisLabelFontFace: String(ir.design?.theme?.bodyFontFace ?? "Aptos"),
                        valAxisLabelFontFace: String(ir.design?.theme?.bodyFontFace ?? "Aptos"),
                        valAxisMinVal: 0,
                        showCatName: true
                    });
                    y += 3.25;
                }
                else {
                    diagnostics.push({
                        code: "RENDER_CHART_DATA_MISSING",
                        severity: "warning",
                        slide: sectionIndex + 1,
                        message: "Chart block was rendered as a labeled placeholder because categories/values or series data were missing."
                    });
                    slide.addText(`Chart: ${block.title ?? block.text ?? "data"}`, { x: 0.65, y, w: 12, h: 0.5, fontSize: 14, color: palette.accent, italic: true });
                    y += 0.65;
                }
                if (y > 7.0)
                    diagnostics.push(renderOverflowDiagnostic(sectionIndex + 1, type, y));
                continue;
            }
            const text = block.text ?? "";
            if (text) {
                slide.addText(text, { x: 0.65, y, w: 12, h: Math.min(1.4, Math.max(0.35, text.length / 120 * 0.35)), fontSize: 13.5, color: palette.text, breakLine: false, fit: "shrink" });
                y += Math.min(1.5, Math.max(0.45, text.length / 120 * 0.38));
                if (y > 7.0)
                    diagnostics.push(renderOverflowDiagnostic(sectionIndex + 1, type, y));
            }
        }
    }
    if (diagnostics.length)
        caveats.push("Some PPTX blocks exceeded the simple layout budget; inspect/view/diagnose the output before using it as a final deck.");
    const bytes = await pptx.write({ outputType: "nodebuffer" });
    await writeOutput(options.out, bytes);
    return {
        schema: "officegen.render.result@1.2",
        target: "pptx",
        out: options.out,
        bytes: options.out ? undefined : bytes,
        caveats,
        diagnostics
    };
}
async function renderDocx(ir, options) {
    const docx = await import("docx");
    const sections = normalizedSections(ir);
    const children = [
        new docx.Paragraph({ text: ir.title ?? "Untitled", heading: docx.HeadingLevel.TITLE }),
        ...sections.flatMap((section) => [
            ...(section.title ? [new docx.Paragraph({ text: section.title, heading: docx.HeadingLevel.HEADING_1 })] : []),
            ...docxChildrenFromBlocks(docx, section.blocks?.length ? section.blocks : blocksFromBody(section.body))
        ])
    ];
    const document = new docx.Document({
        styles: {
            paragraphStyles: [
                {
                    id: "OfficegenBody",
                    name: "Officegen Body",
                    basedOn: "Normal",
                    next: "OfficegenBody",
                    run: { font: "Aptos", size: 22 },
                    paragraph: { spacing: { after: 120 } }
                }
            ]
        },
        sections: [{
                properties: {},
                headers: ir.header ? { default: new docx.Header({ children: [new docx.Paragraph(ir.header)] }) } : undefined,
                footers: ir.footer ? { default: new docx.Footer({ children: [new docx.Paragraph(ir.footer)] }) } : undefined,
                children: children
            }]
    });
    const bytes = await docx.Packer.toBuffer(document);
    await writeOutput(options.out, bytes);
    return {
        schema: "officegen.render.result@1.2",
        target: "docx",
        out: options.out,
        bytes: options.out ? undefined : bytes,
        caveats: ["DOCX generation supports headings, paragraphs, lists, tables, and simple headers/footers; tracked changes/comments are available through edit operations."]
    };
}
async function renderXlsx(ir, options) {
    const ExcelJS = (await import("exceljs"));
    const workbook = new ExcelJS.default.Workbook();
    const sheets = ir.sheets?.length ? ir.sheets : [{ name: ir.title ?? "Sheet1", rows: rowsFromSections(ir) }];
    for (const sheetSpec of sheets) {
        const sheet = workbook.addWorksheet(sanitizeSheetName(sheetSpec.name ?? "Sheet"));
        const rows = sheetSpec.rows ?? [];
        if (rows.length && !Array.isArray(rows[0])) {
            const keys = Object.keys(rows[0]);
            sheet.addRow(keys);
            for (const row of rows)
                sheet.addRow(keys.map((key) => row[key]));
            if (rows.length) {
                sheet.addTable({
                    name: sanitizeTableName(sheetSpec.tableName ?? `${sanitizeSheetName(sheetSpec.name ?? "Sheet")}Table`),
                    ref: "A1",
                    headerRow: true,
                    totalsRow: false,
                    style: { theme: "TableStyleMedium2", showRowStripes: true },
                    columns: keys.map((key) => ({ name: key })),
                    rows: rows.map((row) => keys.map((key) => row[key]))
                });
            }
        }
        else {
            for (const row of rows)
                sheet.addRow(row);
            if (rows.length > 1 && Array.isArray(rows[0])) {
                const headers = rows[0].map((value, index) => String(value ?? `Column${index + 1}`));
                sheet.addTable({
                    name: sanitizeTableName(sheetSpec.tableName ?? `${sanitizeSheetName(sheetSpec.name ?? "Sheet")}Table`),
                    ref: "A1",
                    headerRow: true,
                    totalsRow: false,
                    style: { theme: "TableStyleMedium2", showRowStripes: true },
                    columns: headers.map((name) => ({ name })),
                    rows: rows.slice(1)
                });
            }
        }
        sheet.columns?.forEach((column) => {
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
        caveats: ["XLSX generation supports worksheets, typed rows, and native Excel table objects; charts/pivots/slicers are inspected and guarded but not recalculated internally."]
    };
}
async function renderPdf(ir, options) {
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
            if (y < 54)
                break;
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
function resolveRenderTarget(ir, options) {
    const explicit = options.target ?? ir.kind;
    const outputTarget = inferTargetFromOutput(options.out);
    if (explicit !== undefined) {
        const target = parseRenderTarget(explicit, "render target");
        assertOutputTargetMatches(target, outputTarget, options.out);
        return target;
    }
    if (outputTarget !== undefined)
        return outputTarget;
    if (ir.targets !== undefined && ir.targets.length > 0) {
        return parseRenderTarget(ir.targets[0], "IR targets[0]");
    }
    return "pdf";
}
function inferTargetFromOutput(out) {
    if (!out)
        return undefined;
    const ext = extname(out).slice(1).toLowerCase();
    if (!ext)
        return undefined;
    return parseRenderTarget(ext, "output extension");
}
function isRenderTarget(value) {
    return value === "pptx" || value === "docx" || value === "xlsx" || value === "pdf";
}
function parseRenderTarget(value, source) {
    if (isRenderTarget(value))
        return value;
    throw new OfficegenError("EXPORT_UNSUPPORTED", `Unsupported ${source}: ${String(value)}. Supported render targets are pptx, docx, xlsx, and pdf.`, { source, value: String(value), supported: ["pptx", "docx", "xlsx", "pdf"] });
}
function assertOutputTargetMatches(target, outputTarget, out) {
    if (outputTarget === undefined || outputTarget === target)
        return;
    throw new OfficegenError("TARGET_EXTENSION_MISMATCH", `Render target ${target} does not match output extension .${outputTarget}${out ? ` for ${out}` : ""}.`, { target, outputTarget, ...(out ? { out } : {}) });
}
function toLines(value) {
    if (Array.isArray(value))
        return value.flatMap((item) => String(item).split(/\r?\n/));
    return String(value ?? "").split(/\r?\n/).filter(Boolean);
}
function normalizedSections(ir) {
    const sections = ir.sections?.length ? ir.sections : [{ title: ir.title ?? "Untitled", body: "" }];
    return sections.map((section) => ({
        ...section,
        title: section.title ?? ir.title ?? "Untitled",
        body: section.body ?? bodyFromBlocks(section.blocks)
    }));
}
function bodyFromBlocks(blocks) {
    return (blocks ?? [])
        .map((block) => block.text ?? block.items?.join("\n") ?? (block.rows ? normalizeTableRows(block.rows).map((row) => row.join("\t")).join("\n") : ""))
        .filter(Boolean)
        .join("\n");
}
function rowsFromSections(ir) {
    const rows = (ir.sections ?? []).flatMap((section) => [
        ...(section.rows ?? []),
        ...section.blocks?.flatMap((block) => block.rows ?? []) ?? []
    ]);
    return rows.length ? rows : [["title", "body"], [ir.title ?? "Untitled", bodyFromBlocks(ir.sections?.[0]?.blocks)]];
}
function sanitizeSheetName(name) {
    return name.replace(/[\[\]*?/\\:]/g, " ").slice(0, 31) || "Sheet";
}
function sanitizeTableName(name) {
    const cleaned = name.replace(/[^A-Za-z0-9_]/g, "_").replace(/^[^A-Za-z_]+/, "T_").slice(0, 120);
    return cleaned || "OfficegenTable";
}
function pptxPalette(ir) {
    const colors = ir.design?.colors ?? {};
    return {
        background: cleanColor(colors.background ?? colors.bg ?? "FFFFFF"),
        text: cleanColor(colors.text ?? "111827"),
        accent: cleanColor(colors.accent ?? colors.primary ?? "2563EB"),
        callout: cleanColor(colors.callout ?? colors.muted ?? "EEF2FF")
    };
}
function cleanColor(value) {
    const text = String(value ?? "").replace(/^#/, "").toUpperCase();
    return /^[0-9A-F]{6}$/.test(text) ? text : "111827";
}
function blocksFromBody(body) {
    const lines = toLines(body);
    return lines.length ? [{ type: "paragraph", text: lines.join("\n") }] : [];
}
function normalizeTableRows(rows) {
    if (!rows.length)
        return [];
    if (!Array.isArray(rows[0])) {
        const keys = Object.keys(rows[0]);
        return [keys, ...rows.map((row) => keys.map((key) => row[key]))];
    }
    return rows;
}
function nativeChartData(block) {
    const series = block.series?.length
        ? block.series.map((item, index) => ({
            name: String(item.name ?? `Series ${index + 1}`),
            labels: (item.labels ?? block.categories ?? []).map(String),
            values: (item.values ?? []).map(Number).filter(Number.isFinite)
        }))
        : [{
                name: String(block.title ?? "Series 1"),
                labels: (block.categories ?? []).map(String),
                values: (block.values ?? []).map(Number).filter(Number.isFinite)
            }];
    const valid = series.filter((item) => item.labels.length && item.values.length && item.labels.length === item.values.length);
    return valid.length ? { data: valid } : undefined;
}
function pptxChartType(pptx, chartType) {
    const requested = String(chartType ?? "bar").toLowerCase();
    const types = pptx.ChartType ?? {};
    if (requested === "line")
        return types.line ?? types.bar;
    if (requested === "pie" || requested === "doughnut")
        return types.pie ?? types.bar;
    if (requested === "scatter")
        return types.scatter ?? types.bar;
    if (requested === "area")
        return types.area ?? types.bar;
    return types.bar;
}
function docxChildrenFromBlocks(docx, blocks) {
    const children = [];
    for (const block of blocks) {
        const type = block.type ?? "paragraph";
        if (type === "heading") {
            children.push(new docx.Paragraph({ text: block.text ?? block.title ?? "", heading: docx.HeadingLevel.HEADING_2 }));
        }
        else if (type === "list" || block.items?.length) {
            for (const item of block.items?.length ? block.items : toLines(block.text)) {
                children.push(new docx.Paragraph({ text: item, bullet: { level: 0 } }));
            }
        }
        else if (type === "table" && block.rows?.length) {
            children.push(new docx.Table({
                width: { size: 100, type: docx.WidthType.PERCENTAGE },
                rows: normalizeTableRows(block.rows).map((row) => new docx.TableRow({
                    children: row.map((cell) => new docx.TableCell({
                        children: [new docx.Paragraph(String(cell ?? ""))]
                    }))
                }))
            }));
        }
        else if (type === "callout") {
            children.push(new docx.Paragraph({ children: [new docx.TextRun({ text: block.text ?? "", bold: true })] }));
        }
        else {
            for (const line of toLines(block.text))
                children.push(new docx.Paragraph({ text: line, style: "OfficegenBody" }));
        }
    }
    return children;
}
function renderOverflowDiagnostic(slide, blockType, y) {
    return {
        code: "TEXT_OVERFLOW",
        severity: "warning",
        slide,
        blockType,
        y: Number(y.toFixed(2)),
        message: "Block layout passed the nominal slide height; run diagnose/view and split content if needed."
    };
}
//# sourceMappingURL=render.js.map