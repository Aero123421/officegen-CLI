import { inspect } from "./inspect.js";
import { AGENT_UNTRUSTED_INSTRUCTION, escapeHtml, escapeXml, makeStableObjectId } from "./shared.js";
export async function view(input, options = {}) {
    const inspected = isInspectResult(input) ? input : await inspect(input, { format: undefined, depth: "shallow", config: options.config });
    const pages = toPages(inspected, options);
    return {
        schema: "officegen.view.result@1.2",
        fidelity: "approximate",
        caveats: [
            "Approximate SVG/HTML view only; fonts, wrapping, theme effects, animations, and native layout may differ.",
            ...inspected.trusted.caveats
        ],
        pages,
        objectMap: pages.flatMap((page) => page.objectMap),
        trusted: {
            sourceSchema: inspected.schema,
            sourceFormat: inspected.trusted.format,
            generatedAt: new Date().toISOString()
        },
        agentInstruction: AGENT_UNTRUSTED_INSTRUCTION
    };
}
export const viewDocument = view;
function isInspectResult(value) {
    return Boolean(value && typeof value === "object" && value.schema === "officegen.inspect.result@1.2");
}
function toPages(inspected, options) {
    const format = options.format ?? "svg";
    const maxPages = options.maxPages ?? 50;
    if (inspected.trusted.format === "pptx") {
        const slides = (inspected.untrusted.slides ?? []).slice(0, maxPages);
        return slides.map((slide, index) => buildSlidePage(slide, index + 1, format));
    }
    if (inspected.trusted.format === "docx") {
        return [buildDocxPage((inspected.untrusted.paragraphs ?? []).slice(0, 200), format)];
    }
    if (inspected.trusted.format === "xlsx") {
        const sheets = (inspected.untrusted.sheets ?? []).slice(0, maxPages);
        return sheets.map((sheet, index) => buildSheetPage(sheet, index + 1, format));
    }
    if (inspected.trusted.format === "pdf") {
        const pages = (inspected.untrusted.pages ?? []).slice(0, maxPages);
        return pages.map((page, index) => buildPdfPage(page, index + 1, format));
    }
    return [];
}
function buildSlidePage(slide, page, format) {
    const objects = (slide.textObjects ?? []).map((entry, index) => ({
        ...entry,
        bounds: { x: 48, y: 48 + index * 48, width: 864, height: 40 }
    }));
    const stableObjectId = String(slide.stableObjectId ?? makeStableObjectId("pptx", "deck", "slide", page));
    if (format === "html") {
        return {
            page,
            stableObjectId,
            format,
            content: `<section data-stable-object-id="${escapeHtml(stableObjectId)}" style="width:960px;height:540px;padding:48px;background:#fff;color:#111;font-family:Arial,sans-serif">${objects.map((object) => `<p data-stable-object-id="${escapeHtml(object.stableObjectId)}">${escapeHtml(object.text ?? "")}</p>`).join("")}</section>`,
            objectMap: objects
        };
    }
    return {
        page,
        stableObjectId,
        format,
        content: `<svg xmlns="http://www.w3.org/2000/svg" width="960" height="540" viewBox="0 0 960 540" data-stable-object-id="${escapeXml(stableObjectId)}"><rect width="960" height="540" fill="#fff"/><rect x="0" y="0" width="960" height="540" fill="none" stroke="#d0d7de"/>${objects.map((object) => `<text x="${object.bounds?.x ?? 48}" y="${(object.bounds?.y ?? 48) + 24}" font-family="Arial, sans-serif" font-size="24" fill="#111" data-stable-object-id="${escapeXml(object.stableObjectId)}">${escapeXml(object.text ?? "")}</text>`).join("")}</svg>`,
        objectMap: objects
    };
}
function buildDocxPage(paragraphs, format) {
    const objects = paragraphs
        .filter((paragraph) => paragraph.text)
        .map((paragraph, index) => ({
        stableObjectId: String(paragraph.stableObjectId),
        kind: "paragraph",
        text: String(paragraph.text ?? ""),
        bounds: { x: 72, y: 72 + index * 28, width: 468, height: 24 },
        untrusted: true
    }));
    if (format === "html") {
        return {
            page: 1,
            stableObjectId: makeStableObjectId("docx", "document", "page", 1),
            format,
            content: `<article style="max-width:720px;padding:72px;font-family:Georgia,serif;line-height:1.5">${objects.map((object) => `<p data-stable-object-id="${escapeHtml(object.stableObjectId)}">${escapeHtml(object.text ?? "")}</p>`).join("")}</article>`,
            objectMap: objects
        };
    }
    return {
        page: 1,
        stableObjectId: makeStableObjectId("docx", "document", "page", 1),
        format,
        content: `<svg xmlns="http://www.w3.org/2000/svg" width="612" height="792" viewBox="0 0 612 792"><rect width="612" height="792" fill="#fff"/><rect x="36" y="36" width="540" height="720" fill="none" stroke="#d0d7de"/>${objects.map((object) => `<text x="72" y="${(object.bounds?.y ?? 72) + 16}" font-family="Georgia, serif" font-size="14" fill="#111" data-stable-object-id="${escapeXml(object.stableObjectId)}">${escapeXml(object.text ?? "")}</text>`).join("")}</svg>`,
        objectMap: objects
    };
}
function buildSheetPage(sheet, page, format) {
    const cells = (sheet.cells ?? []).slice(0, 120);
    const objects = cells.map((cell, index) => ({
        stableObjectId: String(cell.stableObjectId),
        kind: "cell",
        label: String(cell.ref ?? ""),
        text: String(cell.value ?? ""),
        bounds: { x: 32 + (index % 6) * 120, y: 48 + Math.floor(index / 6) * 32, width: 120, height: 32 },
        untrusted: true
    }));
    const stableObjectId = String(sheet.stableObjectId ?? makeStableObjectId("xlsx", "workbook", "sheet", page));
    if (format === "html") {
        return {
            page,
            stableObjectId,
            format,
            content: `<table data-stable-object-id="${escapeHtml(stableObjectId)}" style="border-collapse:collapse;font-family:Arial,sans-serif">${objects.map((object) => `<tr><th style="border:1px solid #d0d7de;padding:4px 8px">${escapeHtml(object.label ?? "")}</th><td data-stable-object-id="${escapeHtml(object.stableObjectId)}" style="border:1px solid #d0d7de;padding:4px 8px">${escapeHtml(object.text ?? "")}</td></tr>`).join("")}</table>`,
            objectMap: objects
        };
    }
    return {
        page,
        stableObjectId,
        format,
        content: `<svg xmlns="http://www.w3.org/2000/svg" width="800" height="600" viewBox="0 0 800 600"><rect width="800" height="600" fill="#fff"/>${objects.map((object) => `<g data-stable-object-id="${escapeXml(object.stableObjectId)}"><rect x="${object.bounds?.x}" y="${object.bounds?.y}" width="120" height="32" fill="#fff" stroke="#d0d7de"/><text x="${(object.bounds?.x ?? 0) + 6}" y="${(object.bounds?.y ?? 0) + 21}" font-family="Arial, sans-serif" font-size="12">${escapeXml(`${object.label}: ${object.text}`)}</text></g>`).join("")}</svg>`,
        objectMap: objects
    };
}
function buildPdfPage(pageInfo, page, format) {
    const width = Number(pageInfo.width ?? 612);
    const height = Number(pageInfo.height ?? 792);
    const stableObjectId = String(pageInfo.stableObjectId ?? makeStableObjectId("pdf", "document", "page", page));
    const objectMap = [];
    if (format === "html") {
        return {
            page,
            stableObjectId,
            format,
            content: `<section data-stable-object-id="${escapeHtml(stableObjectId)}" style="width:${width}px;height:${height}px;border:1px solid #d0d7de;background:#fff;font-family:Arial,sans-serif"><p style="padding:24px;color:#57606a">PDF page ${page}</p></section>`,
            objectMap
        };
    }
    return {
        page,
        stableObjectId,
        format,
        content: `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}" data-stable-object-id="${escapeXml(stableObjectId)}"><rect width="${width}" height="${height}" fill="#fff" stroke="#d0d7de"/><text x="24" y="40" font-family="Arial, sans-serif" font-size="16" fill="#57606a">PDF page ${page}</text></svg>`,
        objectMap
    };
}
//# sourceMappingURL=view.js.map