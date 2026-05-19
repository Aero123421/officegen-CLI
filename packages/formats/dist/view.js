import path from "node:path";
import { createRequire } from "node:module";
import { pathToFileURL } from "node:url";
import { createCanvas, Path2D as CanvasPath2D } from "@napi-rs/canvas";
import * as pdfjs from "pdfjs-dist/legacy/build/pdf.mjs";
import { inspect } from "./inspect.js";
import { exportDocument } from "./export.js";
import { buildObjectGraph } from "./graphs/objectGraph.js";
import { AGENT_UNTRUSTED_INSTRUCTION, escapeHtml, escapeXml, makeStableObjectId, normalizeInput } from "./shared.js";
export async function view(input, options = {}) {
    const inspected = isInspectResult(input) ? input : await inspect(input, { format: undefined, depth: "shallow", config: options.config });
    if (isRasterFormat(options.format)) {
        return rasterView(input, inspected, options);
    }
    const pages = toPages(inspected, options);
    const fullObjectMap = pages.flatMap((page) => page.objectMap);
    const crop = buildObjectCrop(pages, fullObjectMap, inspected, options, "officegen-approximate-svg-html", "approximate");
    return withProgressiveDisclosure({
        schema: "officegen.view.result@1.2",
        fidelity: "approximate",
        renderer: {
            id: "officegen-approximate-svg-html",
            mode: "approximate",
            fidelity: "approximate"
        },
        caveats: [
            "Approximate SVG/HTML view only; fonts, wrapping, theme effects, animations, and native layout may differ.",
            ...inspected.trusted.caveats
        ],
        pages: pages.map((page) => ({ ...page, renderer: page.renderer ?? "officegen-approximate-svg-html" })),
        crops: crop.artifacts,
        crop: crop.metadata,
        summary: buildViewSummary(inspected, pages, fullObjectMap, crop.artifacts),
        nextActions: viewNextActions(inspected, options, false),
        objectMap: fullObjectMap,
        trusted: {
            sourceSchema: inspected.schema,
            sourceFormat: inspected.trusted.format,
            generatedAt: new Date().toISOString()
        },
        agentInstruction: AGENT_UNTRUSTED_INSTRUCTION
    }, fullObjectMap, inspected, options);
}
export const viewDocument = view;
function isInspectResult(value) {
    return Boolean(value && typeof value === "object" && value.schema === "officegen.inspect.result@1.2");
}
function toPages(inspected, options) {
    const format = options.format === "html" ? "html" : "svg";
    const maxPages = options.maxPages ?? 50;
    if (inspected.trusted.format === "pptx") {
        const slides = (inspected.untrusted.slides ?? []).slice(0, maxPages);
        return slides.map((slide, index) => buildSlidePage(slide, index + 1, format, inspected.objectMap));
    }
    if (inspected.trusted.format === "docx") {
        return [buildDocxPage((inspected.untrusted.paragraphs ?? []).slice(0, 200), format, inspected.objectMap)];
    }
    if (inspected.trusted.format === "xlsx") {
        const sheets = (inspected.untrusted.sheets ?? []).slice(0, maxPages);
        return sheets.map((sheet, index) => buildSheetPage(sheet, index + 1, format, inspected.objectMap));
    }
    if (inspected.trusted.format === "pdf") {
        const pdfMaxPages = options.maxPages ?? 10;
        const pages = (inspected.untrusted.pages ?? []).slice(0, pdfMaxPages);
        return pages.map((page, index) => buildPdfPage(page, index + 1, format, inspected.objectMap));
    }
    return [];
}
async function rasterView(input, inspected, options) {
    const format = normalizeRasterFormat(options.format);
    const source = isInspectResult(input) ? inspected.trusted.inputPath : input;
    if (!source) {
        throw new Error("VIEW_RASTER_SOURCE_REQUIRED: PNG/JPEG view requires an input file path or bytes, not an inspect-only result without inputPath.");
    }
    const normalized = await normalizeInput(source, inspected.trusted.format);
    const maxPages = options.maxPages ?? 50;
    const dpi = options.dpi ?? 144;
    let pdfBytes = normalized.bytes;
    let fidelity = "internal";
    let renderer = "officegen-pdfjs-canvas";
    const caveats = [...inspected.trusted.caveats];
    if (normalized.format !== "pdf") {
        const exported = await exportDocument(source, {
            to: "pdf",
            mode: options.mode ?? "native",
            config: options.config,
            timeoutMs: options.timeoutMs
        });
        if (!exported.bytes) {
            throw new Error("VIEW_RASTER_EXPORT_EMPTY: native Office-to-PDF export did not return PDF bytes.");
        }
        pdfBytes = exported.bytes;
        fidelity = exported.fidelity === "native" ? "native" : "internal";
        renderer = exported.renderer?.id ? `${exported.renderer.id}+pdfjs-canvas` : "officegen-office-pdfjs-canvas";
        caveats.push(...exported.caveats);
    }
    const rasterPages = await renderPdfToRasterPages(pdfBytes, {
        format,
        dpi,
        maxPages,
        objectMap: inspected.objectMap,
        sourceFormat: inspected.trusted.format
    });
    const pages = rasterPages.map((page) => ({ ...page, renderer }));
    const crop = buildObjectCrop(pages, inspected.objectMap, inspected, options, "officegen-internal-object-crop", fidelity);
    return withProgressiveDisclosure({
        schema: "officegen.view.result@1.2",
        fidelity,
        renderer: {
            id: renderer,
            mode: options.mode ?? "native",
            fidelity
        },
        caveats: [
            normalized.format === "pdf"
                ? "PDF pages were rasterized with PDF.js canvas rendering."
                : "Office pages were converted through the configured native renderer and rasterized with PDF.js canvas rendering.",
            ...caveats
        ],
        pages,
        crops: crop.artifacts,
        crop: crop.metadata,
        summary: buildViewSummary(inspected, pages, inspected.objectMap, crop.artifacts),
        nextActions: viewNextActions(inspected, options, false),
        objectMap: inspected.objectMap,
        trusted: {
            sourceSchema: inspected.schema,
            sourceFormat: inspected.trusted.format,
            generatedAt: new Date().toISOString()
        },
        agentInstruction: AGENT_UNTRUSTED_INSTRUCTION
    }, inspected.objectMap, inspected, options);
}
function isRasterFormat(format) {
    return format === "png" || format === "jpeg" || format === "jpg";
}
function normalizeRasterFormat(format) {
    return format === "jpeg" || format === "jpg" ? "jpeg" : "png";
}
async function renderPdfToRasterPages(pdfBytes, options) {
    const require = createRequire(import.meta.url);
    const pdfjsRoot = path.dirname(require.resolve("pdfjs-dist/package.json"));
    const standardFontDataUrl = pathToFileURL(path.join(pdfjsRoot, "standard_fonts") + path.sep).href;
    ensurePdfjsPath2D();
    const loadingTask = pdfjs.getDocument({
        data: new Uint8Array(pdfBytes),
        disableWorker: true,
        useSystemFonts: true,
        standardFontDataUrl
    });
    const document = await loadingTask.promise;
    const pages = [];
    const pageCount = Math.min(document.numPages, options.maxPages);
    const scale = Math.max(1, options.dpi / 72);
    for (let pageNumber = 1; pageNumber <= pageCount; pageNumber += 1) {
        const page = await document.getPage(pageNumber);
        const viewport = page.getViewport({ scale });
        const width = Math.ceil(viewport.width);
        const height = Math.ceil(viewport.height);
        const canvas = createCanvas(width, height);
        const canvasContext = pdfjsCanvasContext(canvas.getContext("2d"));
        await page.render({ canvasContext: canvasContext, viewport }).promise;
        const bytes = options.format === "png" ? await canvas.encode("png") : await canvas.encode("jpeg");
        const objectMap = pageObjectMap(options.objectMap, options.sourceFormat, pageNumber);
        pages.push({
            page: pageNumber,
            stableObjectId: makeStableObjectId(String(options.sourceFormat), "document", "page", pageNumber),
            format: options.format,
            content: `data:image/${options.format};base64,${Buffer.from(bytes).toString("base64")}`,
            bytes: new Uint8Array(bytes),
            width,
            height,
            renderer: "pdfjs-canvas",
            objectMap
        });
    }
    await document.destroy();
    return pages;
}
function pdfjsCanvasContext(context) {
    const ctx = context;
    // PDF.js uses browser Canvas overloads that @napi-rs/canvas does not fully accept.
    const fill = ctx.fill?.bind(ctx);
    if (fill) {
        ctx.fill = (pathOrRule, fillRule) => {
            if (pathOrRule === undefined || pathOrRule === "nonzero" || pathOrRule === "evenodd")
                return fill();
            return invokePdfjsCanvasPathMethod(fill, fillRule === undefined ? [pathOrRule] : [pathOrRule, fillRule]);
        };
    }
    const stroke = ctx.stroke?.bind(ctx);
    if (stroke) {
        ctx.stroke = (path) => path === undefined ? stroke() : invokePdfjsCanvasPathMethod(stroke, [path]);
    }
    const clip = ctx.clip?.bind(ctx);
    if (clip) {
        ctx.clip = (pathOrRule, fillRule) => {
            if (pathOrRule === undefined || pathOrRule === "nonzero" || pathOrRule === "evenodd")
                return clip();
            return invokePdfjsCanvasPathMethod(clip, fillRule === undefined ? [pathOrRule] : [pathOrRule, fillRule]);
        };
    }
    return context;
}
function invokePdfjsCanvasPathMethod(method, args) {
    try {
        return method(...args);
    }
    catch (error) {
        if (error instanceof Error && /none of these types `String`, `Path`/i.test(error.message))
            return method();
        throw error;
    }
}
function ensurePdfjsPath2D() {
    const globalWithPath = globalThis;
    if (!globalWithPath.Path2D)
        globalWithPath.Path2D = CanvasPath2D;
}
function pageObjectMap(objectMap, sourceFormat, page) {
    if (sourceFormat === "pptx")
        return objectMap.filter((entry) => Number(entry.selectorHints?.slide) === page);
    if (sourceFormat === "xlsx")
        return objectMap.filter((entry) => Number(entry.selectorHints?.sheet) === page);
    if (sourceFormat === "pdf")
        return objectMap.filter((entry) => Number(entry.selectorHints?.page) === page);
    if (sourceFormat === "docx")
        return page === 1 ? objectMap : [];
    return [];
}
function buildObjectCrop(pages, objectMap, inspected, options, renderer, fidelity) {
    if (!options.crop) {
        return { artifacts: [], metadata: { requested: false, status: "not_requested", source: "none", padding: 8 } };
    }
    const objectId = options.objectId;
    const padding = 8;
    if (!objectId) {
        return { artifacts: [], metadata: { requested: true, status: "object_not_found", source: "none", padding } };
    }
    const graph = buildObjectGraph(objectMap, {
        format: inspected.trusted.format,
        inputPath: inspected.trusted.inputPath,
        inputSha256: inspected.trusted.sha256
    });
    const graphNode = graph.nodes.find((node) => node.stableId === objectId);
    const pageWithObject = pages.find((page) => page.objectMap.some((entry) => entry.stableObjectId === objectId));
    const pageObject = pageWithObject?.objectMap.find((entry) => entry.stableObjectId === objectId);
    const target = pageObject ?? objectMap.find((entry) => entry.stableObjectId === objectId);
    if (!target) {
        return {
            artifacts: [],
            metadata: { requested: true, objectId, status: "object_not_found", source: "none", padding }
        };
    }
    const bbox = bboxFromEntry(target) ?? graphNode?.bbox;
    if (!bbox) {
        return {
            artifacts: [],
            metadata: {
                requested: true,
                objectId,
                status: "bbox_unavailable",
                source: graphNode?.bbox ? "objectGraph" : "none",
                padding,
                objectKind: target.kind,
                graphNodeId: graphNode?.nodeId
            }
        };
    }
    const page = pageWithObject?.page ?? pageNumberForObject(target, inspected.trusted.format);
    const cropBox = paddedBBox(bbox, padding);
    const format = pageWithObject?.format === "html" ? "html" : "svg";
    const metadata = {
        requested: true,
        objectId,
        status: "created",
        source: target.bbox || target.bounds ? "objectMap" : "objectGraph",
        bbox,
        page,
        padding,
        objectKind: target.kind,
        graphNodeId: graphNode?.nodeId
    };
    return {
        artifacts: [{
                objectId,
                page,
                format,
                content: format === "html" ? renderCropHtml(target, cropBox) : renderCropSvg(target, cropBox),
                width: Math.ceil(cropBox[2]),
                height: Math.ceil(cropBox[3]),
                renderer,
                fidelity,
                metadata
            }],
        metadata
    };
}
function bboxFromEntry(entry) {
    if (entry.bbox && entry.bbox.every((value) => Number.isFinite(value)))
        return entry.bbox;
    if (entry.bounds)
        return [entry.bounds.x, entry.bounds.y, entry.bounds.width, entry.bounds.height];
    return undefined;
}
function paddedBBox(bbox, padding) {
    const x = Math.max(0, bbox[0] - padding);
    const y = Math.max(0, bbox[1] - padding);
    return [x, y, Math.max(1, bbox[2] + padding * 2), Math.max(1, bbox[3] + padding * 2)];
}
function pageNumberForObject(entry, sourceFormat) {
    if (sourceFormat === "pptx")
        return Number(entry.selectorHints?.slide ?? 1);
    if (sourceFormat === "xlsx")
        return Number(entry.selectorHints?.sheet ?? 1);
    if (sourceFormat === "pdf")
        return Number(entry.selectorHints?.page ?? 1);
    return 1;
}
function renderCropSvg(object, cropBox) {
    const [x, y, width, height] = cropBox;
    const bbox = bboxFromEntry(object) ?? [x, y, width, height];
    const text = object.text ?? object.label ?? object.textPreview ?? "";
    return `<svg xmlns="http://www.w3.org/2000/svg" width="${Math.ceil(width)}" height="${Math.ceil(height)}" viewBox="${x} ${y} ${width} ${height}" data-crop-object-id="${escapeXml(object.stableObjectId)}"><rect x="${x}" y="${y}" width="${width}" height="${height}" fill="#fff"/><rect x="${bbox[0]}" y="${bbox[1]}" width="${bbox[2]}" height="${bbox[3]}" fill="${object.kind === "chart" ? "#f6f8fa" : "#fff"}" stroke="#0969da" stroke-width="2"/><text x="${bbox[0] + 6}" y="${bbox[1] + Math.min(bbox[3] - 6, 22)}" font-family="Arial, sans-serif" font-size="14" fill="#111">${escapeXml(text)}</text></svg>`;
}
function renderCropHtml(object, cropBox) {
    const [, , width, height] = cropBox;
    const text = object.text ?? object.label ?? object.textPreview ?? "";
    return `<section data-crop-object-id="${escapeHtml(object.stableObjectId)}" style="position:relative;width:${Math.ceil(width)}px;height:${Math.ceil(height)}px;background:#fff;color:#111;font-family:Arial,sans-serif;border:1px solid #0969da;box-sizing:border-box;padding:8px;overflow:hidden"><div data-kind="${escapeHtml(object.kind)}">${escapeHtml(text)}</div></section>`;
}
function buildViewSummary(inspected, pages, objectMap, crops) {
    return {
        sourceFormat: inspected.trusted.format,
        sourceSummary: inspected.trusted.summary,
        pageCount: pages.length,
        objectMapEntries: objectMap.length,
        cropArtifacts: crops.length,
        fidelity: pages[0]?.renderer ? undefined : "approximate"
    };
}
function withProgressiveDisclosure(result, fullObjectMap, inspected, options) {
    const offset = Math.max(0, options.objectMapOffset ?? 0);
    const limit = normalizeObjectMapLimit(options.objectMapLimit);
    const returnedObjectMap = fullObjectMap.slice(offset, offset + limit);
    const hasMore = offset + returnedObjectMap.length < fullObjectMap.length;
    const cursor = hasMore || offset > 0 || fullObjectMap.length > limit
        ? {
            objectMapOffset: offset,
            objectMapLimit: limit,
            objectMapReturned: returnedObjectMap.length,
            objectMapTotal: fullObjectMap.length,
            hasMore,
            ...(hasMore ? { nextObjectMapOffset: offset + returnedObjectMap.length } : {})
        }
        : undefined;
    const returnedIds = new Set(returnedObjectMap.map((entry) => entry.stableObjectId));
    return {
        ...result,
        pages: cursor
            ? result.pages.map((page) => ({ ...page, objectMap: page.objectMap.filter((entry) => returnedIds.has(entry.stableObjectId)) }))
            : result.pages,
        objectMap: returnedObjectMap,
        summary: {
            ...result.summary,
            objectMapEntries: fullObjectMap.length,
            objectMapReturned: returnedObjectMap.length,
            truncated: Boolean(cursor?.hasMore)
        },
        ...(cursor ? { cursor } : {}),
        nextActions: viewNextActions(inspected, options, Boolean(cursor?.hasMore))
    };
}
function normalizeObjectMapLimit(limit) {
    if (limit !== undefined && Number.isFinite(limit) && limit > 0)
        return Math.floor(limit);
    return 200;
}
function viewNextActions(inspected, options, hasMore) {
    const input = inspected.trusted.inputPath ?? "<input>";
    const actions = [];
    if (hasMore) {
        actions.push(`officegen inspect ${input} --depth summary --object-map-limit ${normalizeObjectMapLimit(options.objectMapLimit)} --agent --json`);
    }
    if (!options.crop) {
        actions.push(`officegen view ${input} --object <stableObjectId> --crop --out .officegen/runs/object-crop --json`);
    }
    actions.push(`officegen edit ${input} --ops ops.json --dry-run --resolve-selectors --agent --json`);
    return actions;
}
function buildSlidePage(slide, page, format, objectMap) {
    const slideObjects = objectMap.filter((entry) => Number(entry.selectorHints?.slide) === page);
    const sourceObjects = slideObjects.length ? slideObjects : (slide.textObjects ?? []);
    const objects = sourceObjects.map((entry, index) => {
        const bounds = entry.bounds ?? fallbackSlideBounds(entry, index);
        return {
            ...entry,
            bounds,
            bbox: entry.bbox ?? [bounds.x, bounds.y, bounds.width, bounds.height]
        };
    });
    const stableObjectId = String(slide.stableObjectId ?? makeStableObjectId("pptx", "deck", "slide", page));
    if (format === "html") {
        return {
            page,
            stableObjectId,
            format,
            content: `<section data-stable-object-id="${escapeHtml(stableObjectId)}" style="position:relative;width:960px;height:540px;background:#fff;color:#111;font-family:Arial,sans-serif;border:1px solid #d0d7de">${objects.map(renderSlideHtmlObject).join("")}</section>`,
            objectMap: objects
        };
    }
    return {
        page,
        stableObjectId,
        format,
        content: `<svg xmlns="http://www.w3.org/2000/svg" width="960" height="540" viewBox="0 0 960 540" data-stable-object-id="${escapeXml(stableObjectId)}"><rect width="960" height="540" fill="#fff"/><rect x="0" y="0" width="960" height="540" fill="none" stroke="#d0d7de"/>${objects.map(renderSlideSvgObject).join("")}</svg>`,
        objectMap: objects
    };
}
function fallbackSlideBounds(entry, index) {
    if (entry.kind === "chart")
        return { x: 96, y: 96 + index * 12, width: 360, height: 220 };
    if (entry.kind === "picture")
        return { x: 96, y: 96 + index * 12, width: 240, height: 160 };
    if (entry.kind === "tableCell")
        return { x: 72 + (index % 4) * 160, y: 120 + Math.floor(index / 4) * 40, width: 160, height: 40 };
    return { x: 48, y: 48 + index * 48, width: 864, height: 40 };
}
function renderSlideHtmlObject(object) {
    const bounds = object.bounds ?? fallbackSlideBounds(object, 0);
    const text = object.text ?? object.label ?? object.textPreview ?? "";
    const border = object.kind === "shape" ? "none" : "1px solid #8c959f";
    const background = object.kind === "chart" ? "#f6f8fa" : object.kind === "tableCell" ? "#fff" : "transparent";
    return `<div data-stable-object-id="${escapeHtml(object.stableObjectId)}" data-kind="${escapeHtml(object.kind)}" style="position:absolute;left:${bounds.x}px;top:${bounds.y}px;width:${bounds.width}px;height:${bounds.height}px;box-sizing:border-box;border:${border};background:${background};padding:4px 6px;overflow:hidden">${escapeHtml(text)}</div>`;
}
function renderSlideSvgObject(object) {
    const bounds = object.bounds ?? fallbackSlideBounds(object, 0);
    const text = object.text ?? object.label ?? object.textPreview ?? "";
    const fontSize = object.kind === "shape" ? 24 : 12;
    const box = object.kind === "shape"
        ? ""
        : `<rect x="${bounds.x}" y="${bounds.y}" width="${bounds.width}" height="${bounds.height}" fill="${object.kind === "chart" ? "#f6f8fa" : "#fff"}" stroke="#8c959f"/>`;
    return `<g data-stable-object-id="${escapeXml(object.stableObjectId)}" data-kind="${escapeXml(object.kind)}">${box}<text x="${bounds.x + 6}" y="${bounds.y + Math.min(bounds.height - 6, fontSize + 8)}" font-family="Arial, sans-serif" font-size="${fontSize}" fill="#111">${escapeXml(text)}</text></g>`;
}
function buildDocxPage(paragraphs, format, objectMap = []) {
    const mapped = objectMap.filter((entry) => entry.kind === "paragraph");
    const fallback = paragraphs
        .filter((paragraph) => paragraph.text)
        .map((paragraph, index) => ({
        stableObjectId: String(paragraph.stableObjectId),
        kind: "paragraph",
        text: String(paragraph.text ?? ""),
        bounds: { x: 72, y: 72 + index * 28, width: 468, height: 24 },
        trust: { level: "untrusted", reason: "document-content" },
        untrusted: true
    }));
    const objects = (mapped.length ? mapped : fallback).map((entry, index) => {
        const bounds = entry.bounds ?? { x: 72, y: 72 + index * 28, width: 468, height: 24 };
        return { ...entry, bounds, bbox: entry.bbox ?? [bounds.x, bounds.y, bounds.width, bounds.height] };
    });
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
function buildSheetPage(sheet, page, format, objectMap = []) {
    const cells = (sheet.cells ?? []).slice(0, 120);
    const mappedCells = objectMap.filter((entry) => entry.kind === "cell" && Number(entry.selectorHints?.sheet) === page).slice(0, 120);
    const fallbackCells = cells.map((cell, index) => ({
        stableObjectId: String(cell.stableObjectId),
        kind: "cell",
        label: String(cell.ref ?? ""),
        text: String(cell.value ?? ""),
        bounds: { x: 32 + (index % 6) * 120, y: 48 + Math.floor(index / 6) * 32, width: 120, height: 32 },
        trust: { level: "untrusted", reason: "document-content" },
        untrusted: true
    }));
    const objects = (mappedCells.length ? mappedCells : fallbackCells).map((entry, index) => {
        const bounds = entry.bounds ?? { x: 32 + (index % 6) * 120, y: 48 + Math.floor(index / 6) * 32, width: 120, height: 32 };
        return { ...entry, bounds, bbox: entry.bbox ?? [bounds.x, bounds.y, bounds.width, bounds.height] };
    });
    const workbookObjects = page === 1
        ? objectMap
            .filter((entry) => entry.kind !== "cell")
            .map((entry, index) => {
            const bounds = entry.bounds ?? { x: 32, y: 72 + Math.ceil(objects.length / 6) * 32 + index * 40, width: 360, height: 32 };
            return {
                ...entry,
                bounds,
                bbox: entry.bbox ?? [bounds.x, bounds.y, bounds.width, bounds.height]
            };
        })
        : [];
    const pageObjects = objects.concat(workbookObjects);
    const stableObjectId = String(sheet.stableObjectId ?? makeStableObjectId("xlsx", "workbook", "sheet", page));
    if (format === "html") {
        return {
            page,
            stableObjectId,
            format,
            content: `<table data-stable-object-id="${escapeHtml(stableObjectId)}" style="border-collapse:collapse;font-family:Arial,sans-serif">${pageObjects.map((object) => `<tr><th style="border:1px solid #d0d7de;padding:4px 8px">${escapeHtml(object.label ?? object.kind)}</th><td data-stable-object-id="${escapeHtml(object.stableObjectId)}" style="border:1px solid #d0d7de;padding:4px 8px">${escapeHtml(object.text ?? object.textPreview ?? "")}</td></tr>`).join("")}</table>`,
            objectMap: pageObjects
        };
    }
    return {
        page,
        stableObjectId,
        format,
        content: `<svg xmlns="http://www.w3.org/2000/svg" width="800" height="600" viewBox="0 0 800 600"><rect width="800" height="600" fill="#fff"/>${pageObjects.map((object) => `<g data-stable-object-id="${escapeXml(object.stableObjectId)}"><rect x="${object.bounds?.x}" y="${object.bounds?.y}" width="${object.bounds?.width ?? 120}" height="${object.bounds?.height ?? 32}" fill="#fff" stroke="#d0d7de"/><text x="${(object.bounds?.x ?? 0) + 6}" y="${(object.bounds?.y ?? 0) + 21}" font-family="Arial, sans-serif" font-size="12">${escapeXml(`${object.label ?? object.kind}: ${object.text ?? object.textPreview ?? ""}`)}</text></g>`).join("")}</svg>`,
        objectMap: pageObjects
    };
}
function buildPdfPage(pageInfo, page, format, objectMap = []) {
    const width = Number(pageInfo.width ?? 612);
    const height = Number(pageInfo.height ?? 792);
    const stableObjectId = String(pageInfo.stableObjectId ?? makeStableObjectId("pdf", "document", "page", page));
    const pageObjects = objectMap
        .filter((entry) => Number(entry.selectorHints?.page) === page)
        .map((entry, index) => {
        const bounds = entry.bounds ?? { x: 24, y: 56 + index * 24, width: Math.max(120, width - 48), height: 20 };
        return { ...entry, bounds, bbox: entry.bbox ?? [bounds.x, bounds.y, bounds.width, bounds.height] };
    });
    if (format === "html") {
        return {
            page,
            stableObjectId,
            format,
            content: `<section data-stable-object-id="${escapeHtml(stableObjectId)}" style="width:${width}px;height:${height}px;border:1px solid #d0d7de;background:#fff;font-family:Arial,sans-serif"><p style="padding:24px;color:#57606a">PDF page ${page}</p></section>`,
            objectMap: pageObjects
        };
    }
    return {
        page,
        stableObjectId,
        format,
        content: `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}" data-stable-object-id="${escapeXml(stableObjectId)}"><rect width="${width}" height="${height}" fill="#fff" stroke="#d0d7de"/><text x="24" y="40" font-family="Arial, sans-serif" font-size="16" fill="#57606a">PDF page ${page}</text>${pageObjects.map((object) => `<text x="${object.bounds?.x ?? 24}" y="${(object.bounds?.y ?? 56) + 14}" font-family="Arial, sans-serif" font-size="10" fill="#24292f" data-stable-object-id="${escapeXml(object.stableObjectId)}">${escapeXml(object.textPreview ?? object.text ?? "")}</text>`).join("")}</svg>`,
        objectMap: pageObjects
    };
}
//# sourceMappingURL=view.js.map