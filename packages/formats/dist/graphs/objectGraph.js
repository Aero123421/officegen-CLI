export const OBJECT_GRAPH_VERSION = "officegen.objectGraph@2";
export function buildObjectGraph(objectMap, options = {}) {
    const graphVersion = options.graphVersion ?? OBJECT_GRAPH_VERSION;
    const allNodes = objectMap.map((entry, index) => objectMapEntryToNode(entry, index, graphVersion, options));
    const allEdges = buildGeometryEdges(allNodes, graphVersion);
    const nodeOffset = clampOffset(options.nodeOffset);
    const edgeOffset = clampOffset(options.edgeOffset);
    const nodeLimit = normalizeLimit(options.nodeLimit, allNodes.length);
    const edgeLimit = normalizeLimit(options.edgeLimit, allEdges.length);
    const nodes = allNodes.slice(nodeOffset, nodeOffset + nodeLimit);
    const edges = allEdges.slice(edgeOffset, edgeOffset + edgeLimit);
    const riskFlags = options.riskFlags ?? [];
    return {
        schema: graphVersion,
        version: 2,
        graphVersion,
        source: {
            format: options.format,
            inputPath: options.inputPath,
            inputSha256: options.inputSha256,
            objectMapCount: objectMap.length,
            builder: "inspect.objectMap"
        },
        provenance: {
            generatedFrom: "officegen.inspect.result@1.2",
            sourceField: "objectMap"
        },
        confidence: confidenceFromEvidence([...nodes.flatMap((node) => node.evidence), ...edges.flatMap((edge) => edge.evidence)]),
        riskFlags,
        pagination: {
            nodeOffset,
            nodeLimit,
            nodeCount: nodes.length,
            totalNodes: allNodes.length,
            edgeOffset,
            edgeLimit,
            edgeCount: edges.length,
            totalEdges: allEdges.length,
            truncated: nodeOffset + nodes.length < allNodes.length || edgeOffset + edges.length < allEdges.length,
            ...(nodeOffset + nodes.length < allNodes.length ? { nextNodeOffset: nodeOffset + nodes.length } : {}),
            ...(edgeOffset + edges.length < allEdges.length ? { nextEdgeOffset: edgeOffset + edges.length } : {})
        },
        index: buildIndex(nodes, edges),
        nodes,
        edges
    };
}
function objectMapEntryToNode(entry, index, graphVersion, options) {
    const bbox = normalizeBBox(entry);
    const source = sourceFromEntry(entry, options);
    const style = styleFromEntry(entry);
    const evidence = [{
            kind: "object-map",
            confidence: 1,
            message: "Node was normalized from inspect objectMap.",
            sourceField: "objectMap"
        }];
    if (bbox) {
        evidence.push({
            kind: "geometry",
            confidence: 0.95,
            message: "Bounding box was normalized from objectMap geometry.",
            sourceField: entry.bbox ? "bbox" : "bounds"
        });
    }
    for (const key of ["slide", "page", "sheet", "sheetName", "story", "partKind"]) {
        if (entry.selectorHints?.[key] !== undefined) {
            evidence.push({
                kind: "selector-hint",
                confidence: 0.9,
                message: `Source scope includes selector hint ${key}.`,
                sourceField: `selectorHints.${key}`
            });
        }
    }
    return {
        schema: graphVersion,
        version: 2,
        graphVersion,
        index,
        nodeId: `node:${String(index + 1).padStart(4, "0")}`,
        stableId: entry.stableObjectId,
        type: entry.kind,
        label: entry.label,
        bbox,
        text: textFromEntry(entry),
        style,
        source,
        provenance: {
            schema: graphVersion,
            source: "inspect.objectMap",
            objectMapIndex: index,
            stableObjectId: entry.stableObjectId,
            selectorHints: entry.selectorHints,
            editableOps: entry.editableOps,
            media: entry.media,
            trust: entry.trust
        },
        confidence: confidenceFromEvidence(evidence),
        riskFlags: riskFlagsFromEntry(entry),
        evidence
    };
}
function normalizeBBox(entry) {
    if (entry.bbox)
        return entry.bbox;
    return boundsToBBox(entry.bounds);
}
function boundsToBBox(bounds) {
    if (!bounds)
        return undefined;
    return [bounds.x, bounds.y, bounds.width, bounds.height];
}
function textFromEntry(entry) {
    const value = entry.text;
    const preview = entry.textPreview;
    if (value === undefined && preview === undefined)
        return undefined;
    const normalized = normalizeText(value ?? preview ?? "");
    return {
        value,
        preview,
        normalized
    };
}
function sourceFromEntry(entry, options) {
    const hints = entry.selectorHints ?? {};
    return {
        format: options.format,
        inputPath: options.inputPath,
        inputSha256: options.inputSha256,
        sourcePath: entry.sourcePath ?? stringHint(hints.sourcePath),
        xmlPath: entry.xmlPath ?? stringHint(hints.xmlPath),
        slide: numberHint(hints.slide),
        page: numberHint(hints.page),
        sheet: numberHint(hints.sheet),
        sheetName: stringHint(hints.sheetName),
        story: stringHint(hints.story ?? hints.partKind)
    };
}
function styleFromEntry(entry) {
    const hints = entry.selectorHints ?? {};
    const style = hints.style;
    if (style && typeof style === "object" && !Array.isArray(style))
        return style;
    const font = hints.font;
    const fontSize = hints.fontSize;
    const bold = hints.bold;
    if (font === undefined && fontSize === undefined && bold === undefined)
        return undefined;
    return { font, fontSize, bold };
}
function buildGeometryEdges(nodes, graphVersion) {
    const edges = [];
    for (const from of nodes) {
        if (!from.bbox)
            continue;
        for (const to of nodes) {
            if (from.nodeId === to.nodeId || !to.bbox || !sameScope(from, to))
                continue;
            if (contains(from.bbox, to.bbox)) {
                edges.push(makeGeometryEdge(graphVersion, edges.length, from.nodeId, to.nodeId, "contains", 0.86, "Source bbox contains target bbox."));
            }
            if (isRightOf(from.bbox, to.bbox)) {
                edges.push(makeGeometryEdge(graphVersion, edges.length, from.nodeId, to.nodeId, "rightOf", 0.82, "Source bbox is horizontally right of target bbox."));
            }
            if (isBelow(from.bbox, to.bbox)) {
                edges.push(makeGeometryEdge(graphVersion, edges.length, from.nodeId, to.nodeId, "below", 0.78, "Source bbox is vertically below target bbox."));
            }
        }
    }
    return edges;
}
function makeGeometryEdge(graphVersion, index, from, to, relation, confidence, message) {
    return {
        schema: graphVersion,
        version: 2,
        graphVersion,
        index,
        edgeId: `edge:${String(index + 1).padStart(4, "0")}`,
        from,
        to,
        relation,
        confidence,
        riskFlags: [],
        evidence: [{
                kind: "geometry",
                confidence,
                message
            }]
    };
}
function sameScope(left, right) {
    const leftScope = scopeKey(left);
    const rightScope = scopeKey(right);
    return leftScope !== undefined && leftScope === rightScope;
}
function scopeKey(node) {
    if (node.source.slide !== undefined)
        return `slide:${node.source.slide}`;
    if (node.source.page !== undefined)
        return `page:${node.source.page}`;
    if (node.source.sheet !== undefined)
        return `sheet:${node.source.sheet}`;
    if (node.source.story !== undefined)
        return `story:${node.source.story}`;
    return undefined;
}
function contains(outer, inner) {
    const [outerX, outerY, outerWidth, outerHeight] = outer;
    const [innerX, innerY, innerWidth, innerHeight] = inner;
    if (outerWidth * outerHeight <= innerWidth * innerHeight)
        return false;
    return innerX >= outerX && innerY >= outerY && innerX + innerWidth <= outerX + outerWidth && innerY + innerHeight <= outerY + outerHeight;
}
function isRightOf(source, target) {
    const [sourceX] = source;
    const [targetX, , targetWidth] = target;
    return sourceX >= targetX + targetWidth && verticalOverlapRatio(source, target) >= 0.25;
}
function isBelow(source, target) {
    const [, sourceY] = source;
    const [, targetY, , targetHeight] = target;
    return sourceY >= targetY + targetHeight && horizontalOverlapRatio(source, target) >= 0.25;
}
function verticalOverlapRatio(left, right) {
    const top = Math.max(left[1], right[1]);
    const bottom = Math.min(left[1] + left[3], right[1] + right[3]);
    const overlap = Math.max(0, bottom - top);
    return overlap / Math.max(1, Math.min(left[3], right[3]));
}
function horizontalOverlapRatio(left, right) {
    const start = Math.max(left[0], right[0]);
    const end = Math.min(left[0] + left[2], right[0] + right[2]);
    const overlap = Math.max(0, end - start);
    return overlap / Math.max(1, Math.min(left[2], right[2]));
}
export function normalizeText(value) {
    return value.replace(/\s+/g, " ").trim().toLowerCase();
}
function riskFlagsFromEntry(entry) {
    const flags = [];
    if (entry.trust?.level === "untrusted" || entry.untrusted) {
        flags.push({
            code: "UNTRUSTED_DOCUMENT_CONTENT",
            severity: "info",
            message: "Node text and metadata came from user-controlled document content.",
            source: "objectMap.trust"
        });
    }
    if (entry.kind.toLowerCase().includes("annotation")) {
        flags.push({
            code: "ANNOTATION_CONTENT",
            severity: "info",
            message: "Node represents annotation content that may be hidden or review-only.",
            source: "objectMap.kind"
        });
    }
    return flags;
}
function confidenceFromEvidence(evidence) {
    if (!evidence.length)
        return 0;
    return Number((evidence.reduce((sum, item) => sum + item.confidence, 0) / evidence.length).toFixed(2));
}
function clampOffset(value) {
    return Number.isFinite(value) && value && value > 0 ? Math.floor(value) : 0;
}
function normalizeLimit(value, fallback) {
    if (!Number.isFinite(value) || value === undefined)
        return fallback;
    return Math.max(0, Math.floor(value));
}
function buildIndex(nodes, edges) {
    const nodesByStableId = {};
    const nodesByType = {};
    const edgesByRelation = {
        contains: [],
        rightOf: [],
        below: []
    };
    for (const node of nodes) {
        nodesByStableId[node.stableId] = node.nodeId;
        nodesByType[node.type] = [...(nodesByType[node.type] ?? []), node.nodeId];
    }
    for (const edge of edges)
        edgesByRelation[edge.relation].push(edge.edgeId);
    return { nodesByStableId, nodesByType, edgesByRelation };
}
function numberHint(value) {
    const number = Number(value);
    return Number.isFinite(number) ? number : undefined;
}
function stringHint(value) {
    return typeof value === "string" && value.length ? value : undefined;
}
//# sourceMappingURL=objectGraph.js.map