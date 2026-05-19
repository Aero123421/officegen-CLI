import { makeStableObjectId, readZipText, sortedZipFiles, stableHashId } from "../shared.js";
import { bulletParagraphXml, emuToPx, escapeXmlText, exactText, preview, pxToEmu, replaceNthBlock, xmlAttr } from "./xml.js";
import { nextRelationshipId, parseRelationships, relationshipTarget } from "./relationships.js";
export async function getSlidePaths(zip) {
    const files = sortedZipFiles(zip);
    const fallback = files.filter((path) => /^ppt\/slides\/slide\d+\.xml$/i.test(path)).sort(naturalSort);
    const presentationXml = await readZipText(zip, "ppt/presentation.xml");
    const relsXml = await readZipText(zip, "ppt/_rels/presentation.xml.rels");
    if (!presentationXml || !relsXml)
        return fallback;
    const rels = parseRelationships(relsXml);
    const ids = [...presentationXml.matchAll(/<p:sldId\b[^>]*\br:id="([^"]+)"[^>]*\/>/g)].map((match) => match[1] ?? "");
    const ordered = ids
        .map((id) => rels.find((rel) => rel.id === id))
        .filter(Boolean)
        .map((rel) => relationshipTarget("ppt", rel?.target ?? ""))
        .filter((path) => /^ppt\/slides\/slide\d+\.xml$/i.test(path) && zip.file(path));
    return ordered.length ? ordered : fallback;
}
export async function inspectSlides(zip) {
    const slidePaths = await getSlidePaths(zip);
    const objectMap = [];
    const slides = [];
    for (const [slideIndex, slidePath] of slidePaths.entries()) {
        const xml = (await readZipText(zip, slidePath)) ?? "";
        const relsPath = slidePath.replace(/^ppt\/slides\//, "ppt/slides/_rels/") + ".rels";
        const rels = parseRelationships((await readZipText(zip, relsPath)) ?? "");
        const slideNo = slideIndex + 1;
        const slideStableObjectId = stableHashId("pptx", "deck", "slide", slidePath);
        const shapes = extractShapes(xml, slideNo, slideStableObjectId, slidePath);
        const pictures = extractPictures(xml, slideNo, slidePath, rels);
        const charts = extractCharts(xml, slideNo, slidePath, rels);
        const tableCells = extractTableCells(xml, slideNo, slidePath);
        const groups = extractGroups(xml, slideNo, slidePath);
        const connectors = extractConnectors(xml, slideNo, slidePath);
        const smartArts = await extractSmartArts(zip, xml, slideNo, slidePath, rels);
        const chartSeries = await extractChartSeries(zip, charts);
        const textObjects = shapes
            .filter((shape) => shape.text)
            .map((shape) => {
            const entry = {
                stableObjectId: shape.stableObjectId,
                kind: "shape",
                label: shape.name,
                text: shape.text,
                textPreview: shape.textPreview,
                sourcePath: slidePath,
                xmlPath: slidePath,
                bounds: shape.bounds,
                bbox: shape.bounds ? [shape.bounds.x, shape.bounds.y, shape.bounds.width, shape.bounds.height] : undefined,
                selectorHints: {
                    slide: slideNo,
                    shapeIndex: shape.shapeIndex,
                    shapeId: shape.shapeId,
                    name: shape.name,
                    placeholder: shape.placeholderType,
                    placeholderKey: shape.placeholderType,
                    textPreview: shape.textPreview,
                    textHash: simpleHash(shape.text),
                    positionHash: shape.bounds ? simpleHash(`${Math.round(shape.bounds.x)}:${Math.round(shape.bounds.y)}:${Math.round(shape.bounds.width)}:${Math.round(shape.bounds.height)}`) : undefined
                },
                editableOps: [
                    "setText",
                    "pptx.addTextbox",
                    "pptx.formatTitle",
                    "pptx.insertBulletItems",
                    "pptx.replaceBulletItems",
                    "pptx.replaceWithBulletList",
                    "pptx.setFontSize",
                    "pptx.setBold",
                    "pptx.setBulletLevel",
                    "pptx.setNumbering",
                    "pptx.setLineSpacing",
                    "pptx.setSpaceBefore",
                    "pptx.setTextCase"
                ],
                trust: { level: "untrusted", reason: "document-content" },
                untrusted: true
            };
            objectMap.push(entry);
            return entry;
        })
            .concat(pictures.map((picture) => {
            const entry = {
                stableObjectId: picture.stableObjectId,
                kind: "picture",
                label: picture.name,
                sourcePath: slidePath,
                xmlPath: slidePath,
                bounds: picture.bounds,
                bbox: picture.bounds ? [picture.bounds.x, picture.bounds.y, picture.bounds.width, picture.bounds.height] : undefined,
                selectorHints: {
                    slide: slideNo,
                    shapeId: picture.shapeId,
                    name: picture.name,
                    pictureIndex: picture.pictureIndex,
                    relationshipId: picture.relationshipId,
                    assetPath: picture.assetPath
                },
                editableOps: ["asset.replace", "pptx.replaceImageByShape"],
                media: {
                    relationshipId: picture.relationshipId,
                    assetPath: picture.assetPath
                },
                trust: { level: "untrusted", reason: "document-content" },
                untrusted: true
            };
            objectMap.push(entry);
            return entry;
        }))
            .concat(charts.map((chart) => {
            const entry = {
                stableObjectId: chart.stableObjectId,
                kind: "chart",
                label: chart.name,
                sourcePath: slidePath,
                xmlPath: chart.chartPath,
                bounds: chart.bounds,
                bbox: chart.bounds ? [chart.bounds.x, chart.bounds.y, chart.bounds.width, chart.bounds.height] : undefined,
                selectorHints: {
                    slide: slideNo,
                    shapeId: chart.shapeId,
                    name: chart.name,
                    relationshipId: chart.relationshipId,
                    chartPath: chart.chartPath
                },
                editableOps: ["pptx.updateChartData"],
                media: {
                    relationshipId: chart.relationshipId,
                    chartPath: chart.chartPath
                },
                trust: { level: "untrusted", reason: "document-content" },
                untrusted: true
            };
            objectMap.push(entry);
            return entry;
        }))
            .concat(tableCells.map((cell) => {
            const entry = {
                stableObjectId: cell.stableObjectId,
                kind: "tableCell",
                text: cell.text,
                textPreview: cell.textPreview,
                sourcePath: slidePath,
                xmlPath: slidePath,
                bounds: cell.bounds,
                bbox: cell.bounds ? [cell.bounds.x, cell.bounds.y, cell.bounds.width, cell.bounds.height] : undefined,
                selectorHints: {
                    slide: slideNo,
                    tableCell: cell.cellIndex,
                    row: cell.rowIndex,
                    column: cell.columnIndex,
                    textPreview: cell.textPreview,
                    textHash: simpleHash(cell.text),
                    positionHash: cell.bounds ? simpleHash(`${Math.round(cell.bounds.x)}:${Math.round(cell.bounds.y)}:${Math.round(cell.bounds.width)}:${Math.round(cell.bounds.height)}`) : undefined
                },
                editableOps: ["setText", "pptx.setTableCellText"],
                trust: { level: "untrusted", reason: "document-content" },
                untrusted: true
            };
            objectMap.push(entry);
            return entry;
        }))
            .concat(groups.map((group) => {
            const entry = {
                stableObjectId: group.stableObjectId,
                kind: "group",
                label: group.name,
                sourcePath: slidePath,
                xmlPath: slidePath,
                bounds: group.bounds,
                bbox: group.bounds ? [group.bounds.x, group.bounds.y, group.bounds.width, group.bounds.height] : undefined,
                selectorHints: {
                    slide: slideNo,
                    groupIndex: group.groupIndex,
                    shapeId: group.shapeId,
                    name: group.name,
                    childShapeIds: group.childShapeIds
                },
                media: {
                    sourcePath: slidePath
                },
                trust: { level: "untrusted", reason: "document-content" },
                untrusted: true
            };
            objectMap.push(entry);
            return entry;
        }))
            .concat(connectors.map((connector) => {
            const entry = {
                stableObjectId: connector.stableObjectId,
                kind: "connector",
                label: connector.name,
                sourcePath: slidePath,
                xmlPath: slidePath,
                bounds: connector.bounds,
                bbox: connector.bounds ? [connector.bounds.x, connector.bounds.y, connector.bounds.width, connector.bounds.height] : undefined,
                selectorHints: {
                    slide: slideNo,
                    connectorIndex: connector.connectorIndex,
                    shapeId: connector.shapeId,
                    name: connector.name,
                    startConnection: connector.startConnection,
                    endConnection: connector.endConnection
                },
                media: {
                    sourcePath: slidePath
                },
                trust: { level: "untrusted", reason: "document-content" },
                untrusted: true
            };
            objectMap.push(entry);
            return entry;
        }))
            .concat(smartArts.map((smartArt) => {
            const entry = {
                stableObjectId: smartArt.stableObjectId,
                kind: "smartArt",
                label: smartArt.name,
                text: smartArt.text,
                textPreview: smartArt.textPreview,
                sourcePath: slidePath,
                xmlPath: smartArt.dataPath ?? slidePath,
                bounds: smartArt.bounds,
                bbox: smartArt.bounds ? [smartArt.bounds.x, smartArt.bounds.y, smartArt.bounds.width, smartArt.bounds.height] : undefined,
                selectorHints: {
                    slide: slideNo,
                    smartArtIndex: smartArt.smartArtIndex,
                    shapeId: smartArt.shapeId,
                    name: smartArt.name,
                    relationshipIds: smartArt.relationshipIds,
                    dataPath: smartArt.dataPath,
                    layoutPath: smartArt.layoutPath,
                    quickStylePath: smartArt.quickStylePath,
                    colorsPath: smartArt.colorsPath,
                    layoutId: smartArt.layoutId,
                    quickStyleId: smartArt.quickStyleId,
                    colorStyleId: smartArt.colorStyleId
                },
                media: {
                    relationships: smartArt.relationships,
                    smartArt: {
                        dataPath: smartArt.dataPath,
                        layoutPath: smartArt.layoutPath,
                        quickStylePath: smartArt.quickStylePath,
                        colorsPath: smartArt.colorsPath,
                        layoutId: smartArt.layoutId,
                        quickStyleId: smartArt.quickStyleId,
                        colorStyleId: smartArt.colorStyleId,
                        nodeTexts: smartArt.nodes.map((node) => node.text).filter(Boolean),
                        nodes: smartArt.nodes.map(smartArtNodeSummary),
                        nodeTree: smartArt.nodeTree.map(smartArtNodeSummary),
                        graphicFrame: {
                            slide: slideNo,
                            sourcePath: slidePath,
                            shapeId: smartArt.shapeId,
                            name: smartArt.name,
                            bounds: smartArt.bounds
                        }
                    }
                },
                trust: { level: "untrusted", reason: "document-content" },
                untrusted: true
            };
            objectMap.push(entry);
            return entry;
        }))
            .concat(smartArts.flatMap((smartArt) => smartArt.nodes.filter((node) => node.type !== "doc").map((node) => {
            const entry = {
                stableObjectId: node.stableObjectId,
                kind: "smartArtNode",
                label: node.textPreview ?? node.nodeId,
                text: node.text,
                textPreview: node.textPreview,
                sourcePath: slidePath,
                xmlPath: node.dataPath,
                bounds: smartArt.bounds,
                bbox: smartArt.bounds ? [smartArt.bounds.x, smartArt.bounds.y, smartArt.bounds.width, smartArt.bounds.height] : undefined,
                selectorHints: {
                    slide: slideNo,
                    smartArtStableObjectId: smartArt.stableObjectId,
                    smartArtIndex: smartArt.smartArtIndex,
                    smartArtShapeId: smartArt.shapeId,
                    smartArtName: smartArt.name,
                    nodeIndex: node.nodeIndex,
                    nodeId: node.nodeId,
                    nodeType: node.type,
                    parentNodeId: node.parentNodeId,
                    childNodeIds: node.childNodeIds,
                    dataPath: node.dataPath,
                    layoutId: smartArt.layoutId,
                    quickStyleId: smartArt.quickStyleId,
                    colorStyleId: smartArt.colorStyleId
                },
                media: {
                    smartArt: {
                        stableObjectId: smartArt.stableObjectId,
                        relationshipIds: smartArt.relationshipIds,
                        dataPath: smartArt.dataPath,
                        layoutPath: smartArt.layoutPath,
                        quickStylePath: smartArt.quickStylePath,
                        colorsPath: smartArt.colorsPath,
                        layoutId: smartArt.layoutId,
                        quickStyleId: smartArt.quickStyleId,
                        colorStyleId: smartArt.colorStyleId,
                        graphicFrame: {
                            slide: slideNo,
                            sourcePath: slidePath,
                            shapeId: smartArt.shapeId,
                            name: smartArt.name,
                            bounds: smartArt.bounds
                        }
                    }
                },
                trust: { level: "untrusted", reason: "document-content" },
                untrusted: true
            };
            objectMap.push(entry);
            return entry;
        })))
            .concat(chartSeries.map((series) => {
            const entry = {
                stableObjectId: series.stableObjectId,
                kind: "chartSeries",
                label: series.name,
                sourcePath: slidePath,
                xmlPath: series.chartPath,
                selectorHints: {
                    slide: slideNo,
                    seriesIndex: series.seriesIndex,
                    seriesIdx: series.seriesIdx,
                    order: series.order,
                    chartStableObjectId: series.chartStableObjectId,
                    chartShapeId: series.chartShapeId,
                    relationshipId: series.chartRelationshipId,
                    chartPath: series.chartPath,
                    categoryRef: series.categoryRef,
                    valueRef: series.valueRef
                },
                media: {
                    relationshipId: series.chartRelationshipId,
                    chartPath: series.chartPath
                },
                trust: { level: "untrusted", reason: "document-content" },
                untrusted: true
            };
            objectMap.push(entry);
            return entry;
        }));
        slides.push({
            stableObjectId: slideStableObjectId,
            index: slideNo,
            sourcePath: slidePath,
            text: textObjects.map((entry) => entry.text).filter(Boolean).join("\n"),
            textObjects,
            shapeCount: shapes.length,
            pictureCount: pictures.length,
            chartCount: charts.length,
            untrusted: true
        });
    }
    return { slides, objectMap };
}
async function extractChartSeries(zip, charts) {
    const series = [];
    for (const chart of charts) {
        if (!chart.chartPath)
            continue;
        const chartXml = (await readZipText(zip, chart.chartPath)) ?? "";
        let seriesIndex = 0;
        for (const match of chartXml.matchAll(/<c:ser\b[\s\S]*?<\/c:ser>/g)) {
            seriesIndex += 1;
            const block = match[0];
            const seriesIdx = xmlAttr(/<c:idx\b([^>]*)\/>/.exec(block)?.[1] ?? "", "val");
            const order = xmlAttr(/<c:order\b([^>]*)\/>/.exec(block)?.[1] ?? "", "val");
            const name = extractChartSeriesName(block);
            const categoryRef = extractChartFormula(block, "cat") ?? extractChartFormula(block, "xVal");
            const valueRef = extractChartFormula(block, "val") ?? extractChartFormula(block, "yVal");
            series.push({
                stableObjectId: stableHashId("pptx", slideScope(chart.sourcePath), "chartSeries", `${chart.chartPath}#${seriesIdx ?? seriesIndex}`),
                slideIndex: chart.slideIndex,
                seriesIndex,
                chartStableObjectId: chart.stableObjectId,
                chartShapeId: chart.shapeId,
                chartRelationshipId: chart.relationshipId,
                chartPath: chart.chartPath,
                seriesIdx,
                order,
                name,
                categoryRef,
                valueRef,
                sourcePath: chart.sourcePath
            });
        }
    }
    return series;
}
function extractGroups(xml, slideNo, sourcePath) {
    let groupIndex = 0;
    return [...xml.matchAll(/<p:grpSp\b[\s\S]*?<\/p:grpSp>/g)].map((match) => {
        groupIndex += 1;
        const block = match[0];
        const cNvPrs = [...block.matchAll(/<p:cNvPr\b([^>]*?)(?:\/>|>)/g)].map((item) => item[1] ?? "");
        const shapeId = xmlAttr(cNvPrs[0] ?? "", "id");
        const name = xmlAttr(cNvPrs[0] ?? "", "name");
        const childShapeIds = cNvPrs.slice(1).map((attrs) => xmlAttr(attrs, "id")).filter((id) => Boolean(id));
        return {
            stableObjectId: shapeId
                ? stableHashId("pptx", slideScope(sourcePath), "group", `${sourcePath}#${shapeId}`)
                : makeStableObjectId("pptx", slideScope(sourcePath), "group", groupIndex),
            slideIndex: slideNo,
            groupIndex,
            shapeId,
            name,
            childShapeIds,
            bounds: extractBounds(block),
            sourcePath
        };
    });
}
function extractConnectors(xml, slideNo, sourcePath) {
    let connectorIndex = 0;
    return [...xml.matchAll(/<p:cxnSp\b[\s\S]*?<\/p:cxnSp>/g)].map((match) => {
        connectorIndex += 1;
        const block = match[0];
        const cNvPr = /<p:cNvPr\b([^>]*?)(?:\/>|>)/.exec(block)?.[1] ?? "";
        const shapeId = xmlAttr(cNvPr, "id");
        const name = xmlAttr(cNvPr, "name");
        return {
            stableObjectId: shapeId
                ? stableHashId("pptx", slideScope(sourcePath), "connector", `${sourcePath}#${shapeId}`)
                : makeStableObjectId("pptx", slideScope(sourcePath), "connector", connectorIndex),
            slideIndex: slideNo,
            connectorIndex,
            shapeId,
            name,
            startConnection: extractConnection(block, "stCxn"),
            endConnection: extractConnection(block, "endCxn"),
            bounds: extractBounds(block),
            sourcePath
        };
    });
}
async function extractSmartArts(zip, xml, slideNo, sourcePath, rels) {
    let smartArtIndex = 0;
    const smartArts = [];
    for (const match of xml.matchAll(/<p:graphicFrame\b[\s\S]*?<\/p:graphicFrame>/g)) {
        const block = match[0];
        const relIdsAttrs = /<dgm:relIds\b([^>]*?)(?:\/>|>)/.exec(block)?.[1];
        if (!relIdsAttrs)
            continue;
        smartArtIndex += 1;
        const cNvPr = /<p:cNvPr\b([^>]*?)(?:\/>|>)/.exec(block)?.[1] ?? "";
        const shapeId = xmlAttr(cNvPr, "id");
        const name = xmlAttr(cNvPr, "name");
        const relationshipIds = Object.fromEntries(Object.entries(xmlAttrs(relIdsAttrs))
            .filter(([key]) => key.startsWith("r:"))
            .map(([key, value]) => [key.split(":").pop() ?? key, value]));
        const relationships = Object.fromEntries(Object.entries(relationshipIds).map(([role, relationshipId]) => {
            const rel = rels.find((item) => item.id === relationshipId);
            return [role, { relationshipId, target: rel ? relationshipTarget("ppt/slides", rel.target) : undefined, type: rel?.type }];
        }));
        const dataPath = relationships.dm?.target;
        const layoutPath = relationships.lo?.target;
        const quickStylePath = relationships.qs?.target;
        const colorsPath = relationships.cs?.target;
        const stableObjectId = shapeId
            ? stableHashId("pptx", slideScope(sourcePath), "smartArt", `${sourcePath}#${shapeId}`)
            : makeStableObjectId("pptx", slideScope(sourcePath), "smartArt", smartArtIndex);
        const [dataXml, layoutXml, quickStyleXml, colorsXml, diagramRels] = await Promise.all([
            dataPath ? readZipText(zip, dataPath) : Promise.resolve(undefined),
            layoutPath ? readZipText(zip, layoutPath) : Promise.resolve(undefined),
            quickStylePath ? readZipText(zip, quickStylePath) : Promise.resolve(undefined),
            colorsPath ? readZipText(zip, colorsPath) : Promise.resolve(undefined),
            readSmartArtPartRelationships(zip, [dataPath, layoutPath, quickStylePath, colorsPath])
        ]);
        const nodes = extractSmartArtNodes(dataXml ?? "", {
            slideNo,
            sourcePath,
            dataPath,
            smartArtStableObjectId: stableObjectId,
            smartArtShapeId: shapeId,
            smartArtName: name
        });
        const nodeTree = smartArtNodeTree(nodes);
        const text = nodes.map((node) => node.text).filter(Boolean).join("\n") || undefined;
        const relationshipsWithPartRels = Object.fromEntries(Object.entries(relationships).map(([role, rel]) => [
            role,
            {
                ...rel,
                partRelationships: rel.target ? diagramRels[rel.target] : undefined
            }
        ]));
        smartArts.push({
            stableObjectId,
            slideIndex: slideNo,
            smartArtIndex,
            shapeId,
            name,
            relationshipIds,
            relationships: relationshipsWithPartRels,
            dataPath,
            layoutPath,
            quickStylePath,
            colorsPath,
            layoutId: extractDiagramDefinitionId(layoutXml ?? "", "layoutDef"),
            quickStyleId: extractDiagramDefinitionId(quickStyleXml ?? "", "styleDef"),
            colorStyleId: extractDiagramDefinitionId(colorsXml ?? "", "colorsDef"),
            text,
            textPreview: preview(text),
            nodes,
            nodeTree,
            bounds: extractBounds(block),
            sourcePath
        });
    }
    return smartArts;
}
function extractSmartArtNodes(xml, context) {
    if (!xml)
        return [];
    const points = [...xml.matchAll(/<dgm:pt\b([^>]*?)(?:\/>|>([\s\S]*?)<\/dgm:pt>)/g)];
    const parentByChild = new Map();
    const childrenByParent = new Map();
    for (const match of xml.matchAll(/<dgm:cxn\b([^>]*?)(?:\/>|>[\s\S]*?<\/dgm:cxn>)/g)) {
        const attrs = xmlAttrs(match[1] ?? "");
        if (attrs.type !== "parOf" || !attrs.srcId || !attrs.destId)
            continue;
        const order = Number(attrs.srcOrd ?? attrs.destOrd ?? childrenByParent.get(attrs.srcId)?.length ?? 0);
        parentByChild.set(attrs.destId, attrs.srcId);
        const children = childrenByParent.get(attrs.srcId) ?? [];
        children.push({ childId: attrs.destId, order: Number.isFinite(order) ? order : children.length });
        childrenByParent.set(attrs.srcId, children);
    }
    const rawNodes = points
        .map((match, index) => {
        const attrs = xmlAttrs(match[1] ?? "");
        const nodeId = attrs.modelId;
        if (!nodeId)
            return undefined;
        const body = match[2] ?? "";
        const text = exactText(body, "a:t").join("") || exactText(body, "dgm:t").join("") || undefined;
        return {
            stableObjectId: stableHashId("pptx", slideScope(context.sourcePath), "smartArtNode", `${context.dataPath ?? context.sourcePath}#${context.smartArtShapeId ?? context.smartArtStableObjectId}#${nodeId}`),
            slideIndex: context.slideNo,
            nodeIndex: index + 1,
            nodeId,
            type: attrs.type,
            text,
            textPreview: preview(text),
            parentNodeId: parentByChild.get(nodeId),
            childNodeIds: (childrenByParent.get(nodeId) ?? [])
                .sort((a, b) => a.order - b.order)
                .map((child) => child.childId),
            children: [],
            smartArtStableObjectId: context.smartArtStableObjectId,
            smartArtShapeId: context.smartArtShapeId,
            smartArtName: context.smartArtName,
            dataPath: context.dataPath,
            sourcePath: context.sourcePath
        };
    })
        .filter((node) => Boolean(node));
    const nodeMap = new Map(rawNodes.map((node) => [node.nodeId, node]));
    for (const node of rawNodes) {
        node.children = node.childNodeIds.map((id) => nodeMap.get(id)).filter((child) => Boolean(child));
    }
    return rawNodes.filter((node) => node.type !== "doc" || Boolean(node.text) || node.childNodeIds.length > 0);
}
function smartArtNodeTree(nodes) {
    const nodeIds = new Set(nodes.map((node) => node.nodeId));
    const roots = nodes.filter((node) => !node.parentNodeId || !nodeIds.has(node.parentNodeId));
    if (roots.length)
        return roots;
    return nodes.filter((node) => node.text);
}
function smartArtNodeSummary(node) {
    return {
        stableObjectId: node.stableObjectId,
        nodeIndex: node.nodeIndex,
        nodeId: node.nodeId,
        type: node.type,
        text: node.text,
        textPreview: node.textPreview,
        parentNodeId: node.parentNodeId,
        childNodeIds: node.childNodeIds,
        children: node.children.map(smartArtNodeSummary)
    };
}
function extractDiagramDefinitionId(xml, localName) {
    const attrs = new RegExp(`<dgm:${localName}\\b([^>]*)`).exec(xml)?.[1] ?? "";
    return xmlAttr(attrs, "uniqueId") ?? xmlAttr(attrs, "id") ?? xmlAttr(attrs, "name") ?? xmlAttr(attrs, "defStyle");
}
async function readSmartArtPartRelationships(zip, paths) {
    const entries = {};
    await Promise.all(paths.filter((path) => Boolean(path)).map(async (path) => {
        const relsXml = await readZipText(zip, zipRelationshipsPath(path));
        if (!relsXml)
            return;
        entries[path] = parseRelationships(relsXml).map((rel) => ({
            relationshipId: rel.id,
            target: relationshipTarget(zipDirname(path), rel.target),
            type: rel.type,
            targetMode: rel.targetMode
        }));
    }));
    return entries;
}
function zipDirname(path) {
    return path.split("/").slice(0, -1).join("/");
}
function zipBasename(path) {
    return path.split("/").pop() ?? path;
}
function zipRelationshipsPath(path) {
    const dir = zipDirname(path);
    return `${dir}/_rels/${zipBasename(path)}.rels`;
}
function extractCharts(xml, slideNo, sourcePath, rels) {
    let chartIndex = 0;
    const charts = [];
    for (const match of xml.matchAll(/<p:graphicFrame\b[\s\S]*?<\/p:graphicFrame>/g)) {
        const block = match[0];
        chartIndex += 1;
        const relationshipId = /<c:chart\b[^>]*\br:id="([^"]+)"/.exec(block)?.[1];
        const cNvPr = /<p:cNvPr\b([^>]*)\/>/.exec(block)?.[1] ?? "";
        const shapeId = xmlAttr(cNvPr, "id");
        const name = xmlAttr(cNvPr, "name");
        const rel = rels.find((item) => item.id === relationshipId);
        const chartPath = rel ? relationshipTarget("ppt/slides", rel.target) : undefined;
        charts.push({
            stableObjectId: shapeId
                ? stableHashId("pptx", slideScope(sourcePath), "chart", `${sourcePath}#${shapeId}`)
                : makeStableObjectId("pptx", slideScope(sourcePath), "chart", chartIndex),
            slideIndex: slideNo,
            chartIndex,
            shapeId,
            name,
            relationshipId,
            chartPath,
            bounds: extractBounds(block),
            sourcePath
        });
    }
    return charts;
}
function extractPictures(xml, slideNo, sourcePath, rels) {
    let pictureIndex = 0;
    return [...xml.matchAll(/<p:pic\b[\s\S]*?<\/p:pic>/g)].map((match) => {
        pictureIndex += 1;
        const block = match[0];
        const cNvPr = /<p:cNvPr\b([^>]*)\/>/.exec(block)?.[1] ?? "";
        const shapeId = xmlAttr(cNvPr, "id");
        const name = xmlAttr(cNvPr, "name");
        const relationshipId = /<a:blip\b[^>]*(?:r:embed|r:link)="([^"]+)"/.exec(block)?.[1];
        const rel = relationshipId ? rels.find((item) => item.id === relationshipId) : undefined;
        const assetPath = rel ? relationshipTarget("ppt/slides", rel.target) : undefined;
        return {
            stableObjectId: shapeId
                ? stableHashId("pptx", slideScope(sourcePath), "picture", `${sourcePath}#${shapeId}`)
                : makeStableObjectId("pptx", slideScope(sourcePath), "picture", pictureIndex),
            slideIndex: slideNo,
            pictureIndex,
            shapeId,
            name,
            relationshipId,
            assetPath,
            bounds: extractBounds(block),
            sourcePath
        };
    });
}
function extractTableCells(xml, slideNo, sourcePath) {
    let cellIndex = 0;
    const tableCells = [];
    for (const frameMatch of xml.matchAll(/<p:graphicFrame\b[\s\S]*?<\/p:graphicFrame>/g)) {
        const frame = frameMatch[0];
        const table = /<a:tbl\b[\s\S]*?<\/a:tbl>/.exec(frame)?.[0];
        if (!table)
            continue;
        const tableBounds = extractBounds(frame);
        const rows = [...table.matchAll(/<a:tr\b([^>]*)>([\s\S]*?)<\/a:tr>/g)];
        const rowHeights = rows.map((row) => positiveNumber(xmlAttr(row[1] ?? "", "h")));
        const columnWeights = [...table.matchAll(/<a:gridCol\b([^>]*)\/>/g)]
            .map((column) => positiveNumber(xmlAttr(column[1] ?? "", "w")))
            .filter((value) => value !== undefined);
        const maxColumns = Math.max(1, ...rows.map((row) => [...(row[2] ?? "").matchAll(/<a:tc\b[^>]*>[\s\S]*?<\/a:tc>/g)].length));
        const columnCount = Math.max(columnWeights.length, maxColumns);
        const totalColumnWeight = sumWeights(columnWeights, columnCount);
        const totalRowWeight = sumWeights(rowHeights, rows.length || 1);
        let rowOffsetWeight = 0;
        rows.forEach((row, rowIndex) => {
            const rowBody = row[2] ?? "";
            const rowWeight = rowHeights[rowIndex] ?? 1;
            let columnOffsetWeight = 0;
            let columnIndex = 0;
            for (const cellMatch of rowBody.matchAll(/<a:tc\b([^>]*)>([\s\S]*?)<\/a:tc>/g)) {
                const attrs = cellMatch[1] ?? "";
                const body = cellMatch[2] ?? "";
                const span = Math.max(1, Number(xmlAttr(attrs, "gridSpan") ?? 1));
                const text = exactText(body, "a:t").join("");
                const cellWeight = spanWeights(columnWeights, columnIndex, span);
                cellIndex += 1;
                if (text) {
                    const bounds = tableBounds
                        ? {
                            x: tableBounds.x + (tableBounds.width * columnOffsetWeight) / totalColumnWeight,
                            y: tableBounds.y + (tableBounds.height * rowOffsetWeight) / totalRowWeight,
                            width: (tableBounds.width * cellWeight) / totalColumnWeight,
                            height: (tableBounds.height * rowWeight) / totalRowWeight
                        }
                        : undefined;
                    tableCells.push({
                        stableObjectId: stableHashId("pptx", slideScope(sourcePath), "tableCell", `${sourcePath}#${cellIndex}`),
                        cellIndex,
                        rowIndex: rowIndex + 1,
                        columnIndex: columnIndex + 1,
                        text,
                        textPreview: preview(text),
                        bounds
                    });
                }
                columnOffsetWeight += cellWeight;
                columnIndex += span;
            }
            rowOffsetWeight += rowWeight;
        });
    }
    if (tableCells.length)
        return tableCells;
    return [...xml.matchAll(/<a:tc\b[\s\S]*?<\/a:tc>/g)]
        .map((match) => {
        cellIndex += 1;
        const text = exactText(match[0], "a:t").join("");
        return {
            stableObjectId: stableHashId("pptx", slideScope(sourcePath), "tableCell", `${sourcePath}#${cellIndex}`),
            cellIndex,
            text,
            textPreview: preview(text)
        };
    })
        .filter((cell) => cell.text);
}
export function extractShapes(xml, slideNo, slideStableObjectId, sourcePath) {
    let shapeIndex = 0;
    return [...xml.matchAll(/<p:sp\b[\s\S]*?<\/p:sp>/g)].map((match) => {
        shapeIndex += 1;
        const block = match[0];
        const cNvPr = /<p:cNvPr\b([^>]*)\/>/.exec(block)?.[1] ?? "";
        const shapeId = xmlAttr(cNvPr, "id");
        const name = xmlAttr(cNvPr, "name");
        const placeholderType = /<p:ph\b([^>]*?)\/>/.exec(block)?.[1];
        const text = exactText(block, "a:t").join("");
        const scope = slideScope(sourcePath);
        const stableObjectId = shapeId
            ? stableHashId("pptx", scope, "shape", `${sourcePath}#${shapeId}`)
            : makeStableObjectId("pptx", scope, "shape", shapeIndex);
        return {
            stableObjectId,
            slideStableObjectId,
            slideIndex: slideNo,
            shapeIndex,
            shapeId,
            name,
            placeholderType: placeholderType ? xmlAttr(placeholderType, "type") ?? "body" : undefined,
            text,
            textPreview: preview(text),
            bounds: extractBounds(block),
            sourcePath
        };
    });
}
function slideScope(sourcePath) {
    return `slide-${stablePathToken(sourcePath)}`;
}
function stablePathToken(sourcePath) {
    let hash = 2166136261;
    for (const char of sourcePath.toLowerCase()) {
        hash ^= char.charCodeAt(0);
        hash = Math.imul(hash, 16777619);
    }
    return (hash >>> 0).toString(16).padStart(8, "0").slice(0, 8);
}
export function replaceShapeBulletItems(xml, ordinal, items, mode) {
    return replaceNthBlock(xml, /<p:sp\b[\s\S]*?<\/p:sp>/g, ordinal, (shape) => {
        const bulletXml = items.map((item) => bulletParagraphXml(item)).join("");
        if (!/<p:txBody\b[\s\S]*?<\/p:txBody>/.test(shape))
            return shape;
        if (mode === "replace") {
            return shape.replace(/(<p:txBody\b[^>]*>)([\s\S]*?)(<\/p:txBody>)/, (_match, open, body, close) => {
                const bodyPr = /<a:bodyPr\b[\s\S]*?\/>/.exec(body)?.[0] ?? "<a:bodyPr/>";
                const lstStyle = /<a:lstStyle\b[\s\S]*?\/>/.exec(body)?.[0] ?? "<a:lstStyle/>";
                return `${open}${bodyPr}${lstStyle}${bulletXml}${close}`;
            });
        }
        return shape.replace(/<\/p:txBody>/, `${bulletXml}</p:txBody>`);
    });
}
export async function duplicateSlide(zip, slideNumber, after) {
    const slidePaths = await getSlidePaths(zip);
    const sourcePath = slidePaths[slideNumber - 1];
    if (!sourcePath)
        throw new Error(`SELECTOR_NOT_FOUND: pptx slide ${slideNumber} not found.`);
    const nextNo = nextSlideNumber(sortedZipFiles(zip));
    const nextPath = `ppt/slides/slide${nextNo}.xml`;
    zip.file(nextPath, (await readZipText(zip, sourcePath)) ?? "");
    const sourceRels = sourcePath.replace(/^ppt\/slides\//, "ppt/slides/_rels/") + ".rels";
    const targetRels = nextPath.replace(/^ppt\/slides\//, "ppt/slides/_rels/") + ".rels";
    const relsText = await readZipText(zip, sourceRels);
    if (relsText !== undefined)
        zip.file(targetRels, relsText);
    await addPresentationSlide(zip, nextNo, after ?? slideNumber);
    await addSlideContentType(zip, nextNo);
}
export async function addBlankSlide(zip, after) {
    const slidePaths = await getSlidePaths(zip);
    const nextNo = nextSlideNumber(sortedZipFiles(zip));
    const nextPath = `ppt/slides/slide${nextNo}.xml`;
    zip.file(nextPath, blankSlideXml());
    zip.file(nextPath.replace(/^ppt\/slides\//, "ppt/slides/_rels/") + ".rels", await blankSlideRelsXml(zip, slidePaths[sourceSlideIndexForInsert(slidePaths.length, after)]));
    await addPresentationSlide(zip, nextNo, after ?? slidePaths.length);
    await addSlideContentType(zip, nextNo);
    return nextNo;
}
export async function addTextBox(zip, slideNumber, spec) {
    const slidePaths = await getSlidePaths(zip);
    const slidePath = slidePaths[slideNumber - 1];
    if (!slidePath)
        throw new Error(`SELECTOR_NOT_FOUND: pptx slide ${slideNumber} not found.`);
    const xml = (await readZipText(zip, slidePath)) ?? "";
    const nextId = nextShapeId(xml);
    const shape = textBoxShapeXml(nextId, spec);
    const next = /<p:spTree\b[\s\S]*?<\/p:spTree>/.test(xml)
        ? xml.replace(/<p:spTree\b[\s\S]*?<\/p:spTree>/, (spTree) => insertShapeIntoSpTree(spTree, shape))
        : xml.replace(/<\/p:cSld>/, `<p:spTree>${defaultGroupShapeTreeHead()}${shape}</p:spTree></p:cSld>`);
    if (next === xml)
        throw new Error("SELECTOR_NOT_FOUND: pptx slide shape tree not found.");
    zip.file(slidePath, next);
}
export async function reorderSlides(zip, order) {
    const presentationXml = (await readZipText(zip, "ppt/presentation.xml")) ?? "";
    const ids = [...presentationXml.matchAll(/<p:sldId\b[^>]*\br:id="([^"]+)"[^>]*\/>/g)].map((match) => match[1] ?? "");
    if (!ids.length)
        throw new Error("SELECTOR_NOT_FOUND: pptx presentation slide list not found.");
    if (order.length !== ids.length || new Set(order).size !== ids.length || order.some((item) => item < 1 || item > ids.length)) {
        throw new Error(`SELECTOR_NOT_FOUND: pptx reorderSlides requires a permutation of 1..${ids.length}.`);
    }
    const sldIdTags = [...presentationXml.matchAll(/<p:sldId\b[^>]*\/>/g)].map((match) => match[0]);
    const reordered = order.map((slideNo) => sldIdTags[slideNo - 1]).join("");
    zip.file("ppt/presentation.xml", presentationXml.replace(/(<p:sldIdLst\b[^>]*>)[\s\S]*?(<\/p:sldIdLst>)/, `$1${reordered}$2`));
}
function extractBounds(block) {
    const off = /<a:off\b([^>]*)\/>/.exec(block)?.[1];
    const ext = /<a:ext\b([^>]*)\/>/.exec(block)?.[1];
    if (!off || !ext)
        return undefined;
    const x = Number(xmlAttr(off, "x"));
    const y = Number(xmlAttr(off, "y"));
    const cx = Number(xmlAttr(ext, "cx"));
    const cy = Number(xmlAttr(ext, "cy"));
    if (![x, y, cx, cy].every(Number.isFinite))
        return undefined;
    return { x: emuToPx(x), y: emuToPx(y), width: emuToPx(cx), height: emuToPx(cy) };
}
function positiveNumber(value) {
    const parsed = Number(value);
    return Number.isFinite(parsed) && parsed > 0 ? parsed : undefined;
}
function sumWeights(weights, count) {
    let total = 0;
    for (let index = 0; index < count; index += 1)
        total += weights[index] ?? 1;
    return total || 1;
}
function spanWeights(weights, start, span) {
    let total = 0;
    for (let index = start; index < start + span; index += 1)
        total += weights[index] ?? 1;
    return total || 1;
}
async function addPresentationSlide(zip, slideNo, after) {
    const presentationXml = (await readZipText(zip, "ppt/presentation.xml")) ?? "";
    if (!/<p:sldIdLst\b[\s\S]*?<\/p:sldIdLst>/.test(presentationXml)) {
        throw new Error("SELECTOR_NOT_FOUND: pptx presentation slide list not found.");
    }
    const relsXml = (await readZipText(zip, "ppt/_rels/presentation.xml.rels")) ?? '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
    const rId = nextRelationshipId(relsXml);
    const sldIds = [...presentationXml.matchAll(/<p:sldId\b[^>]*\bid="(\d+)"[^>]*\/>/g)].map((match) => Number(match[1])).filter(Number.isFinite);
    const newId = Math.max(255, ...sldIds) + 1;
    const slideTag = `<p:sldId id="${newId}" r:id="${rId}"/>`;
    const tags = [...presentationXml.matchAll(/<p:sldId\b[^>]*\/>/g)].map((match) => match[0]);
    const insertAt = Math.min(Math.max(after, 0), tags.length);
    tags.splice(insertAt, 0, slideTag);
    zip.file("ppt/presentation.xml", presentationXml.replace(/(<p:sldIdLst\b[^>]*>)[\s\S]*?(<\/p:sldIdLst>)/, `$1${tags.join("")}$2`));
    zip.file("ppt/_rels/presentation.xml.rels", relsXml.replace(/<\/Relationships>\s*$/, `<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${slideNo}.xml"/></Relationships>`));
}
async function addSlideContentType(zip, slideNo) {
    const xml = await readZipText(zip, "[Content_Types].xml");
    if (!xml || xml.includes(`/ppt/slides/slide${slideNo}.xml`))
        return;
    const override = `<Override PartName="/ppt/slides/slide${slideNo}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`;
    zip.file("[Content_Types].xml", xml.replace(/<\/Types>\s*$/, `${override}</Types>`));
}
function nextSlideNumber(paths) {
    const numbers = paths
        .map((path) => /^ppt\/slides\/slide(\d+)\.xml$/i.exec(path)?.[1])
        .filter(Boolean)
        .map(Number);
    return Math.max(0, ...numbers) + 1;
}
function insertShapeIntoSpTree(spTree, shape) {
    return /<p:extLst\b[\s\S]*?<\/p:extLst>\s*<\/p:spTree>$/.test(spTree)
        ? spTree.replace(/(<p:extLst\b[\s\S]*?<\/p:extLst>\s*<\/p:spTree>)$/, `${shape}$1`)
        : spTree.replace(/<\/p:spTree>$/, `${shape}</p:spTree>`);
}
function nextShapeId(xml) {
    const ids = [...xml.matchAll(/<p:cNvPr\b[^>]*\bid="(\d+)"/g)].map((match) => Number(match[1])).filter(Number.isFinite);
    return Math.max(1, ...ids) + 1;
}
function blankSlideXml() {
    return [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">',
        '<p:cSld><p:spTree>',
        defaultGroupShapeTreeHead(),
        '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>'
    ].join("");
}
function defaultGroupShapeTreeHead() {
    return [
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>',
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
    ].join("");
}
function sourceSlideIndexForInsert(slideCount, after) {
    if (slideCount <= 0)
        return -1;
    if (after === undefined)
        return slideCount - 1;
    return Math.max(0, Math.min(slideCount - 1, after === 0 ? 0 : after - 1));
}
async function blankSlideRelsXml(zip, sourceSlidePath) {
    const empty = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
    if (!sourceSlidePath)
        return empty;
    const sourceRelsPath = sourceSlidePath.replace(/^ppt\/slides\//, "ppt/slides/_rels/") + ".rels";
    const sourceRels = (await readZipText(zip, sourceRelsPath)) ?? "";
    const layoutRel = /<Relationship\b[^>]*\bType="[^"]*\/slideLayout"[^>]*\/>/.exec(sourceRels)?.[0];
    if (!layoutRel)
        return empty;
    const target = xmlAttr(layoutRel, "Target");
    if (!target)
        return empty;
    return `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="${escapeXmlText(target)}"/></Relationships>`;
}
function textBoxShapeXml(id, spec) {
    const fontSize = spec.fontSize ? ` sz="${Math.round(spec.fontSize * 100)}"` : "";
    const bold = spec.bold ? ' b="1"' : "";
    return [
        "<p:sp>",
        `<p:nvSpPr><p:cNvPr id="${id}" name="${escapeXmlText(spec.name ?? `TextBox ${id}`)}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>`,
        `<p:spPr><a:xfrm><a:off x="${pxToEmu(spec.bounds.x)}" y="${pxToEmu(spec.bounds.y)}"/><a:ext cx="${pxToEmu(spec.bounds.width)}" cy="${pxToEmu(spec.bounds.height)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr>`,
        `<p:txBody><a:bodyPr wrap="square"/><a:lstStyle/><a:p><a:r><a:rPr${fontSize}${bold}/><a:t>${escapeXmlText(spec.text)}</a:t></a:r></a:p></p:txBody>`,
        "</p:sp>"
    ].join("");
}
function extractChartSeriesName(block) {
    const tx = /<c:tx\b[\s\S]*?<\/c:tx>/.exec(block)?.[0] ?? "";
    return exactText(tx, "c:v")[0] ?? extractChartFormula(block, "tx");
}
function extractChartFormula(block, container) {
    const pattern = new RegExp(`<c:${container}\\b[\\s\\S]*?<c:f(?:\\s[^>]*)?>([\\s\\S]*?)<\\/c:f>[\\s\\S]*?<\\/c:${container}>`);
    return pattern.exec(block)?.[1]?.trim();
}
function extractConnection(block, tagName) {
    const attrs = new RegExp(`<a:${tagName}\\b([^>]*)\\/>`).exec(block)?.[1];
    if (!attrs)
        return undefined;
    return {
        shapeId: xmlAttr(attrs, "id"),
        index: xmlAttr(attrs, "idx")
    };
}
function xmlAttrs(attrs) {
    const values = {};
    for (const match of attrs.matchAll(/([A-Za-z_][\w:.-]*)\s*=\s*(?:"([^"]*)"|'([^']*)')/g)) {
        values[match[1] ?? ""] = match[2] ?? match[3] ?? "";
    }
    return values;
}
function simpleHash(value) {
    let hash = 2166136261;
    for (const char of value) {
        hash ^= char.charCodeAt(0);
        hash = Math.imul(hash, 16777619);
    }
    return (hash >>> 0).toString(16).padStart(8, "0");
}
function naturalSort(a, b) {
    return a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" });
}
//# sourceMappingURL=pptx.js.map