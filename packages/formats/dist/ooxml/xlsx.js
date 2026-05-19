import { decodeXmlEntities, makeStableObjectId, readZipText, sortedZipFiles, stableHashId } from "../shared.js";
import { exactText, localText, preview, xmlAttr } from "./xml.js";
export async function inspectSheets(zip) {
    const paths = sortedZipFiles(zip);
    const sheetPaths = paths.filter((path) => /^xl\/worksheets\/sheet\d+\.xml$/i.test(path)).sort(naturalSort);
    const sharedStrings = await readSharedStrings(zip);
    const workbookXml = (await readZipText(zip, "xl/workbook.xml")) ?? "";
    const sheetNames = readWorkbookSheetNames(workbookXml);
    const definedNames = readDefinedNameRefs(workbookXml);
    const externalLinkPaths = paths.filter((path) => /^xl\/externalLinks\//i.test(path));
    const workbookObjects = await readWorkbookObjectInventory(zip, paths);
    const objectMap = [];
    const sheets = [];
    for (const [sheetIndex, sheetPath] of sheetPaths.entries()) {
        const xml = (await readZipText(zip, sheetPath)) ?? "";
        const sheetName = sheetNames[sheetIndex] ?? `Sheet${sheetIndex + 1}`;
        const formulaCells = [];
        const sharedFormulas = new Map();
        const cells = extractWorksheetCells(xml).map((cell) => {
            const type = xmlAttr(cell.attrs, "t");
            const raw = exactText(cell.body, "v")[0] ?? "";
            const inlineText = localText(cell.body, "t").join("");
            const formulaInfo = readFormulaInfo(cell.body);
            if (formulaInfo?.type === "shared" && formulaInfo.sharedIndex && formulaInfo.formula) {
                sharedFormulas.set(formulaInfo.sharedIndex, formulaInfo.formula);
            }
            if (formulaInfo?.type === "shared" && formulaInfo.sharedIndex && !formulaInfo.formula) {
                formulaInfo.formula = sharedFormulas.get(formulaInfo.sharedIndex) ?? "";
                formulaInfo.unsupported = !formulaInfo.formula;
            }
            const value = type === "s" ? sharedStrings[Number(raw)] ?? raw : type === "inlineStr" ? inlineText : type === "b" ? booleanText(raw) : raw;
            const sheetScope = `s${String(sheetIndex + 1).padStart(3, "0")}`;
            const stableObjectId = stableHashId("xlsx", sheetScope, "cell", `${sheetPath}#${cell.ref}`);
            const bounds = boundsFromRef(cell.ref);
            const formulaCell = formulaInfo
                ? buildFormulaCell({
                    stableObjectId,
                    sheetIndex: sheetIndex + 1,
                    sheetName,
                    cellRef: cell.ref,
                    sourcePath: sheetPath,
                    formulaInfo,
                    definedNames,
                    externalLinkPaths,
                    workbookObjects
                })
                : undefined;
            if (formulaCell)
                formulaCells.push(formulaCell);
            const entry = {
                stableObjectId,
                kind: "cell",
                label: cell.ref,
                text: value,
                textPreview: preview(value),
                sourcePath: sheetPath,
                xmlPath: sheetPath,
                bounds,
                bbox: bounds ? [bounds.x, bounds.y, bounds.width, bounds.height] : undefined,
                selectorHints: {
                    sheet: sheetIndex + 1,
                    cell: cell.ref,
                    formula: formulaCell?.formula,
                    formulaType: formulaCell?.formulaType,
                    sharedFormula: formulaCell?.formulaType === "shared" ? true : undefined,
                    sharedIndex: formulaCell?.sharedIndex,
                    sharedRef: formulaCell?.sharedRef,
                    formulaDependencies: formulaCell?.dependencies,
                    formulaUnsafeFlags: formulaCell?.unsafeFlags,
                    formulaRelatedObjects: formulaCell?.relatedObjects,
                    regionRole: formulaCell ? "formula" : undefined
                },
                editableOps: ["setText", "xlsx.setCell"],
                trust: { level: "untrusted", reason: "document-content" },
                untrusted: true
            };
            objectMap.push(entry);
            return {
                stableObjectId,
                ref: cell.ref,
                value,
                sourcePath: sheetPath,
                untrusted: true,
                ...(formulaCell
                    ? {
                        formula: formulaCell.formula,
                        formulaType: formulaCell.formulaType,
                        sharedIndex: formulaCell.sharedIndex,
                        sharedRef: formulaCell.sharedRef,
                        dependencies: formulaCell.dependencies,
                        unsafeFlags: formulaCell.unsafeFlags,
                        relatedObjects: formulaCell.relatedObjects
                    }
                    : {})
            };
        });
        for (const [validationIndex, validation] of [...xml.matchAll(/<dataValidation\b([^>]*)/g)].entries()) {
            const attrs = validation[1] ?? "";
            const range = xmlAttr(attrs, "sqref");
            objectMap.push({
                stableObjectId: stableHashId("xlsx", `s${String(sheetIndex + 1).padStart(3, "0")}`, "validation", `${sheetPath}#${validationIndex + 1}`),
                kind: "validation",
                label: range,
                sourcePath: sheetPath,
                xmlPath: sheetPath,
                selectorHints: { sheet: sheetIndex + 1, range, regionRole: "validation" },
                editableOps: ["xlsx.validation.set", "xlsx.validation.delete"],
                trust: { level: "untrusted", reason: "document-content" },
                untrusted: true
            });
        }
        sheets.push({
            stableObjectId: makeStableObjectId("xlsx", "workbook", "sheet", sheetIndex + 1),
            index: sheetIndex + 1,
            name: sheetName,
            sourcePath: sheetPath,
            cells,
            formulaGraph: buildSheetFormulaGraph(sheetIndex + 1, sheetName, formulaCells),
            untrusted: true
        });
    }
    for (const [index, path] of paths.filter((path) => /^xl\/tables\/table\d+\.xml$/i.test(path)).sort(naturalSort).entries()) {
        const xml = (await readZipText(zip, path)) ?? "";
        const attrs = /<table\b([^>]*)/.exec(xml)?.[1] ?? "";
        const name = xmlAttr(attrs, "displayName") ?? xmlAttr(attrs, "name") ?? `Table${index + 1}`;
        const ref = xmlAttr(attrs, "ref");
        objectMap.push({
            stableObjectId: stableHashId("xlsx", "workbook", "table", path),
            kind: "table",
            label: name,
            sourcePath: path,
            xmlPath: path,
            selectorHints: { tableName: name, ref },
            editableOps: ["xlsx.writeTable", "xlsx.updateTable", "xlsx.appendRows", "xlsx.table.resize"],
            trust: { level: "untrusted", reason: "document-content" },
            untrusted: true
        });
    }
    for (const definedName of workbookXml.matchAll(/<definedName\b([^>]*)>([\s\S]*?)<\/definedName>/g)) {
        const attrs = definedName[1] ?? "";
        const name = xmlAttr(attrs, "name");
        const ref = (definedName[2] ?? "").replace(/&apos;/g, "'").replace(/&amp;/g, "&");
        if (!name)
            continue;
        objectMap.push({
            stableObjectId: stableHashId("xlsx", "workbook", "namedRange", name),
            kind: "namedRange",
            label: name,
            text: ref,
            textPreview: preview(ref),
            sourcePath: "xl/workbook.xml",
            xmlPath: "xl/workbook.xml",
            selectorHints: { namedRange: name, range: ref, regionRole: "namedRange" },
            editableOps: ["xlsx.definedName.set", "xlsx.definedName.delete"],
            trust: { level: "untrusted", reason: "document-content" },
            untrusted: true
        });
    }
    for (const [index, path] of paths.filter((path) => /^xl\/charts\/chart\d+\.xml$/i.test(path)).sort(naturalSort).entries()) {
        objectMap.push({
            stableObjectId: stableHashId("xlsx", "workbook", "chart", path),
            kind: "chart",
            label: `Chart ${index + 1}`,
            sourcePath: path,
            xmlPath: path,
            selectorHints: { chartPath: path },
            editableOps: ["xlsx.chart.setData"],
            trust: { level: "untrusted", reason: "document-content" },
            untrusted: true
        });
    }
    for (const [index, path] of paths.filter((path) => /^xl\/pivotTables\/pivotTable\d+\.xml$/i.test(path)).sort(naturalSort).entries()) {
        objectMap.push({
            stableObjectId: stableHashId("xlsx", "workbook", "pivotTable", path),
            kind: "pivotTable",
            label: `PivotTable ${index + 1}`,
            sourcePath: path,
            xmlPath: path,
            selectorHints: { pivotTablePath: path },
            editableOps: ["xlsx.pivot.refreshDefinition", "xlsx.pivot.refreshAll"],
            trust: { level: "untrusted", reason: "document-content" },
            untrusted: true
        });
    }
    for (const [index, path] of paths.filter((path) => /^xl\/slicers\//i.test(path) || /^xl\/slicerCaches\//i.test(path)).sort(naturalSort).entries()) {
        objectMap.push({
            stableObjectId: stableHashId("xlsx", "workbook", "slicer", path),
            kind: "slicer",
            label: `Slicer ${index + 1}`,
            sourcePath: path,
            xmlPath: path,
            selectorHints: { slicerPath: path },
            editableOps: ["xlsx.slicer.setSelection"],
            trust: { level: "untrusted", reason: "document-content" },
            untrusted: true
        });
    }
    return { sheets, objectMap, sharedStrings };
}
function buildFormulaCell(input) {
    const dependencies = extractFormulaDependencies(input.formulaInfo.formula, input.definedNames, input.workbookObjects.tables);
    const volatileFunctions = extractVolatileFunctions(input.formulaInfo.formula);
    const unsafeFlags = formulaUnsafeFlags(input.formulaInfo, volatileFunctions, input.externalLinkPaths);
    return {
        stableObjectId: input.stableObjectId,
        sheetIndex: input.sheetIndex,
        sheetName: input.sheetName,
        ref: input.cellRef,
        formula: input.formulaInfo.formula,
        formulaType: input.formulaInfo.type,
        sharedIndex: input.formulaInfo.sharedIndex,
        sharedRef: input.formulaInfo.sharedRef,
        dependencies,
        unsafeFlags,
        volatileFunctions: volatileFunctions.length ? volatileFunctions : undefined,
        relatedObjects: relatedObjectsForFormula(input.sheetIndex, dependencies, input.workbookObjects),
        sourcePath: input.sourcePath,
        untrusted: true
    };
}
function buildSheetFormulaGraph(sheetIndex, sheetName, formulaCells) {
    return {
        schema: "officegen.xlsx.formulaGraph@1.0",
        sheetIndex,
        sheetName,
        formulaCells,
        dependencies: uniqueDependencies(formulaCells.flatMap((cell) => cell.dependencies)),
        unsafeFlags: uniqueStrings(formulaCells.flatMap((cell) => cell.unsafeFlags)),
        relatedObjects: uniqueRelatedObjects(formulaCells.flatMap((cell) => cell.relatedObjects)),
        untrusted: true
    };
}
export async function readSharedStrings(zip) {
    const sharedStringsXml = (await readZipText(zip, "xl/sharedStrings.xml")) ?? "";
    return [...sharedStringsXml.matchAll(/<si\b[\s\S]*?<\/si>/g)].map((match) => localText(match[0], "t").join(""));
}
export function setCell(xml, ref, value) {
    const cells = extractWorksheetCells(xml);
    const existing = cells.find((cell) => cell.ref.toUpperCase() === ref.toUpperCase());
    if (existing) {
        const pattern = new RegExp(`<c\\b[^>]*\\br=["']${escapeRegExp(existing.ref)}["'][^>]*?(?:\\/>|>[\\s\\S]*?<\\/c>)`);
        const next = xml.replace(pattern, inlineCellXml(existing.ref, value, existing.attrs));
        return { changed: next !== xml, xml: next };
    }
    const rowNo = rowFromRef(ref);
    const rowPattern = new RegExp(`<row\\b([^>]*)\\br=["']${rowNo}["'][^>]*>[\\s\\S]*?<\\/row>`);
    if (rowPattern.test(xml)) {
        const next = xml.replace(rowPattern, (row) => row.replace(/<\/row>$/, `${inlineCellXml(ref, value)}</row>`));
        return { changed: next !== xml, xml: next };
    }
    const rowXml = `<row r="${rowNo}">${inlineCellXml(ref, value)}</row>`;
    const next = xml.replace(/<\/sheetData>/, `${rowXml}</sheetData>`);
    return { changed: next !== xml, xml: next };
}
export function insertRows(xml, rowIndex, rows) {
    if (!Number.isInteger(rowIndex) || rowIndex < 1 || !rows.length)
        return { changed: false, xml };
    const shifted = xml.replace(/(<row\b[^>]*\br=")(\d+)("[^>]*>[\s\S]*?<\/row>)/g, (_match, open, row, close) => {
        const rowNo = Number(row);
        return `${open}${rowNo >= rowIndex ? rowNo + rows.length : rowNo}${close}`;
    }).replace(/\br="([A-Z]+)(\d+)"/g, (_match, col, row) => {
        const rowNo = Number(row);
        return `r="${col}${rowNo >= rowIndex ? rowNo + rows.length : rowNo}"`;
    }).replace(/<f\b([^>]*)>([\s\S]*?)<\/f>/g, (_match, attrs, formula) => {
        return `<f${attrs}>${shiftFormulaRows(formula, rowIndex, rows.length)}</f>`;
    });
    const rowXml = rows
        .map((row, offset) => {
        const rowNo = rowIndex + offset;
        return `<row r="${rowNo}">${row.map((value, index) => inlineCellXml(`${columnName(index + 1)}${rowNo}`, value)).join("")}</row>`;
    })
        .join("");
    const next = shifted.replace(/<\/sheetData>/, `${rowXml}</sheetData>`);
    return { changed: next !== xml, xml: next };
}
function shiftFormulaRows(formula, rowIndex, delta) {
    return formula.replace(/(\$?[A-Z]{1,3})(\$?)(\d+)/g, (_match, col, absolute, row) => {
        const rowNo = Number(row);
        if (!Number.isFinite(rowNo) || rowNo < rowIndex)
            return `${col}${absolute}${row}`;
        return `${col}${absolute}${rowNo + delta}`;
    });
}
export function appendRows(xml, rows) {
    const existing = extractWorksheetCells(xml);
    const maxRow = Math.max(0, ...existing.map((cell) => rowFromRef(cell.ref)));
    const startRow = maxRow + 1;
    const inserted = insertRows(xml, startRow, rows);
    return { ...inserted, startRow };
}
export function extractWorksheetCells(xml) {
    const rows = [...xml.matchAll(/<row\b([^>]*)>([\s\S]*?)<\/row>/g)];
    if (!rows.length)
        return extractCellTags(xml, 1);
    return rows.flatMap((rowMatch, rowIndex) => {
        const rowAttrs = rowMatch[1] ?? "";
        const rowBody = rowMatch[2] ?? "";
        const rowNumber = Number(xmlAttr(rowAttrs, "r") ?? rowIndex + 1);
        return extractCellTags(rowBody, Number.isFinite(rowNumber) && rowNumber > 0 ? rowNumber : rowIndex + 1);
    });
}
export function sheetPath(index) {
    return `xl/worksheets/sheet${index && index > 0 ? index : 1}.xml`;
}
function readFormulaInfo(cellBody) {
    const match = /<f\b([^>]*?)(?:\/>|>([\s\S]*?)<\/f>)/i.exec(cellBody);
    if (!match)
        return undefined;
    const attrs = match[1] ?? "";
    const type = xmlAttr(attrs, "t");
    return {
        formula: decodeXmlEntities(match[2] ?? "").trim(),
        type,
        sharedIndex: xmlAttr(attrs, "si"),
        sharedRef: xmlAttr(attrs, "ref")
    };
}
function extractFormulaDependencies(formula, definedNames, tables) {
    if (!formula)
        return [];
    const dependencies = [];
    const structuredRefs = extractStructuredRefs(formula);
    for (const ref of structuredRefs) {
        dependencies.push({
            kind: "tableStructuredRef",
            tableName: ref.tableName,
            sourceText: ref.sourceText,
            untrusted: true
        });
    }
    const sanitized = stripFormulaStrings(replaceLiteralRanges(formula, structuredRefs.map((ref) => ref.sourceText)));
    const a1Pattern = /(?:\[([^\]]+)\])?(?:(?:'([^']+)'|([A-Za-z_][\w .]*))!)?(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)/g;
    for (const match of sanitized.matchAll(a1Pattern)) {
        const workbook = match[1];
        const sheet = match[2] ?? match[3];
        const ref = normalizeCellRef(match[4] ?? "");
        const sourceText = match[0];
        if (!ref || isLikelyFunctionPrefix(sanitized, match.index ?? 0))
            continue;
        dependencies.push({
            kind: sheet || workbook ? "threeD" : ref.includes(":") ? "range" : "cell",
            ref,
            sheet: sheet ? decodeXmlEntities(sheet) : undefined,
            workbook: workbook ? decodeXmlEntities(workbook) : undefined,
            sourceText,
            untrusted: true
        });
    }
    for (const name of definedNames) {
        if (!name || tables.some((table) => table.name.toLowerCase() === name.toLowerCase()))
            continue;
        const pattern = new RegExp(`(^|[^A-Za-z0-9_.])${escapeRegExp(name)}($|[^A-Za-z0-9_.])`, "i");
        if (pattern.test(sanitized)) {
            dependencies.push({
                kind: "namedRange",
                name,
                sourceText: name,
                untrusted: true
            });
        }
    }
    return uniqueDependencies(dependencies);
}
function extractStructuredRefs(formula) {
    const refs = [];
    const pattern = /\b([A-Za-z_][A-Za-z0-9_.]*)\s*\[/g;
    let match;
    while ((match = pattern.exec(formula))) {
        const tableName = match[1] ?? "";
        const start = match.index + tableName.length;
        let depth = 0;
        let end = -1;
        for (let index = start; index < formula.length; index += 1) {
            const char = formula[index];
            if (char === "[")
                depth += 1;
            if (char === "]") {
                depth -= 1;
                if (depth === 0) {
                    end = index + 1;
                    break;
                }
            }
        }
        if (end > start) {
            const sourceText = `${tableName}${formula.slice(start, end)}`;
            refs.push({ tableName, sourceText });
            pattern.lastIndex = end;
        }
    }
    return refs;
}
function extractVolatileFunctions(formula) {
    const volatileFunctions = ["INDIRECT", "OFFSET", "NOW", "TODAY", "RAND", "RANDBETWEEN", "INFO", "CELL"];
    const found = new Set();
    for (const name of volatileFunctions) {
        if (new RegExp(`\\b${name}\\s*\\(`, "i").test(formula))
            found.add(name);
    }
    return [...found].sort();
}
function formulaUnsafeFlags(formulaInfo, volatileFunctions, externalLinkPaths) {
    const flags = new Set();
    if (/\[[^\]]+\]/.test(formulaInfo.formula) || externalLinkPaths.length > 0 && /!/.test(formulaInfo.formula))
        flags.add("external");
    if (volatileFunctions.length > 0)
        flags.add("volatile");
    if (volatileFunctions.includes("INDIRECT"))
        flags.add("indirect");
    if (formulaInfo.unsupported || /_xlfn\.|GET\.CELL\s*\(/i.test(formulaInfo.formula))
        flags.add("unsupported");
    return [...flags].sort();
}
async function readWorkbookObjectInventory(zip, paths) {
    const tableSheetIndexes = await readTableSheetIndexes(zip, paths);
    const tables = await Promise.all(paths.filter((path) => /^xl\/tables\/table\d+\.xml$/i.test(path)).sort(naturalSort).map(async (path, index) => {
        const xml = (await readZipText(zip, path)) ?? "";
        const attrs = /<table\b([^>]*)/.exec(xml)?.[1] ?? "";
        return {
            kind: "table",
            name: xmlAttr(attrs, "displayName") ?? xmlAttr(attrs, "name") ?? `Table${index + 1}`,
            path,
            ref: xmlAttr(attrs, "ref"),
            sheetIndex: tableSheetIndexes.get(path),
            untrusted: true
        };
    }));
    const charts = await Promise.all(paths.filter((path) => /^xl\/charts\/chart\d+\.xml$/i.test(path)).sort(naturalSort).map(async (path) => {
        const xml = (await readZipText(zip, path)) ?? "";
        const formulas = [...xml.matchAll(/<c:f\b[^>]*>([\s\S]*?)<\/c:f>/g)].map((match) => decodeXmlEntities(match[1] ?? "").trim()).filter(Boolean);
        return {
            kind: "chart",
            path,
            formulas,
            ranges: formulas.flatMap((formula) => extractFormulaDependencies(formula, [], tables)),
            untrusted: true
        };
    }));
    const pivotTables = await Promise.all(paths.filter((path) => /^xl\/pivotTables\/pivotTable\d+\.xml$/i.test(path)).sort(naturalSort).map(async (path, index) => {
        const xml = (await readZipText(zip, path)) ?? "";
        const attrs = /<pivotTableDefinition\b([^>]*)/.exec(xml)?.[1] ?? "";
        const sourceAttrs = /<worksheetSource\b([^>]*)/.exec(xml)?.[1] ?? "";
        return {
            kind: "pivotTable",
            name: xmlAttr(attrs, "name") ?? `PivotTable${index + 1}`,
            path,
            sourceRef: xmlAttr(sourceAttrs, "ref") ?? xmlAttr(sourceAttrs, "name"),
            sourceSheet: xmlAttr(sourceAttrs, "sheet"),
            untrusted: true
        };
    }));
    const slicers = await Promise.all(paths.filter((path) => /^xl\/slicers\//i.test(path) || /^xl\/slicerCaches\//i.test(path)).sort(naturalSort).map(async (path, index) => {
        const xml = (await readZipText(zip, path)) ?? "";
        const attrs = /<(?:slicer|slicerCacheDefinition)\b([^>]*)/.exec(xml)?.[1] ?? "";
        return {
            kind: "slicer",
            name: xmlAttr(attrs, "name") ?? xmlAttr(attrs, "r:id") ?? `Slicer${index + 1}`,
            path,
            tableName: xmlAttr(attrs, "tableId") ?? xmlAttr(attrs, "table"),
            untrusted: true
        };
    }));
    return { tables, charts, pivotTables, slicers };
}
async function readTableSheetIndexes(zip, paths) {
    const result = new Map();
    const sheetPaths = paths.filter((path) => /^xl\/worksheets\/sheet\d+\.xml$/i.test(path)).sort(naturalSort);
    for (const [index, worksheetPath] of sheetPaths.entries()) {
        const relsPath = worksheetPath.replace(/^xl\/worksheets\//i, "xl/worksheets/_rels/") + ".rels";
        const relsXml = (await readZipText(zip, relsPath)) ?? "";
        for (const rel of relsXml.matchAll(/<Relationship\b([^>]*)/g)) {
            const attrs = rel[1] ?? "";
            const target = xmlAttr(attrs, "Target");
            if (!target || !/tables\/table\d+\.xml$/i.test(target))
                continue;
            result.set(resolveRelationshipTarget(worksheetPath, target), index + 1);
        }
    }
    return result;
}
function relatedObjectsForFormula(sheetIndex, dependencies, workbookObjects) {
    const related = [];
    for (const table of workbookObjects.tables) {
        if (dependencies.some((dependency) => dependency.kind === "tableStructuredRef" && sameName(dependency.tableName, table.name))) {
            related.push({ kind: "table", name: table.name, path: table.path, ...(table.ref ? { ref: table.ref } : {}), reason: "structured-ref", untrusted: true });
            continue;
        }
        const tableRef = table.ref;
        if (tableRef && dependencies.some((dependency) => rangesMayOverlap(dependency, tableRef, sheetIndex, table.sheetIndex))) {
            related.push({ kind: "table", name: table.name, path: table.path, ref: tableRef, reason: "range-overlap", untrusted: true });
        }
    }
    for (const chart of workbookObjects.charts) {
        if (chart.ranges.some((chartRange) => dependencies.some((dependency) => dependencyRangesMayOverlap(dependency, chartRange)))) {
            related.push({ kind: "chart", path: chart.path, reason: "range-overlap", untrusted: true });
        }
    }
    for (const pivotTable of workbookObjects.pivotTables) {
        if (pivotTable.sourceRef && dependencies.some((dependency) => rangesMayOverlap(dependency, pivotTable.sourceRef, sheetIndex, undefined, pivotTable.sourceSheet))) {
            related.push({ kind: "pivotTable", ...(pivotTable.name ? { name: pivotTable.name } : {}), path: pivotTable.path, ref: pivotTable.sourceRef, reason: "source-overlap", untrusted: true });
        }
    }
    for (const slicer of workbookObjects.slicers) {
        if (!slicer.tableName)
            continue;
        if (related.some((item) => item.kind === "table" && (sameName(item.name, slicer.tableName) || item.path.endsWith(`${slicer.tableName}.xml`)))) {
            related.push({ kind: "slicer", ...(slicer.name ? { name: slicer.name } : {}), path: slicer.path, reason: "table-slicer", untrusted: true });
        }
    }
    return uniqueRelatedObjects(related);
}
function extractCellTags(xml, rowNumber) {
    let ordinalInRow = 0;
    return [...xml.matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)].map((match) => {
        ordinalInRow += 1;
        const attrs = match[1] ?? "";
        return {
            attrs,
            body: match[2] ?? "",
            ref: xmlAttr(attrs, "r") ?? `${columnName(ordinalInRow)}${rowNumber}`
        };
    });
}
function inlineCellXml(ref, value, existingAttrs = "") {
    const preservedAttrs = preserveCellAttrs(existingAttrs);
    const open = (type) => `<c r="${ref}"${type ? ` t="${type}"` : ""}${preservedAttrs}>`;
    if (value === null || value === undefined)
        return `<c r="${ref}"${preservedAttrs}/>`;
    if (typeof value === "number" && Number.isFinite(value))
        return `${open()}<v>${value}</v></c>`;
    if (typeof value === "boolean")
        return `${open("b")}<v>${value ? 1 : 0}</v></c>`;
    return `${open("inlineStr")}<is><t>${escapeXmlText(String(value))}</t></is></c>`;
}
function preserveCellAttrs(attrs) {
    const preserved = [];
    for (const name of ["s", "cm", "vm", "ph"]) {
        const value = xmlAttr(attrs, name);
        if (value !== undefined)
            preserved.push(`${name}="${escapeXmlText(value)}"`);
    }
    return preserved.length ? ` ${preserved.join(" ")}` : "";
}
function boundsFromRef(ref) {
    const match = /^([A-Z]+)(\d+)$/i.exec(ref);
    if (!match)
        return undefined;
    const col = columnIndex(match[1] ?? "A");
    const row = Number(match[2]);
    return { x: 32 + (col - 1) * 120, y: 48 + (row - 1) * 32, width: 120, height: 32 };
}
function columnIndex(name) {
    let value = 0;
    for (const char of name.toUpperCase())
        value = value * 26 + (char.charCodeAt(0) - 64);
    return value || 1;
}
function rowFromRef(ref) {
    return Number(/\d+/.exec(ref)?.[0] ?? 1);
}
function columnName(index) {
    let value = index;
    let name = "";
    while (value > 0) {
        value -= 1;
        name = String.fromCharCode(65 + (value % 26)) + name;
        value = Math.floor(value / 26);
    }
    return name || "A";
}
function escapeXmlText(value) {
    return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}
function readWorkbookSheetNames(workbookXml) {
    return [...workbookXml.matchAll(/<sheet\b([^>]*)/g)].map((match, index) => decodeXmlEntities(xmlAttr(match[1] ?? "", "name") ?? `Sheet${index + 1}`));
}
function readDefinedNameRefs(workbookXml) {
    return [...workbookXml.matchAll(/<definedName\b([^>]*)>/g)]
        .map((match) => xmlAttr(match[1] ?? "", "name"))
        .filter((name) => Boolean(name));
}
function stripFormulaStrings(formula) {
    return formula.replace(/"(?:[^"]|"")*"/g, (value) => " ".repeat(value.length));
}
function replaceLiteralRanges(input, ranges) {
    let output = input;
    for (const range of ranges.sort((a, b) => b.length - a.length)) {
        output = output.replace(new RegExp(escapeRegExp(range), "g"), " ".repeat(range.length));
    }
    return output;
}
function normalizeCellRef(ref) {
    return ref.replace(/\$/g, "").toUpperCase();
}
function isLikelyFunctionPrefix(formula, index) {
    const prefix = formula.slice(Math.max(0, index - 20), index);
    return /[A-Z][A-Z0-9_.]*\s*$/i.test(prefix) && formula[index - 1] !== "!" && formula[index - 1] !== "'";
}
function uniqueStrings(values) {
    return [...new Set(values.filter(Boolean))].sort();
}
function uniqueDependencies(dependencies) {
    const seen = new Set();
    return dependencies.filter((dependency) => {
        const key = [dependency.kind, dependency.workbook, dependency.sheet, dependency.ref, dependency.name, dependency.tableName, dependency.sourceText].join("|").toLowerCase();
        if (seen.has(key))
            return false;
        seen.add(key);
        return true;
    });
}
function uniqueRelatedObjects(objects) {
    const seen = new Set();
    return objects.filter((object) => {
        const key = [object.kind, object.path, object.name, object.ref, object.reason].join("|").toLowerCase();
        if (seen.has(key))
            return false;
        seen.add(key);
        return true;
    });
}
function resolveRelationshipTarget(sourcePath, target) {
    if (target.startsWith("/"))
        return target.replace(/^\/+/, "");
    const sourceDir = sourcePath.split("/").slice(0, -1).join("/");
    const parts = `${sourceDir}/${target}`.split("/");
    const normalized = [];
    for (const part of parts) {
        if (!part || part === ".")
            continue;
        if (part === "..")
            normalized.pop();
        else
            normalized.push(part);
    }
    return normalized.join("/");
}
function sameName(left, right) {
    return Boolean(left && right && left.toLowerCase() === right.toLowerCase());
}
function dependencyRangesMayOverlap(left, right) {
    if (left.kind === "tableStructuredRef" || right.kind === "tableStructuredRef")
        return sameName(left.tableName, right.tableName);
    if (left.sheet && right.sheet && !sameName(left.sheet, right.sheet))
        return false;
    const leftBounds = parseRangeBounds(left.ref);
    const rightBounds = parseRangeBounds(right.ref);
    return Boolean(leftBounds && rightBounds && boundsOverlap(leftBounds, rightBounds));
}
function rangesMayOverlap(dependency, ref, currentSheetIndex, targetSheetIndex, targetSheetName) {
    if (targetSheetIndex && !dependency.sheet && currentSheetIndex !== targetSheetIndex)
        return false;
    if (targetSheetName && dependency.sheet && !sameName(dependency.sheet, targetSheetName))
        return false;
    const dependencyBounds = parseRangeBounds(dependency.ref);
    const targetBounds = parseRangeBounds(ref);
    return Boolean(dependencyBounds && targetBounds && boundsOverlap(dependencyBounds, targetBounds));
}
function parseRangeBounds(ref) {
    const normalized = normalizeCellRef(String(ref ?? "").split("!").pop() ?? "");
    const match = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(normalized);
    if (!match)
        return undefined;
    const left = { col: columnIndex(match[1] ?? "A"), row: Number(match[2]) };
    const right = { col: columnIndex(match[3] ?? match[1] ?? "A"), row: Number(match[4] ?? match[2]) };
    return {
        minCol: Math.min(left.col, right.col),
        maxCol: Math.max(left.col, right.col),
        minRow: Math.min(left.row, right.row),
        maxRow: Math.max(left.row, right.row)
    };
}
function boundsOverlap(left, right) {
    return left.minCol <= right.maxCol && left.maxCol >= right.minCol && left.minRow <= right.maxRow && left.maxRow >= right.minRow;
}
function booleanText(value) {
    return value === "1" || value.toLowerCase() === "true" ? "TRUE" : "FALSE";
}
function escapeRegExp(value) {
    return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}
function naturalSort(a, b) {
    return a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" });
}
//# sourceMappingURL=xlsx.js.map