import type JSZip from "jszip";
import type { ObjectMapEntry } from "../shared.js";
import { makeStableObjectId, readZipText, sortedZipFiles, stableHashId } from "../shared.js";
import { exactText, localText, preview, xmlAttr } from "./xml.js";

export interface XlsxCell {
  stableObjectId: string;
  ref: string;
  value: string;
  sourcePath: string;
  untrusted: true;
}

export interface XlsxSheet {
  stableObjectId: string;
  index: number;
  sourcePath: string;
  cells: XlsxCell[];
  untrusted: true;
}

interface WorksheetCell {
  attrs: string;
  body: string;
  ref: string;
}

export async function inspectSheets(zip: JSZip): Promise<{ sheets: XlsxSheet[]; objectMap: ObjectMapEntry[]; sharedStrings: string[] }> {
  const paths = sortedZipFiles(zip);
  const sheetPaths = paths.filter((path) => /^xl\/worksheets\/sheet\d+\.xml$/i.test(path)).sort(naturalSort);
  const sharedStrings = await readSharedStrings(zip);
  const objectMap: ObjectMapEntry[] = [];
  const sheets: XlsxSheet[] = [];
  for (const [sheetIndex, sheetPath] of sheetPaths.entries()) {
    const xml = (await readZipText(zip, sheetPath)) ?? "";
    const cells = extractWorksheetCells(xml).map((cell) => {
      const type = xmlAttr(cell.attrs, "t");
      const raw = exactText(cell.body, "v")[0] ?? "";
      const inlineText = localText(cell.body, "t").join("");
      const value = type === "s" ? sharedStrings[Number(raw)] ?? raw : type === "inlineStr" ? inlineText : type === "b" ? booleanText(raw) : raw;
      const sheetScope = `s${String(sheetIndex + 1).padStart(3, "0")}`;
      const stableObjectId = stableHashId("xlsx", sheetScope, "cell", `${sheetPath}#${cell.ref}`);
      const bounds = boundsFromRef(cell.ref);
      const entry: ObjectMapEntry = {
        stableObjectId,
        kind: "cell",
        label: cell.ref,
        text: value,
        textPreview: preview(value),
        sourcePath: sheetPath,
        xmlPath: sheetPath,
        bounds,
        bbox: bounds ? [bounds.x, bounds.y, bounds.width, bounds.height] : undefined,
        selectorHints: { sheet: sheetIndex + 1, cell: cell.ref },
        editableOps: ["setText", "xlsx.setCell"],
        trust: { level: "untrusted", reason: "document-content" },
        untrusted: true
      };
      objectMap.push(entry);
      return { stableObjectId, ref: cell.ref, value, sourcePath: sheetPath, untrusted: true as const };
    });
    sheets.push({
      stableObjectId: makeStableObjectId("xlsx", "workbook", "sheet", sheetIndex + 1),
      index: sheetIndex + 1,
      sourcePath: sheetPath,
      cells,
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

export async function readSharedStrings(zip: JSZip): Promise<string[]> {
  const sharedStringsXml = (await readZipText(zip, "xl/sharedStrings.xml")) ?? "";
  return [...sharedStringsXml.matchAll(/<si\b[\s\S]*?<\/si>/g)].map((match) => localText(match[0], "t").join(""));
}

export function setCell(xml: string, ref: string, value: unknown): { changed: boolean; xml: string } {
  const cells = extractWorksheetCells(xml);
  const existing = cells.find((cell) => cell.ref.toUpperCase() === ref.toUpperCase());
  if (existing) {
    const pattern = new RegExp(`<c\\b[^>]*\\br=["']${escapeRegExp(existing.ref)}["'][^>]*?(?:\\/>|>[\\s\\S]*?<\\/c>)`);
    const next = xml.replace(pattern, inlineCellXml(existing.ref, value));
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

export function insertRows(xml: string, rowIndex: number, rows: unknown[][]): { changed: boolean; xml: string } {
  if (!Number.isInteger(rowIndex) || rowIndex < 1 || !rows.length) return { changed: false, xml };
  const shifted = xml.replace(/(<row\b[^>]*\br=")(\d+)("[^>]*>[\s\S]*?<\/row>)/g, (_match, open: string, row: string, close: string) => {
    const rowNo = Number(row);
    return `${open}${rowNo >= rowIndex ? rowNo + rows.length : rowNo}${close}`;
  }).replace(/\br="([A-Z]+)(\d+)"/g, (_match, col: string, row: string) => {
    const rowNo = Number(row);
    return `r="${col}${rowNo >= rowIndex ? rowNo + rows.length : rowNo}"`;
  }).replace(/<f\b([^>]*)>([\s\S]*?)<\/f>/g, (_match, attrs: string, formula: string) => {
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

function shiftFormulaRows(formula: string, rowIndex: number, delta: number): string {
  return formula.replace(/(\$?[A-Z]{1,3})(\$?)(\d+)/g, (_match, col: string, absolute: string, row: string) => {
    const rowNo = Number(row);
    if (!Number.isFinite(rowNo) || rowNo < rowIndex) return `${col}${absolute}${row}`;
    return `${col}${absolute}${rowNo + delta}`;
  });
}

export function appendRows(xml: string, rows: unknown[][]): { changed: boolean; xml: string; startRow: number } {
  const existing = extractWorksheetCells(xml);
  const maxRow = Math.max(0, ...existing.map((cell) => rowFromRef(cell.ref)));
  const startRow = maxRow + 1;
  const inserted = insertRows(xml, startRow, rows);
  return { ...inserted, startRow };
}

export function extractWorksheetCells(xml: string): WorksheetCell[] {
  const rows = [...xml.matchAll(/<row\b([^>]*)>([\s\S]*?)<\/row>/g)];
  if (!rows.length) return extractCellTags(xml, 1);
  return rows.flatMap((rowMatch, rowIndex) => {
    const rowAttrs = rowMatch[1] ?? "";
    const rowBody = rowMatch[2] ?? "";
    const rowNumber = Number(xmlAttr(rowAttrs, "r") ?? rowIndex + 1);
    return extractCellTags(rowBody, Number.isFinite(rowNumber) && rowNumber > 0 ? rowNumber : rowIndex + 1);
  });
}

export function sheetPath(index: number | undefined): string {
  return `xl/worksheets/sheet${index && index > 0 ? index : 1}.xml`;
}

function extractCellTags(xml: string, rowNumber: number): WorksheetCell[] {
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

function inlineCellXml(ref: string, value: unknown): string {
  if (value === null || value === undefined) return `<c r="${ref}"/>`;
  if (typeof value === "number" && Number.isFinite(value)) return `<c r="${ref}"><v>${value}</v></c>`;
  if (typeof value === "boolean") return `<c r="${ref}" t="b"><v>${value ? 1 : 0}</v></c>`;
  return `<c r="${ref}" t="inlineStr"><is><t>${escapeXmlText(String(value))}</t></is></c>`;
}

function boundsFromRef(ref: string): { x: number; y: number; width: number; height: number } | undefined {
  const match = /^([A-Z]+)(\d+)$/i.exec(ref);
  if (!match) return undefined;
  const col = columnIndex(match[1] ?? "A");
  const row = Number(match[2]);
  return { x: 32 + (col - 1) * 120, y: 48 + (row - 1) * 32, width: 120, height: 32 };
}

function columnIndex(name: string): number {
  let value = 0;
  for (const char of name.toUpperCase()) value = value * 26 + (char.charCodeAt(0) - 64);
  return value || 1;
}

function rowFromRef(ref: string): number {
  return Number(/\d+/.exec(ref)?.[0] ?? 1);
}

function columnName(index: number): string {
  let value = index;
  let name = "";
  while (value > 0) {
    value -= 1;
    name = String.fromCharCode(65 + (value % 26)) + name;
    value = Math.floor(value / 26);
  }
  return name || "A";
}

function escapeXmlText(value: string): string {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function booleanText(value: string): string {
  return value === "1" || value.toLowerCase() === "true" ? "TRUE" : "FALSE";
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function naturalSort(a: string, b: string): number {
  return a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" });
}
