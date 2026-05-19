import type JSZip from "jszip";
import type { ObjectMapEntry } from "../shared.js";
export type XlsxFormulaDependencyKind = "cell" | "range" | "threeD" | "namedRange" | "tableStructuredRef";
export type XlsxFormulaUnsafeFlag = "external" | "volatile" | "indirect" | "unsupported";
export interface XlsxFormulaDependency {
    kind: XlsxFormulaDependencyKind;
    ref?: string;
    sheet?: string;
    workbook?: string;
    name?: string;
    tableName?: string;
    sourceText: string;
    untrusted: true;
}
export interface XlsxFormulaRelatedObject {
    kind: "table" | "chart" | "pivotTable" | "slicer";
    name?: string;
    path: string;
    ref?: string;
    reason: string;
    untrusted: true;
}
export interface XlsxFormulaCell {
    stableObjectId: string;
    sheetIndex: number;
    sheetName?: string;
    ref: string;
    formula: string;
    formulaType?: string;
    sharedIndex?: string;
    sharedRef?: string;
    dependencies: XlsxFormulaDependency[];
    unsafeFlags: XlsxFormulaUnsafeFlag[];
    volatileFunctions?: string[];
    relatedObjects: XlsxFormulaRelatedObject[];
    sourcePath: string;
    untrusted: true;
}
export interface XlsxFormulaGraph {
    schema: "officegen.xlsx.formulaGraph@1.0";
    sheetIndex: number;
    sheetName?: string;
    formulaCells: XlsxFormulaCell[];
    dependencies: XlsxFormulaDependency[];
    unsafeFlags: XlsxFormulaUnsafeFlag[];
    relatedObjects: XlsxFormulaRelatedObject[];
    untrusted: true;
}
export interface XlsxCell {
    stableObjectId: string;
    ref: string;
    value: string;
    formula?: string;
    formulaType?: string;
    sharedIndex?: string;
    sharedRef?: string;
    dependencies?: XlsxFormulaDependency[];
    unsafeFlags?: XlsxFormulaUnsafeFlag[];
    relatedObjects?: XlsxFormulaRelatedObject[];
    sourcePath: string;
    untrusted: true;
}
export interface XlsxSheet {
    stableObjectId: string;
    index: number;
    name?: string;
    sourcePath: string;
    cells: XlsxCell[];
    formulaGraph?: XlsxFormulaGraph;
    untrusted: true;
}
interface WorksheetCell {
    attrs: string;
    body: string;
    ref: string;
}
export declare function inspectSheets(zip: JSZip): Promise<{
    sheets: XlsxSheet[];
    objectMap: ObjectMapEntry[];
    sharedStrings: string[];
}>;
export declare function readSharedStrings(zip: JSZip): Promise<string[]>;
export declare function setCell(xml: string, ref: string, value: unknown): {
    changed: boolean;
    xml: string;
};
export declare function insertRows(xml: string, rowIndex: number, rows: unknown[][]): {
    changed: boolean;
    xml: string;
};
export declare function appendRows(xml: string, rows: unknown[][]): {
    changed: boolean;
    xml: string;
    startRow: number;
};
export declare function extractWorksheetCells(xml: string): WorksheetCell[];
export declare function sheetPath(index: number | undefined): string;
export {};
