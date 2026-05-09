import type JSZip from "jszip";
import type { ObjectMapEntry } from "../shared.js";
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
