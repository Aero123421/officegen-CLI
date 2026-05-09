import { type InputLike, type OfficegenConfig, type ObjectMapEntry } from "./shared.js";
export type EditSelector = {
    stableObjectId?: string;
    contains?: string;
    placeholderKey?: string;
    shapeName?: string;
    contentControlTag?: string;
    namedRange?: string;
    textMatch?: {
        text: string;
        exact?: boolean;
    };
};
export type EditOperation = {
    type: "replaceText";
    from: string;
    to: string;
    selector?: EditSelector;
} | {
    type: "setText";
    text: string;
    selector: EditSelector;
} | {
    type: "pdf.textOverlay";
    page: number;
    text: string;
    x: number;
    y: number;
    size?: number;
    color?: string;
} | {
    type: "pdf.annotation";
    page: number;
    text: string;
    x: number;
    y: number;
    width?: number;
    height?: number;
} | {
    op: "replaceText";
    from: string;
    to: string;
    selector?: EditSelector;
} | {
    op: "setText";
    text: string;
    selector: EditSelector;
} | {
    op: "pptx.duplicateSlide";
    slide?: number;
    after?: number;
    selector?: EditSelector;
} | {
    op: "pptx.reorderSlides";
    order: number[];
    selector?: EditSelector;
} | {
    op: "pptx.insertBulletItems";
    items: string[];
    selector: EditSelector;
} | {
    op: "pptx.replaceBulletItems";
    items: string[];
    selector: EditSelector;
} | {
    op: "pptx.replaceImageByShape";
    selector: EditSelector;
    replacementBase64: string;
    replacementPath?: string;
    fit?: "contain" | "cover" | "stretch";
    crop?: CropRect;
} | {
    op: "pptx.updateChartData";
    selector: EditSelector;
    categories: string[];
    values: number[];
    seriesName?: string;
} | {
    op: "pptx.setBounds";
    selector: EditSelector;
    bounds: {
        x: number;
        y: number;
        width: number;
        height: number;
    };
} | {
    op: "docx.insertParagraphAfter";
    text: string;
    selector: EditSelector;
} | {
    op: "docx.setHeader";
    text: string;
} | {
    op: "docx.setFooter";
    text: string;
} | {
    op: "docx.setStyle";
    styleId: string;
    font?: string;
    size?: number;
    bold?: boolean;
} | {
    op: "docx.addComment";
    text: string;
    selector: EditSelector;
    author?: string;
} | {
    op: "docx.addRedline";
    text: string;
    selector: EditSelector;
    author?: string;
} | {
    op: "docx.redline.insert";
    text: string;
    selector: EditSelector;
    author?: string;
} | {
    op: "docx.redline.delete";
    selector: EditSelector;
    author?: string;
} | {
    op: "docx.redline.replace";
    text: string;
    selector: EditSelector;
    author?: string;
} | {
    op: "docx.applyStyle";
    styleId: string;
    selector: EditSelector;
} | {
    op: "docx.headerFooter.setText";
    kind: "header" | "footer";
    text: string;
} | {
    op: "xlsx.insertRows";
    sheet?: number;
    rowIndex: number;
    rows: unknown[][];
    selector?: EditSelector;
} | {
    op: "xlsx.appendRows";
    sheet?: number;
    rows: unknown[][];
    selector?: EditSelector;
} | {
    op: "xlsx.setCell";
    sheet?: number;
    cell: string;
    value: unknown;
    selector?: EditSelector;
} | {
    op: "xlsx.setFormula";
    sheet?: number;
    cell: string;
    formula: string;
    selector?: EditSelector;
} | {
    op: "xlsx.updateTable";
    sheet?: number;
    startCell: string;
    rows: unknown[][];
    selector?: EditSelector;
} | {
    op: "xlsx.writeTable";
    sheet?: number;
    startCell: string;
    rows: unknown[][];
    tableName?: string;
    selector?: EditSelector;
} | {
    op: "xlsx.table.resize";
    selector: EditSelector;
    ref: string;
} | {
    op: "xlsx.chart.setData";
    selector: EditSelector;
    categories: string[];
    values: number[];
    seriesName?: string;
} | {
    op: "xlsx.pivot.refreshDefinition";
    selector: EditSelector;
} | {
    op: "xlsx.pivot.refreshAll";
} | {
    op: "xlsx.slicer.setSelection";
    selector: EditSelector;
    selected: string[];
};
interface CropRect {
    left?: number;
    right?: number;
    top?: number;
    bottom?: number;
}
export interface EditOptions {
    out?: string;
    dryRun?: boolean;
    resolveSelectors?: boolean;
    format?: "pptx" | "docx" | "xlsx" | "pdf" | "unknown";
    atomic?: boolean;
    validateFirst?: boolean;
    idempotencyKey?: string;
    continueOnError?: boolean;
    config?: OfficegenConfig;
}
export interface EditSelectorResolution {
    operationIndex: number;
    selector: EditSelector;
    stableObjectId?: string;
    matched: boolean;
    matchCount: number;
    matches: Array<{
        stableObjectId: string;
        kind: string;
        label?: string;
        text?: string;
        sourcePath?: string;
        xmlPath?: string;
    }>;
    reason?: "not-found" | "ambiguous" | "unsupported-selector";
}
export interface ResolveEditSelectorsResult {
    schema: "officegen.edit.selectors@1.2";
    format: string;
    resolutions: EditSelectorResolution[];
    objectMap: ObjectMapEntry[];
    caveats: string[];
}
export interface EditOperationResult {
    operationIndex: number;
    op: string;
    applied: boolean;
    reason?: "not-found" | "ambiguous" | "unsupported" | "validation-failed" | "idempotency-replay" | "skipped-after-error";
    message?: string;
}
export interface EditResult {
    schema: "officegen.edit.result@1.2";
    format: string;
    changed: boolean;
    applied: number;
    skipped: number;
    out?: string;
    bytes?: Uint8Array;
    resolvedSelectors?: EditSelectorResolution[];
    opResults?: EditOperationResult[];
    errors?: EditOperationResult[];
    caveats: string[];
}
export declare function edit(input: InputLike, operations: EditOperation[], options?: EditOptions): Promise<EditResult>;
export declare const editDocument: typeof edit;
export declare function resolveEditSelectors(input: InputLike, operations: EditOperation[], options?: Pick<EditOptions, "format" | "config">): Promise<ResolveEditSelectorsResult>;
export {};
