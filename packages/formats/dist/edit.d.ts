import { type InputLike, type OfficegenConfig, type ObjectMapEntry } from "./shared.js";
export type EditSelector = {
    stableObjectId?: string;
    slide?: number;
    shapeId?: string;
    contains?: string;
    placeholderKey?: string;
    placeholder?: string;
    shapeName?: string;
    contentControlTag?: string;
    namedRange?: string;
    sheetName?: string;
    cell?: string;
    tableName?: string;
    chartPath?: string;
    textMatch?: {
        text: string;
        exact?: boolean;
    };
    textHash?: string;
    positionHash?: string;
    sourcePath?: string;
    xmlPath?: string;
    page?: number;
    story?: string;
    paragraph?: number;
    table?: number;
    row?: number;
    column?: number;
    range?: string;
    relationshipId?: string;
    assetPath?: string;
    commentId?: string;
    revisionId?: string;
    nearestTo?: {
        slide?: number;
        x: number;
        y: number;
    };
    rightOf?: string | {
        text: string;
        slide?: number;
    };
    largestTextOnSlide?: number | boolean;
    nthBodyShape?: {
        slide: number;
        n: number;
    };
};
export type PptxBulletListItem = string | {
    text: string;
    level?: number;
    bold?: boolean;
    numbering?: boolean;
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
    op: "pptx.addSlide";
    after?: number;
} | {
    op: "pptx.addSlideFromLayout";
    after?: number;
    layout?: string | number;
} | {
    op: "pptx.reorderSlides";
    order: number[];
    selector?: EditSelector;
} | {
    op: "pptx.addTextbox";
    slide: number;
    text: string;
    bounds: {
        x: number;
        y: number;
        width: number;
        height: number;
    };
    name?: string;
    fontSize?: number;
    bold?: boolean;
} | {
    op: "pptx.formatTitle";
    selector: EditSelector;
    fontSize?: number;
    bold?: boolean;
    textCase?: "upper" | "lower" | "title" | "sentence";
} | {
    op: "pptx.formatAllTitles";
    fontSize?: number;
    bold?: boolean;
    textCase?: "upper" | "lower" | "title" | "sentence";
} | {
    op: "pptx.replaceBodyBullets";
    slide: number;
    items: PptxBulletListItem[];
    spaceBeforeForLevel1ExceptFirst?: number;
} | {
    op: "pptx.fitContentToPlaceholder";
    selector: EditSelector;
    minFontSize?: number;
} | {
    op: "pptx.alignObjects";
    selectors: EditSelector[];
    mode: "left" | "right" | "center" | "top" | "bottom" | "middle";
} | {
    op: "pptx.distributeObjects";
    selectors: EditSelector[];
    axis: "x" | "y";
} | {
    op: "pptx.setAltText";
    selector: EditSelector;
    title?: string;
    description?: string;
    decorative?: boolean;
} | {
    op: "pptx.setSpeakerNotes";
    slide: number;
    text: string;
    mode?: "replace" | "append";
} | {
    op: "pptx.replaceWithBulletList";
    items: PptxBulletListItem[];
    selector: EditSelector;
    spaceBeforeForLevel1ExceptFirst?: number;
} | {
    op: "pptx.insertBulletItems";
    items: string[];
    selector: EditSelector;
} | {
    op: "pptx.replaceBulletItems";
    items: string[];
    selector: EditSelector;
} | {
    op: "pptx.setFontSize";
    selector: EditSelector;
    fontSize: number;
} | {
    op: "pptx.setBold";
    selector: EditSelector;
    bold: boolean;
} | {
    op: "pptx.setBulletLevel";
    selector: EditSelector;
    level: number;
} | {
    op: "pptx.setNumbering";
    selector: EditSelector;
    level?: number;
    startAt?: number;
} | {
    op: "pptx.setLineSpacing";
    selector: EditSelector;
    lineSpacing: number;
} | {
    op: "pptx.setSpaceBefore";
    selector: EditSelector;
    spaceBefore: number;
} | {
    op: "pptx.setTextCase";
    selector: EditSelector;
    textCase: "upper" | "lower" | "title" | "sentence";
} | {
    op: "pptx.setTableCellText";
    selector: EditSelector;
    text: string;
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
    op: "docx.replaceTextSmart";
    from: string;
    to: string;
    selector?: EditSelector;
} | {
    op: "docx.setTableCellText";
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
    op: "xlsx.definedName.set";
    name: string;
    ref: string;
} | {
    op: "xlsx.definedName.delete";
    name: string;
} | {
    op: "xlsx.setRange";
    sheet?: number;
    startCell: string;
    values: unknown[][];
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
    expectedInputSha256?: string;
    expectedObjectMapHash?: string;
    minSelectorConfidence?: number;
    continueOnError?: boolean;
    config?: OfficegenConfig;
}
export interface EditSelectorResolution {
    operationIndex: number;
    selector: EditSelector;
    stableObjectId?: string;
    matched: boolean;
    matchCount: number;
    confidence?: number;
    matches: Array<{
        stableObjectId: string;
        kind: string;
        confidence?: number;
        label?: string;
        text?: string;
        sourcePath?: string;
        xmlPath?: string;
    }>;
    reason?: "not-found" | "ambiguous" | "low-confidence" | "unsupported-selector";
}
export interface ResolveEditSelectorsResult {
    schema: "officegen.edit.selectors@1.2";
    format: string;
    inputSha256: string;
    objectMapHash: string;
    resolutions: EditSelectorResolution[];
    objectMap: ObjectMapEntry[];
    caveats: string[];
}
export interface EditOperationResult {
    operationIndex: number;
    op: string;
    applied: boolean;
    reason?: "not-found" | "ambiguous" | "low-confidence" | "unsupported" | "validation-failed" | "idempotency-replay" | "skipped-after-error" | "stale-plan";
    message?: string;
}
export interface EditResult {
    schema: "officegen.edit.result@1.2";
    format: string;
    dryRun?: boolean;
    inputSha256?: string;
    objectMapHash?: string;
    rolledBack?: boolean;
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
