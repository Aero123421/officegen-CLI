import { OptionalContext } from "./common.js";
export interface LayoutBox {
    id: string;
    x: number;
    y: number;
    width: number;
    height: number;
}
export interface LayoutConstraint {
    id: string;
    minWidth?: number;
    maxWidth?: number;
    minHeight?: number;
    maxHeight?: number;
    alignX?: "left" | "center" | "right";
    alignY?: "top" | "middle" | "bottom";
    snap?: number;
    bounds?: {
        x: number;
        y: number;
        width: number;
        height: number;
    };
}
export interface LayoutApplyOptions extends OptionalContext {
    boxes: LayoutBox[];
    constraints: LayoutConstraint[];
    outputPath?: string;
    planId?: string;
}
export interface LayoutApplyResult {
    kind: "officegen.layout.apply";
    generatedAt: string;
    boxes: LayoutBox[];
    changes: Array<{
        id: string;
        before: LayoutBox;
        after: LayoutBox;
    }>;
    note: string;
}
export declare function applyLayoutConstraints(options: LayoutApplyOptions): Promise<LayoutApplyResult>;
