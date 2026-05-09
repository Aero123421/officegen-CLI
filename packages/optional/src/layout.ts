import path from "node:path";

import { OptionalContext, featureRoot, nowIso, requireFeature, slugify, writeJsonFile } from "./common.js";

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
  changes: Array<{ id: string; before: LayoutBox; after: LayoutBox }>;
  note: string;
}

export async function applyLayoutConstraints(options: LayoutApplyOptions): Promise<LayoutApplyResult> {
  requireFeature(options, "layout", "layout apply");
  const constraints = new Map(options.constraints.map((constraint) => [constraint.id, constraint]));
  const changes: LayoutApplyResult["changes"] = [];
  const boxes = options.boxes.map((box) => {
    const after = applyConstraint(box, constraints.get(box.id));
    if (!sameBox(box, after)) {
      changes.push({ id: box.id, before: box, after });
    }
    return after;
  });

  const result: LayoutApplyResult = {
    kind: "officegen.layout.apply",
    generatedAt: nowIso(),
    boxes,
    changes,
    note: "Simple layout constraints only; Office file mutation is delegated to @officegen/formats."
  };

  const outputPath =
    options.outputPath ??
    path.join(featureRoot(options, "layout"), "runs", `${slugify(options.planId ?? "layout")}.apply.json`);
  await writeJsonFile(outputPath, result);
  return result;
}

function applyConstraint(box: LayoutBox, constraint: LayoutConstraint | undefined): LayoutBox {
  if (!constraint) {
    return box;
  }

  let next: LayoutBox = { ...box };
  next.width = clamp(next.width, constraint.minWidth, constraint.maxWidth);
  next.height = clamp(next.height, constraint.minHeight, constraint.maxHeight);

  if (constraint.bounds) {
    if (constraint.alignX === "center") next.x = constraint.bounds.x + (constraint.bounds.width - next.width) / 2;
    if (constraint.alignX === "right") next.x = constraint.bounds.x + constraint.bounds.width - next.width;
    if (constraint.alignX === "left") next.x = constraint.bounds.x;
    if (constraint.alignY === "middle") next.y = constraint.bounds.y + (constraint.bounds.height - next.height) / 2;
    if (constraint.alignY === "bottom") next.y = constraint.bounds.y + constraint.bounds.height - next.height;
    if (constraint.alignY === "top") next.y = constraint.bounds.y;

    next.x = clamp(next.x, constraint.bounds.x, constraint.bounds.x + constraint.bounds.width - next.width);
    next.y = clamp(next.y, constraint.bounds.y, constraint.bounds.y + constraint.bounds.height - next.height);
  }

  if (constraint.snap && constraint.snap > 0) {
    next = {
      ...next,
      x: Math.round(next.x / constraint.snap) * constraint.snap,
      y: Math.round(next.y / constraint.snap) * constraint.snap,
      width: Math.round(next.width / constraint.snap) * constraint.snap,
      height: Math.round(next.height / constraint.snap) * constraint.snap
    };
  }

  return next;
}

function clamp(value: number, min: number | undefined, max: number | undefined): number {
  return Math.min(max ?? Number.POSITIVE_INFINITY, Math.max(min ?? Number.NEGATIVE_INFINITY, value));
}

function sameBox(left: LayoutBox, right: LayoutBox): boolean {
  return left.x === right.x && left.y === right.y && left.width === right.width && left.height === right.height;
}
