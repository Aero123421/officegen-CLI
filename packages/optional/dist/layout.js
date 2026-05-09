import path from "node:path";
import { edit } from "../../formats/dist/index.js";
import { featureRoot, nowIso, requireFeature, slugify, writeJsonFile } from "./common.js";
export async function applyLayoutConstraints(options) {
    requireFeature(options, "layout", "layout apply");
    const constraints = new Map(options.constraints.map((constraint) => [constraint.id, constraint]));
    const changes = [];
    const boxes = options.boxes.map((box) => {
        const after = applyConstraint(box, constraints.get(box.id));
        if (!sameBox(box, after)) {
            changes.push({ id: box.id, before: box, after });
        }
        return after;
    });
    if (options.targetPath && options.outputPath && ["pptx"].includes(path.extname(options.outputPath).replace(/^\./, "").toLowerCase())) {
        const ops = changes.map((change) => ({
            op: "pptx.setBounds",
            selector: { stableObjectId: change.id },
            bounds: { x: change.after.x, y: change.after.y, width: change.after.width, height: change.after.height }
        }));
        const editResult = await edit(path.resolve(options.cwd ?? process.cwd(), options.targetPath), ops, {
            out: options.outputPath,
            format: "pptx",
            resolveSelectors: true,
            validateFirst: true,
            atomic: true
        });
        return {
            kind: "officegen.layout.apply",
            generatedAt: nowIso(),
            boxes,
            changes,
            note: "Applied layout constraints directly to PPTX object bounds.",
            ...{ targetPath: options.targetPath, out: options.outputPath, mutatesOffice: true, editResult }
        };
    }
    const result = {
        kind: "officegen.layout.apply",
        generatedAt: nowIso(),
        boxes,
        changes,
        note: "Simple layout constraints computed. Provide targetPath and an Office --out path to mutate PPTX bounds."
    };
    const outputPath = options.outputPath ??
        path.join(featureRoot(options, "layout"), "runs", `${slugify(options.planId ?? "layout")}.apply.json`);
    await writeJsonFile(outputPath, result);
    return result;
}
function applyConstraint(box, constraint) {
    if (!constraint) {
        return box;
    }
    let next = { ...box };
    next.width = clamp(next.width, constraint.minWidth, constraint.maxWidth);
    next.height = clamp(next.height, constraint.minHeight, constraint.maxHeight);
    if (constraint.bounds) {
        if (constraint.alignX === "center")
            next.x = constraint.bounds.x + (constraint.bounds.width - next.width) / 2;
        if (constraint.alignX === "right")
            next.x = constraint.bounds.x + constraint.bounds.width - next.width;
        if (constraint.alignX === "left")
            next.x = constraint.bounds.x;
        if (constraint.alignY === "middle")
            next.y = constraint.bounds.y + (constraint.bounds.height - next.height) / 2;
        if (constraint.alignY === "bottom")
            next.y = constraint.bounds.y + constraint.bounds.height - next.height;
        if (constraint.alignY === "top")
            next.y = constraint.bounds.y;
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
    next.width = clamp(next.width, constraint.minWidth, constraint.maxWidth);
    next.height = clamp(next.height, constraint.minHeight, constraint.maxHeight);
    if (constraint.bounds) {
        next.x = clamp(next.x, constraint.bounds.x, constraint.bounds.x + constraint.bounds.width - next.width);
        next.y = clamp(next.y, constraint.bounds.y, constraint.bounds.y + constraint.bounds.height - next.height);
    }
    return next;
}
function clamp(value, min, max) {
    return Math.min(max ?? Number.POSITIVE_INFINITY, Math.max(min ?? Number.NEGATIVE_INFINITY, value));
}
function sameBox(left, right) {
    return left.x === right.x && left.y === right.y && left.width === right.width && left.height === right.height;
}
//# sourceMappingURL=layout.js.map