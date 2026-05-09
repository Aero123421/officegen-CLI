import { inspect } from "./inspect.js";
import { view } from "./view.js";
export async function diffDocuments(before, after, options = {}) {
    const beforeInspect = await inspect(before, { depth: "shallow", config: options.config });
    const afterInspect = await inspect(after, { depth: "shallow", config: options.config });
    const semantic = semanticDiff(beforeInspect, afterInspect);
    const visual = options.visual ? await visualDiff(beforeInspect, afterInspect, options) : undefined;
    const visualRegressionScore = visual?.pageScores.length
        ? Number((visual.pageScores.reduce((sum, page) => sum + page.score, 0) / visual.pageScores.length).toFixed(4))
        : undefined;
    const changed = semantic.added.length > 0 || semantic.removed.length > 0 || semantic.changedText.length > 0 || (visualRegressionScore ?? 0) > 0;
    return {
        schema: "officegen.diff.result@1.2",
        formatBefore: beforeInspect.trusted.format,
        formatAfter: afterInspect.trusted.format,
        changed,
        summary: {
            addedObjects: semantic.added.length,
            removedObjects: semantic.removed.length,
            changedTextObjects: semantic.changedText.length,
            visualRegressionScore
        },
        semantic,
        visual,
        caveats: [
            "Visual diff is based on officegen's approximate SVG/HTML view, not a native Office rasterization.",
            "StableObjectId matching is best-effort across generated files and preserves strongest value for edits within the same document lineage."
        ]
    };
}
function semanticDiff(before, after) {
    const beforeMap = new Map(before.objectMap.map((entry) => [entry.stableObjectId, entry]));
    const afterMap = new Map(after.objectMap.map((entry) => [entry.stableObjectId, entry]));
    const added = [...afterMap.values()].filter((entry) => !beforeMap.has(entry.stableObjectId));
    const removed = [...beforeMap.values()].filter((entry) => !afterMap.has(entry.stableObjectId));
    const changedText = [...beforeMap.entries()]
        .map(([stableObjectId, beforeEntry]) => {
        const afterEntry = afterMap.get(stableObjectId);
        if (!afterEntry || (beforeEntry.text ?? beforeEntry.textPreview) === (afterEntry.text ?? afterEntry.textPreview))
            return undefined;
        return {
            stableObjectId,
            kind: beforeEntry.kind,
            before: beforeEntry.text ?? beforeEntry.textPreview,
            after: afterEntry.text ?? afterEntry.textPreview
        };
    })
        .filter((entry) => Boolean(entry));
    return { added, removed, changedText };
}
async function visualDiff(before, after, options) {
    const beforeView = await view(before, { format: "svg", maxPages: options.maxPages, config: options.config });
    const afterView = await view(after, { format: "svg", maxPages: options.maxPages, config: options.config });
    const pagesCompared = Math.min(beforeView.pages.length, afterView.pages.length);
    const pageScores = [];
    for (let index = 0; index < pagesCompared; index += 1) {
        const beforeHash = textHash(beforeView.pages[index]?.content ?? "");
        const afterHash = textHash(afterView.pages[index]?.content ?? "");
        pageScores.push({
            page: index + 1,
            beforeHash,
            afterHash,
            score: beforeHash === afterHash ? 0 : normalizedStringDistance(beforeView.pages[index]?.content ?? "", afterView.pages[index]?.content ?? "")
        });
    }
    return {
        fidelity: "approximate",
        pagesCompared,
        pageScores
    };
}
function normalizedStringDistance(before, after) {
    if (!before && !after)
        return 0;
    const max = Math.max(before.length, after.length, 1);
    let changed = Math.abs(before.length - after.length);
    const limit = Math.min(before.length, after.length);
    for (let index = 0; index < limit; index += 1) {
        if (before.charCodeAt(index) !== after.charCodeAt(index))
            changed += 1;
    }
    return Number(Math.min(1, changed / max).toFixed(4));
}
function textHash(value) {
    let hash = 2166136261;
    for (const char of value) {
        hash ^= char.charCodeAt(0);
        hash = Math.imul(hash, 16777619);
    }
    return (hash >>> 0).toString(16).padStart(8, "0");
}
export const diff = diffDocuments;
//# sourceMappingURL=diff.js.map