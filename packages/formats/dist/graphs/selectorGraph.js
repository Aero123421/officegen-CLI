import { normalizeText } from "./objectGraph.js";
export const SELECTOR_GRAPH_LOW_CONFIDENCE_THRESHOLD = 0.65;
export const SELECTOR_GRAPH_AMBIGUITY_DISTANCE_DELTA = 24;
export const SELECTOR_GRAPH_AMBIGUITY_DISTANCE_RATIO = 1.15;
export function resolveGraphSelector(graph, selector) {
    const result = applySelector(graph, graph.nodes.map((node) => ({ node, evidence: [], confidenceParts: [] })), selector);
    const matches = result.candidates.map(candidateToMatch);
    const ambiguity = result.forcedAmbiguity ?? detectAmbiguity(result.candidates);
    const confidence = matches.length ? Number(Math.max(...matches.map((match) => match.confidence)).toFixed(2)) : undefined;
    const lowConfidence = matches.length > 0 && !ambiguity.ambiguous && confidence !== undefined && confidence < SELECTOR_GRAPH_LOW_CONFIDENCE_THRESHOLD;
    const status = matches.length === 0
        ? "not-found"
        : ambiguity.ambiguous
            ? "ambiguous"
            : lowConfidence
                ? "low-confidence"
                : "matched";
    return {
        graphVersion: graph.graphVersion,
        selector,
        status,
        matched: status === "matched" || status === "low-confidence",
        matchCount: matches.length,
        confidence,
        lowConfidence,
        ambiguity,
        matches,
        evidence: result.evidence
    };
}
export function resolveGraphSelectors(graph, selectors) {
    return selectors.map((selector) => resolveGraphSelector(graph, selector));
}
function applySelector(graph, initialCandidates, selector) {
    let candidates = initialCandidates;
    const evidence = [];
    if (selector.nodeId) {
        const item = criterion("object-map", 1, "Matched exact graph nodeId.", "nodeId");
        candidates = addCriterion(candidates.filter(({ node }) => node.nodeId === selector.nodeId), item);
        evidence.push(item);
    }
    if (selector.stableId) {
        const item = criterion("object-map", 1, "Matched exact stable object id.", "stableId");
        candidates = addCriterion(candidates.filter(({ node }) => node.stableId === selector.stableId), item);
        evidence.push(item);
    }
    if (selector.type) {
        const types = new Set((Array.isArray(selector.type) ? selector.type : [selector.type]).map((type) => type.toLowerCase()));
        const item = criterion("object-map", 0.62, "Matched object type.", "type");
        candidates = addCriterion(candidates.filter(({ node }) => types.has(node.type.toLowerCase())), item);
        evidence.push(item);
    }
    if (selector.slide !== undefined) {
        const item = criterion("selector-hint", 0.58, "Matched slide scope.", "source.slide");
        candidates = addCriterion(candidates.filter(({ node }) => node.source.slide === selector.slide), item);
        evidence.push(item);
    }
    if (selector.page !== undefined) {
        const item = criterion("selector-hint", 0.58, "Matched page scope.", "source.page");
        candidates = addCriterion(candidates.filter(({ node }) => node.source.page === selector.page), item);
        evidence.push(item);
    }
    if (selector.sheet !== undefined) {
        const item = criterion("selector-hint", 0.58, "Matched sheet scope.", "source.sheet");
        candidates = addCriterion(candidates.filter(({ node }) => node.source.sheet === selector.sheet), item);
        evidence.push(item);
    }
    if (selector.text) {
        const textSelector = normalizeTextSelector(selector.text);
        const item = criterion("object-map", textSelector.exact ? 0.84 : 0.74, textSelector.exact ? "Matched exact text." : "Matched text containment.", "text");
        candidates = addCriterion(candidates.filter(({ node }) => textMatches(node, textSelector)), item);
        evidence.push(item);
    }
    if (selector.bbox) {
        const bboxSelector = normalizeBBoxSelector(selector.bbox);
        const item = criterion("geometry", 0.7, `Matched bbox using ${bboxSelector.mode} mode.`, "bbox");
        candidates = addCriterion(candidates.filter(({ node }) => node.bbox !== undefined && bboxMatches(node.bbox, bboxSelector)), item);
        evidence.push(item);
    }
    if (selector.relation) {
        const relationResult = applyRelation(graph, candidates, selector.relation);
        candidates = relationResult.candidates;
        evidence.push(...relationResult.evidence);
        if (relationResult.forcedAmbiguity)
            return { candidates, evidence, forcedAmbiguity: relationResult.forcedAmbiguity };
    }
    if (selector.rightOf) {
        const rightOfResult = applyRightOf(graph, candidates, selector.rightOf, selector.slide);
        candidates = rightOfResult.candidates;
        evidence.push(...rightOfResult.evidence);
        if (rightOfResult.forcedAmbiguity)
            return { candidates, evidence, forcedAmbiguity: rightOfResult.forcedAmbiguity };
    }
    if (selector.nearestTo) {
        const nearestResult = applyNearestTo(candidates, selector.nearestTo);
        candidates = nearestResult.candidates;
        evidence.push(...nearestResult.evidence);
    }
    return { candidates, evidence };
}
function applyRelation(graph, candidates, relation) {
    const anchor = relationAnchor(graph, relation);
    const direction = relation.direction ?? "outgoing";
    const item = criterion("derived", 0.82, `Matched graph relation ${relation.relation}.`, "edges");
    if (anchor.forcedAmbiguity) {
        return {
            candidates: addCriterion(anchor.candidates ?? [], item),
            evidence: [item],
            forcedAmbiguity: anchor.forcedAmbiguity
        };
    }
    const related = candidates.filter(({ node }) => graph.edges.some((edge) => {
        if (edge.relation !== relation.relation)
            return false;
        const nodeSide = direction === "outgoing" ? edge.from : edge.to;
        const anchorSide = direction === "outgoing" ? edge.to : edge.from;
        return nodeSide === node.nodeId && (!anchor.ids || anchor.ids.has(anchorSide));
    }));
    return {
        candidates: addCriterion(related, item),
        evidence: [item]
    };
}
function applyRightOf(graph, candidates, selector, fallbackSlide) {
    const spec = typeof selector === "string" ? { text: selector, slide: fallbackSlide } : { ...selector, slide: selector.slide ?? fallbackSlide };
    const anchors = graph.nodes.filter((node) => {
        if (!node.bbox || !nodeTextContains(node, spec.text))
            return false;
        if (spec.slide !== undefined && node.source.slide !== spec.slide)
            return false;
        if (spec.type === undefined)
            return true;
        const types = new Set((Array.isArray(spec.type) ? spec.type : [spec.type]).map((type) => type.toLowerCase()));
        return types.has(node.type.toLowerCase());
    });
    const item = criterion("geometry", 0.82, "Matched nearest object to the right of an anchor.", "rightOf");
    if (anchors.length === 0)
        return { candidates: [], evidence: [item] };
    if (anchors.length > 1) {
        return {
            candidates: addCriterion(anchors.map((node) => ({ node, evidence: [], confidenceParts: [] })), item),
            evidence: [item],
            forcedAmbiguity: {
                ambiguous: true,
                reason: "rightOf-anchor-ambiguous",
                candidateNodeIds: anchors.map((node) => node.nodeId)
            }
        };
    }
    const anchor = anchors[0];
    const fromEdges = new Set(graph.edges.filter((edge) => edge.relation === "rightOf" && edge.to === anchor.nodeId).map((edge) => edge.from));
    const ranked = candidates
        .filter(({ node }) => node.bbox && node.nodeId !== anchor.nodeId && sameSlideOrUnscoped(node, anchor) && (fromEdges.has(node.nodeId) || isRightOf(node.bbox, anchor.bbox)))
        .map((candidate) => ({ ...candidate, distance: horizontalGap(candidate.node.bbox, anchor.bbox) }))
        .sort((left, right) => (left.distance ?? 0) - (right.distance ?? 0));
    const best = ranked[0];
    if (!best)
        return { candidates: [], evidence: [item] };
    const close = ranked.filter((candidate) => (candidate.distance ?? 0) - (best.distance ?? 0) <= SELECTOR_GRAPH_AMBIGUITY_DISTANCE_DELTA);
    return {
        candidates: addCriterion(close, item),
        evidence: [item]
    };
}
function applyNearestTo(candidates, point) {
    const maxDistance = point.maxDistance ?? 1000;
    const ranked = candidates
        .filter(({ node }) => node.bbox && (point.slide === undefined || node.source.slide === point.slide))
        .map((candidate) => ({ ...candidate, distance: centerDistance(candidate.node, point.x, point.y) }))
        .filter((candidate) => (candidate.distance ?? Number.MAX_SAFE_INTEGER) <= maxDistance)
        .sort((left, right) => (left.distance ?? 0) - (right.distance ?? 0));
    const best = ranked[0];
    const item = criterion("geometry", 0.82, "Matched nearest object to point.", "nearestTo");
    if (!best)
        return { candidates: [], evidence: [item] };
    const close = ranked.filter((candidate) => {
        const distance = candidate.distance ?? 0;
        const bestDistance = best.distance ?? 0;
        return distance - bestDistance <= SELECTOR_GRAPH_AMBIGUITY_DISTANCE_DELTA || distance <= bestDistance * SELECTOR_GRAPH_AMBIGUITY_DISTANCE_RATIO;
    });
    return {
        candidates: close.map((candidate) => {
            const distance = candidate.distance ?? 0;
            const confidence = Number(Math.max(0.55, Math.min(0.98, 1 - distance / maxDistance)).toFixed(2));
            const distanceEvidence = criterion("geometry", confidence, `Center distance ${Number(distance.toFixed(2))}.`, "nearestTo");
            return {
                ...candidate,
                evidence: [...candidate.evidence, item, distanceEvidence],
                confidenceParts: [...candidate.confidenceParts, confidence]
            };
        }),
        evidence: [item]
    };
}
function relationAnchor(graph, relation) {
    if (relation.nodeId)
        return { ids: new Set([relation.nodeId]) };
    if (relation.stableId) {
        const node = graph.nodes.find((item) => item.stableId === relation.stableId);
        return { ids: node ? new Set([node.nodeId]) : new Set() };
    }
    if (!relation.anchor)
        return {};
    const resolution = resolveGraphSelector(graph, relation.anchor);
    if (resolution.ambiguity.ambiguous) {
        const anchorIds = new Set(resolution.ambiguity.candidateNodeIds);
        return {
            ids: anchorIds,
            candidates: graph.nodes
                .filter((node) => anchorIds.has(node.nodeId))
                .map((node) => ({ node, evidence: [], confidenceParts: [] })),
            forcedAmbiguity: {
                ambiguous: true,
                reason: "relation-anchor-ambiguous",
                candidateNodeIds: resolution.ambiguity.candidateNodeIds
            }
        };
    }
    return { ids: new Set(resolution.matches.map((match) => match.nodeId)) };
}
function addCriterion(candidates, evidence) {
    return candidates.map((candidate) => ({
        ...candidate,
        evidence: [...candidate.evidence, evidence],
        confidenceParts: [...candidate.confidenceParts, evidence.confidence]
    }));
}
function candidateToMatch(candidate) {
    return {
        nodeId: candidate.node.nodeId,
        stableId: candidate.node.stableId,
        type: candidate.node.type,
        label: candidate.node.label,
        text: candidate.node.text?.value ?? candidate.node.text?.preview,
        bbox: candidate.node.bbox,
        source: candidate.node.source,
        confidence: candidateConfidence(candidate),
        evidence: candidate.evidence
    };
}
function candidateConfidence(candidate) {
    if (!candidate.confidenceParts.length)
        return 0;
    if (candidate.confidenceParts.includes(1))
        return 1;
    const average = candidate.confidenceParts.reduce((sum, value) => sum + value, 0) / candidate.confidenceParts.length;
    const bonus = Math.min(0.12, Math.max(0, candidate.confidenceParts.length - 1) * 0.04);
    return Number(Math.min(0.98, average + bonus).toFixed(2));
}
function detectAmbiguity(candidates) {
    if (candidates.length <= 1) {
        return { ambiguous: false, candidateNodeIds: candidates.map(({ node }) => node.nodeId) };
    }
    const nearestTie = candidates.every((candidate) => candidate.distance !== undefined);
    return {
        ambiguous: true,
        reason: nearestTie ? "nearest-distance-tie" : "multiple-matches",
        candidateNodeIds: candidates.map(({ node }) => node.nodeId)
    };
}
function normalizeTextSelector(selector) {
    return typeof selector === "string"
        ? { text: selector, exact: false, caseSensitive: false }
        : { text: selector.text, exact: selector.exact ?? false, caseSensitive: selector.caseSensitive ?? false };
}
function textMatches(node, selector) {
    const text = node.text?.value ?? node.text?.preview ?? "";
    const actual = selector.caseSensitive ? text : normalizeText(text);
    const expected = selector.caseSensitive ? selector.text : normalizeText(selector.text);
    return selector.exact ? actual === expected : actual.includes(expected);
}
function nodeTextContains(node, text) {
    return normalizeText(node.text?.value ?? node.text?.preview ?? "").includes(normalizeText(text));
}
function normalizeBBoxSelector(selector) {
    if (Array.isArray(selector)) {
        return { x: selector[0], y: selector[1], width: selector[2], height: selector[3], tolerance: 0, mode: "near" };
    }
    return {
        x: selector.x,
        y: selector.y,
        width: selector.width,
        height: selector.height,
        tolerance: selector.tolerance ?? 0,
        mode: selector.mode ?? "near"
    };
}
function bboxMatches(actual, selector) {
    if (!actual)
        return false;
    const expected = [selector.x, selector.y, selector.width, selector.height];
    if (selector.mode === "intersects")
        return intersects(actual, expected);
    if (selector.mode === "contains")
        return contains(actual, expected);
    return Math.abs(actual[0] - selector.x) <= selector.tolerance &&
        Math.abs(actual[1] - selector.y) <= selector.tolerance &&
        Math.abs(actual[2] - selector.width) <= selector.tolerance &&
        Math.abs(actual[3] - selector.height) <= selector.tolerance;
}
function criterion(kind, confidence, message, sourceField) {
    return { kind, confidence, message, sourceField };
}
function sameSlideOrUnscoped(left, right) {
    return left.source.slide === undefined || right.source.slide === undefined || left.source.slide === right.source.slide;
}
function isRightOf(source, target) {
    if (!source || !target)
        return false;
    return source[0] >= target[0] + target[2] && verticalOverlapRatio(source, target) >= 0.25;
}
function horizontalGap(source, target) {
    return Math.max(0, source[0] - (target[0] + target[2]));
}
function centerDistance(node, x, y) {
    if (!node.bbox)
        return Number.MAX_SAFE_INTEGER;
    const cx = node.bbox[0] + node.bbox[2] / 2;
    const cy = node.bbox[1] + node.bbox[3] / 2;
    return Math.hypot(cx - x, cy - y);
}
function contains(outer, inner) {
    return inner[0] >= outer[0] &&
        inner[1] >= outer[1] &&
        inner[0] + inner[2] <= outer[0] + outer[2] &&
        inner[1] + inner[3] <= outer[1] + outer[3];
}
function intersects(left, right) {
    return left[0] < right[0] + right[2] &&
        left[0] + left[2] > right[0] &&
        left[1] < right[1] + right[3] &&
        left[1] + left[3] > right[1];
}
function verticalOverlapRatio(left, right) {
    const top = Math.max(left[1], right[1]);
    const bottom = Math.min(left[1] + left[3], right[1] + right[3]);
    const overlap = Math.max(0, bottom - top);
    return overlap / Math.max(1, Math.min(left[3], right[3]));
}
//# sourceMappingURL=selectorGraph.js.map