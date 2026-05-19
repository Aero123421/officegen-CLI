import { byteOffsetToCharIndex, isValidByteRange, verifySourceFingerprint } from "./sourceSpan.js";
import { checkXmlWellFormed } from "./tokenIndex.js";
export class PatchEngineError extends Error {
    code;
    constructor(code, message) {
        super(message);
        this.code = code;
        this.name = "PatchEngineError";
    }
}
export function applyXmlPatches(source, patches, options = {}) {
    const normalized = patches.map((patch, originalIndex) => normalizePatch(source, patch, originalIndex));
    for (const item of normalized) {
        const fingerprint = item.patch.fingerprint;
        if (fingerprint && !verifySourceFingerprint(source, fingerprint)) {
            throw new PatchEngineError("PATCH_STALE_FINGERPRINT", "Patch fingerprint is stale for the current XML source.");
        }
    }
    rejectOverlaps(normalized);
    const ordered = [...normalized].sort((left, right) => {
        if (left.start !== right.start)
            return left.start - right.start;
        if (left.isInsert !== right.isInsert)
            return left.isInsert ? -1 : 1;
        return left.originalIndex - right.originalIndex;
    });
    let result = "";
    let cursorChar = 0;
    for (const item of ordered) {
        const charStart = byteOffsetToCharIndex(source, item.start);
        const charEnd = byteOffsetToCharIndex(source, item.end);
        if (charStart < cursorChar) {
            throw new PatchEngineError("PATCH_OVERLAP", "Patches overlap after offset ordering.");
        }
        result += source.slice(cursorChar, charStart);
        result += item.value;
        cursorChar = charEnd;
    }
    result += source.slice(cursorChar);
    if (options.validateWellFormed !== false) {
        const wellFormed = checkXmlWellFormed(result);
        if (!wellFormed.ok) {
            throw new PatchEngineError("PATCH_NOT_WELL_FORMED", `Patched XML is not well-formed: ${wellFormed.issues[0]?.message ?? "unknown error"}`);
        }
    }
    return result;
}
function normalizePatch(source, patch, originalIndex) {
    if (patch.type === "insert") {
        if (!isValidByteRange(source, patch.offset, patch.offset)) {
            throw new PatchEngineError("PATCH_OFFSET_OUT_OF_RANGE", `Patch insert offset ${patch.offset} is outside the XML source.`);
        }
        return { patch, originalIndex, start: patch.offset, end: patch.offset, value: patch.value, isInsert: true };
    }
    if (!isValidByteRange(source, patch.span.start, patch.span.end)) {
        throw new PatchEngineError("PATCH_OFFSET_OUT_OF_RANGE", `Patch span ${patch.span.start}..${patch.span.end} is outside the XML source.`);
    }
    return {
        patch,
        originalIndex,
        start: patch.span.start,
        end: patch.span.end,
        value: patch.type === "replace" ? patch.value : "",
        isInsert: false
    };
}
function rejectOverlaps(patches) {
    for (let leftIndex = 0; leftIndex < patches.length; leftIndex += 1) {
        for (let rightIndex = leftIndex + 1; rightIndex < patches.length; rightIndex += 1) {
            const left = patches[leftIndex];
            const right = patches[rightIndex];
            if (!left || !right)
                continue;
            if (overlaps(left, right)) {
                throw new PatchEngineError("PATCH_OVERLAP", `Patch ${left.originalIndex} overlaps patch ${right.originalIndex}.`);
            }
        }
    }
}
function overlaps(left, right) {
    if (left.start < right.end && right.start < left.end)
        return true;
    if (left.isInsert && !right.isInsert)
        return left.start > right.start && left.start < right.end;
    if (right.isInsert && !left.isInsert)
        return right.start > left.start && right.start < left.end;
    return false;
}
//# sourceMappingURL=patchEngine.js.map