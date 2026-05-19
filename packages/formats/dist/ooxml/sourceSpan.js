import { Buffer } from "node:buffer";
import { createHash } from "node:crypto";
export function sourceSpanFromCharRange(source, charStart, charEnd) {
    assertCharRange(source, charStart, charEnd);
    return {
        start: charIndexToByteOffset(source, charStart),
        end: charIndexToByteOffset(source, charEnd),
        charStart,
        charEnd
    };
}
export function sourceSpanFromByteRange(source, start, end) {
    assertByteRange(source, start, end);
    return {
        start,
        end,
        charStart: byteOffsetToCharIndex(source, start),
        charEnd: byteOffsetToCharIndex(source, end)
    };
}
export function sliceSource(source, span) {
    const charStart = span.charStart ?? byteOffsetToCharIndex(source, span.start);
    const charEnd = span.charEnd ?? byteOffsetToCharIndex(source, span.end);
    return source.slice(charStart, charEnd);
}
export function createSourceFingerprint(source, span) {
    const exact = sliceSource(source, span);
    return {
        algorithm: "sha256",
        hash: createHash("sha256").update(exact, "utf8").digest("hex"),
        span,
        byteLength: span.end - span.start
    };
}
export function verifySourceFingerprint(source, fingerprint) {
    if (!isValidByteRange(source, fingerprint.span.start, fingerprint.span.end))
        return false;
    if (fingerprint.byteLength !== fingerprint.span.end - fingerprint.span.start)
        return false;
    const current = createSourceFingerprint(source, sourceSpanFromByteRange(source, fingerprint.span.start, fingerprint.span.end));
    return current.algorithm === fingerprint.algorithm && current.hash === fingerprint.hash;
}
export function charIndexToByteOffset(source, charIndex) {
    if (!Number.isInteger(charIndex) || charIndex < 0 || charIndex > source.length) {
        throw new RangeError(`Character offset ${charIndex} is outside the source range.`);
    }
    return buildUtf8OffsetMap(source)[charIndex] ?? 0;
}
export function byteOffsetToCharIndex(source, byteOffset) {
    if (!Number.isInteger(byteOffset) || byteOffset < 0 || byteOffset > Buffer.byteLength(source, "utf8")) {
        throw new RangeError(`Byte offset ${byteOffset} is outside the source range.`);
    }
    const offsets = buildUtf8OffsetMap(source);
    for (let index = 0; index < offsets.length; index += 1) {
        if (offsets[index] === byteOffset && (index === 0 || offsets[index - 1] !== byteOffset))
            return index;
    }
    throw new RangeError(`Byte offset ${byteOffset} does not align to a UTF-8 character boundary.`);
}
export function isValidByteRange(source, start, end) {
    try {
        assertByteRange(source, start, end);
        byteOffsetToCharIndex(source, start);
        byteOffsetToCharIndex(source, end);
        return true;
    }
    catch {
        return false;
    }
}
function assertCharRange(source, charStart, charEnd) {
    if (!Number.isInteger(charStart) || !Number.isInteger(charEnd) || charStart < 0 || charEnd < charStart || charEnd > source.length) {
        throw new RangeError(`Character range ${charStart}..${charEnd} is outside the source range.`);
    }
}
function assertByteRange(source, start, end) {
    const byteLength = Buffer.byteLength(source, "utf8");
    if (!Number.isInteger(start) || !Number.isInteger(end) || start < 0 || end < start || end > byteLength) {
        throw new RangeError(`Byte range ${start}..${end} is outside the source range.`);
    }
}
function buildUtf8OffsetMap(source) {
    const offsets = new Array(source.length + 1);
    let byteOffset = 0;
    for (let index = 0; index < source.length;) {
        offsets[index] = byteOffset;
        const codePoint = source.codePointAt(index);
        if (codePoint === undefined)
            break;
        const character = String.fromCodePoint(codePoint);
        for (let unit = 1; unit < character.length; unit += 1) {
            offsets[index + unit] = byteOffset;
        }
        byteOffset += Buffer.byteLength(character, "utf8");
        index += character.length;
    }
    offsets[source.length] = byteOffset;
    return offsets;
}
//# sourceMappingURL=sourceSpan.js.map