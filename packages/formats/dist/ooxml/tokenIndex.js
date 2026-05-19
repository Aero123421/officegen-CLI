import { sourceSpanFromCharRange, sliceSource } from "./sourceSpan.js";
export class TokenIndex {
    source;
    elements;
    attributes;
    textRuns;
    constructor(source, elements, attributes, textRuns) {
        this.source = source;
        this.elements = elements;
        this.attributes = attributes;
        this.textRuns = textRuns;
    }
    findElementsByName(name) {
        return this.elements.filter((element) => element.name === name);
    }
    findElementsByLocalName(localName) {
        return this.elements.filter((element) => element.localName === localName);
    }
    findAttributesByName(name) {
        return this.attributes.filter((attribute) => attribute.name === name);
    }
    sourceFor(span) {
        return sliceSource(this.source, span);
    }
}
export function buildTokenIndex(source) {
    const elements = [];
    const attributes = [];
    const textRuns = [];
    const stack = [];
    let cursor = 0;
    while (cursor < source.length) {
        const tagStart = source.indexOf("<", cursor);
        if (tagStart < 0) {
            addTextRun(source, cursor, source.length, stack, textRuns, false);
            break;
        }
        addTextRun(source, cursor, tagStart, stack, textRuns, false);
        if (source.startsWith("<!--", tagStart)) {
            cursor = requireTerminator(source, tagStart, "-->", 4) + 3;
            continue;
        }
        if (source.startsWith("<?", tagStart)) {
            cursor = requireTerminator(source, tagStart, "?>", 2) + 2;
            continue;
        }
        if (source.startsWith("<![CDATA[", tagStart)) {
            const cdataEnd = requireTerminator(source, tagStart, "]]>", 9);
            addCdataRun(source, tagStart, cdataEnd + 3, stack, textRuns);
            cursor = cdataEnd + 3;
            continue;
        }
        if (source.startsWith("</", tagStart)) {
            const tagEnd = findTagEnd(source, tagStart);
            closeElement(source, tagStart, tagEnd + 1, stack, elements);
            cursor = tagEnd + 1;
            continue;
        }
        if (source.startsWith("<!", tagStart)) {
            cursor = findTagEnd(source, tagStart) + 1;
            continue;
        }
        const tagEnd = findTagEnd(source, tagStart);
        openElement(source, tagStart, tagEnd + 1, stack, elements, attributes);
        cursor = tagEnd + 1;
    }
    return new TokenIndex(source, elements, attributes, textRuns);
}
export function checkXmlWellFormed(source) {
    const stack = [];
    const issues = [];
    let cursor = 0;
    try {
        while (cursor < source.length) {
            const tagStart = source.indexOf("<", cursor);
            if (tagStart < 0) {
                collectInvalidEntityReferences(source, cursor, source.length, issues);
                break;
            }
            collectInvalidEntityReferences(source, cursor, tagStart, issues);
            if (source.startsWith("<!--", tagStart)) {
                cursor = requireTerminator(source, tagStart, "-->", 4) + 3;
                continue;
            }
            if (source.startsWith("<?", tagStart)) {
                cursor = requireTerminator(source, tagStart, "?>", 2) + 2;
                continue;
            }
            if (source.startsWith("<![CDATA[", tagStart)) {
                cursor = requireTerminator(source, tagStart, "]]>", 9) + 3;
                continue;
            }
            const tagEnd = findTagEnd(source, tagStart);
            if (source.startsWith("</", tagStart)) {
                const name = readCloseName(source, tagStart, tagEnd + 1);
                const open = stack.pop();
                if (!open || open.name !== name) {
                    issues.push({
                        code: "XML_MISMATCHED_CLOSE_TAG",
                        message: `Expected closing tag for ${open?.name ?? "(none)"} but found ${name}.`,
                        offset: sourceSpanFromCharRange(source, tagStart, tagStart).start
                    });
                }
            }
            else if (source.startsWith("<!", tagStart)) {
                cursor = tagEnd + 1;
                continue;
            }
            else {
                const tag = readOpenTag(source, tagStart, tagEnd + 1);
                if (!tag.name) {
                    issues.push({
                        code: "XML_INVALID_TAG",
                        message: "Found an opening tag without a name.",
                        offset: sourceSpanFromCharRange(source, tagStart, tagStart).start
                    });
                }
                else if (!tag.selfClosing) {
                    stack.push({ name: tag.name, offset: sourceSpanFromCharRange(source, tagStart, tagStart).start });
                }
                for (const attribute of readAttributes(source, tag.attributeStartChar, tag.attributeEndChar, -1)) {
                    collectInvalidEntityReferences(source, attribute.valueSpan.charStart, attribute.valueSpan.charEnd, issues);
                }
            }
            cursor = tagEnd + 1;
        }
    }
    catch (error) {
        issues.push({
            code: "XML_UNTERMINATED_MARKUP",
            message: error instanceof Error ? error.message : "XML markup is unterminated.",
            offset: sourceSpanFromCharRange(source, Math.min(cursor, source.length), Math.min(cursor, source.length)).start
        });
    }
    for (const open of stack.reverse()) {
        issues.push({
            code: "XML_UNCLOSED_TAG",
            message: `Tag ${open.name} was not closed.`,
            offset: open.offset
        });
    }
    return { ok: issues.length === 0, issues };
}
function collectInvalidEntityReferences(source, charStart, charEnd, issues) {
    let cursor = charStart;
    while (cursor < charEnd) {
        const ampersand = source.indexOf("&", cursor);
        if (ampersand < 0 || ampersand >= charEnd)
            return;
        const semicolon = source.indexOf(";", ampersand + 1);
        const name = semicolon >= 0 && semicolon < charEnd ? source.slice(ampersand + 1, semicolon) : "";
        if (!isValidXmlEntityReference(name)) {
            issues.push({
                code: "XML_INVALID_ENTITY_REFERENCE",
                message: "Found an invalid or unescaped XML entity reference.",
                offset: sourceSpanFromCharRange(source, ampersand, ampersand).start
            });
            cursor = ampersand + 1;
            continue;
        }
        cursor = semicolon + 1;
    }
}
function isValidXmlEntityReference(name) {
    if (["amp", "lt", "gt", "apos", "quot"].includes(name))
        return true;
    if (/^#[0-9]+$/.test(name))
        return isValidXmlCodePoint(Number(name.slice(1)));
    if (/^#x[0-9a-fA-F]+$/.test(name))
        return isValidXmlCodePoint(Number.parseInt(name.slice(2), 16));
    return false;
}
function isValidXmlCodePoint(value) {
    return Number.isInteger(value) && (value === 0x9 ||
        value === 0xA ||
        value === 0xD ||
        (value >= 0x20 && value <= 0xD7FF) ||
        (value >= 0xE000 && value <= 0xFFFD) ||
        (value >= 0x10000 && value <= 0x10FFFF));
}
function openElement(source, tagStart, tagEndExclusive, stack, elements, allAttributes) {
    const tag = readOpenTag(source, tagStart, tagEndExclusive);
    const name = splitName(tag.name);
    const elementIndex = elements.length;
    const parent = stack.at(-1);
    const element = {
        kind: "element",
        index: elementIndex,
        depth: stack.length,
        parentIndex: parent?.index,
        ...name,
        span: sourceSpanFromCharRange(source, tagStart, tagEndExclusive),
        openTagSpan: sourceSpanFromCharRange(source, tagStart, tagEndExclusive),
        contentSpan: sourceSpanFromCharRange(source, tagEndExclusive, tagEndExclusive),
        selfClosing: tag.selfClosing,
        attributes: []
    };
    elements.push(element);
    const parsedAttributes = readAttributes(source, tag.attributeStartChar, tag.attributeEndChar, elementIndex);
    element.attributes.push(...parsedAttributes);
    allAttributes.push(...parsedAttributes);
    if (!tag.selfClosing) {
        stack.push({ index: elementIndex, name: tag.name, contentStartChar: tagEndExclusive });
    }
}
function closeElement(source, tagStart, tagEndExclusive, stack, elements) {
    const name = readCloseName(source, tagStart, tagEndExclusive);
    const pending = stack.pop();
    if (!pending)
        return;
    const element = elements[pending.index];
    if (!element)
        return;
    element.closeTagSpan = sourceSpanFromCharRange(source, tagStart, tagEndExclusive);
    element.contentSpan = sourceSpanFromCharRange(source, pending.contentStartChar, tagStart);
    element.span = sourceSpanFromCharRange(source, element.openTagSpan.charStart, tagEndExclusive);
    if (element.name !== name) {
        stack.push(pending);
    }
}
function addTextRun(source, charStart, charEnd, stack, textRuns, cdata) {
    if (charEnd <= charStart)
        return;
    const span = sourceSpanFromCharRange(source, charStart, charEnd);
    textRuns.push({
        kind: "text",
        span,
        valueSpan: span,
        parentIndex: stack.at(-1)?.index,
        text: source.slice(charStart, charEnd),
        cdata
    });
}
function addCdataRun(source, charStart, charEnd, stack, textRuns) {
    textRuns.push({
        kind: "text",
        span: sourceSpanFromCharRange(source, charStart, charEnd),
        valueSpan: sourceSpanFromCharRange(source, charStart + 9, charEnd - 3),
        parentIndex: stack.at(-1)?.index,
        text: source.slice(charStart + 9, charEnd - 3),
        cdata: true
    });
}
function readOpenTag(source, tagStart, tagEndExclusive) {
    let cursor = tagStart + 1;
    while (cursor < tagEndExclusive && /\s/.test(source[cursor] ?? ""))
        cursor += 1;
    const nameStart = cursor;
    while (cursor < tagEndExclusive && !/[\s/>]/.test(source[cursor] ?? ""))
        cursor += 1;
    const name = source.slice(nameStart, cursor);
    let attributeEndChar = tagEndExclusive - 1;
    while (attributeEndChar > cursor && /\s/.test(source[attributeEndChar - 1] ?? ""))
        attributeEndChar -= 1;
    const selfClosing = source[attributeEndChar - 1] === "/";
    if (selfClosing)
        attributeEndChar -= 1;
    return { name, selfClosing, attributeStartChar: cursor, attributeEndChar };
}
function readCloseName(source, tagStart, tagEndExclusive) {
    let cursor = tagStart + 2;
    while (cursor < tagEndExclusive && /\s/.test(source[cursor] ?? ""))
        cursor += 1;
    const nameStart = cursor;
    while (cursor < tagEndExclusive && !/[\s>]/.test(source[cursor] ?? ""))
        cursor += 1;
    return source.slice(nameStart, cursor);
}
function readAttributes(source, charStart, charEnd, elementIndex) {
    const attributes = [];
    let cursor = charStart;
    while (cursor < charEnd) {
        while (cursor < charEnd && /\s/.test(source[cursor] ?? ""))
            cursor += 1;
        if (cursor >= charEnd || source[cursor] === "/")
            break;
        const nameStart = cursor;
        while (cursor < charEnd && !/[\s=/>]/.test(source[cursor] ?? ""))
            cursor += 1;
        const nameEnd = cursor;
        while (cursor < charEnd && /\s/.test(source[cursor] ?? ""))
            cursor += 1;
        if (source[cursor] !== "=")
            break;
        cursor += 1;
        while (cursor < charEnd && /\s/.test(source[cursor] ?? ""))
            cursor += 1;
        const quote = source[cursor];
        if (quote !== "\"" && quote !== "'")
            break;
        const rawValueStart = cursor;
        cursor += 1;
        const valueStart = cursor;
        while (cursor < charEnd && source[cursor] !== quote)
            cursor += 1;
        const valueEnd = cursor;
        if (source[cursor] !== quote)
            break;
        cursor += 1;
        const rawValueEnd = cursor;
        const xmlName = splitName(source.slice(nameStart, nameEnd));
        attributes.push({
            kind: "attribute",
            elementIndex,
            ...xmlName,
            span: sourceSpanFromCharRange(source, nameStart, rawValueEnd),
            nameSpan: sourceSpanFromCharRange(source, nameStart, nameEnd),
            rawValueSpan: sourceSpanFromCharRange(source, rawValueStart, rawValueEnd),
            valueSpan: sourceSpanFromCharRange(source, valueStart, valueEnd),
            quote,
            value: source.slice(valueStart, valueEnd)
        });
    }
    return attributes;
}
function findTagEnd(source, tagStart) {
    let quote;
    for (let cursor = tagStart + 1; cursor < source.length; cursor += 1) {
        const character = source[cursor];
        if (quote) {
            if (character === quote)
                quote = undefined;
            continue;
        }
        if (character === "\"" || character === "'") {
            quote = character;
            continue;
        }
        if (character === ">")
            return cursor;
    }
    throw new Error(`Unterminated XML tag at character offset ${tagStart}.`);
}
function requireTerminator(source, start, terminator, contentOffset) {
    const found = source.indexOf(terminator, start + contentOffset);
    if (found < 0)
        throw new Error(`Unterminated XML markup at character offset ${start}.`);
    return found;
}
function splitName(name) {
    const colon = name.indexOf(":");
    if (colon < 0)
        return { name, localName: name };
    return { name, prefix: name.slice(0, colon), localName: name.slice(colon + 1) };
}
//# sourceMappingURL=tokenIndex.js.map