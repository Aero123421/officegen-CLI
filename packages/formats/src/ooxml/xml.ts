import { XMLBuilder, XMLParser } from "fast-xml-parser";
import { decodeXmlEntities, escapeXml } from "../shared.js";

export const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  preserveOrder: false,
  parseTagValue: false,
  parseAttributeValue: false,
  trimValues: false
});

export const builder = new XMLBuilder({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  suppressEmptyNode: true,
  format: false
});

export function parseXml<T = unknown>(xml: string): T {
  return parser.parse(xml) as T;
}

export function buildXml(value: unknown): string {
  return builder.build(value);
}

export function ensureArray<T>(value: T | T[] | undefined): T[] {
  if (value === undefined) return [];
  return Array.isArray(value) ? value : [value];
}

export function xmlAttr(attrs: string, name: string): string | undefined {
  const pattern = new RegExp(`(?:^|\\s)${escapeRegExp(name)}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`);
  const match = pattern.exec(attrs);
  return match?.[1] ?? match?.[2];
}

export function localText(xml: string, localName: string): string[] {
  const tag = escapeRegExp(localName);
  const pattern = new RegExp(
    `<(?:[A-Za-z_][\\w.-]*:)?${tag}\\b[^>]*>([\\s\\S]*?)<\\/(?:[A-Za-z_][\\w.-]*:)?${tag}>`,
    "g"
  );
  return [...xml.matchAll(pattern)]
    .map((match) => decodeXmlEntities(stripTags(match[1] ?? "")).trim())
    .filter(Boolean);
}

export function exactText(xml: string, tagName: string): string[] {
  const tag = escapeRegExp(tagName);
  const pattern = new RegExp(`<${tag}(?:\\s[^>]*)?>([\\s\\S]*?)<\\/${tag}>`, "g");
  return [...xml.matchAll(pattern)]
    .map((match) => decodeXmlEntities(stripTags(match[1] ?? "")).trim())
    .filter(Boolean);
}

export function replaceAllXmlText(input: string, from: string, to: string): string {
  return input.split(escapeXmlText(from)).join(escapeXmlText(to));
}

export function escapeXmlText(value: string): string {
  return String(value).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

export function paragraphXml(text: string, namespace: "a" | "w" = "a"): string {
  if (namespace === "w") return `<w:p><w:r><w:t>${escapeXmlText(text)}</w:t></w:r></w:p>`;
  return `<a:p><a:r><a:t>${escapeXmlText(text)}</a:t></a:r></a:p>`;
}

export function bulletParagraphXml(text: string): string {
  return `<a:p><a:pPr><a:buChar char="&#8226;"/></a:pPr><a:r><a:t>${escapeXmlText(text)}</a:t></a:r></a:p>`;
}

export function stripTags(xml: string): string {
  return xml.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ");
}

export function replaceNthBlock(
  xml: string,
  pattern: RegExp,
  ordinal: number,
  replacer: (block: string) => string
): { changed: boolean; matched: boolean; xml: string } {
  let index = 0;
  let matched = false;
  const next = xml.replace(pattern, (match) => {
    index += 1;
    if (index !== ordinal) return match;
    matched = true;
    return replacer(match);
  });
  return { changed: next !== xml, matched, xml: next };
}

export function replaceFirstBlock(
  xml: string,
  pattern: RegExp,
  predicate: (block: string, ordinal: number) => boolean,
  replacer: (block: string, ordinal: number) => string
): { changed: boolean; matchCount: number; xml: string } {
  let ordinal = 0;
  let matchCount = 0;
  let replaced = false;
  const next = xml.replace(pattern, (match) => {
    ordinal += 1;
    if (!predicate(match, ordinal)) return match;
    matchCount += 1;
    if (replaced) return match;
    replaced = true;
    return replacer(match, ordinal);
  });
  return { changed: next !== xml, matchCount, xml: next };
}

export function setFirstTextInBlock(block: string, tagName: string, text: string): string {
  const tag = escapeRegExp(tagName);
  const pattern = new RegExp(`<${tag}(?=\\s|>)([^>]*)>[\\s\\S]*?<\\/${tag}>`, "g");
  let replaced = false;
  let sawText = false;
  const next = block.replace(pattern, (_match, attrs: string) => {
    sawText = true;
    if (!replaced) {
      replaced = true;
      return `<${tagName}${attrs}>${escapeXmlText(text)}</${tagName}>`;
    }
    return `<${tagName}${attrs}></${tagName}>`;
  });
  return sawText ? next : block;
}

export function preview(text: string | undefined, limit = 120): string | undefined {
  if (!text) return undefined;
  const normalized = text.replace(/\s+/g, " ").trim();
  return normalized.length > limit ? `${normalized.slice(0, limit - 1)}…` : normalized;
}

export function bboxFromBounds(bounds?: { x: number; y: number; width: number; height: number }): [number, number, number, number] | undefined {
  return bounds ? [bounds.x, bounds.y, bounds.width, bounds.height] : undefined;
}

export function emuToPx(value: number): number {
  return Math.round((value / 914400) * 96);
}

export function pxToEmu(value: number): number {
  return Math.round((value / 96) * 914400);
}

export function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

export function xmlData(value: unknown): string {
  return escapeXml(String(value ?? ""));
}
