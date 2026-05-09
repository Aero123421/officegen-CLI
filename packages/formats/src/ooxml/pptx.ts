import type JSZip from "jszip";
import type { ObjectBounds, ObjectMapEntry } from "../shared.js";
import { makeStableObjectId, readZipText, sortedZipFiles, stableHashId } from "../shared.js";
import { bulletParagraphXml, emuToPx, exactText, preview, replaceNthBlock, xmlAttr } from "./xml.js";
import { nextRelationshipId, parseRelationships, relationshipTarget } from "./relationships.js";

export interface PptxShape {
  stableObjectId: string;
  slideStableObjectId: string;
  slideIndex: number;
  shapeIndex: number;
  shapeId?: string;
  name?: string;
  text: string;
  textPreview?: string;
  bounds?: ObjectBounds;
  sourcePath: string;
}

interface PptxPicture {
  stableObjectId: string;
  slideIndex: number;
  pictureIndex: number;
  shapeId?: string;
  name?: string;
  relationshipId?: string;
  assetPath?: string;
  bounds?: ObjectBounds;
  sourcePath: string;
}

export interface PptxSlide {
  stableObjectId: string;
  index: number;
  sourcePath: string;
  relationshipId?: string;
  text: string;
  textObjects: ObjectMapEntry[];
  shapeCount: number;
  pictureCount: number;
  untrusted: true;
}

export async function getSlidePaths(zip: JSZip): Promise<string[]> {
  const files = sortedZipFiles(zip);
  const fallback = files.filter((path) => /^ppt\/slides\/slide\d+\.xml$/i.test(path)).sort(naturalSort);
  const presentationXml = await readZipText(zip, "ppt/presentation.xml");
  const relsXml = await readZipText(zip, "ppt/_rels/presentation.xml.rels");
  if (!presentationXml || !relsXml) return fallback;

  const rels = parseRelationships(relsXml);
  const ids = [...presentationXml.matchAll(/<p:sldId\b[^>]*\br:id="([^"]+)"[^>]*\/>/g)].map((match) => match[1] ?? "");
  const ordered = ids
    .map((id) => rels.find((rel) => rel.id === id))
    .filter(Boolean)
    .map((rel) => relationshipTarget("ppt", rel?.target ?? ""))
    .filter((path) => /^ppt\/slides\/slide\d+\.xml$/i.test(path) && zip.file(path));
  return ordered.length ? ordered : fallback;
}

export async function inspectSlides(zip: JSZip): Promise<{ slides: PptxSlide[]; objectMap: ObjectMapEntry[] }> {
  const slidePaths = await getSlidePaths(zip);
  const objectMap: ObjectMapEntry[] = [];
  const slides: PptxSlide[] = [];
  for (const [slideIndex, slidePath] of slidePaths.entries()) {
    const xml = (await readZipText(zip, slidePath)) ?? "";
    const relsPath = slidePath.replace(/^ppt\/slides\//, "ppt/slides/_rels/") + ".rels";
    const rels = parseRelationships((await readZipText(zip, relsPath)) ?? "");
    const slideNo = slideIndex + 1;
    const slideStableObjectId = stableHashId("pptx", "deck", "slide", slidePath);
    const shapes = extractShapes(xml, slideNo, slideStableObjectId, slidePath);
    const pictures = extractPictures(xml, slideNo, slidePath, rels);
    const tableCells = extractTableCells(xml, slideNo, slidePath);
    const textObjects = shapes
      .filter((shape) => shape.text)
      .map((shape) => {
        const entry: ObjectMapEntry = {
          stableObjectId: shape.stableObjectId,
          kind: "shape",
          label: shape.name,
          text: shape.text,
          textPreview: shape.textPreview,
          sourcePath: slidePath,
          xmlPath: slidePath,
          bounds: shape.bounds,
          bbox: shape.bounds ? [shape.bounds.x, shape.bounds.y, shape.bounds.width, shape.bounds.height] : undefined,
          selectorHints: {
            slide: slideNo,
            shapeId: shape.shapeId,
            name: shape.name,
            textPreview: shape.textPreview
          },
          trust: { level: "untrusted", reason: "document-content" },
          untrusted: true
        };
        objectMap.push(entry);
        return entry;
      })
      .concat(pictures.map((picture) => {
        const entry: ObjectMapEntry = {
          stableObjectId: picture.stableObjectId,
          kind: "picture",
          label: picture.name,
          sourcePath: slidePath,
          xmlPath: slidePath,
          bounds: picture.bounds,
          bbox: picture.bounds ? [picture.bounds.x, picture.bounds.y, picture.bounds.width, picture.bounds.height] : undefined,
          selectorHints: {
            slide: slideNo,
            shapeId: picture.shapeId,
            name: picture.name,
            relationshipId: picture.relationshipId,
            assetPath: picture.assetPath
          },
          editableOps: ["asset.replace", "pptx.replaceImageByShape"],
          media: {
            relationshipId: picture.relationshipId,
            assetPath: picture.assetPath
          },
          trust: { level: "untrusted", reason: "document-content" },
          untrusted: true
        };
        objectMap.push(entry);
        return entry;
      }))
      .concat(tableCells.map((cell) => {
        const entry: ObjectMapEntry = {
          stableObjectId: cell.stableObjectId,
          kind: "tableCell",
          text: cell.text,
          textPreview: cell.textPreview,
          sourcePath: slidePath,
          xmlPath: slidePath,
          selectorHints: { slide: slideNo, tableCell: cell.cellIndex, textPreview: cell.textPreview },
          trust: { level: "untrusted", reason: "document-content" },
          untrusted: true
        };
        objectMap.push(entry);
        return entry;
      }));
    slides.push({
      stableObjectId: slideStableObjectId,
      index: slideNo,
      sourcePath: slidePath,
      text: textObjects.map((entry) => entry.text).filter(Boolean).join("\n"),
      textObjects,
      shapeCount: shapes.length,
      pictureCount: pictures.length,
      untrusted: true
    });
  }
  return { slides, objectMap };
}

function extractPictures(xml: string, slideNo: number, sourcePath: string, rels: ReturnType<typeof parseRelationships>): PptxPicture[] {
  let pictureIndex = 0;
  return [...xml.matchAll(/<p:pic\b[\s\S]*?<\/p:pic>/g)].map((match) => {
    pictureIndex += 1;
    const block = match[0];
    const cNvPr = /<p:cNvPr\b([^>]*)\/>/.exec(block)?.[1] ?? "";
    const shapeId = xmlAttr(cNvPr, "id");
    const name = xmlAttr(cNvPr, "name");
    const relationshipId = /<a:blip\b[^>]*(?:r:embed|r:link)="([^"]+)"/.exec(block)?.[1];
    const rel = relationshipId ? rels.find((item) => item.id === relationshipId) : undefined;
    const assetPath = rel ? relationshipTarget("ppt/slides", rel.target) : undefined;
    return {
      stableObjectId: shapeId
        ? stableHashId("pptx", slideScope(sourcePath), "picture", `${sourcePath}#${shapeId}`)
        : makeStableObjectId("pptx", slideScope(sourcePath), "picture", pictureIndex),
      slideIndex: slideNo,
      pictureIndex,
      shapeId,
      name,
      relationshipId,
      assetPath,
      bounds: extractBounds(block),
      sourcePath
    };
  });
}

function extractTableCells(xml: string, slideNo: number, sourcePath: string): Array<{ stableObjectId: string; cellIndex: number; text: string; textPreview?: string }> {
  let cellIndex = 0;
  return [...xml.matchAll(/<a:tc\b[\s\S]*?<\/a:tc>/g)]
    .map((match) => {
      cellIndex += 1;
      const text = exactText(match[0], "a:t").join("");
      return {
        stableObjectId: stableHashId("pptx", slideScope(sourcePath), "tableCell", `${sourcePath}#${cellIndex}`),
        cellIndex,
        text,
        textPreview: preview(text)
      };
    })
    .filter((cell) => cell.text);
}

export function extractShapes(xml: string, slideNo: number, slideStableObjectId: string, sourcePath: string): PptxShape[] {
  let shapeIndex = 0;
  return [...xml.matchAll(/<p:sp\b[\s\S]*?<\/p:sp>/g)].map((match) => {
    shapeIndex += 1;
    const block = match[0];
    const cNvPr = /<p:cNvPr\b([^>]*)\/>/.exec(block)?.[1] ?? "";
    const shapeId = xmlAttr(cNvPr, "id");
    const name = xmlAttr(cNvPr, "name");
    const text = exactText(block, "a:t").join("");
    const scope = slideScope(sourcePath);
    const stableObjectId = shapeId
      ? stableHashId("pptx", scope, "shape", `${sourcePath}#${shapeId}`)
      : makeStableObjectId("pptx", scope, "shape", shapeIndex);
    return {
      stableObjectId,
      slideStableObjectId,
      slideIndex: slideNo,
      shapeIndex,
      shapeId,
      name,
      text,
      textPreview: preview(text),
      bounds: extractBounds(block),
      sourcePath
    };
  });
}

function slideScope(sourcePath: string): string {
  return `slide-${stablePathToken(sourcePath)}`;
}

function stablePathToken(sourcePath: string): string {
  let hash = 2166136261;
  for (const char of sourcePath.toLowerCase()) {
    hash ^= char.charCodeAt(0);
    hash = Math.imul(hash, 16777619);
  }
  return (hash >>> 0).toString(16).padStart(8, "0").slice(0, 8);
}

export function replaceShapeBulletItems(
  xml: string,
  ordinal: number,
  items: string[],
  mode: "insert" | "replace"
): { changed: boolean; matched: boolean; xml: string } {
  return replaceNthBlock(xml, /<p:sp\b[\s\S]*?<\/p:sp>/g, ordinal, (shape) => {
    const bulletXml = items.map((item) => bulletParagraphXml(item)).join("");
    if (!/<p:txBody\b[\s\S]*?<\/p:txBody>/.test(shape)) return shape;
    if (mode === "replace") {
      return shape.replace(/(<p:txBody\b[^>]*>)([\s\S]*?)(<\/p:txBody>)/, (_match, open: string, body: string, close: string) => {
        const bodyPr = /<a:bodyPr\b[\s\S]*?\/>/.exec(body)?.[0] ?? "<a:bodyPr/>";
        const lstStyle = /<a:lstStyle\b[\s\S]*?\/>/.exec(body)?.[0] ?? "<a:lstStyle/>";
        return `${open}${bodyPr}${lstStyle}${bulletXml}${close}`;
      });
    }
    return shape.replace(/<\/p:txBody>/, `${bulletXml}</p:txBody>`);
  });
}

export async function duplicateSlide(zip: JSZip, slideNumber: number, after?: number): Promise<void> {
  const slidePaths = await getSlidePaths(zip);
  const sourcePath = slidePaths[slideNumber - 1];
  if (!sourcePath) throw new Error(`SELECTOR_NOT_FOUND: pptx slide ${slideNumber} not found.`);
  const nextNo = nextSlideNumber(sortedZipFiles(zip));
  const nextPath = `ppt/slides/slide${nextNo}.xml`;
  zip.file(nextPath, (await readZipText(zip, sourcePath)) ?? "");
  const sourceRels = sourcePath.replace(/^ppt\/slides\//, "ppt/slides/_rels/") + ".rels";
  const targetRels = nextPath.replace(/^ppt\/slides\//, "ppt/slides/_rels/") + ".rels";
  const relsText = await readZipText(zip, sourceRels);
  if (relsText !== undefined) zip.file(targetRels, relsText);
  await addPresentationSlide(zip, nextNo, after ?? slideNumber);
  await addSlideContentType(zip, nextNo);
}

export async function reorderSlides(zip: JSZip, order: number[]): Promise<void> {
  const presentationXml = (await readZipText(zip, "ppt/presentation.xml")) ?? "";
  const ids = [...presentationXml.matchAll(/<p:sldId\b[^>]*\br:id="([^"]+)"[^>]*\/>/g)].map((match) => match[1] ?? "");
  if (!ids.length) throw new Error("SELECTOR_NOT_FOUND: pptx presentation slide list not found.");
  if (order.length !== ids.length || new Set(order).size !== ids.length || order.some((item) => item < 1 || item > ids.length)) {
    throw new Error(`SELECTOR_NOT_FOUND: pptx reorderSlides requires a permutation of 1..${ids.length}.`);
  }
  const sldIdTags = [...presentationXml.matchAll(/<p:sldId\b[^>]*\/>/g)].map((match) => match[0]);
  const reordered = order.map((slideNo) => sldIdTags[slideNo - 1]).join("");
  zip.file("ppt/presentation.xml", presentationXml.replace(/(<p:sldIdLst\b[^>]*>)[\s\S]*?(<\/p:sldIdLst>)/, `$1${reordered}$2`));
}

function extractBounds(block: string): ObjectBounds | undefined {
  const off = /<a:off\b([^>]*)\/>/.exec(block)?.[1];
  const ext = /<a:ext\b([^>]*)\/>/.exec(block)?.[1];
  if (!off || !ext) return undefined;
  const x = Number(xmlAttr(off, "x"));
  const y = Number(xmlAttr(off, "y"));
  const cx = Number(xmlAttr(ext, "cx"));
  const cy = Number(xmlAttr(ext, "cy"));
  if (![x, y, cx, cy].every(Number.isFinite)) return undefined;
  return { x: emuToPx(x), y: emuToPx(y), width: emuToPx(cx), height: emuToPx(cy) };
}

async function addPresentationSlide(zip: JSZip, slideNo: number, after: number): Promise<void> {
  const presentationXml = (await readZipText(zip, "ppt/presentation.xml")) ?? "";
  const relsXml = (await readZipText(zip, "ppt/_rels/presentation.xml.rels")) ?? '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
  const rId = nextRelationshipId(relsXml);
  const sldIds = [...presentationXml.matchAll(/<p:sldId\b[^>]*\bid="(\d+)"[^>]*\/>/g)].map((match) => Number(match[1])).filter(Number.isFinite);
  const newId = Math.max(255, ...sldIds) + 1;
  const slideTag = `<p:sldId id="${newId}" r:id="${rId}"/>`;
  const tags = [...presentationXml.matchAll(/<p:sldId\b[^>]*\/>/g)].map((match) => match[0]);
  const insertAt = Math.min(Math.max(after, 0), tags.length);
  tags.splice(insertAt, 0, slideTag);
  zip.file("ppt/presentation.xml", presentationXml.replace(/(<p:sldIdLst\b[^>]*>)[\s\S]*?(<\/p:sldIdLst>)/, `$1${tags.join("")}$2`));
  zip.file(
    "ppt/_rels/presentation.xml.rels",
    relsXml.replace(
      /<\/Relationships>\s*$/,
      `<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${slideNo}.xml"/></Relationships>`
    )
  );
}

async function addSlideContentType(zip: JSZip, slideNo: number): Promise<void> {
  const xml = await readZipText(zip, "[Content_Types].xml");
  if (!xml || xml.includes(`/ppt/slides/slide${slideNo}.xml`)) return;
  const override = `<Override PartName="/ppt/slides/slide${slideNo}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`;
  zip.file("[Content_Types].xml", xml.replace(/<\/Types>\s*$/, `${override}</Types>`));
}

function nextSlideNumber(paths: string[]): number {
  const numbers = paths
    .map((path) => /^ppt\/slides\/slide(\d+)\.xml$/i.exec(path)?.[1])
    .filter(Boolean)
    .map(Number);
  return Math.max(0, ...numbers) + 1;
}

function naturalSort(a: string, b: string): number {
  return a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" });
}
