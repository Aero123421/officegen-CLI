import { createHash } from "node:crypto";
import { createCanvas, loadImage } from "@napi-rs/canvas";

export type VisualDiffStatus = "compared" | "skipped" | "blocked";

export interface RasterImageData {
  width: number;
  height: number;
  data: Uint8Array | Uint8ClampedArray;
  channels?: 3 | 4;
}

export interface RasterPixelDiffOptions {
  threshold?: number;
}

export interface RasterPixelDiffResult {
  kind: "raster-pixel";
  status: VisualDiffStatus;
  changed: boolean;
  beforeHash: string;
  afterHash: string;
  beforeWidth: number;
  beforeHeight: number;
  afterWidth: number;
  afterHeight: number;
  width: number;
  height: number;
  totalPixels: number;
  pixelsCompared: number;
  changedPixels: number;
  changedRatio: number;
  threshold: number;
  boundingBox?: { x: number; y: number; width: number; height: number };
  message?: string;
}

export async function comparePngPixels(
  before: Uint8Array,
  after: Uint8Array,
  options: RasterPixelDiffOptions = {}
): Promise<RasterPixelDiffResult> {
  try {
    return compareRasterPixels(await decodePngRaster(before), await decodePngRaster(after), options);
  } catch (error) {
    return blockedRasterDiff(before, after, error instanceof Error ? error.message : String(error), options);
  }
}

export async function decodePngRaster(bytes: Uint8Array): Promise<RasterImageData> {
  const image = await loadImage(Buffer.from(bytes));
  const canvas = createCanvas(image.width, image.height);
  const context = canvas.getContext("2d");
  context.drawImage(image, 0, 0);
  const imageData = context.getImageData(0, 0, image.width, image.height);
  return {
    width: image.width,
    height: image.height,
    data: imageData.data,
    channels: 4
  };
}

export function compareRasterPixels(
  before: RasterImageData,
  after: RasterImageData,
  options: RasterPixelDiffOptions = {}
): RasterPixelDiffResult {
  const threshold = normalizeThreshold(options.threshold);
  const beforeWidth = Math.max(0, Math.floor(before.width));
  const beforeHeight = Math.max(0, Math.floor(before.height));
  const afterWidth = Math.max(0, Math.floor(after.width));
  const afterHeight = Math.max(0, Math.floor(after.height));
  const width = Math.max(beforeWidth, afterWidth);
  const height = Math.max(beforeHeight, afterHeight);
  const totalPixels = width * height;
  const pixelsCompared = Math.min(beforeWidth, afterWidth) * Math.min(beforeHeight, afterHeight);
  let changedPixels = 0;
  let minX = Number.POSITIVE_INFINITY;
  let minY = Number.POSITIVE_INFINITY;
  let maxX = -1;
  let maxY = -1;

  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      const pixelChanged = x >= beforeWidth || y >= beforeHeight || x >= afterWidth || y >= afterHeight
        ? true
        : pixelsDiffer(before, after, x, y, threshold);
      if (!pixelChanged) continue;
      changedPixels += 1;
      minX = Math.min(minX, x);
      minY = Math.min(minY, y);
      maxX = Math.max(maxX, x);
      maxY = Math.max(maxY, y);
    }
  }

  return {
    kind: "raster-pixel",
    status: "compared",
    changed: changedPixels > 0,
    beforeHash: bytesHash(before.data),
    afterHash: bytesHash(after.data),
    beforeWidth,
    beforeHeight,
    afterWidth,
    afterHeight,
    width,
    height,
    totalPixels,
    pixelsCompared,
    changedPixels,
    changedRatio: totalPixels ? Number((changedPixels / totalPixels).toFixed(6)) : 0,
    threshold,
    boundingBox: changedPixels ? { x: minX, y: minY, width: maxX - minX + 1, height: maxY - minY + 1 } : undefined
  };
}

function pixelsDiffer(before: RasterImageData, after: RasterImageData, x: number, y: number, threshold: number): boolean {
  for (let channel = 0; channel < 4; channel += 1) {
    if (Math.abs(pixelChannel(before, x, y, channel) - pixelChannel(after, x, y, channel)) > threshold) return true;
  }
  return false;
}

function pixelChannel(image: RasterImageData, x: number, y: number, channel: number): number {
  const channels = image.channels ?? 4;
  const offset = (y * image.width + x) * channels;
  if (channel === 3 && channels === 3) return 255;
  return image.data[offset + Math.min(channel, channels - 1)] ?? 0;
}

function blockedRasterDiff(before: Uint8Array, after: Uint8Array, message: string, options: RasterPixelDiffOptions): RasterPixelDiffResult {
  return {
    kind: "raster-pixel",
    status: "blocked",
    changed: false,
    beforeHash: bytesHash(before),
    afterHash: bytesHash(after),
    beforeWidth: 0,
    beforeHeight: 0,
    afterWidth: 0,
    afterHeight: 0,
    width: 0,
    height: 0,
    totalPixels: 0,
    pixelsCompared: 0,
    changedPixels: 0,
    changedRatio: 0,
    threshold: normalizeThreshold(options.threshold),
    message
  };
}

function normalizeThreshold(value: number | undefined): number {
  if (value === undefined || !Number.isFinite(value)) return 0;
  return Math.max(0, Math.min(255, Math.floor(value)));
}

function bytesHash(bytes: Uint8Array | Uint8ClampedArray): string {
  return `sha256:${createHash("sha256").update(bytes).digest("hex")}`;
}
