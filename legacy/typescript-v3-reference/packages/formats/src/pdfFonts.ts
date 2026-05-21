import { existsSync } from "node:fs";
import { readFile } from "node:fs/promises";
import { dirname, join } from "node:path";
import { createRequire } from "node:module";
import fontkit from "@pdf-lib/fontkit";
import { OfficegenError } from "@officegen/core";
import { type PDFDocument, type PDFFont, StandardFonts } from "pdf-lib";

export interface PdfFontSet {
  font: PDFFont;
  bold: PDFFont;
  embeddedUnicode: boolean;
  fontPath?: string;
  caveats: string[];
}

export async function embedPdfFonts(pdf: PDFDocument, textSamples: string[]): Promise<PdfFontSet> {
  const text = textSamples.join("\n");
  if (!needsUnicodeFont(text)) {
    return {
      font: await pdf.embedFont(StandardFonts.Helvetica),
      bold: await pdf.embedFont(StandardFonts.HelveticaBold),
      embeddedUnicode: false,
      caveats: []
    };
  }

  const attempts: string[] = [];
  for (const fontPath of findCjkFontPaths()) {
    try {
      pdf.registerFontkit(fontkit);
      const fontBytes = await readFile(fontPath);
      const font = await pdf.embedFont(fontBytes, { subset: true });
      return {
        font,
        bold: font,
        embeddedUnicode: true,
        fontPath,
        caveats: [`Embedded Unicode PDF font: ${fontPath}.`]
      };
    } catch (error) {
      attempts.push(`${fontPath}: ${error instanceof Error ? error.message : String(error)}`);
    }
  }
  {
    throw new OfficegenError(
      "RENDER_FONT_UNSUPPORTED",
      "PDF text contains CJK or other Unicode characters, but no embeddable CJK font was found. Set OFFICEGEN_PDF_FONT to a .ttf/.otf font path or use native export.",
      { searched: cjkFontCandidates(), attempts, env: "OFFICEGEN_PDF_FONT", preview: text.slice(0, 120) }
    );
  }
}

export function ensurePdfTextEncodable(value: string, font: PDFFont, context: string): string {
  try {
    font.encodeText(value);
    return value;
  } catch {
    throw new OfficegenError(
      "RENDER_FONT_UNSUPPORTED",
      "PDF text contains glyphs that cannot be encoded by the active PDF font.",
      { context, preview: value.slice(0, 80) }
    );
  }
}

function needsUnicodeFont(value: string): boolean {
  return /[^\u0009\u000a\u000d\u0020-\u007e]/.test(value);
}

function findCjkFontPaths(): string[] {
  const env = process.env.OFFICEGEN_PDF_FONT;
  return [...(env ? [env] : []), ...cjkFontCandidates()].filter((candidate, index, list) => existsSync(candidate) && list.indexOf(candidate) === index);
}

function cjkFontCandidates(): string[] {
  const bundled = bundledJapaneseFontCandidates();
  if (process.platform === "win32") {
    const win = process.env.WINDIR ?? "C:\\Windows";
    return [
      ...bundled,
      `${win}\\Fonts\\YuGothR.ttc`,
      `${win}\\Fonts\\YuGothM.ttc`,
      `${win}\\Fonts\\meiryo.ttc`,
      `${win}\\Fonts\\msgothic.ttc`,
      `${win}\\Fonts\\msyh.ttc`
    ];
  }
  if (process.platform === "darwin") {
    return [
      ...bundled,
      "/System/Library/Fonts/ヒラギノ角ゴシック W3.ttc",
      "/System/Library/Fonts/ヒラギノ角ゴシック W6.ttc",
      "/System/Library/Fonts/AppleSDGothicNeo.ttc",
      "/System/Library/Fonts/PingFang.ttc"
    ];
  }
  return [
    ...bundled,
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    "/usr/share/fonts/opentype/noto/NotoSansCJKjp-Regular.otf",
    "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
    "/usr/share/fonts/truetype/noto/NotoSansJP-Regular.otf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
  ];
}

function bundledJapaneseFontCandidates(): string[] {
  try {
    const require = createRequire(import.meta.url);
    const entry = require.resolve("@embedpdf/fonts-jp");
    const root = dirname(dirname(entry));
    return [
      join(root, "fonts", "NotoSansJP-Regular.otf"),
      join(root, "fonts", "NotoSansJP-Bold.otf")
    ];
  } catch {
    return [];
  }
}
