import { type PDFDocument, type PDFFont } from "pdf-lib";
export interface PdfFontSet {
    font: PDFFont;
    bold: PDFFont;
    embeddedUnicode: boolean;
    fontPath?: string;
    caveats: string[];
}
export declare function embedPdfFonts(pdf: PDFDocument, textSamples: string[]): Promise<PdfFontSet>;
export declare function ensurePdfTextEncodable(value: string, font: PDFFont, context: string): string;
