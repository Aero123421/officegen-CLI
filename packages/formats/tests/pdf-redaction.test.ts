import { describe, expect, it } from "vitest";
import { PDFDocument, rgb } from "pdf-lib";
import { edit, inspect, scanPdfForForbiddenText } from "../src/index.js";

describe("PDF object graph and redaction safety", () => {
  it("inspects PDF text, metadata, annotations, and risk flags", async () => {
    const pdf = await PDFDocument.create();
    pdf.setTitle("Inspection Fixture");
    pdf.setAuthor("Officegen Test");
    const page = pdf.addPage([300, 200]);
    page.drawText("Visible account SECRET-123", { x: 32, y: 150, size: 12 });
    page.drawRectangle({ x: 28, y: 126, width: 140, height: 18, color: rgb(1, 0.9, 0.7), borderColor: rgb(1, 0.6, 0.1), borderWidth: 1 });
    page.drawText("Review note", { x: 32, y: 132, size: 10 });
    const bytes = await pdf.save({ useObjectStreams: false });

    const inspected = await inspect({ data: bytes, format: "pdf" }, { depth: "full" });
    const graph = inspected.untrusted.pdfGraph as {
      pageCount: number;
      textBlocks: Array<{ text: string }>;
      metadata: Record<string, unknown>;
      scan: { filters: Array<{ name: string }>; embeddedFiles: number };
      riskFlags: Array<{ code: string }>;
    };

    expect(inspected.trusted.summary.pages).toBe(1);
    expect(graph.pageCount).toBe(1);
    expect(graph.textBlocks.map((block) => block.text).join(" ")).toContain("SECRET-123");
    expect(JSON.stringify(graph.metadata)).toContain("Inspection Fixture");
    expect(graph.scan.filters.some((filter) => filter.name === "FlateDecode")).toBe(true);
    expect(graph.scan.embeddedFiles).toBe(0);
    expect(graph.riskFlags.map((flag) => flag.code)).toContain("PDF_EXTRACTABLE_TEXT_PRESENT");
    expect(inspected.objectMap.some((entry) => entry.kind === "pdfText" && entry.text?.includes("SECRET-123"))).toBe(true);
  });

  it("blocks pdf.redact atomically instead of treating an overlay as redaction", async () => {
    const pdf = await PDFDocument.create();
    const page = pdf.addPage([300, 200]);
    page.drawText("Classified SECRET-456", { x: 32, y: 150, size: 12 });
    const bytes = await pdf.save({ useObjectStreams: false });

    const edited = await edit(
      { data: bytes, format: "pdf" },
      [
        { op: "pdf.textOverlay", page: 1, text: "REDACTED", x: 32, y: 150 },
        { op: "pdf.redact", page: 1, text: "SECRET-456" }
      ]
    );

    expect(edited.changed).toBe(false);
    expect(edited.applied).toBe(0);
    expect(edited.skipped).toBe(2);
    expect(edited.bytes).toBeUndefined();
    expect(edited.errors?.[0]).toMatchObject({ op: "pdf.redact", reason: "unsupported" });
    expect(edited.opResults?.[0]).toMatchObject({ op: "pdf.textOverlay", applied: false, reason: "skipped-after-error" });
  });

  it("finds forbidden text after additive PDF overlays", async () => {
    const pdf = await PDFDocument.create();
    const page = pdf.addPage([300, 200]);
    page.drawText("Do not leak SECRET-789", { x: 32, y: 150, size: 12 });
    const bytes = await pdf.save({ useObjectStreams: false });

    const edited = await edit(
      { data: bytes, format: "pdf" },
      [{ op: "pdf.textOverlay", page: 1, text: "REDACTED", x: 32, y: 150 }]
    );
    const scan = scanPdfForForbiddenText(edited.bytes as Uint8Array, ["SECRET-789"]);

    expect(edited.changed).toBe(true);
    expect(scan.found).toEqual([
      expect.objectContaining({ pattern: "SECRET-789" })
    ]);
  });
});
