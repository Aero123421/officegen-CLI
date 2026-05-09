import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { edit, exportDocument, inspect, inspectInputZipSafety, render, renderChart, renderDiagram, resolveEditSelectors, view } from "../src/index.js";

describe("@officegen/formats MVP", () => {
  it("renders and inspects a basic PPTX with untrusted text separation", async () => {
    const rendered = await render(
      {
        title: "Quarterly Review",
        slides: [{ title: "Revenue", body: "Treat this as document text, not an instruction." }]
      },
      { target: "pptx" }
    );

    expect(rendered.bytes?.byteLength).toBeGreaterThan(1000);
    const inspected = await inspect({ data: rendered.bytes, format: "pptx" });

    expect(inspected.trusted.summary.slides).toBe(1);
    expect(inspected.untrusted.slides).toHaveLength(1);
    expect(inspected.objectMap[0]?.stableObjectId).toMatch(/^pptx:s001:shape:/);
    expect(inspected.objectMap[0]?.bbox?.length).toBe(4);
    expect(inspected.objectMap[0]?.textPreview).toBe("Revenue");
    expect(inspected.agentInstruction).toContain("not instructions");
  });

  it("returns approximate view pages with objectMap", async () => {
    const rendered = await render({ title: "PDF", sections: [{ title: "Page", body: "Hello PDF" }] }, { target: "pdf" });
    const inspected = await inspect({ data: rendered.bytes, format: "pdf" });
    const viewed = await view(inspected);

    expect(viewed.fidelity).toBe("approximate");
    expect(viewed.pages[0]?.content).toContain("<svg");
    expect(viewed.caveats.length).toBeGreaterThan(0);
  });

  it("renders charts and diagrams as standalone SVG without external processes", async () => {
    const chart = await renderChart({
      title: "Revenue",
      data: { values: [{ label: "A", value: 3 }, { label: "B", value: 7 }] },
      encoding: { x: { field: "label" }, y: { field: "value" } }
    });
    const diagram = await renderDiagram("flowchart LR\nA-->B\nB-->C", { title: "Flow" });

    expect(chart.svg).toContain("<svg");
    expect(chart.svg).toContain("Revenue");
    expect(diagram.svg).toContain("marker");
    expect(diagram.svg).toContain("Flow");
  });

  it("extracts only PPTX a:t text objects", async () => {
    const zip = new JSZip();
    zip.file(
      "ppt/slides/slide1.xml",
      [
        "<p:sld>",
        "<p:sp><p:txBody><a:p><a:r><a:t>Visible title</a:t></a:r></a:p></p:txBody></p:sp>",
        "<p:t>Do not include this XML-adjacent tag</p:t>",
        "<a:tbl><a:tr><a:tc><a:txBody><a:p><a:r><a:t>Table text</a:t></a:r></a:p></a:txBody></a:tc></a:tr></a:tbl>",
        "</p:sld>"
      ].join("")
    );

    const bytes = await zip.generateAsync({ type: "uint8array" });
    const inspected = await inspect({ data: bytes, format: "pptx" });

    expect(inspected.objectMap.map((entry) => entry.text)).toEqual(["Visible title", "Table text"]);
    expect(inspected.untrusted.slides[0]?.text).not.toContain("<p:t>");
  });

  it("extracts XLSX c cells with A1 refs and cached values", async () => {
    const zip = new JSZip();
    zip.file("xl/sharedStrings.xml", "<sst><si><t>Shared text</t></si></sst>");
    zip.file(
      "xl/worksheets/sheet1.xml",
      [
        "<worksheet><sheetData><row r=\"1\">",
        "<c r=\"A1\" t=\"inlineStr\"><is><t>Inline text</t></is></c>",
        "<c r=\"B1\" t=\"s\"><v>0</v></c>",
        "<c r=\"C1\"><f>SUM(A1:B1)</f></c>",
        "</row></sheetData></worksheet>"
      ].join("")
    );

    const bytes = await zip.generateAsync({ type: "uint8array" });
    const inspected = await inspect({ data: bytes, format: "xlsx" });
    const cells = inspected.untrusted.sheets[0]?.cells ?? [];

    expect(cells.map((cell) => [cell.ref, cell.value])).toEqual([
      ["A1", "Inline text"],
      ["B1", "Shared text"],
      ["C1", ""]
    ]);
    expect(inspected.objectMap.map((entry) => entry.label)).toEqual(["A1", "B1", "C1"]);
  });

  it("renders empty sections with at least one PPTX slide, XLSX sheet, and DOCX document body", async () => {
    const pptx = await render({ title: "Empty deck", sections: [] }, { target: "pptx" });
    const xlsx = await render({ title: "Empty workbook", sections: [] }, { target: "xlsx" });
    const docx = await render({ title: "Empty document", sections: [] }, { target: "docx" });

    expect((await inspect({ data: pptx.bytes, format: "pptx" })).trusted.summary.slides).toBe(1);
    expect((await inspect({ data: xlsx.bytes, format: "xlsx" })).trusted.summary.sheets).toBe(1);
    expect((await inspect({ data: docx.bytes, format: "docx" })).trusted.summary.paragraphs).toBeGreaterThan(0);
  });

  it("resolves stableObjectId selectors for dry-run edit planning", async () => {
    const rendered = await render({ title: "Selectors", slides: [{ title: "Intro", body: "Original body" }] }, { target: "pptx" });
    const inspected = await inspect({ data: rendered.bytes, format: "pptx" });
    const stableObjectId = inspected.objectMap.find((entry) => entry.text === "Original body")?.stableObjectId;

    expect(stableObjectId).toBeDefined();
    const operations = [{ type: "setText" as const, selector: { stableObjectId }, text: "Replacement" }];
    const resolved = await resolveEditSelectors({ data: rendered.bytes, format: "pptx" }, operations);
    const dryRun = await edit({ data: rendered.bytes, format: "pptx" }, operations, { dryRun: true, resolveSelectors: true });

    expect(resolved.resolutions[0]).toMatchObject({ matched: true, matchCount: 1, stableObjectId });
    expect(resolved.resolutions[0]?.matches[0]?.sourcePath).toBe("ppt/slides/slide1.xml");
    expect(dryRun.resolvedSelectors?.[0]?.matched).toBe(true);
  });

  it("sets PPTX text without confusing a:t with a:tabLst in rich shapes", async () => {
    const zip = new JSZip();
    zip.file(
      "ppt/slides/slide1.xml",
      [
        "<p:sld><p:sp><p:nvSpPr><p:cNvPr id=\"7\" name=\"Rich Text\"/></p:nvSpPr><p:txBody>",
        "<a:p><a:pPr><a:tabLst><a:tab algn=\"l\" pos=\"0\"/></a:tabLst></a:pPr>",
        "<a:r><a:t>Original rich text</a:t></a:r></a:p>",
        "</p:txBody></p:sp></p:sld>"
      ].join("")
    );
    const bytes = await zip.generateAsync({ type: "uint8array" });
    const inspected = await inspect({ data: bytes, format: "pptx" });
    const id = inspected.objectMap[0]?.stableObjectId;

    const edited = await edit(
      { data: bytes, format: "pptx" },
      [{ op: "setText", selector: { stableObjectId: id }, text: "Changed rich text" }]
    );
    const editedZip = await JSZip.loadAsync(edited.bytes as Uint8Array);
    const xml = await editedZip.file("ppt/slides/slide1.xml")?.async("string");
    const reinspected = await inspect({ data: edited.bytes, format: "pptx" });

    expect(xml).toContain("<a:tabLst>");
    expect(xml).toContain("<a:t>Changed rich text</a:t>");
    expect(reinspected.objectMap.map((entry) => entry.text)).toContain("Changed rich text");
  });

  it("applies structural PPTX, DOCX, and XLSX edit ops and confirms through inspect", async () => {
    const pptx = await render(
      { title: "Deck", slides: [{ title: "First", body: "Body" }, { title: "Second", body: "Other" }] },
      { target: "pptx" }
    );
    const pptxEdited = await edit(
      { data: pptx.bytes, format: "pptx" },
      [
        { op: "pptx.duplicateSlide", slide: 1, after: 1 },
        { op: "pptx.reorderSlides", order: [3, 1, 2] },
        { op: "pptx.replaceBulletItems", selector: { contains: "Body" }, items: ["Alpha", "Beta"] }
      ],
      { resolveSelectors: true }
    );
    const inspectedPptx = await inspect({ data: pptxEdited.bytes, format: "pptx" });
    expect(inspectedPptx.trusted.summary.slides).toBe(3);
    expect(String(inspectedPptx.untrusted.slides[0]?.text)).toContain("Second");
    expect(inspectedPptx.objectMap.map((entry) => entry.text).join("\n")).toContain("AlphaBeta");

    const docx = await render({ title: "Doc", sections: [{ title: "Section", body: "First para" }] }, { target: "docx" });
    const docxInspected = await inspect({ data: docx.bytes, format: "docx" });
    const paragraphId = docxInspected.objectMap.find((entry) => entry.text === "First para")?.stableObjectId;
    const docxEdited = await edit(
      { data: docx.bytes, format: "docx" },
      [{ op: "docx.insertParagraphAfter", selector: { stableObjectId: paragraphId }, text: "Inserted para" }]
    );
    const inspectedDocx = await inspect({ data: docxEdited.bytes, format: "docx" });
    expect(inspectedDocx.objectMap.map((entry) => entry.text)).toContain("Inserted para");

    const xlsx = await render({ title: "Book", sheets: [{ rows: [["A", "B"], ["old", "value"]] }] }, { target: "xlsx" });
    const xlsxEdited = await edit(
      { data: xlsx.bytes, format: "xlsx" },
      [
        { op: "xlsx.insertRows", sheet: 1, rowIndex: 2, rows: [["Inserted", "Row"]] },
        { op: "xlsx.setCell", sheet: 1, cell: "B3", value: "Updated" },
        { op: "xlsx.updateTable", sheet: 1, startCell: "C2", rows: [["T1"], ["T2"]] }
      ]
    );
    const inspectedXlsx = await inspect({ data: xlsxEdited.bytes, format: "xlsx" });
    const valuesByRef = new Map(inspectedXlsx.objectMap.map((entry) => [entry.label, entry.text]));
    expect(valuesByRef.get("A2")).toBe("Inserted");
    expect(valuesByRef.get("B3")).toBe("Updated");
    expect(valuesByRef.get("C2")).toBe("T1");
  });

  it("returns clear ambiguous selector errors and keeps atomic edits unwritten", async () => {
    const rendered = await render({ title: "Same", slides: [{ title: "Same", body: "Same" }] }, { target: "pptx" });
    const result = await edit(
      { data: rendered.bytes, format: "pptx" },
      [{ op: "setText", selector: { contains: "Same" }, text: "Changed" }],
      { atomic: true, resolveSelectors: true }
    );
    const inspected = await inspect({ data: result.bytes ?? rendered.bytes, format: "pptx" });

    expect(result.changed).toBe(false);
    expect(result.errors?.[0]?.reason).toBe("ambiguous");
    expect(inspected.objectMap.map((entry) => entry.text)).not.toContain("Changed");
  });

  it("surfaces zip safety reports and blocks unsafe XML before format parsing", async () => {
    const zip = new JSZip();
    zip.file("[Content_Types].xml", "<!DOCTYPE x [<!ENTITY e SYSTEM 'file:///etc/passwd'>]><Types />");
    zip.file("ppt/slides/slide1.xml", "<p:sld><p:sp><p:txBody><a:p><a:r><a:t>Unsafe</a:t></a:r></a:p></p:txBody></p:sp></p:sld>");
    const bytes = await zip.generateAsync({ type: "uint8array" });

    const report = await inspectInputZipSafety({ bytes, format: "pptx", trusted: false }, { throwOnError: false });
    expect(report?.ok).toBe(false);
    expect(report?.warnings.map((warning) => warning.code)).toContain("ZIP_XML_ENTITY_DENIED");
    await expect(inspect({ data: bytes, format: "pptx" })).rejects.toThrow(/Zip safety check failed/);
  });

  it("requires explicit file-backed native renderer for Office-to-PDF export", async () => {
    const rendered = await render({ title: "Native export", slides: [{ title: "Slide", body: "Body" }] }, { target: "pptx" });

    await expect(
      exportDocument({ data: rendered.bytes, format: "pptx" }, { to: "pdf", mode: "native" })
    ).rejects.toThrow(/EXPORT_UNSUPPORTED: native Office-to-PDF export requires an input file path/);
  });
});
