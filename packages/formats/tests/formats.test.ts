import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { PDFDocument } from "pdf-lib";
import { getBuiltinConfig } from "@officegen/core";
import { DEFAULT_NATIVE_RENDERER_TIMEOUT_MS, MIN_NATIVE_RENDERER_TIMEOUT_MS, diffDocuments, diagnose, edit, exportDocument, extractAssets, inspect, inspectInputZipSafety, render, renderChart, renderDiagram, replaceAsset, resolveEditSelectors, resolveNativeRendererTimeoutMs, verify, view } from "../src/index.js";

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
    expect(inspected.objectMap[0]?.stableObjectId).toMatch(/^pptx:slide-[a-f0-9]{8}:shape:/);
    expect(inspected.objectMap[0]?.bbox?.length).toBe(4);
    expect(inspected.objectMap[0]?.textPreview).toBe("Revenue");
    expect(inspected.agentInstruction).toContain("not instructions");
  });

  it("renders rich PPTX blocks as separate text and table objects", async () => {
    const rendered = await render(
      {
        title: "Board KPI",
        targets: ["pptx"],
        sections: [{
          title: "Executive Summary",
          blocks: [
            { type: "heading", text: "Highlights" },
            { type: "list", items: ["Revenue up", "Margin stable"] },
            { type: "table", rows: [{ metric: "Revenue", value: "$10M" }, { metric: "Margin", value: "42%" }] },
            { type: "callout", text: "Board-reviewed summary" }
          ]
        }]
      },
      { target: "pptx" }
    );
    const inspected = await inspect({ data: rendered.bytes, format: "pptx" });

    expect(inspected.objectMap.map((entry) => entry.text).join("\n")).toContain("Highlights");
    const tableCell = inspected.objectMap.find((entry) => entry.kind === "tableCell" && entry.text === "Revenue");
    const tableCellId = tableCell?.stableObjectId ?? "";
    expect(tableCell?.bbox?.length).toBe(4);
    expect(tableCell?.bounds?.width).toBeGreaterThan(0);
    const viewed = await view(inspected);
    expect(viewed.objectMap.find((entry) => entry.stableObjectId === tableCellId)?.bbox?.length).toBe(4);
    expect(viewed.pages[0]?.content).toContain(tableCellId);
    expect(rendered.caveats[0]).toContain("tables");
  });

  it("renders native editable PPTX charts and exposes them in the object map", async () => {
    const rendered = await render(
      {
        title: "Chart Deck",
        targets: ["pptx"],
        sections: [{
          title: "Revenue",
          blocks: [
            { type: "chart", title: "Revenue", chartType: "bar", categories: ["Q1", "Q2"], values: [10, 15] }
          ]
        }]
      },
      { target: "pptx" }
    );
    const inspected = await inspect({ data: rendered.bytes, format: "pptx" }, { depth: "full" });

    expect(inspected.trusted.summary.charts).toBeGreaterThan(0);
    expect(inspected.objectMap.some((entry) => entry.kind === "chart" && entry.editableOps?.includes("pptx.updateChartData"))).toBe(true);
    const viewed = await view({ data: rendered.bytes, format: "pptx" });
    expect(viewed.objectMap.some((entry) => entry.kind === "chart" && entry.editableOps?.includes("pptx.updateChartData"))).toBe(true);
    expect(viewed.pages[0]?.content).toContain('data-kind="chart"');
    expect(rendered.caveats[0]).toContain("Office charts");
  });

  it("updates PPTX chart caches and the embedded chart workbook together", async () => {
    const rendered = await render(
      {
        title: "Chart Deck",
        targets: ["pptx"],
        sections: [{ title: "Revenue", blocks: [{ type: "chart", title: "Revenue", categories: ["Q1", "Q2"], values: [10, 15] }] }]
      },
      { target: "pptx" }
    );
    const inspected = await inspect({ data: rendered.bytes, format: "pptx" }, { depth: "full" });
    const chartId = inspected.objectMap.find((entry) => entry.kind === "chart")?.stableObjectId;
    const edited = await edit(
      { data: rendered.bytes, format: "pptx" },
      [{ op: "pptx.updateChartData", selector: { stableObjectId: chartId }, seriesName: "Bookings", categories: ["H1", "H2", "H3"], values: [11, 22, 33] }]
    );
    const editedZip = await JSZip.loadAsync(edited.bytes as Uint8Array);
    const chartPath = Object.keys(editedZip.files).find((path) => /^ppt\/charts\/chart\d+\.xml$/i.test(path));
    const chartXml = chartPath ? await editedZip.file(chartPath)?.async("string") : undefined;
    const embeddedPath = Object.keys(editedZip.files).find((path) => /^ppt\/embeddings\/.*\.xlsx$/i.test(path));
    const embedded = embeddedPath ? await editedZip.file(embeddedPath)?.async("uint8array") : undefined;
    expect(embedded).toBeDefined();
    const workbookZip = await JSZip.loadAsync(embedded as Uint8Array);
    const sheetXml = await workbookZip.file("xl/worksheets/sheet1.xml")?.async("string");

    expect(chartXml).toContain("Bookings");
    expect(chartXml).toContain("Sheet1!$A$2:$A$4");
    expect(chartXml).toContain("<c:v>H3</c:v>");
    expect(chartXml).toContain("<c:v>33</c:v>");
    expect(sheetXml).toContain("<t>Bookings</t>");
    expect(sheetXml).toContain("<t>H3</t>");
    expect(sheetXml).toContain('<c r="B4"><v>33</v></c>');
  });

  it("replaces PPTX images by shape selector and writes crop metadata", async () => {
    const zip = new JSZip();
    zip.file("ppt/slides/slide1.xml", [
      "<p:sld><p:pic><p:nvPicPr><p:cNvPr id=\"10\" name=\"Logo\"/></p:nvPicPr>",
      "<p:blipFill><a:blip r:embed=\"rId1\"/></p:blipFill>",
      "<p:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"914400\" cy=\"914400\"/></a:xfrm></p:spPr>",
      "</p:pic></p:sld>"
    ].join(""));
    zip.file("ppt/slides/_rels/slide1.xml.rels", "<Relationships><Relationship Id=\"rId1\" Target=\"../media/image1.png\"/></Relationships>");
    zip.file("ppt/media/image1.png", pngBytes(1, 1));
    const bytes = await zip.generateAsync({ type: "uint8array" });
    const inspected = await inspect({ data: bytes, format: "pptx" });
    const pictureId = inspected.objectMap.find((entry) => entry.kind === "picture")?.stableObjectId;

    const edited = await edit(
      { data: bytes, format: "pptx" },
      [{ op: "pptx.replaceImageByShape", selector: { stableObjectId: pictureId }, replacementBase64: Buffer.from(pngBytes(2, 1)).toString("base64"), fit: "cover" }]
    );
    const editedZip = await JSZip.loadAsync(edited.bytes as Uint8Array);
    const slideXml = await editedZip.file("ppt/slides/slide1.xml")?.async("string");
    const imageBytes = await editedZip.file("ppt/media/image1.png")?.async("uint8array");

    expect(edited.changed).toBe(true);
    expect(slideXml).toContain("<a:srcRect");
    expect(imageBytes?.byteLength).toBe(pngBytes(2, 1).byteLength);
  });

  it("returns approximate view pages with objectMap", async () => {
    const rendered = await render({ title: "PDF", sections: [{ title: "Page", body: "Hello PDF" }] }, { target: "pdf" });
    const inspected = await inspect({ data: rendered.bytes, format: "pdf" });
    const viewed = await view(inspected);

    expect(viewed.fidelity).toBe("approximate");
    expect(viewed.pages[0]?.content).toContain("<svg");
    expect(viewed.caveats.length).toBeGreaterThan(0);
  });

  it("rejects unsupported render target inference instead of falling back to PDF", async () => {
    await expect(render({ title: "Web", targets: ["html"] })).rejects.toMatchObject({
      payload: { code: "EXPORT_UNSUPPORTED" }
    });
    await expect(render({ title: "Vector", kind: "svg" })).rejects.toMatchObject({
      payload: { code: "EXPORT_UNSUPPORTED" }
    });
  });

  it("rejects render target and output extension mismatches", async () => {
    await expect(render({ title: "Mismatch" }, { target: "pdf", out: "mismatch.pptx" })).rejects.toMatchObject({
      payload: { code: "TARGET_EXTENSION_MISMATCH" }
    });
  });

  it("renders Japanese/CJK PDF text with the bundled fallback font", async () => {
    const rendered = await render({ title: "\u65e5\u672c\u8a9e", sections: [{ title: "\u6982\u8981", body: "\u58f2\u4e0a\u306f\u9806\u8abf\u3067\u3059" }] }, { target: "pdf" });
    const inspected = await inspect({ data: rendered.bytes, format: "pdf" });

    expect(rendered.bytes?.byteLength).toBeGreaterThan(1000);
    expect(rendered.caveats.join(" ")).toContain("Embedded Unicode PDF font");
    expect(inspected.trusted.summary.pages).toBe(1);
  });

  it("renders tabular ASCII PDF text by normalizing tabs to spaces", async () => {
    const rendered = await render({ title: "Table PDF", sections: [{ title: "Rows", body: "Company\tSignal\tRisk" }] }, { target: "pdf" });
    const inspected = await inspect({ data: rendered.bytes, format: "pdf" }, { depth: "full" });

    expect(rendered.bytes?.byteLength).toBeGreaterThan(500);
    expect(inspected.trusted.summary.pages).toBe(1);
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

  it("sets XLSX cells with numeric, boolean, and null scalar values", async () => {
    const rendered = await render({ title: "Scalars", sheets: [{ rows: [["A", "B", "C"]] }] }, { target: "xlsx" });
    const edited = await edit(
      { data: rendered.bytes, format: "xlsx" },
      [
        { op: "xlsx.setCell", sheet: 1, cell: "A2", value: 123.45 },
        { op: "xlsx.setCell", sheet: 1, cell: "B2", value: true },
        { op: "xlsx.setCell", sheet: 1, cell: "C2", value: null }
      ]
    );
    const editedZip = await JSZip.loadAsync(edited.bytes as Uint8Array);
    const sheetXml = await editedZip.file("xl/worksheets/sheet1.xml")?.async("string");
    const inspected = await inspect({ data: edited.bytes, format: "xlsx" });
    const valuesByRef = new Map(inspected.objectMap.map((entry) => [entry.label, entry.text]));

    expect(sheetXml).toContain('<c r="A2"><v>123.45</v></c>');
    expect(sheetXml).toContain('<c r="B2" t="b"><v>1</v></c>');
    expect(sheetXml).toContain('<c r="C2"/>');
    expect(valuesByRef.get("A2")).toBe("123.45");
    expect(valuesByRef.get("B2")).toBe("TRUE");
    expect(valuesByRef.get("C2")).toBe("");
  });

  it("preserves XLSX styled cell attributes when setting a scalar value", async () => {
    const rendered = await render({ title: "Styled", sheets: [{ rows: [["Old"]] }] }, { target: "xlsx" });
    const zip = await JSZip.loadAsync(rendered.bytes as Uint8Array);
    const sheetXml = await zip.file("xl/worksheets/sheet1.xml")?.async("string");
    zip.file("xl/worksheets/sheet1.xml", String(sheetXml).replace('<c r="A1"', '<c r="A1" s="7"'));
    const styledBytes = await zip.generateAsync({ type: "uint8array" });

    const edited = await edit(
      { data: styledBytes, format: "xlsx" },
      [{ op: "xlsx.setCell", sheet: 1, cell: "A1", value: "New" }]
    );
    const editedZip = await JSZip.loadAsync(edited.bytes as Uint8Array);
    const editedSheetXml = await editedZip.file("xl/worksheets/sheet1.xml")?.async("string");

    expect(editedSheetXml).toContain('<c r="A1" t="inlineStr" s="7"><is><t>New</t></is></c>');
  });

  it("sets XLSX formulas and exposes them as valid Office XML", async () => {
    const rendered = await render({ title: "Formula", sheets: [{ rows: [["A", "B"], [2, 3]] }] }, { target: "xlsx" });
    const edited = await edit(
      { data: rendered.bytes, format: "xlsx" },
      [{ op: "xlsx.setFormula", sheet: 1, cell: "C2", formula: "=SUM(A2:B2)" }]
    );
    const zip = await JSZip.loadAsync(edited.bytes as Uint8Array);
    const sheetXml = await zip.file("xl/worksheets/sheet1.xml")?.async("string");

    expect(sheetXml).toContain('<c r="C2"><f>SUM(A2:B2)</f></c>');
    expect(sheetXml?.match(/<row\b[^>]*\br="2"/g)).toHaveLength(1);
  });

  it("updates shifted XLSX formula references when inserting rows", async () => {
    const rendered = await render({ title: "Formula", sheets: [{ rows: [["A", "B", "Total"], [2, 3, ""]] }] }, { target: "xlsx" });
    const withFormula = await edit(
      { data: rendered.bytes, format: "xlsx" },
      [{ op: "xlsx.setFormula", sheet: 1, cell: "C2", formula: "=SUM(A2:B2)" }]
    );
    const edited = await edit(
      { data: withFormula.bytes, format: "xlsx" },
      [{ op: "xlsx.insertRows", sheet: 1, rowIndex: 2, rows: [["Inserted", 1, 2]] }]
    );
    const zip = await JSZip.loadAsync(edited.bytes as Uint8Array);
    const sheetXml = await zip.file("xl/worksheets/sheet1.xml")?.async("string");

    expect(sheetXml).toContain('<c r="C3"><f>SUM(A3:B3)</f></c>');
  });

  it("keeps XLSX chart objects in the approximate view object map", async () => {
    const zip = new JSZip();
    zip.file("xl/worksheets/sheet1.xml", '<worksheet><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Revenue</t></is></c></row></sheetData></worksheet>');
    zip.file("xl/charts/chart1.xml", "<c:chartSpace><c:chart/></c:chartSpace>");
    const inspected = await inspect({ data: await zip.generateAsync({ type: "uint8array" }), format: "xlsx" });
    const viewed = await view(inspected);

    expect(viewed.objectMap.some((entry) => entry.kind === "chart" && entry.selectorHints?.chartPath === "xl/charts/chart1.xml")).toBe(true);
  });

  it("updates XLSX chart caches and backing worksheet cells together", async () => {
    const zip = new JSZip();
    zip.file("xl/worksheets/sheet1.xml", [
      '<worksheet><sheetData>',
      '<row r="1"><c r="A1" t="inlineStr"><is><t>Quarter</t></is></c><c r="B1" t="inlineStr"><is><t>Old</t></is></c></row>',
      '<row r="2"><c r="A2" t="inlineStr"><is><t>Q1</t></is></c><c r="B2"><v>10</v></c></row>',
      '<row r="3"><c r="A3" t="inlineStr"><is><t>Q2</t></is></c><c r="B3"><v>20</v></c></row>',
      '</sheetData></worksheet>'
    ].join(""));
    zip.file("xl/charts/chart1.xml", [
      '<c:chartSpace><c:chart><c:plotArea><c:barChart><c:ser>',
      '<c:tx><c:v>Old</c:v></c:tx>',
      '<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$3</c:f><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt></c:strCache></c:strRef></c:cat>',
      '<c:val><c:numRef><c:f>Sheet1!$B$2:$B$3</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt></c:numCache></c:numRef></c:val>',
      '</c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>'
    ].join(""));
    const bytes = await zip.generateAsync({ type: "uint8array" });
    const inspected = await inspect({ data: bytes, format: "xlsx" });
    const chartId = inspected.objectMap.find((entry) => entry.kind === "chart")?.stableObjectId;
    const edited = await edit(
      { data: bytes, format: "xlsx" },
      [{ op: "xlsx.chart.setData", selector: { stableObjectId: chartId }, seriesName: "New", categories: ["N1", "N2"], values: [111, 222] }]
    );
    const editedZip = await JSZip.loadAsync(edited.bytes as Uint8Array);
    const sheetXml = await editedZip.file("xl/worksheets/sheet1.xml")?.async("string");
    const chartXml = await editedZip.file("xl/charts/chart1.xml")?.async("string");

    expect(sheetXml).toContain("<t>New</t>");
    expect(sheetXml).toContain("<t>N1</t>");
    expect(sheetXml).toContain("<v>222</v>");
    expect(chartXml).toContain("<c:v>New</c:v>");
    expect(chartXml).toContain("<c:v>N2</c:v>");
    expect(chartXml).toContain("<c:v>222</c:v>");
  });

  it("renders native Excel table objects and inspects workbook objects compactly", async () => {
    const rendered = await render(
      { title: "Book", targets: ["xlsx"], sheets: [{ name: "Data", tableName: "RevenueTable", rows: [{ quarter: "Q1", revenue: 10 }, { quarter: "Q2", revenue: 15 }] }] },
      { target: "xlsx" }
    );
    const inspected = await inspect({ data: rendered.bytes, format: "xlsx" }, { depth: "summary" });

    expect(inspected.trusted.summary.tables).toBeGreaterThan(0);
    expect(inspected.untrusted.workbookObjects?.tables.length).toBeGreaterThan(0);
    expect(inspected.objectMap.some((entry) => entry.kind === "table")).toBe(true);
  });

  it("writes and updates XLSX table parts, relationships, and worksheet table references", async () => {
    const rendered = await render({ title: "Book", sheets: [{ name: "Data", rows: [["A", "B"], ["old", "value"]] }] }, { target: "xlsx" });
    const edited = await edit(
      { data: rendered.bytes, format: "xlsx" },
      [{ op: "xlsx.writeTable", sheet: 1, startCell: "D4", tableName: "ManualTable", rows: [["Name", "Value"], ["A", 1], ["B", 2]] }]
    );
    const editedZip = await JSZip.loadAsync(edited.bytes as Uint8Array);
    const tableXml = await editedZip.file("xl/tables/table2.xml")?.async("string") ?? await editedZip.file("xl/tables/table1.xml")?.async("string");
    const worksheetXml = await editedZip.file("xl/worksheets/sheet1.xml")?.async("string");
    const relsXml = await editedZip.file("xl/worksheets/_rels/sheet1.xml.rels")?.async("string");
    const contentTypes = await editedZip.file("[Content_Types].xml")?.async("string");

    expect(tableXml).toContain('displayName="ManualTable"');
    expect(tableXml).toContain('ref="D4:E6"');
    expect(worksheetXml).toContain("<tableParts");
    expect(relsXml).toContain("relationships/table");
    expect(contentTypes).toContain("spreadsheetml.table+xml");
  });

  it("resizes XLSX tables without leaving stale autoFilter ranges", async () => {
    const rendered = await render({ title: "Book", sheets: [{ name: "Data", tableName: "ManualTable", rows: [["Name", "Value"], ["A", 1]] }] }, { target: "xlsx" });
    const inspected = await inspect({ data: rendered.bytes, format: "xlsx" });
    const tableId = inspected.objectMap.find((entry) => entry.kind === "table")?.stableObjectId;
    const edited = await edit(
      { data: rendered.bytes, format: "xlsx" },
      [{ op: "xlsx.table.resize", selector: { stableObjectId: tableId }, ref: "A1:B4" }]
    );
    const zip = await JSZip.loadAsync(edited.bytes as Uint8Array);
    const tableXml = await zip.file("xl/tables/table1.xml")?.async("string");

    expect(tableXml).toContain('ref="A1:B4"');
    expect(tableXml).toContain('<autoFilter ref="A1:B4"');
  });

  it("edits DOCX headers, comments, and tracked insertions without native Office", async () => {
    const rendered = await render({ title: "Doc", sections: [{ title: "Body", body: "Paragraph" }] }, { target: "docx" });
    const inspected = await inspect({ data: rendered.bytes, format: "docx" });
    const paragraphId = inspected.objectMap.find((entry) => entry.text === "Paragraph")?.stableObjectId;
    const edited = await edit(
      { data: rendered.bytes, format: "docx" },
      [
        { op: "docx.setHeader", text: "Confidential" },
        { op: "docx.setFooter", text: "Page footer" },
        { op: "docx.addComment", selector: { stableObjectId: paragraphId }, text: "Review this", author: "QA" },
        { op: "docx.addRedline", selector: { stableObjectId: paragraphId }, text: "Inserted with tracking", author: "QA" }
      ]
    );
    const reinspected = await inspect({ data: edited.bytes, format: "docx" }, { depth: "full" });
    const text = reinspected.objectMap.map((entry) => entry.text).join("\n");

    expect(reinspected.trusted.summary.headers).toBe(1);
    expect(reinspected.trusted.summary.footers).toBe(1);
    expect(reinspected.trusted.summary.comments).toBe(1);
    expect(text).toContain("Confidential");
    expect(text).toContain("Review this");
    expect(text).toContain("Inserted with tracking");
    const zip = await JSZip.loadAsync(edited.bytes as Uint8Array);
    const documentXml = await zip.file("word/document.xml")?.async("string");
    expect(documentXml).toContain("commentRangeStart");
    expect(documentXml).toContain("commentRangeEnd");
  });

  it("assigns unique DOCX redline ids across multiple tracked insertions", async () => {
    const rendered = await render({ title: "Doc", sections: [{ title: "Body", body: "Paragraph" }] }, { target: "docx" });
    const inspected = await inspect({ data: rendered.bytes, format: "docx" });
    const paragraphId = inspected.objectMap.find((entry) => entry.text === "Paragraph")?.stableObjectId;
    const edited = await edit(
      { data: rendered.bytes, format: "docx" },
      [
        { op: "docx.addRedline", selector: { stableObjectId: paragraphId }, text: "First", author: "QA" },
        { op: "docx.addRedline", selector: { stableObjectId: paragraphId }, text: "Second", author: "QA" }
      ]
    );
    const zip = await JSZip.loadAsync(edited.bytes as Uint8Array);
    const documentXml = await zip.file("word/document.xml")?.async("string") ?? "";
    const ids = [...documentXml.matchAll(/<w:ins\b[^>]*\bw:id="(\d+)"/g)].map((match) => match[1]);

    expect(ids.sort()).toEqual(["1", "2"]);
  });

  it("applies DOCX style, header/footer, and comments as first-class structure edits", async () => {
    const rendered = await render({ title: "Doc", sections: [{ title: "Body", body: "Original paragraph" }] }, { target: "docx" });
    const inspected = await inspect({ data: rendered.bytes, format: "docx" });
    const paragraphId = inspected.objectMap.find((entry) => entry.text === "Original paragraph")?.stableObjectId;
    const edited = await edit(
      { data: rendered.bytes, format: "docx" },
      [
        { op: "docx.headerFooter.setText", kind: "header", text: "Review header" },
        { op: "docx.headerFooter.setText", kind: "footer", text: "Review footer" },
        { op: "docx.applyStyle", selector: { stableObjectId: paragraphId }, styleId: "Heading1" },
        { op: "docx.addComment", selector: { stableObjectId: paragraphId }, text: "Range comment", author: "QA" }
      ]
    );
    const zip = await JSZip.loadAsync(edited.bytes as Uint8Array);
    const documentXml = await zip.file("word/document.xml")?.async("string") ?? "";
    const headerXml = await zip.file("word/header1.xml")?.async("string") ?? "";
    const footerXml = await zip.file("word/footer1.xml")?.async("string") ?? "";
    const structured = await inspect({ data: edited.bytes, format: "docx" }, { structure: true, depth: "full" });

    expect(headerXml).toContain("Review header");
    expect(footerXml).toContain("Review footer");
    expect(documentXml).toContain('<w:pStyle w:val="Heading1"/>');
    expect(documentXml).toContain("commentRangeStart");
    expect(structured.objectMap.some((entry) => entry.kind === "header" && entry.editableOps?.includes("docx.headerFooter.setText"))).toBe(true);
    expect((structured.untrusted.structureMap as any).comments).toBe(1);
  });

  it("replaces DOCX paragraphs with paired tracked deletion and insertion", async () => {
    const rendered = await render({ title: "Doc", sections: [{ title: "Body", body: "Original paragraph" }] }, { target: "docx" });
    const inspected = await inspect({ data: rendered.bytes, format: "docx" });
    const paragraphId = inspected.objectMap.find((entry) => entry.text === "Original paragraph")?.stableObjectId;
    const edited = await edit(
      { data: rendered.bytes, format: "docx" },
      [{ op: "docx.redline.replace", selector: { stableObjectId: paragraphId }, text: "Replacement paragraph", author: "QA" }]
    );
    const zip = await JSZip.loadAsync(edited.bytes as Uint8Array);
    const documentXml = await zip.file("word/document.xml")?.async("string") ?? "";
    const structured = await inspect({ data: edited.bytes, format: "docx" }, { structure: true, depth: "full" });

    expect(documentXml).toContain("<w:del");
    expect(documentXml).toContain("<w:ins");
    expect(documentXml).toContain("Replacement paragraph");
    expect((structured.untrusted.structureMap as any).trackedChanges).toBeGreaterThanOrEqual(2);
  });

  it("keeps DOCX stableObjectIds unique across repeated header/footer parts", async () => {
    const zip = new JSZip();
    zip.file("word/document.xml", "<w:document><w:body><w:p><w:r><w:t>Body</w:t></w:r></w:p></w:body></w:document>");
    zip.file("word/header1.xml", "<w:hdr><w:p><w:r><w:t>Header A</w:t></w:r></w:p></w:hdr>");
    zip.file("word/header2.xml", "<w:hdr><w:p><w:r><w:t>Header B</w:t></w:r></w:p></w:hdr>");
    const inspected = await inspect({ data: await zip.generateAsync({ type: "uint8array" }), format: "docx" }, { depth: "full" });
    const ids = inspected.objectMap.map((entry) => entry.stableObjectId);

    expect(new Set(ids).size).toBe(ids.length);
    expect(inspected.objectMap.find((entry) => entry.text === "Header A")?.stableObjectId).not.toBe(inspected.objectMap.find((entry) => entry.text === "Header B")?.stableObjectId);
  });

  it("reports DOCX structure maps for styles, fields, content controls, comments, and redlines", async () => {
    const zip = new JSZip();
    zip.file("word/document.xml", [
      '<w:document><w:body>',
      '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Heading</w:t></w:r></w:p>',
      '<w:p><w:sdt><w:sdtPr><w:tag w:val="client"/></w:sdtPr><w:sdtContent><w:r><w:t>{{client}}</w:t></w:r></w:sdtContent></w:sdt></w:p>',
      '<w:p><w:fldChar w:fldCharType="begin"/><w:instrText>DATE</w:instrText></w:p>',
      '<w:p><w:ins w:id="1"><w:r><w:t>Inserted</w:t></w:r></w:ins></w:p>',
      '<w:sectPr/></w:body></w:document>'
    ].join(""));
    zip.file("word/header1.xml", "<w:hdr><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:hdr>");
    zip.file("word/footer1.xml", "<w:ftr><w:p><w:r><w:t>Footer</w:t></w:r></w:p></w:ftr>");
    zip.file("word/comments.xml", '<w:comments><w:comment w:id="0"><w:p><w:r><w:t>Comment</w:t></w:r></w:p></w:comment></w:comments>');
    zip.file("word/styles.xml", '<w:styles><w:style w:type="paragraph" w:styleId="Heading1"/></w:styles>');
    const inspected = await inspect({ data: await zip.generateAsync({ type: "uint8array" }), format: "docx" }, { structure: true, depth: "full" });
    const structure = inspected.untrusted.structureMap as any;

    expect(structure.headingTree[0].text).toBe("Heading");
    expect(structure.headerFooterVariants.headers).toEqual(["word/header1.xml"]);
    expect(structure.contentControls).toBe(1);
    expect(structure.fields[0].text).toContain("fldChar");
    expect(structure.fillablePlaceholders[0].field).toBe("client");
    expect(structure.comments).toBe(1);
    expect(structure.trackedChanges).toBe(1);
    expect(inspected.objectMap.some((entry) => entry.kind === "style" && entry.label === "Heading1")).toBe(true);
  });

  it("sets XLSX pivot refresh flags and slicer selections through OOXML guards", async () => {
    const zip = new JSZip();
    zip.file("xl/workbook.xml", '<workbook><sheets><sheet name="Data"/></sheets></workbook>');
    zip.file("xl/worksheets/sheet1.xml", '<worksheet><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Region</t></is></c></row></sheetData></worksheet>');
    zip.file("xl/pivotTables/pivotTable1.xml", '<pivotTableDefinition name="Pivot1"></pivotTableDefinition>');
    zip.file("xl/pivotCache/pivotCacheDefinition1.xml", '<pivotCacheDefinition></pivotCacheDefinition>');
    zip.file("xl/slicers/slicer1.xml", '<slicers><slicer><slicerItem n="East"/><slicerItem n="West" h="0"/></slicer></slicers>');
    const bytes = await zip.generateAsync({ type: "uint8array" });
    const inspected = await inspect({ data: bytes, format: "xlsx" }, { depth: "full" });
    const slicerId = inspected.objectMap.find((entry) => entry.kind === "slicer")?.stableObjectId;
    const edited = await edit(
      { data: bytes, format: "xlsx" },
      [
        { op: "xlsx.pivot.refreshAll" },
        { op: "xlsx.slicer.setSelection", selector: { stableObjectId: slicerId }, selected: ["East"] }
      ]
    );
    const editedZip = await JSZip.loadAsync(edited.bytes as Uint8Array);
    const pivotXml = await editedZip.file("xl/pivotTables/pivotTable1.xml")?.async("string") ?? "";
    const cacheXml = await editedZip.file("xl/pivotCache/pivotCacheDefinition1.xml")?.async("string") ?? "";
    const slicerXml = await editedZip.file("xl/slicers/slicer1.xml")?.async("string") ?? "";
    const reinspected = await inspect({ data: edited.bytes, format: "xlsx" }, { depth: "summary", sheet: "Data", range: "A1:A1" });

    expect(inspected.objectMap.find((entry) => entry.kind === "pivotTable")?.editableOps).toContain("xlsx.pivot.refreshAll");
    expect(inspected.objectMap.find((entry) => entry.kind === "slicer")?.editableOps).toContain("xlsx.slicer.setSelection");
    expect(pivotXml).toContain('refreshOnLoad="1"');
    expect(cacheXml).toContain('refreshOnLoad="1"');
    expect(slicerXml).toContain('n="East" h="0"');
    expect(slicerXml).toContain('n="West" h="1"');
    expect((reinspected.untrusted.workbookMap as any).pivotTables).toEqual(["xl/pivotTables/pivotTable1.xml"]);
    expect((reinspected.untrusted.workbookMap as any).slicers).toEqual(["xl/slicers/slicer1.xml"]);
    expect(reinspected.untrusted.sheets[0]?.cellCount).toBe(1);
  });

  it("reports semantic and approximate visual differences", async () => {
    const before = await render({ title: "Before", slides: [{ title: "Slide", body: "Alpha" }] }, { target: "pptx" });
    const after = await edit(
      { data: before.bytes, format: "pptx" },
      [{ op: "setText", selector: { contains: "Alpha" }, text: "Beta" }]
    );
    const diff = await diffDocuments({ data: before.bytes, format: "pptx" }, { data: after.bytes, format: "pptx" }, { visual: true });

    expect(diff.changed).toBe(true);
    expect(diff.summary.changedTextObjects).toBe(1);
    expect(diff.summary.visualRegressionScore).toBeGreaterThan(0);
  });

  it("does not treat different PDF drawing content as a zero visual diff", async () => {
    const before = await render({ title: "PDF", sections: [{ title: "Page", body: "Alpha" }] }, { target: "pdf" });
    const after = await render({ title: "PDF", sections: [{ title: "Page", body: "Beta" }] }, { target: "pdf" });
    const diff = await diffDocuments({ data: before.bytes, format: "pdf" }, { data: after.bytes, format: "pdf" }, { visual: true });

    expect(diff.changed).toBe(true);
    expect(diff.visual?.renderer).toBe("pdf-bytes");
    expect(diff.summary.visualRegressionScore).toBeGreaterThan(0);
  });

  it("updates PPTX chart graphicFrame bounds", async () => {
    const zip = new JSZip();
    zip.file(
      "ppt/slides/slide1.xml",
      [
        '<p:sld><p:cSld>',
        '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Revenue Chart"/></p:nvGraphicFramePr>',
        '<p:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></p:xfrm>',
        '<a:graphic><a:graphicData><c:chart r:id="rId1"/></a:graphicData></a:graphic></p:graphicFrame>',
        '</p:cSld></p:sld>'
      ].join("")
    );
    zip.file("ppt/slides/_rels/slide1.xml.rels", '<Relationships><Relationship Id="rId1" Target="../charts/chart1.xml"/></Relationships>');
    zip.file("ppt/charts/chart1.xml", "<c:chartSpace><c:chart/></c:chartSpace>");
    const bytes = await zip.generateAsync({ type: "uint8array" });
    const inspected = await inspect({ data: bytes, format: "pptx" });
    const chartId = inspected.objectMap.find((entry) => entry.kind === "chart")?.stableObjectId;
    const edited = await edit(
      { data: bytes, format: "pptx" },
      [{ op: "pptx.setBounds", selector: { stableObjectId: chartId }, bounds: { x: 10, y: 20, width: 30, height: 40 } }]
    );
    const editedZip = await JSZip.loadAsync(edited.bytes as Uint8Array);
    const slideXml = await editedZip.file("ppt/slides/slide1.xml")?.async("string");

    expect(slideXml).toContain('<p:xfrm><a:off x="95250" y="190500"/><a:ext cx="285750" cy="381000"/></p:xfrm>');
  });

  it("flags OOXML relationship repair risks without launching Office", async () => {
    const zip = new JSZip();
    zip.file("[Content_Types].xml", "<Types/>");
    zip.file("_rels/.rels", "<Relationships/>");
    zip.file("ppt/presentation.xml", "<p:presentation/>");
    zip.file("ppt/slides/slide1.xml", "<p:sld/>");
    zip.file("ppt/slides/_rels/slide1.xml.rels", "<Relationships><Relationship Id=\"rId1\" Target=\"../media/missing.png\"/></Relationships>");
    const bytes = await zip.generateAsync({ type: "uint8array" });
    const result = await diagnose({ data: bytes, format: "pptx" });

    expect(result.issues.map((issue) => issue.code)).toContain("OFFICE_REPAIR_RISK_BROKEN_RELATIONSHIP");
  });

  it("keeps XLSX summary inspect compact while full inspect keeps cells", async () => {
    const zip = new JSZip();
    const cells = Array.from({ length: 150 }, (_item, index) => `<c r="A${index + 1}" t="inlineStr"><is><t>${"Value ".repeat(index === 0 ? 1000 : 1)}${index + 1}</t></is></c>`).join("");
    zip.file("xl/worksheets/sheet1.xml", `<worksheet><sheetData><row r="1">${cells}</row></sheetData></worksheet>`);
    const bytes = await zip.generateAsync({ type: "uint8array" });

    const summary = await inspect({ data: bytes, format: "xlsx" }, { depth: "summary" });
    const full = await inspect({ data: bytes, format: "xlsx" }, { depth: "full" });

    expect(summary.untrusted.sheets[0]?.cellCount).toBe(150);
    expect(summary.untrusted.sheets[0]?.cells).toBeUndefined();
    expect(summary.untrusted.sheets[0]?.previewCells[0]?.valuePreview.length).toBeLessThanOrEqual(120);
    expect(summary.objectMap.length).toBeLessThanOrEqual(50);
    expect(full.untrusted.sheets[0]?.cells).toHaveLength(150);
  });

  it("rejects asset replacement when bytes do not match the target media type", async () => {
    const zip = new JSZip();
    zip.file("ppt/media/image1.png", new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 1]));
    const bytes = await zip.generateAsync({ type: "uint8array" });
    const svg = Buffer.from('<svg xmlns="http://www.w3.org/2000/svg"></svg>');

    await expect(replaceAsset({ data: bytes, format: "pptx" }, { assetPath: "ppt/media/image1.png", replacement: svg, replacementPath: "logo.svg" }))
      .rejects.toThrow(/ASSET_UNSUPPORTED_FORMAT/);
  });

  it("classifies asset extraction for non-OOXML inputs as unsupported format before zip parsing", async () => {
    const rendered = await render({ title: "PDF", sections: [{ title: "Page", body: "Body" }] }, { target: "pdf" });

    await expect(extractAssets({ data: rendered.bytes, format: "pdf" }))
      .rejects.toMatchObject({
        payload: expect.objectContaining({ code: "UNSUPPORTED_FORMAT" })
      });
  });

  it("repairs media relationship targets when replacing a mismatched PNG path containing SVG bytes", async () => {
    const zip = new JSZip();
    zip.file("[Content_Types].xml", '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="png" ContentType="image/png"/></Types>');
    zip.file("ppt/slides/slide1.xml", "<p:sld/>");
    zip.file("ppt/slides/_rels/slide1.xml.rels", '<Relationships><Relationship Id="rId1" Target="../media/image1.png"/></Relationships>');
    zip.file("ppt/media/image1.png", Buffer.from('<svg xmlns="http://www.w3.org/2000/svg"></svg>'));
    const bytes = await zip.generateAsync({ type: "uint8array" });
    const replaced = await replaceAsset(
      { data: bytes, format: "pptx" },
      { assetPath: "ppt/media/image1.png", replacement: Buffer.from('<svg xmlns="http://www.w3.org/2000/svg"></svg>'), replacementPath: "logo.svg" }
    );
    const outZip = await JSZip.loadAsync(replaced.bytes as Uint8Array);
    const rels = await outZip.file("ppt/slides/_rels/slide1.xml.rels")?.async("string");

    expect(replaced.media.targetAssetPath).toBe("ppt/media/image1.svg");
    expect(outZip.file("ppt/media/image1.png")).toBeNull();
    expect(outZip.file("ppt/media/image1.svg")).not.toBeNull();
    expect(rels).toContain("../media/image1.svg");
  });

  it("does not trust GIF media type from extension alone during asset replacement", async () => {
    const zip = new JSZip();
    zip.file("ppt/media/image1.gif", Buffer.from("GIF89a"));
    const bytes = await zip.generateAsync({ type: "uint8array" });

    await expect(replaceAsset({ data: bytes, format: "pptx" }, { assetPath: "ppt/media/image1.gif", replacement: Buffer.from("not-a-gif"), replacementPath: "logo.gif" }))
      .rejects.toThrow(/ASSET_UNSUPPORTED_FORMAT/);
  });

  it("returns best-effort PDF text previews when plain text operators are present", async () => {
    const rendered = await render({ title: "PDF", sections: [{ title: "Page", body: "Body" }] }, { target: "pdf" });
    const pdfBytes = Buffer.concat([Buffer.from(rendered.bytes as Uint8Array), Buffer.from("\nBT (Hello PDF) Tj ET\n", "latin1")]);
    const inspected = await inspect({ data: pdfBytes, format: "pdf" });

    expect(inspected.trusted.summary.textBlocks).toBeGreaterThan(0);
    expect(inspected.objectMap[0]?.text).toContain("Hello PDF");
  });

  it("keeps PDF summary inspect to previews and exposes full text only in full depth", async () => {
    const rendered = await render({ title: "PDF", sections: [{ title: "Page", body: "Body" }] }, { target: "pdf" });
    const pdfBytes = Buffer.concat([Buffer.from(rendered.bytes as Uint8Array), Buffer.from(`\nBT (${"Long PDF text ".repeat(800)}) Tj ET\n`, "latin1")]);

    const summary = await inspect({ data: pdfBytes, format: "pdf" }, { depth: "summary" });
    const full = await inspect({ data: pdfBytes, format: "pdf" }, { depth: "full" });

    expect(summary.untrusted.text).toBeUndefined();
    expect(summary.objectMap[0]?.text).toBeUndefined();
    expect(summary.untrusted.pages[0]?.textPreview.length).toBeLessThanOrEqual(300);
    expect(full.objectMap[0]?.text).toContain("Long PDF text");
  });

  it("limits PDF view to a bounded first page sample by default", async () => {
    const pdfDoc = await PDFDocument.create();
    for (let index = 0; index < 12; index += 1) pdfDoc.addPage([200, 200]);
    const bytes = await pdfDoc.save();
    const viewed = await view({ data: bytes, format: "pdf" });

    expect(viewed.pages).toHaveLength(10);
  });

  it("verifies openability, repair risk, and approximate visual readiness", async () => {
    const rendered = await render({ title: "Verify", slides: [{ title: "Ready", body: "Body" }] }, { target: "pptx" });
    const result = await verify({ data: rendered.bytes, format: "pptx" }, { visual: true });

    expect(result.openable).toBe(true);
    expect(result.noRepairDialogExpected).toBe(true);
    expect(result.visual?.pagesChecked).toBeGreaterThan(0);
    expect(result.readiness).not.toBe("blocked");
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
    expect(dryRun.bytes).toBeUndefined();
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

  it("replaces all text runs in a selected PPTX shape instead of leaving stale run text behind", async () => {
    const zip = new JSZip();
    zip.file(
      "ppt/slides/slide1.xml",
      [
        "<p:sld><p:sp><p:nvSpPr><p:cNvPr id=\"9\" name=\"Split Title\"/></p:nvSpPr><p:txBody>",
        "<a:p><a:r><a:t>Hello </a:t></a:r><a:r><a:t>world</a:t></a:r></a:p>",
        "</p:txBody></p:sp></p:sld>"
      ].join("")
    );
    const bytes = await zip.generateAsync({ type: "uint8array" });
    const inspected = await inspect({ data: bytes, format: "pptx" });
    const id = inspected.objectMap[0]?.stableObjectId;
    const edited = await edit(
      { data: bytes, format: "pptx" },
      [{ op: "setText", selector: { stableObjectId: id }, text: "Changed" }]
    );
    const reinspected = await inspect({ data: edited.bytes, format: "pptx" });

    expect(reinspected.objectMap[0]?.text).toBe("Changed");
    expect(reinspected.objectMap[0]?.text).not.toContain("world");
  });

  it("replaces all text runs in a selected DOCX paragraph instead of leaving stale run text behind", async () => {
    const zip = new JSZip();
    zip.file(
      "word/document.xml",
      "<w:document><w:body><w:p><w:r><w:t>Hello </w:t></w:r><w:r><w:t>world</w:t></w:r></w:p></w:body></w:document>"
    );
    const bytes = await zip.generateAsync({ type: "uint8array" });
    const inspected = await inspect({ data: bytes, format: "docx" });
    const id = inspected.objectMap[0]?.stableObjectId;
    const edited = await edit(
      { data: bytes, format: "docx" },
      [{ op: "setText", selector: { stableObjectId: id }, text: "Changed" }]
    );
    const reinspected = await inspect({ data: edited.bytes, format: "docx" });

    expect(reinspected.objectMap[0]?.text).toBe("Changed");
    expect(reinspected.objectMap[0]?.text).not.toContain("world");
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

  it("keeps PPTX shape stableObjectId values stable when slides are reordered", async () => {
    const zip = new JSZip();
    zip.file(
      "ppt/presentation.xml",
      "<p:presentation><p:sldIdLst><p:sldId id=\"256\" r:id=\"rId1\"/><p:sldId id=\"257\" r:id=\"rId2\"/></p:sldIdLst></p:presentation>"
    );
    zip.file(
      "ppt/_rels/presentation.xml.rels",
      "<Relationships><Relationship Id=\"rId1\" Target=\"slides/slide1.xml\"/><Relationship Id=\"rId2\" Target=\"slides/slide2.xml\"/></Relationships>"
    );
    zip.file(
      "ppt/slides/slide1.xml",
      "<p:sld><p:sp><p:nvSpPr><p:cNvPr id=\"7\" name=\"A\"/></p:nvSpPr><p:txBody><a:p><a:r><a:t>First</a:t></a:r></a:p></p:txBody></p:sp></p:sld>"
    );
    zip.file(
      "ppt/slides/slide2.xml",
      "<p:sld><p:sp><p:nvSpPr><p:cNvPr id=\"7\" name=\"B\"/></p:nvSpPr><p:txBody><a:p><a:r><a:t>Second</a:t></a:r></a:p></p:txBody></p:sp></p:sld>"
    );
    const bytes = await zip.generateAsync({ type: "uint8array" });
    const before = await inspect({ data: bytes, format: "pptx" });
    const beforeIds = new Map(before.objectMap.map((entry) => [entry.text, entry.stableObjectId]));
    const edited = await edit(
      { data: bytes, format: "pptx" },
      [{ op: "pptx.reorderSlides", order: [2, 1] }]
    );
    const after = await inspect({ data: edited.bytes, format: "pptx" });
    const afterIds = new Map(after.objectMap.map((entry) => [entry.text, entry.stableObjectId]));

    expect(after.untrusted.slides[0]?.text).toBe("Second");
    expect(afterIds.get("First")).toBe(beforeIds.get("First"));
    expect(afterIds.get("Second")).toBe(beforeIds.get("Second"));
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

  it("denies native renderer export unless active config explicitly allows it", async () => {
    const rendered = await render({ title: "Native export", slides: [{ title: "Slide", body: "Body" }] }, { target: "pptx" });

    await expect(
      exportDocument({ data: rendered.bytes, format: "pptx" }, { to: "pdf", mode: "native" })
    ).rejects.toMatchObject({
      payload: { code: "SECURITY_EXTERNAL_PROCESS_DENIED" }
    });
  });

  it("requires explicit file-backed native renderer for Office-to-PDF export", async () => {
    const rendered = await render({ title: "Native export", slides: [{ title: "Slide", body: "Body" }] }, { target: "pptx" });
    const config = getBuiltinConfig("enterprise");
    config.security.externalProcess = "allow";
    config.security.renderers = "enabled";

    await expect(
      exportDocument({ data: rendered.bytes, format: "pptx" }, { to: "pdf", mode: "native", config })
    ).rejects.toMatchObject({
      payload: { code: "EXPORT_UNSUPPORTED" }
    });
  });

  it("normalizes native renderer timeouts to a bounded default", () => {
    expect(resolveNativeRendererTimeoutMs()).toBe(DEFAULT_NATIVE_RENDERER_TIMEOUT_MS);
    expect(resolveNativeRendererTimeoutMs(Number.NaN)).toBe(DEFAULT_NATIVE_RENDERER_TIMEOUT_MS);
    expect(resolveNativeRendererTimeoutMs(10)).toBe(MIN_NATIVE_RENDERER_TIMEOUT_MS);
    expect(resolveNativeRendererTimeoutMs(1500.9)).toBe(1500);
  });

  it("rejects export target and output extension mismatches", async () => {
    const rendered = await render({ title: "Export mismatch", slides: [{ title: "Slide", body: "Body" }] }, { target: "pptx" });

    await expect(
      exportDocument({ data: rendered.bytes, format: "pptx" }, { to: "pdf", out: "export-mismatch.html" })
    ).rejects.toMatchObject({
      payload: { code: "TARGET_EXTENSION_MISMATCH" }
    });
  });

  it("skips PDF edit operations that target pages outside the document", async () => {
    const rendered = await render({ title: "One page", sections: [{ title: "Page", body: "Body" }] }, { target: "pdf" });
    const edited = await edit(
      { data: rendered.bytes, format: "pdf" },
      [{ op: "pdf.textOverlay", page: 99, text: "Out of range", x: 72, y: 72 }]
    );

    expect(edited.changed).toBe(false);
    expect(edited.applied).toBe(0);
    expect(edited.skipped).toBe(1);
  });
});

function pngBytes(width: number, height: number): Uint8Array {
  const bytes = new Uint8Array(24);
  bytes.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], 0);
  const view = new DataView(bytes.buffer);
  view.setUint32(16, width);
  view.setUint32(20, height);
  return bytes;
}
