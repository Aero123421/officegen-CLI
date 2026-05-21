import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { inspect } from "../src/index.js";

describe("inspect workbookMap formula samples", () => {
  it("attributes each formula only to its enclosing c cell", async () => {
    const zip = new JSZip();
    zip.file("xl/workbook.xml", '<workbook><sheets><sheet name="Model"/></sheets></workbook>');
    zip.file(
      "xl/worksheets/sheet1.xml",
      [
        "<worksheet><sheetData>",
        '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row>',
        '<row r="2"><c r="D2"><f>SUM(A1:B1)</f><v>3</v></c></row>',
        "</sheetData></worksheet>"
      ].join("")
    );

    const inspected = await inspect({ data: await zip.generateAsync({ type: "uint8array" }), format: "xlsx" });
    const samples = (inspected.untrusted.workbookMap as { formulas: Array<{ samples: Array<{ ref: string; formula: string }> }> })
      .formulas[0]?.samples ?? [];

    expect(samples).toEqual([{ ref: "D2", formula: "SUM(A1:B1)", untrusted: true }]);
    expect(samples.some((sample) => sample.ref === "A1")).toBe(false);
  });

  it("keeps scoped XLSX summary and sheets aligned with scoped objectMap cells", async () => {
    const zip = new JSZip();
    zip.file(
      "xl/workbook.xml",
      [
        '<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        '<sheets><sheet name="Data" sheetId="1" r:id="rIdData"/></sheets>',
        "</workbook>"
      ].join("")
    );
    zip.file(
      "xl/_rels/workbook.xml.rels",
      [
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
        '<Relationship Id="rIdData" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>',
        "</Relationships>"
      ].join("")
    );
    zip.file("xl/worksheets/sheet1.xml", '<worksheet><sheetData><row r="1"><c r="A1"><v>999</v></c></row></sheetData></worksheet>');
    zip.file(
      "xl/worksheets/sheet3.xml",
      [
        "<worksheet><sheetData>",
        '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c><c r="C1"><v>3</v></c><c r="D1"><v>4</v></c></row>',
        '<row r="2"><c r="A2"><v>5</v></c><c r="B2"><v>6</v></c><c r="C2"><f>SUM(A1:B2)</f><v>14</v></c></row>',
        "</sheetData></worksheet>"
      ].join("")
    );

    const inspected = await inspect(
      { data: await zip.generateAsync({ type: "uint8array" }), format: "xlsx" },
      { depth: "summary", sheet: "Data", range: "A1:C2" }
    );

    expect(inspected.objectMap.map((entry) => entry.label)).toEqual(["A1", "B1", "C1", "A2", "B2", "C2"]);
    expect(inspected.trusted.summary.cells).toBe(6);
    expect(inspected.trusted.summary.formulas).toBe(1);
    expect(inspected.untrusted.sheets).toHaveLength(1);
    expect(inspected.untrusted.sheets[0]).toMatchObject({ name: "Data", cellCount: 6, usedRange: "A1:C2" });
  });
});
