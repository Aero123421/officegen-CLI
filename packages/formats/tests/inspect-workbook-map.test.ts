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
});
