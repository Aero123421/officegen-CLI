import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { validateSchema } from "@officegen/core";
import { inspect } from "../src/index.js";

describe("XLSX formula graph inspect", () => {
  it("detects shared formulas and basic A1/range/3D/named dependencies", async () => {
    const zip = new JSZip();
    zip.file(
      "xl/workbook.xml",
      [
        '<workbook><sheets><sheet name="Model"/><sheet name="Lookup"/></sheets>',
        '<definedNames><definedName name="TaxRate">Lookup!$B$1</definedName></definedNames></workbook>'
      ].join("")
    );
    zip.file(
      "xl/worksheets/sheet1.xml",
      [
        '<worksheet><sheetData>',
        '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c><c r="C1"><f t="shared" si="0" ref="C1:C2">SUM(A1:B1)+Lookup!A1+TaxRate</f><v>3</v></c></row>',
        '<row r="2"><c r="C2"><f t="shared" si="0"/><v>4</v></c></row>',
        '</sheetData></worksheet>'
      ].join("")
    );
    zip.file("xl/worksheets/sheet2.xml", '<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>0.1</v></c></row></sheetData></worksheet>');

    const inspected = await inspect({ data: await zip.generateAsync({ type: "uint8array" }), format: "xlsx" }, { depth: "full" });
    const graph = (inspected.untrusted.sheets[0] as any).formulaGraph;
    const base = graph.formulaCells.find((cell: any) => cell.ref === "C1");
    const sharedChild = graph.formulaCells.find((cell: any) => cell.ref === "C2");

    expect(graph.formulaCells).toHaveLength(2);
    expect(graph.schema).toBe("officegen.xlsx.formulaGraph@1.0");
    expect(validateSchema("officegen.xlsx.formulaGraph@1.0", graph).ok).toBe(true);
    expect(sharedChild).toMatchObject({ formulaType: "shared", sharedIndex: "0", formula: "SUM(A1:B1)+Lookup!A1+TaxRate" });
    expect(base.dependencies.map((dependency: any) => dependency.kind)).toEqual(expect.arrayContaining(["range", "threeD", "namedRange"]));
    expect(base.dependencies.find((dependency: any) => dependency.kind === "threeD")).toMatchObject({ sheet: "Lookup", ref: "A1" });
    expect(base.dependencies.find((dependency: any) => dependency.kind === "namedRange")).toMatchObject({ name: "TaxRate" });
  });

  it("classifies external, volatile, indirect, and unsupported formulas as unsafe", async () => {
    const zip = new JSZip();
    zip.file("xl/workbook.xml", '<workbook><sheets><sheet name="Model"/></sheets></workbook>');
    zip.file(
      "xl/worksheets/sheet1.xml",
      '<worksheet><sheetData><row r="1"><c r="A1"><f>[Rates.xlsx]Sheet1!A1+NOW()+INDIRECT(&quot;A1&quot;)+_xlfn.UNKNOWN(A1)</f><v>0</v></c></row></sheetData></worksheet>'
    );
    zip.file("xl/externalLinks/externalLink1.xml", "<externalLink/>");

    const inspected = await inspect({ data: await zip.generateAsync({ type: "uint8array" }), format: "xlsx" }, { depth: "full" });
    const formulaCell = (inspected.untrusted.sheets[0] as any).formulaGraph.formulaCells[0];

    expect(formulaCell.unsafeFlags).toEqual(["external", "indirect", "unsupported", "volatile"]);
    expect(formulaCell.volatileFunctions).toEqual(["INDIRECT", "NOW"]);
    expect(formulaCell.dependencies.find((dependency: any) => dependency.kind === "threeD")).toMatchObject({ workbook: "Rates.xlsx", sheet: "Sheet1", ref: "A1" });
  });

  it("detects structured references and relates formulas to tables and slicers", async () => {
    const zip = new JSZip();
    zip.file("xl/workbook.xml", '<workbook><sheets><sheet name="Data"/></sheets></workbook>');
    zip.file(
      "xl/worksheets/sheet1.xml",
      '<worksheet><sheetData><row r="1"><c r="C1"><f>SUM(SalesTable[Amount])</f><v>30</v></c></row></sheetData></worksheet>'
    );
    zip.file("xl/tables/table1.xml", '<table name="SalesTable" displayName="SalesTable" ref="A1:B3"/>');
    zip.file("xl/slicers/slicer1.xml", '<slicer name="RegionSlicer" table="SalesTable"/>');

    const inspected = await inspect({ data: await zip.generateAsync({ type: "uint8array" }), format: "xlsx" }, { depth: "full" });
    const formulaCell = (inspected.untrusted.sheets[0] as any).formulaGraph.formulaCells[0];

    expect(formulaCell.dependencies).toContainEqual(expect.objectContaining({ kind: "tableStructuredRef", tableName: "SalesTable" }));
    expect(formulaCell.relatedObjects).toContainEqual(expect.objectContaining({ kind: "table", name: "SalesTable", reason: "structured-ref" }));
    expect(formulaCell.relatedObjects).toContainEqual(expect.objectContaining({ kind: "slicer", name: "RegionSlicer", reason: "table-slicer" }));
    expect(inspected.objectMap.some((entry) => entry.kind === "table" && entry.label === "SalesTable")).toBe(true);
  });

  it("relates range formulas to chart and pivot source ranges", async () => {
    const zip = new JSZip();
    zip.file("xl/workbook.xml", '<workbook><sheets><sheet name="Data"/></sheets></workbook>');
    zip.file(
      "xl/worksheets/sheet1.xml",
      '<worksheet><sheetData><row r="1"><c r="B1"><f>SUM(A2:A3)</f><v>30</v></c></row></sheetData></worksheet>'
    );
    zip.file("xl/charts/chart1.xml", '<c:chartSpace><c:chart><c:plotArea><c:barChart><c:ser><c:val><c:numRef><c:f>Data!$A$2:$A$3</c:f></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>');
    zip.file("xl/pivotTables/pivotTable1.xml", '<pivotTableDefinition name="Pivot1"><pivotCacheDefinition><cacheSource><worksheetSource sheet="Data" ref="A2:A3"/></cacheSource></pivotCacheDefinition></pivotTableDefinition>');

    const inspected = await inspect({ data: await zip.generateAsync({ type: "uint8array" }), format: "xlsx" }, { depth: "full" });
    const relatedObjects = (inspected.untrusted.sheets[0] as any).formulaGraph.formulaCells[0].relatedObjects;

    expect(relatedObjects).toContainEqual(expect.objectContaining({ kind: "chart", path: "xl/charts/chart1.xml", reason: "range-overlap" }));
    expect(relatedObjects).toContainEqual(expect.objectContaining({ kind: "pivotTable", name: "Pivot1", reason: "source-overlap" }));
  });
});
