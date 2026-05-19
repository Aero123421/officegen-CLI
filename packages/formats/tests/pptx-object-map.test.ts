import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { inspect } from "../src/index.js";

describe("PPTX object map inspection", () => {
  it("detects groups, connectors, SmartArt relationships, and chart series", async () => {
    const zip = new JSZip();
    zip.file(
      "ppt/slides/slide1.xml",
      [
        '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram">',
        "<p:cSld><p:spTree>",
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>',
        '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="20" name="Group 1"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/><a:chOff x="0" y="0"/><a:chExt cx="914400" cy="914400"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="21" name="Grouped child"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:txBody><a:p><a:r><a:t>Grouped text</a:t></a:r></a:p></p:txBody></p:sp></p:grpSp>',
        '<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="30" name="Connector 1"/><p:cNvCxnSpPr><a:stCxn id="20" idx="0"/><a:endCxn id="21" idx="1"/></p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr><p:spPr><a:xfrm><a:off x="914400" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm><a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom></p:spPr></p:cxnSp>',
        '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="40" name="Sales Chart"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="914400"/><a:ext cx="1828800" cy="914400"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rIdChart"/></a:graphicData></a:graphic></p:graphicFrame>',
        '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="50" name="Process SmartArt"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="1828800"/><a:ext cx="1828800" cy="914400"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds r:dm="rIdData" r:lo="rIdLayout" r:qs="rIdQuickStyle" r:cs="rIdColors"/></a:graphicData></a:graphic></p:graphicFrame>',
        "</p:spTree></p:cSld></p:sld>"
      ].join("")
    );
    zip.file(
      "ppt/slides/_rels/slide1.xml.rels",
      [
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
        '<Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>',
        '<Relationship Id="rIdData" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/data1.xml"/>',
        '<Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout" Target="../diagrams/layout1.xml"/>',
        '<Relationship Id="rIdQuickStyle" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle" Target="../diagrams/quickStyle1.xml"/>',
        '<Relationship Id="rIdColors" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors" Target="../diagrams/colors1.xml"/>',
        "</Relationships>"
      ].join("")
    );
    zip.file(
      "ppt/charts/chart1.xml",
      [
        '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:ser>',
        '<c:idx val="0"/><c:order val="0"/>',
        "<c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:pt idx=\"0\"><c:v>Revenue</c:v></c:pt></c:strCache></c:strRef></c:tx>",
        "<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$3</c:f></c:strRef></c:cat>",
        "<c:val><c:numRef><c:f>Sheet1!$B$2:$B$3</c:f></c:numRef></c:val>",
        "</c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"
      ].join("")
    );
    zip.file(
      "ppt/diagrams/data1.xml",
      [
        '<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">',
        '<dgm:ptLst>',
        '<dgm:pt modelId="doc" type="doc"/>',
        '<dgm:pt modelId="n1" type="node"><dgm:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Discover</a:t></a:r></a:p></dgm:txBody></dgm:pt>',
        '<dgm:pt modelId="n2" type="node"><dgm:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Ship</a:t></a:r></a:p></dgm:txBody></dgm:pt>',
        "</dgm:ptLst>",
        '<dgm:cxnLst><dgm:cxn modelId="c1" type="parOf" srcId="doc" destId="n1" srcOrd="0" destOrd="0"/><dgm:cxn modelId="c2" type="parOf" srcId="n1" destId="n2" srcOrd="0" destOrd="0"/></dgm:cxnLst>',
        "</dgm:dataModel>"
      ].join("")
    );
    zip.file(
      "ppt/diagrams/layout1.xml",
      '<dgm:layoutDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" uniqueId="urn:layout:basicProcess" defStyle="simple1"/>'
    );
    zip.file(
      "ppt/diagrams/quickStyle1.xml",
      '<dgm:styleDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" uniqueId="urn:quickStyle:subtleEffect"/>'
    );
    zip.file(
      "ppt/diagrams/colors1.xml",
      '<dgm:colorsDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" uniqueId="urn:colors:accent1"/>'
    );
    zip.file(
      "ppt/diagrams/_rels/layout1.xml.rels",
      [
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
        '<Relationship Id="rIdLayoutImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>',
        "</Relationships>"
      ].join("")
    );

    const inspected = await inspect({ data: await zip.generateAsync({ type: "uint8array" }), format: "pptx" }, { depth: "full" });

    const group = inspected.objectMap.find((entry) => entry.kind === "group");
    const connector = inspected.objectMap.find((entry) => entry.kind === "connector");
    const smartArt = inspected.objectMap.find((entry) => entry.kind === "smartArt");
    const smartArtNode = inspected.objectMap.find((entry) => entry.kind === "smartArtNode" && entry.text === "Ship");
    const chartSeries = inspected.objectMap.find((entry) => entry.kind === "chartSeries");
    const smartArtMedia = smartArt?.media?.smartArt as Record<string, unknown> | undefined;

    expect(group?.selectorHints).toMatchObject({ slide: 1, shapeId: "20", childShapeIds: ["21"] });
    expect(connector?.selectorHints).toMatchObject({ slide: 1, shapeId: "30", startConnection: { shapeId: "20" }, endConnection: { shapeId: "21" } });
    expect(smartArt?.editableOps).toBeUndefined();
    expect(smartArt?.text).toBe("Discover\nShip");
    expect(smartArt?.selectorHints).toMatchObject({
      slide: 1,
      shapeId: "50",
      relationshipIds: { dm: "rIdData", lo: "rIdLayout", qs: "rIdQuickStyle", cs: "rIdColors" },
      dataPath: "ppt/diagrams/data1.xml",
      layoutId: "urn:layout:basicProcess",
      quickStyleId: "urn:quickStyle:subtleEffect",
      colorStyleId: "urn:colors:accent1"
    });
    expect(smartArt?.media?.relationships).toMatchObject({
      dm: { relationshipId: "rIdData", target: "ppt/diagrams/data1.xml" },
      lo: { relationshipId: "rIdLayout", target: "ppt/diagrams/layout1.xml", partRelationships: [{ relationshipId: "rIdLayoutImage", target: "ppt/media/image1.png" }] }
    });
    expect(smartArtMedia).toMatchObject({
      nodeTexts: ["Discover", "Ship"],
      graphicFrame: { slide: 1, sourcePath: "ppt/slides/slide1.xml", shapeId: "50" }
    });
    expect(smartArtMedia?.nodeTree).toMatchObject([
      { nodeId: "doc", childNodeIds: ["n1"], children: [{ nodeId: "n1", text: "Discover", childNodeIds: ["n2"], children: [{ nodeId: "n2", text: "Ship" }] }] }
    ]);
    expect(smartArtNode?.editableOps).toBeUndefined();
    expect(smartArtNode?.selectorHints).toMatchObject({
      slide: 1,
      smartArtShapeId: "50",
      nodeId: "n2",
      parentNodeId: "n1",
      dataPath: "ppt/diagrams/data1.xml",
      layoutId: "urn:layout:basicProcess"
    });
    expect(smartArtNode?.media?.smartArt).toMatchObject({
      dataPath: "ppt/diagrams/data1.xml",
      graphicFrame: { shapeId: "50", name: "Process SmartArt" }
    });
    expect(chartSeries?.label).toBe("Revenue");
    expect(chartSeries?.selectorHints).toMatchObject({
      slide: 1,
      chartShapeId: "40",
      relationshipId: "rIdChart",
      chartPath: "ppt/charts/chart1.xml",
      categoryRef: "Sheet1!$A$2:$A$3",
      valueRef: "Sheet1!$B$2:$B$3"
    });
  });
});
