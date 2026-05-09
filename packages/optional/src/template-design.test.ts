import { mkdtemp, readFile, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { inspect } from "@officegen/formats";
import { createOptionalCapabilities } from "./common.js";
import { applyDesign, captureDesign, capturePptxDesignSignals, initDesign, type DesignSourceCapture } from "./design.js";
import { applyTemplateMap, createTemplate, fillTemplate, templateCandidates } from "./template.js";

describe("@officegen/optional PPTX template and design signals", () => {
  it("extracts PPTX design signals, writes evidence artifacts, and suggests template maps", async () => {
    const cwd = await mkdtemp(path.join(os.tmpdir(), "officegen-optional-"));
    const deckPath = path.join(cwd, "source.pptx");
    await writeFile(deckPath, await makePptxFixture());

    const signals = await capturePptxDesignSignals(deckPath, {
      cwd,
      artifactsDir: ".officegen/optional/design/smoke"
    });

    expect(signals?.metadata.slides).toBe(2);
    expect(signals?.colorRoleCandidates.map((candidate) => candidate.role)).toEqual(
      expect.arrayContaining(["background", "text", "accent"])
    );
    expect(signals?.textSizeDistribution.length).toBeGreaterThan(0);
    expect(signals?.bboxPatterns.map((pattern) => pattern.kind)).toEqual(expect.arrayContaining(["title", "body", "image"]));
    expect(signals?.chartPresence).toMatchObject({ count: 1, slides: [1] });
    expect(signals?.diagramPresence).toMatchObject({ count: 1, slides: [2] });
    expect(signals?.slideSignals[0]?.slideType).toBe("chart");
    expect(signals?.schemaCandidates.map((field) => field.name)).toContain("quarter");
    expect(signals?.schemaCandidates.find((field) => field.name === "title")?.type).toBe("string");
    expect(signals?.schemaCandidates.find((field) => field.name === "product_update")?.type).toBe("string");
    expect(signals?.schemaCandidates.find((field) => field.name === "launch_date")?.type).toBe("date");
    expect(signals?.schemaCandidates.find((field) => field.name === "hero_image")?.type).toBe("image");
    expect(signals?.schemaCandidates.find((field) => field.name === "revenue_chart")?.type).toBe("chartData");
    expect(signals?.templateMapSuggested.mapping.quarter).toMatch(/^pptx:slide-[a-f0-9]{8}:shape:/);
    expect(signals?.trust.trusted.schema).toBe("officegen.design.signals.trusted@1.2");
    expect(signals?.trust.untrusted.textSamples.join(" ")).toContain("Quarterly Review");
    expect(await readFile(signals?.artifactPaths?.contextPath ?? "", "utf8")).toContain("PPTX design context");
    expect(await readFile(signals?.artifactPaths?.templateMapSuggestedPath ?? "", "utf8")).toContain("quarter");
  });

  it("emits suggested template selectors that match inspect/edit stable IDs", async () => {
    const cwd = await mkdtemp(path.join(os.tmpdir(), "officegen-optional-selectors-"));
    const deckPath = path.join(cwd, "source.pptx");
    await writeFile(deckPath, await makePptxFixture());
    const optional = {
      cwd,
      capabilities: createOptionalCapabilities(["template", "design"])
    };

    const signals = await capturePptxDesignSignals(deckPath, {
      cwd,
      artifactsDir: ".officegen/optional/design/selectors"
    });
    const inspected = await inspect(deckPath, { format: "pptx" });
    const inspectedQuarterId = inspected.objectMap.find((entry) => entry.text === "Quarter: {{quarter}}")?.stableObjectId;
    const inspectedChartId = inspected.objectMap.find((entry) => entry.kind === "chart")?.stableObjectId;
    const candidates = await templateCandidates({ ...optional, sourcePath: deckPath });

    expect(signals?.templateMapSuggested.mapping.quarter).toBe(inspectedQuarterId);
    expect(signals?.templateMapSuggested.mapping.revenue_chart).toBe(inspectedChartId);
    expect(candidates[0]?.template.mapping?.quarter).toBe(inspectedQuarterId);
    expect(inspectedQuarterId).toMatch(/^pptx:slide-[a-f0-9]{8}:shape:/);
  });

  it("creates and captures source-derived template/design records from a PPTX", async () => {
    const cwd = await mkdtemp(path.join(os.tmpdir(), "officegen-optional-api-"));
    const deckPath = path.join(cwd, "source.pptx");
    await writeFile(deckPath, await makePptxFixture());
    const optional = {
      cwd,
      capabilities: createOptionalCapabilities(["template", "design"])
    };

    const candidates = await templateCandidates({ ...optional, sourcePath: deckPath });
    expect(candidates[0]?.generatedFromSource).toBe(true);
    expect(candidates[0]?.template.fields?.map((field) => field.name)).toContain("quarter");
    expect(candidates[0]?.artifactPaths?.evidencePath).toContain("evidence.json");

    const created = await createTemplate({
      ...optional,
      sourcePath: deckPath,
      template: {
        id: "quarterly-review",
        name: "Quarterly Review"
      }
    });
    expect(created.fields?.map((field) => field.name)).toContain("quarter");
    expect(created.mapping?.quarter).toMatch(/^pptx:slide-[a-f0-9]{8}:shape:/);
    expect(created.sourceCapture?.trust?.agentInstruction).toContain("untrusted");

    await initDesign({ ...optional, id: "deck-design" });
    const design = await captureDesign({ ...optional, id: "deck-design", sourcePath: deckPath });
    const capture = design.sourceCapture as DesignSourceCapture;
    expect(capture.densityScore).toBeGreaterThan(0);
    expect((design.tokens.color as Record<string, unknown>).palette).toEqual(expect.any(Array));
    expect((design.tokens.layout as Record<string, unknown>).archetypes).toEqual(expect.any(Array));
    expect(capture.artifactPaths?.previewPaths[0]).toContain("preview-slide-001.svg");
    expect(capture.schemaCandidates?.map((field) => field.name)).toContain("quarter");
  });

  it("returns explicit plan-only JSON and preserves object template mappings", async () => {
    const cwd = await mkdtemp(path.join(os.tmpdir(), "officegen-optional-plans-"));
    const optional = {
      cwd,
      capabilities: createOptionalCapabilities(["template", "design"])
    };
    const template = await createTemplate({
      ...optional,
      template: {
        id: "plan-template",
        name: "Plan Template",
        fields: [
          { name: "title", type: "string", required: true },
          { name: "metadata", type: "json" }
        ]
      }
    });

    const mapPlan = await applyTemplateMap({
      ...optional,
      id: template.id,
      mapping: {
        title: {
          selector: {
            stableObjectId: "pptx:slide-00000000:shape:0001"
          },
          transform: { trim: true }
        }
      }
    });
    const fillPlan = await fillTemplate({
      ...optional,
      id: template.id,
      values: {
        title: "Quarterly Review",
        metadata: { audience: "exec" }
      }
    });
    const fillValidation = await fillTemplate({
      ...optional,
      id: template.id,
      values: {
        title: "Quarterly Review",
        metadata: { audience: "exec" }
      },
      validateOnly: true
    });
    await initDesign({ ...optional, id: "plan-design" });
    const designPlan = await applyDesign({ ...optional, id: "plan-design" });

    expect(mapPlan).toMatchObject({
      kind: "officegen.template.apply-map",
      planOnly: true,
      mapping: {
        title: {
          selector: {
            stableObjectId: "pptx:slide-00000000:shape:0001"
          },
          transform: { trim: true }
        }
      }
    });
    expect(fillPlan).toMatchObject({
      kind: "officegen.template.fill",
      planOnly: true,
      values: {
        metadata: { audience: "exec" }
      }
    });
    expect(fillValidation).toMatchObject({
      validateOnly: true,
      bindings: expect.any(Array)
    });
    expect(designPlan).toMatchObject({
      kind: "officegen.design.apply",
      planOnly: true
    });
  });

  it("fills a source-backed PPTX template into a real Office file", async () => {
    const cwd = await mkdtemp(path.join(os.tmpdir(), "officegen-template-fill-"));
    const deckPath = path.join(cwd, "source.pptx");
    const outPath = path.join(cwd, "filled.pptx");
    await writeFile(deckPath, await makePptxFixture());
    const optional = {
      cwd,
      capabilities: createOptionalCapabilities(["template"])
    };
    const template = await createTemplate({
      ...optional,
      sourcePath: deckPath,
      template: { id: "deck", name: "Deck", fields: [{ name: "quarter", type: "string" }] }
    });
    expect(template.mapping?.quarter).toBeTruthy();

    const result = await fillTemplate({ ...optional, id: "deck", values: { quarter: "Q4 FY27" }, outputPath: outPath });
    const inspected = await inspect(outPath);

    expect(result.planOnly).toBe(false);
    expect(result.mutatesOffice).toBe(true);
    expect((result.artifacts as Array<{ exists?: boolean }>)[0]?.exists).toBe(true);
    expect(inspected.objectMap.map((entry) => entry.text).join(" ")).toContain("Q4 FY27");
  });

  it("does not report Office mutation when template fill has no resolved operations", async () => {
    const cwd = await mkdtemp(path.join(os.tmpdir(), "officegen-template-empty-fill-"));
    const deckPath = path.join(cwd, "source.pptx");
    const outPath = path.join(cwd, "filled.pptx");
    await writeFile(deckPath, await makePptxFixture());
    const optional = {
      cwd,
      capabilities: createOptionalCapabilities(["template"])
    };
    await createTemplate({
      ...optional,
      sourcePath: deckPath,
      template: { id: "empty", name: "Empty", fields: [{ name: "missing", type: "string" }] }
    });

    await expect(fillTemplate({ ...optional, id: "empty", values: { missing: "value" }, outputPath: outPath }))
      .rejects.toThrow(/no Office edit operations/i);
  });
});

async function makePptxFixture(): Promise<Buffer> {
  const zip = new JSZip();
  zip.file("docProps/core.xml", "<cp:coreProperties><dc:title>Quarterly Review</dc:title><dc:creator>Officegen Test</dc:creator></cp:coreProperties>");
  zip.file("ppt/presentation.xml", '<p:presentation><p:sldSz cx="9144000" cy="6858000"/></p:presentation>');
  zip.file("ppt/theme/theme1.xml", '<a:theme><a:srgbClr val="2563EB"/><a:srgbClr val="F8FAFC"/></a:theme>');
  zip.file(
    "ppt/slides/slide1.xml",
    [
      '<p:sld><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="F8FAFC"/></a:solidFill></p:bgPr></p:bg>',
      shapeXml({
        id: 1,
        name: "Title Placeholder 1",
        placeholder: "title",
        x: 457200,
        y: 274320,
        cx: 8229600,
        cy: 731520,
        color: "111827",
        size: 3600,
        text: "Quarterly Review"
      }),
      shapeXml({
        id: 2,
        name: "Project KPI",
        placeholder: "body",
        x: 685800,
        y: 1371600,
        cx: 4114800,
        cy: 1828800,
        color: "2563EB",
        size: 2000,
        text: "Quarter: {{quarter}}"
      }),
      picXml(3, "Hero Image", 5029200, 1371600, 3429000, 2286000),
      '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Revenue Chart"/></p:nvGraphicFramePr><p:xfrm><a:off x="685800" y="3657600"/><a:ext cx="7772400" cy="2286000"/></p:xfrm><a:graphic><a:graphicData><c:chart r:id="rId1"/></a:graphicData></a:graphic></p:graphicFrame>',
      "</p:cSld></p:sld>"
    ].join("")
  );
  zip.file("ppt/slides/_rels/slide1.xml.rels", '<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>');
  zip.file("ppt/charts/chart1.xml", "<c:chartSpace><c:chart><c:title/></c:chart></c:chartSpace>");
  zip.file(
    "ppt/slides/slide2.xml",
    [
      "<p:sld><p:cSld>",
      shapeXml({
        id: 1,
        name: "Section Header",
        placeholder: "title",
        x: 457200,
        y: 457200,
        cx: 8229600,
        cy: 914400,
        color: "111827",
        size: 3200,
        text: "Operating Model"
      }),
      shapeXml({
        id: 3,
        name: "Product Update Field",
        placeholder: "body",
        x: 685800,
        y: 1371600,
        cx: 3657600,
        cy: 548640,
        color: "111827",
        size: 1800,
        text: "Product Update"
      }),
      shapeXml({
        id: 4,
        name: "Launch Date Field",
        placeholder: "body",
        x: 685800,
        y: 2057400,
        cx: 3657600,
        cy: 548640,
        color: "111827",
        size: 1800,
        text: "Launch Date: {{launch_date}}"
      }),
      '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="2" name="Process Diagram"/></p:nvGraphicFramePr><p:xfrm><a:off x="914400" y="1828800"/><a:ext cx="7315200" cy="3200400"/></p:xfrm><a:graphic><a:graphicData><dgm:relIds/></a:graphicData></a:graphic></p:graphicFrame>',
      "</p:cSld></p:sld>"
    ].join("")
  );
  return zip.generateAsync({ type: "nodebuffer" });
}

function shapeXml(input: {
  id: number;
  name: string;
  placeholder: string;
  x: number;
  y: number;
  cx: number;
  cy: number;
  color: string;
  size: number;
  text: string;
}): string {
  return [
    "<p:sp>",
    `<p:nvSpPr><p:cNvPr id="${input.id}" name="${input.name}"/><p:nvPr><p:ph type="${input.placeholder}"/></p:nvPr></p:nvSpPr>`,
    `<p:spPr><a:xfrm><a:off x="${input.x}" y="${input.y}"/><a:ext cx="${input.cx}" cy="${input.cy}"/></a:xfrm><a:solidFill><a:srgbClr val="${input.color}"/></a:solidFill></p:spPr>`,
    `<p:txBody><a:p><a:r><a:rPr sz="${input.size}"><a:solidFill><a:srgbClr val="${input.color}"/></a:solidFill></a:rPr><a:t>${input.text}</a:t></a:r></a:p></p:txBody>`,
    "</p:sp>"
  ].join("");
}

function picXml(id: number, name: string, x: number, y: number, cx: number, cy: number): string {
  return [
    "<p:pic>",
    `<p:nvPicPr><p:cNvPr id="${id}" name="${name}"/></p:nvPicPr>`,
    `<p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm></p:spPr>`,
    "</p:pic>"
  ].join("");
}
