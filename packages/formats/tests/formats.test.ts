import { describe, expect, it } from "vitest";
import { inspect, render, renderChart, renderDiagram, view } from "../src/index.js";

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
    expect(inspected.objectMap[0]?.stableObjectId).toMatch(/^pptx:s001:text:/);
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
});

