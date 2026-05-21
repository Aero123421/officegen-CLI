import { describe, expect, it } from "vitest";
import { renderDiagram } from "../src/index.js";

describe("diagram SVG render", () => {
  it("keeps the first node rect inside the viewport with symmetric horizontal padding", async () => {
    const diagram = await renderDiagram("flowchart LR\nA-->B\nB-->C", { title: "Flow", width: 800, height: 420 });
    const rects = [...diagram.svg.matchAll(/<rect x="(-?\d+(?:\.\d+)?)"/g)].map((match) => Number(match[1]));

    expect(rects.length).toBeGreaterThanOrEqual(3);
    expect(rects[0]).toBeGreaterThanOrEqual(0);
    expect(rects[rects.length - 1]).toBeLessThanOrEqual(800 - 128);
  });

  it("clamps node padding for small custom widths", async () => {
    const diagram = await renderDiagram("flowchart LR\nA-->B", { width: 96, height: 180 });
    const rects = [...diagram.svg.matchAll(/<rect x="(-?\d+(?:\.\d+)?)"[^>]*width="(\d+(?:\.\d+)?)"/g)]
      .map((match) => ({ x: Number(match[1]), width: Number(match[2]) }))
      .filter((rect) => rect.width > 1);

    expect(rects.length).toBeGreaterThanOrEqual(2);
    for (const rect of rects) {
      expect(rect.x).toBeGreaterThanOrEqual(0);
      expect(rect.x + rect.width).toBeLessThanOrEqual(96);
    }
  });

  it("splits Mermaid node ids from bracket labels", async () => {
    const diagram = await renderDiagram("flowchart LR\nA[Inspect] --> B[Edit]", { width: 420, height: 220 });

    expect(diagram.svg).toContain("Inspect");
    expect(diagram.svg).toContain("Edit");
    expect(diagram.svg).not.toContain("AInspect");
    expect(diagram.svg).not.toContain("BEdit");
  });
});
