import { escapeXml, sha256, writeOutput } from "./shared.js";

export interface DiagramRenderOptions {
  width?: number;
  height?: number;
  out?: string;
  title?: string;
}

export interface DiagramRenderResult {
  schema: "officegen.diagram.render.result@1.2";
  format: "svg";
  svg: string;
  out?: string;
  sha256: string;
  caveats: string[];
}

export async function renderDiagram(source: string, options: DiagramRenderOptions = {}): Promise<DiagramRenderResult> {
  const width = options.width ?? 800;
  const height = options.height ?? 420;
  const graph = parseSimpleMermaid(source);
  const svg = buildDiagramSvg(graph, width, height, options.title);
  await writeOutput(options.out, svg);
  return {
    schema: "officegen.diagram.render.result@1.2",
    format: "svg",
    svg,
    out: options.out,
    sha256: sha256(svg),
    caveats: ["MVP diagram renderer supports simple Mermaid-like A-->B edges as static SVG without external processes."]
  };
}

export const diagramRender = renderDiagram;

interface DiagramGraph {
  nodes: string[];
  edges: Array<[string, string]>;
}

function parseSimpleMermaid(source: string): DiagramGraph {
  const edges: Array<[string, string]> = [];
  const nodes = new Set<string>();
  for (const rawLine of source.split(/\r?\n/)) {
    const line = rawLine.trim();
    if (!line || /^(graph|flowchart|sequenceDiagram|classDiagram)\b/.test(line)) continue;
    const match = /^(.+?)\s*[-=.]+>\s*(.+)$/.exec(line);
    if (match) {
      const from = cleanNode(match[1] ?? "");
      const to = cleanNode(match[2] ?? "");
      if (from && to) {
        nodes.add(from);
        nodes.add(to);
        edges.push([from, to]);
      }
    } else {
      const node = cleanNode(line);
      if (node) nodes.add(node);
    }
  }
  if (!nodes.size) {
    nodes.add("Start");
    nodes.add("Finish");
    edges.push(["Start", "Finish"]);
  }
  return { nodes: [...nodes], edges };
}

function cleanNode(value: string): string {
  return value.replace(/[\[\]{}()"]/g, "").replace(/\|.*?\|/g, "").trim();
}

const DIAGRAM_NODE_HALF_WIDTH = 64;

function buildDiagramSvg(graph: DiagramGraph, width: number, height: number, title?: string): string {
  const top = title ? 64 : 32;
  const nodeHalfWidth = Math.min(DIAGRAM_NODE_HALF_WIDTH, Math.max(1, width / 2));
  const horizontalPad = nodeHalfWidth;
  const layoutWidth = Math.max(0, width - horizontalPad * 2);
  const gap = layoutWidth / Math.max(1, graph.nodes.length - 1);
  const positions = new Map(
    graph.nodes.map((node, index) => [
      node,
      {
        x: graph.nodes.length === 1 ? width / 2 : horizontalPad + index * gap,
        y: top + (index % 2) * 120
      }
    ])
  );
  const edges = graph.edges
    .map(([from, to]) => {
      const a = positions.get(from);
      const b = positions.get(to);
      if (!a || !b) return "";
      return `<line x1="${a.x}" y1="${a.y + 28}" x2="${b.x}" y2="${b.y + 28}" stroke="#57606a" stroke-width="2" marker-end="url(#arrow)"/>`;
    })
    .join("");
  const nodes = graph.nodes
    .map((node) => {
      const pos = positions.get(node) ?? { x: 0, y: 0 };
      return `<g><rect x="${pos.x - nodeHalfWidth}" y="${pos.y}" width="${nodeHalfWidth * 2}" height="56" rx="6" fill="#f6f8fa" stroke="#2f6f73" stroke-width="2"/><text x="${pos.x}" y="${pos.y + 34}" text-anchor="middle" font-family="Arial, sans-serif" font-size="14" fill="#111">${escapeXml(node)}</text></g>`;
    })
    .join("");
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}"><defs><marker id="arrow" markerWidth="10" markerHeight="10" refX="8" refY="3" orient="auto"><path d="M0,0 L0,6 L9,3 z" fill="#57606a"/></marker></defs><rect width="${width}" height="${height}" fill="#fff"/>${title ? `<text x="32" y="36" font-family="Arial, sans-serif" font-size="22" font-weight="700">${escapeXml(title)}</text>` : ""}${edges}${nodes}</svg>`;
}
