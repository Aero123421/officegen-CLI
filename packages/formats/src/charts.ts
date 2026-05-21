import { escapeXml, sha256, writeOutput } from "./shared.js";

export interface ChartRenderOptions {
  width?: number;
  height?: number;
  out?: string;
  title?: string;
}

export interface ChartRenderResult {
  schema: "officegen.chart.render.result@1.2";
  format: "svg";
  width: number;
  height: number;
  svg: string;
  out?: string;
  sha256: string;
  caveats: string[];
}

export async function renderChart(spec: unknown, options: ChartRenderOptions = {}): Promise<ChartRenderResult> {
  const width = options.width ?? 640;
  const height = options.height ?? 360;
  const normalized = normalizeChartSpec(spec, options);
  const svg = buildBarLikeSvg(normalized, width, height);
  await writeOutput(options.out, svg);
  return {
    schema: "officegen.chart.render.result@1.2",
    format: "svg",
    width,
    height,
    svg,
    out: options.out,
    sha256: sha256(svg),
    caveats: ["MVP chart renderer supports simple inline SVG bar/line-like output without external Vega processing."]
  };
}

export const chartRender = renderChart;

interface NormalizedChart {
  title: string;
  labels: string[];
  values: number[];
}

function normalizeChartSpec(spec: unknown, options: ChartRenderOptions): NormalizedChart {
  const record = asRecord(spec);
  const rowData = chartRows(record);
  const pairData = rowData ? undefined : chartLabelValuePairs(record);
  const xField = readField(record.encoding, "x") ?? "label";
  const yField = readField(record.encoding, "y") ?? "value";
  const labels = rowData
    ? rowData.map((row, index) => String(row[xField] ?? row.label ?? row.name ?? index + 1))
    : pairData?.labels ?? [];
  const values = rowData
    ? rowData.map((row) => Number(row[yField] ?? row.value ?? row.y))
    : pairData?.values ?? [];

  assertChartData(labels, values);
  return {
    title: options.title ?? String(record.title ?? "Chart"),
    labels,
    values
  };
}

function asRecord(value: unknown): Record<string, unknown> {
  if (typeof value === "object" && value !== null && !Array.isArray(value)) return value as Record<string, unknown>;
  throw new Error("SCHEMA_INVALID: chart render requires a JSON object with data.values rows, data rows, or labels/values arrays.");
}

function chartRows(record: Record<string, unknown>): Array<Record<string, unknown>> | undefined {
  const data = record.data;
  if (Array.isArray(data)) return rowsOrThrow(data, "data");
  if (isRecord(data) && Array.isArray(data.values) && data.values.every(isRecord)) return rowsOrThrow(data.values, "data.values");
  if (Array.isArray(record.values) && record.values.every(isRecord)) return rowsOrThrow(record.values, "values");
  return undefined;
}

function chartLabelValuePairs(record: Record<string, unknown>): { labels: string[]; values: number[] } | undefined {
  const nested = isRecord(record.data) ? record.data : undefined;
  const labels = arrayValue(record.labels) ?? arrayValue(record.categories) ?? arrayValue(nested?.labels) ?? arrayValue(nested?.categories);
  const values = arrayValue(record.values) ?? arrayValue(nested?.values);
  if (!labels && !values) return undefined;
  return {
    labels: (labels ?? []).map((label) => String(label)),
    values: (values ?? []).map((value) => Number(value))
  };
}

function rowsOrThrow(rows: unknown[], path: string): Array<Record<string, unknown>> {
  if (!rows.every(isRecord)) {
    throw new Error(`SCHEMA_INVALID: chart render ${path} must contain row objects with label/value fields.`);
  }
  return rows;
}

function assertChartData(labels: string[], values: number[]): void {
  if (!labels.length || !values.length) {
    throw new Error("SCHEMA_INVALID: chart render requires chart data; supported shapes are {labels, values}, {data:[{label,value}]}, or {data:{values:[...]}}.");
  }
  if (labels.length !== values.length) {
    throw new Error(`SCHEMA_INVALID: chart render labels/values length mismatch (${labels.length} labels, ${values.length} values).`);
  }
  if (values.some((value) => !Number.isFinite(value))) {
    throw new Error("SCHEMA_INVALID: chart render values must be finite numbers.");
  }
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function arrayValue(value: unknown): unknown[] | undefined {
  return Array.isArray(value) ? value : undefined;
}

function readField(encoding: unknown, channel: string): string | undefined {
  const channelSpec = typeof encoding === "object" && encoding !== null ? (encoding as Record<string, unknown>)[channel] : undefined;
  return typeof channelSpec === "object" && channelSpec !== null
    ? String((channelSpec as Record<string, unknown>).field ?? "")
    : undefined;
}

function buildBarLikeSvg(chart: NormalizedChart, width: number, height: number): string {
  const margin = { top: 48, right: 24, bottom: 48, left: 48 };
  const plotWidth = width - margin.left - margin.right;
  const plotHeight = height - margin.top - margin.bottom;
  const max = Math.max(...chart.values, 1);
  const barWidth = plotWidth / chart.values.length;
  const bars = chart.values
    .map((value, index) => {
      const barHeight = (value / max) * plotHeight;
      const x = margin.left + index * barWidth + 8;
      const y = margin.top + plotHeight - barHeight;
      return `<g><rect x="${x}" y="${y}" width="${Math.max(4, barWidth - 16)}" height="${barHeight}" fill="#2f6f73"/><text x="${x + Math.max(4, barWidth - 16) / 2}" y="${height - 20}" text-anchor="middle" font-size="12" font-family="Arial">${escapeXml(chart.labels[index])}</text><text x="${x + Math.max(4, barWidth - 16) / 2}" y="${y - 6}" text-anchor="middle" font-size="12" font-family="Arial">${escapeXml(value)}</text></g>`;
    })
    .join("");
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}"><rect width="${width}" height="${height}" fill="#fff"/><text x="${margin.left}" y="28" font-family="Arial, sans-serif" font-size="20" font-weight="700">${escapeXml(chart.title)}</text><line x1="${margin.left}" y1="${margin.top + plotHeight}" x2="${width - margin.right}" y2="${margin.top + plotHeight}" stroke="#57606a"/><line x1="${margin.left}" y1="${margin.top}" x2="${margin.left}" y2="${margin.top + plotHeight}" stroke="#57606a"/>${bars}</svg>`;
}
