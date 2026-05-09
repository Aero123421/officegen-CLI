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
export declare function renderChart(spec: unknown, options?: ChartRenderOptions): Promise<ChartRenderResult>;
export declare const chartRender: typeof renderChart;
