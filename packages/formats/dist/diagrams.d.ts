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
export declare function renderDiagram(source: string, options?: DiagramRenderOptions): Promise<DiagramRenderResult>;
export declare const diagramRender: typeof renderDiagram;
