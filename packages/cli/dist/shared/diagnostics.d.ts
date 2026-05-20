import type { RuntimeContext } from "./types.js";
export declare function withJsonStdoutDiagnosticsRedirect<T>(context: Pick<RuntimeContext, "json">, stderr: (text: string) => void, action: () => Promise<T>): Promise<T>;
