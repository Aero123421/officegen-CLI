import type { RuntimeContext } from "./types.js";

export async function withJsonStdoutDiagnosticsRedirect<T>(
  context: Pick<RuntimeContext, "json">,
  stderr: (text: string) => void,
  action: () => Promise<T>
): Promise<T> {
  if (!context.json) return action();

  const originalConsole = { debug: console.debug, info: console.info, log: console.log, warn: console.warn };
  const originalStdoutWrite = process.stdout.write.bind(process.stdout);
  const diagnosticWriter = (...values: unknown[]) => stderr(values.map((value) => typeof value === "string" ? value : String(value)).join(" "));
  console.debug = diagnosticWriter;
  console.info = diagnosticWriter;
  console.log = diagnosticWriter;
  console.warn = diagnosticWriter;
  process.stdout.write = ((chunk: string | Uint8Array, encodingOrCallback?: BufferEncoding | ((error?: Error | null) => void), callback?: (error?: Error | null) => void): boolean => {
    const encoding = typeof encodingOrCallback === "string" ? encodingOrCallback : undefined;
    const done = typeof encodingOrCallback === "function" ? encodingOrCallback : callback;
    const text = typeof chunk === "string" ? chunk : Buffer.from(chunk).toString(encoding ?? "utf8");
    stderr(text.endsWith("\n") ? text.slice(0, -1) : text);
    done?.();
    return true;
  }) as typeof process.stdout.write;

  try {
    return await action();
  } finally {
    console.debug = originalConsole.debug;
    console.info = originalConsole.info;
    console.log = originalConsole.log;
    console.warn = originalConsole.warn;
    process.stdout.write = originalStdoutWrite as typeof process.stdout.write;
  }
}
