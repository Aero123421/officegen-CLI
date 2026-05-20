export async function withJsonStdoutDiagnosticsRedirect(context, stderr, action) {
    if (!context.json)
        return action();
    const originalConsole = { debug: console.debug, info: console.info, log: console.log, warn: console.warn };
    const originalStdoutWrite = process.stdout.write.bind(process.stdout);
    const diagnosticWriter = (...values) => stderr(values.map((value) => typeof value === "string" ? value : String(value)).join(" "));
    console.debug = diagnosticWriter;
    console.info = diagnosticWriter;
    console.log = diagnosticWriter;
    console.warn = diagnosticWriter;
    process.stdout.write = ((chunk, encodingOrCallback, callback) => {
        const encoding = typeof encodingOrCallback === "string" ? encodingOrCallback : undefined;
        const done = typeof encodingOrCallback === "function" ? encodingOrCallback : callback;
        const text = typeof chunk === "string" ? chunk : Buffer.from(chunk).toString(encoding ?? "utf8");
        stderr(text.endsWith("\n") ? text.slice(0, -1) : text);
        done?.();
        return true;
    });
    try {
        return await action();
    }
    finally {
        await flushDeferredStdoutDiagnostics();
        console.debug = originalConsole.debug;
        console.info = originalConsole.info;
        console.log = originalConsole.log;
        console.warn = originalConsole.warn;
        process.stdout.write = originalStdoutWrite;
    }
}
async function flushDeferredStdoutDiagnostics() {
    await new Promise((resolve) => setImmediate(resolve));
}
//# sourceMappingURL=diagnostics.js.map