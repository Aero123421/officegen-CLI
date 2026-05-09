export const VALUE_OPTIONS = new Set([
    "--out",
    "--schema",
    "--kind",
    "--title",
    "--target",
    "--scope",
    "--from",
    "--to",
    "--mode",
    "--renderer",
    "--ops",
    "--views",
    "--include",
    "--slides",
    "--pages",
    "--object-map-limit",
    "--fields",
    "--depth",
    "--role",
    "--name",
    "--map",
    "--data",
    "--strategy",
    "--selector",
    "--asset",
    "--format",
    "--max-pages",
    "--issues",
    "--sha256",
    "--trust",
    "--allow-root",
    "--capabilities-hash",
    "--json-budget-bytes",
    "--report-out",
    "--log-jsonl",
    "--manifest",
    "--summary",
    "--output-root",
    "--expected-artifacts",
    "--timeout-ms",
    "--profile"
]);
export function commandFromArgv(argv) {
    const parts = [];
    for (let index = 2; index < argv.length; index += 1) {
        const arg = argv[index];
        if (arg.startsWith("-")) {
            if (VALUE_OPTIONS.has(arg) && index + 1 < argv.length) {
                index += 1;
            }
            continue;
        }
        parts.push(arg);
    }
    return parts.join(" ") || "help";
}
export function getTopCommand(argv) {
    for (let index = 2; index < argv.length; index += 1) {
        const value = argv[index];
        if (value.startsWith("-")) {
            if (VALUE_OPTIONS.has(value) && !value.includes("=") && index + 1 < argv.length)
                index += 1;
            continue;
        }
        return value;
    }
    return undefined;
}
export function positionalArgs(argv, start) {
    const args = [];
    for (let index = start; index < argv.length; index += 1) {
        const value = argv[index];
        if (value.startsWith("-")) {
            if (VALUE_OPTIONS.has(value) && !value.includes("=") && index + 1 < argv.length) {
                index += 1;
            }
            continue;
        }
        args.push(value);
    }
    return args;
}
export function hasFlag(argv, flag) {
    return argv.some((arg) => arg === flag);
}
export function optionValue(argv, name) {
    for (let index = 0; index < argv.length; index += 1) {
        const value = argv[index];
        if (value === name) {
            return argv[index + 1];
        }
        if (value.startsWith(`${name}=`)) {
            return value.slice(name.length + 1);
        }
    }
    return undefined;
}
export function secondCommandToken(argv) {
    const top = getTopCommand(argv);
    if (!top)
        return undefined;
    const topIndex = argv.indexOf(top, 2);
    for (let index = topIndex + 1; index < argv.length; index += 1) {
        const value = argv[index];
        if (value.startsWith("-")) {
            if (VALUE_OPTIONS.has(value) && index + 1 < argv.length)
                index += 1;
            continue;
        }
        return value;
    }
    return undefined;
}
//# sourceMappingURL=argv.js.map