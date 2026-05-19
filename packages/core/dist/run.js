import { createHash, randomBytes } from "node:crypto";
import { appendFile, mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { computeCapabilitiesHash } from "./capabilities.js";
import { OFFICEGEN_CLI_VERSION } from "./types.js";
function safeTimestamp(date) {
    return date.toISOString().replace(/\.\d{3}Z$/, "Z").replace(/:/g, "-");
}
export function createRunId(date = new Date()) {
    return `${safeTimestamp(date)}_${randomBytes(3).toString("hex")}`;
}
export function getRunFolder(config, runId) {
    const root = path.resolve(config.paths.projectRoot, config.paths.defaultRunsDir, runId);
    return {
        runId,
        root,
        inputDir: path.join(root, "input"),
        irDir: path.join(root, "ir"),
        opsDir: path.join(root, "ops"),
        viewsDir: path.join(root, "views"),
        diagnosticsDir: path.join(root, "diagnostics"),
        outputDir: path.join(root, "output"),
        backupDir: path.join(root, "backup"),
        logsDir: path.join(root, "logs"),
        tracePath: path.join(root, "trace.jsonl"),
        runJsonPath: path.join(root, "run.json"),
        manifestPath: path.join(root, "manifest.json")
    };
}
export function createEmptyManifest(config, runId) {
    return {
        schema: "officegen.manifest@1.2",
        runId,
        cliVersion: OFFICEGEN_CLI_VERSION,
        profile: config.profile,
        capabilitiesHash: computeCapabilitiesHash(config),
        inputs: [],
        outputs: [],
        security: {
            network: config.security.network,
            externalProcess: config.security.externalProcess,
            redactedPaths: config.security.redactAbsolutePathsInJson,
            macrosPreserved: false,
            externalRelationshipsDropped: true
        },
        warnings: []
    };
}
export async function createRunFolder(config, runId = createRunId()) {
    const folder = getRunFolder(config, runId);
    await Promise.all([
        folder.root,
        folder.inputDir,
        folder.irDir,
        folder.opsDir,
        folder.viewsDir,
        folder.diagnosticsDir,
        folder.outputDir,
        folder.backupDir,
        folder.logsDir
    ].map((dir) => mkdir(dir, { recursive: true })));
    await writeFile(folder.runJsonPath, JSON.stringify({ runId, createdAt: new Date().toISOString() }, null, 2));
    await writeManifest(folder, createEmptyManifest(config, runId));
    return folder;
}
export async function writeManifest(folder, manifest) {
    await writeFile(folder.manifestPath, `${JSON.stringify(manifest, null, 2)}\n`, "utf8");
}
export async function readManifest(folder) {
    return JSON.parse(await readFile(folder.manifestPath, "utf8"));
}
export async function updateManifest(folder, updater) {
    const manifest = await readManifest(folder);
    const updated = updater(manifest) ?? manifest;
    await writeManifest(folder, updated);
    return updated;
}
export async function appendTrace(folder, record) {
    await appendFile(folder.tracePath, `${JSON.stringify(record)}\n`, "utf8");
}
export async function sha256File(filePath) {
    const hash = createHash("sha256");
    hash.update(await readFile(filePath));
    return hash.digest("hex");
}
//# sourceMappingURL=run.js.map