import { readdir, readFile } from "node:fs/promises";
import path from "node:path";
import process from "node:process";

const root = process.cwd();
const failures = [];

const schemasPath = path.join(root, "packages/core/src/schemas.ts");
const editPath = path.join(root, "packages/formats/src/edit.ts");
const schemasSource = await readFile(schemasPath, "utf8");
const editSource = await readFile(editPath, "utf8");

const registryIds = new Set([...schemasSource.matchAll(/entry\("([^"]+)"/g)].map((match) => match[1]));
const schemaEditOps = new Set([...schemasSource.matchAll(/op\("([^"]+)"/g)].map((match) => match[1]));
const implementedEditOps = readImplementedEditOps(editSource);
const sourceSchemaIds = await readRuntimeSchemaIds();
const schemaExcludedUnsupportedEditOps = new Set([
  "pdf.redact"
]);

const requiredResultSchemas = [
  "officegen.capabilities@1.2",
  "officegen.help@1.2",
  "officegen.config@1.2",
  "officegen.doctor@1.2",
  "officegen.schema.list@1.2",
  "officegen.schema.definition@1.2",
  "officegen.validation.result@1.2",
  "officegen.schema.migration.result@1.2",
  "officegen.template.candidates.result@2.5",
  "officegen.errors@1.2",
  "officegen.error@1.2",
  "officegen.progressive-disclosure@1.2",
  "officegen.inspect.result@1.2",
  "officegen.view.result@1.2",
  "officegen.edit.result@1.2",
  "officegen.render.result@1.2",
  "officegen.export.result@1.2",
  "officegen.diagnose.result@1.2",
  "officegen.verify.result@1.2",
  "officegen.repair.result@1.2",
  "officegen.diff.result@1.2",
  "officegen.scaffold.result@1.2",
  "officegen.chart.render.result@1.2",
  "officegen.diagram.render.result@1.2",
  "officegen.asset.info@1.2",
  "officegen.asset.embedded.info@2.5",
  "officegen.asset.embedded.result@2.5",
  "officegen.asset.embedded.trusted@2.5",
  "officegen.asset.embedded.untrusted@2.5",
  "officegen.asset.extract.result@1.2",
  "officegen.asset.replace.result@1.2",
  "officegen.critique.result@2.3",
  "officegen.improve.plan@2.5",
  "officegen.improve.plan@2.3",
  "officegen.benchmark.run.result@2.5",
  "officegen.benchmark.run.result@2.3",
  "officegen.benchmark.compare.result@2.3",
  "officegen.run.manifest@2.3",
  "officegen.run.result@2.3",
  "officegen.mcp.tools@1.2",
  "officegen.renderer.doctor@2.2"
];

for (const id of requiredResultSchemas) {
  if (!registryIds.has(id)) failures.push(`representative result schema is not registered: ${id}`);
}

for (const id of sourceSchemaIds) {
  if (!registryIds.has(id)) failures.push(`runtime schema literal is not registered: ${id}`);
}

for (const op of implementedEditOps) {
  if (schemaExcludedUnsupportedEditOps.has(op)) continue;
  if (!schemaEditOps.has(op)) failures.push(`EditOperation is implemented but missing from officegen.edit.ops@1.2: ${op}`);
}

for (const op of schemaExcludedUnsupportedEditOps) {
  if (schemaEditOps.has(op)) failures.push(`unsupported EditOperation must stay out of officegen.edit.ops@1.2: ${op}`);
  if (!implementedEditOps.has(op)) failures.push(`schema coverage unsupported exclusion no longer matches an implemented EditOperation: ${op}`);
}

if (failures.length) {
  console.error("schema:coverage failed:");
  for (const failure of failures) console.error(`- ${failure}`);
  process.exit(1);
}

console.log(`schema:coverage ok (${registryIds.size} registered schemas, ${schemaEditOps.size} edit ops)`);

function readImplementedEditOps(source) {
  const start = source.indexOf("export type EditOperation");
  const end = source.indexOf("export interface EditResult", start);
  const typeBlock = source.slice(start, end > start ? end : undefined);
  return new Set([...typeBlock.matchAll(/op:\s*"([^"]+)"/g)].map((match) => match[1]));
}

async function readRuntimeSchemaIds() {
  const files = [
    path.join(root, "packages/cli/src/commands/payloads.ts"),
    path.join(root, "packages/cli/src/shared/envelope.ts"),
    ...(await sourceFiles(path.join(root, "packages/formats/src"))),
    ...(await sourceFiles(path.join(root, "packages/optional/src")))
  ];
  const ids = new Set();
  for (const file of files) {
    const source = await readFile(file, "utf8");
    for (const match of source.matchAll(/officegen\.[A-Za-z0-9_.-]+@[0-9.]+/g)) {
      ids.add(match[0].replace(/[.,;:]+$/, ""));
    }
  }
  return ids;
}

async function sourceFiles(dir) {
  const entries = await readdir(dir, { withFileTypes: true });
  const files = [];
  for (const entry of entries) {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) files.push(...await sourceFiles(full));
    else if (entry.isFile() && full.endsWith(".ts")) files.push(full);
  }
  return files;
}
