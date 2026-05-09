import { readdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";

const root = process.cwd();
const distRoots = ["packages/cli/dist", "packages/formats/dist", "packages/optional/dist"];
const targets = {
  "@officegen/core": path.resolve(root, "packages/core/dist/index.js"),
  "@officegen/formats": path.resolve(root, "packages/formats/dist/index.js"),
  "@officegen/optional": path.resolve(root, "packages/optional/dist/index.js")
};

for (const distRoot of distRoots) {
  for (const file of await listJsFiles(path.resolve(root, distRoot))) {
    const before = await readFile(file, "utf8");
    let after = before;
    for (const [specifier, target] of Object.entries(targets)) {
      after = after.replaceAll(JSON.stringify(specifier), JSON.stringify(relativeModuleSpecifier(path.dirname(file), target)));
    }
    if (after !== before) await writeFile(file, after, "utf8");
  }
}

async function listJsFiles(directory) {
  const entries = await readdir(directory, { withFileTypes: true });
  const files = [];
  for (const entry of entries) {
    const fullPath = path.join(directory, entry.name);
    if (entry.isDirectory()) {
      files.push(...await listJsFiles(fullPath));
    } else if (entry.isFile() && entry.name.endsWith(".js") && !entry.name.endsWith(".test.js")) {
      files.push(fullPath);
    }
  }
  return files;
}

function relativeModuleSpecifier(fromDir, target) {
  const relative = path.relative(fromDir, target).replace(/\\/g, "/");
  return relative.startsWith(".") ? relative : `./${relative}`;
}
