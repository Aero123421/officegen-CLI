import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";

const replacements = [
  {
    file: "packages/cli/dist/program.js",
    from: /"@officegen\/core"/g,
    to: "\"../../core/dist/index.js\""
  },
  {
    file: "packages/cli/dist/commands/register.js",
    from: /"@officegen\/core"/g,
    to: "\"../../../core/dist/index.js\""
  },
  {
    file: "packages/cli/dist/commands/payloads.js",
    from: /"@officegen\/core"/g,
    to: "\"../../../core/dist/index.js\""
  },
  {
    file: "packages/cli/dist/commands/payloads.js",
    from: /"@officegen\/formats"/g,
    to: "\"../../../formats/dist/index.js\""
  },
  {
    file: "packages/cli/dist/commands/payloads.js",
    from: /"@officegen\/optional"/g,
    to: "\"../../../optional/dist/index.js\""
  },
  {
    file: "packages/cli/dist/shared/context.js",
    from: /"@officegen\/core"/g,
    to: "\"../../../core/dist/index.js\""
  },
  {
    file: "packages/cli/dist/shared/envelope.js",
    from: /"@officegen\/core"/g,
    to: "\"../../../core/dist/index.js\""
  },
  {
    file: "packages/cli/dist/shared/io.js",
    from: /"@officegen\/core"/g,
    to: "\"../../../core/dist/index.js\""
  },
  {
    file: "packages/cli/dist/shared/io.js",
    from: /"@officegen\/optional"/g,
    to: "\"../../../optional/dist/index.js\""
  },
  {
    file: "packages/formats/dist/shared.js",
    from: /"@officegen\/core"/g,
    to: "\"../../core/dist/index.js\""
  },
  {
    file: "packages/optional/dist/design.js",
    from: /"@officegen\/formats"/g,
    to: "\"../../formats/dist/index.js\""
  }
];

for (const replacement of replacements) {
  const filePath = path.resolve(replacement.file);
  const before = await readFile(filePath, "utf8");
  const after = before.replace(replacement.from, replacement.to);
  if (after !== before) await writeFile(filePath, after, "utf8");
}
