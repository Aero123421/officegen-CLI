import { rm } from "node:fs/promises";

for (const path of [
  "legacy/typescript-v3-reference/packages/core/dist",
  "legacy/typescript-v3-reference/packages/formats/dist",
  "legacy/typescript-v3-reference/packages/optional/dist",
  "legacy/typescript-v3-reference/packages/cli/dist",
  "legacy/typescript-v3-reference/packages/core/tsconfig.tsbuildinfo",
  "legacy/typescript-v3-reference/packages/formats/tsconfig.tsbuildinfo",
  "legacy/typescript-v3-reference/packages/optional/tsconfig.tsbuildinfo",
  "legacy/typescript-v3-reference/packages/cli/tsconfig.tsbuildinfo"
]) {
  await rm(path, { recursive: true, force: true });
}
