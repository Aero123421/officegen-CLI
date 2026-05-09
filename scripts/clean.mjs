import { rm } from "node:fs/promises";

for (const path of [
  "packages/core/dist",
  "packages/formats/dist",
  "packages/optional/dist",
  "packages/cli/dist",
  "packages/core/tsconfig.tsbuildinfo",
  "packages/formats/tsconfig.tsbuildinfo",
  "packages/optional/tsconfig.tsbuildinfo",
  "packages/cli/tsconfig.tsbuildinfo"
]) {
  await rm(path, { recursive: true, force: true });
}
