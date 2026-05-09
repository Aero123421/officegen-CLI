import { rm } from "node:fs/promises";

for (const path of ["packages/core/dist", "packages/formats/dist", "packages/optional/dist", "packages/cli/dist"]) {
  await rm(path, { recursive: true, force: true });
}
