import { fileURLToPath } from "node:url";
import { defineConfig } from "vitest/config";

export default defineConfig({
  resolve: {
    alias: {
      "@officegen/core": fileURLToPath(new URL("./packages/core/src/index.ts", import.meta.url)),
      "@officegen/formats": fileURLToPath(new URL("./packages/formats/src/index.ts", import.meta.url)),
      "@officegen/optional": fileURLToPath(new URL("./packages/optional/src/index.ts", import.meta.url))
    }
  },
  test: {
    include: ["packages/**/*.test.ts"],
    testTimeout: 20000
  }
});
