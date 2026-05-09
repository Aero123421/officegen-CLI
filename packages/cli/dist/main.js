#!/usr/bin/env node
import { runCli } from "./program.js";
runCli(process.argv).catch((error) => {
    const message = error instanceof Error ? error.message : String(error);
    console.error(message);
    process.exitCode = 1;
});
//# sourceMappingURL=main.js.map