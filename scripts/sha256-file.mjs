import { createHash } from "node:crypto";
import { createReadStream } from "node:fs";
import path from "node:path";

const filePath = process.argv[2];
if (!filePath) {
  console.error("usage: node scripts/sha256-file.mjs <file>");
  process.exit(2);
}

const hash = createHash("sha256");
const stream = createReadStream(filePath);

stream.on("data", (chunk) => hash.update(chunk));
stream.on("error", (error) => {
  console.error(error instanceof Error ? error.message : String(error));
  process.exit(1);
});
stream.on("end", () => {
  console.log(`${hash.digest("hex")}  ${path.basename(filePath)}`);
});
