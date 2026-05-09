import { link, mkdir, readFile, symlink, writeFile } from "node:fs/promises";
import { mkdtemp } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { getBuiltinConfig } from "./config.js";
import { validatePath } from "./paths.js";
import { createRunFolder, updateManifest } from "./run.js";

async function tempProject() {
  const root = await mkdtemp(path.join(tmpdir(), "officegen-core-"));
  await mkdir(path.join(root, ".officegen"), { recursive: true });
  const config = getBuiltinConfig("substrate");
  config.paths.projectRoot = root;
  config.paths.projectConfigDir = path.join(root, ".officegen");
  config.paths.userConfigDir = path.join(root, "user");
  config.security.trustedRoots = [".", ".officegen", config.paths.userConfigDir];
  return { root, config };
}

describe("path policy and run folders", () => {
  it("denies absolute output paths by default", async () => {
    const { root, config } = await tempProject();

    await expect(validatePath(config, { kind: "output", path: path.join(root, "out.pptx") })).rejects.toMatchObject({
      payload: { code: "SECURITY_ABSOLUTE_OUT_DENIED" }
    });
  });

  it("denies out-of-project traversal", async () => {
    const { config } = await tempProject();

    await expect(validatePath(config, { kind: "output", path: "../outside.pptx", allowAbsoluteOut: true })).rejects.toMatchObject({
      payload: { code: "SECURITY_PATH_OUTSIDE_ROOT" }
    });
  });

  it("requires explicit overwrite for existing outputs", async () => {
    const { root, config } = await tempProject();
    await writeFile(path.join(root, "out.pptx"), "old");

    await expect(validatePath(config, { kind: "output", path: "out.pptx" })).rejects.toMatchObject({
      payload: { code: "EDIT_TRANSACTION_FAILED" }
    });
    await expect(validatePath(config, { kind: "output", path: "out.pptx", overwrite: true })).resolves.toMatchObject({
      existed: true
    });
  });

  it("denies hardlinked outputs best-effort", async () => {
    const { root, config } = await tempProject();
    const original = path.join(root, "original.txt");
    const linked = path.join(root, "linked.txt");
    await writeFile(original, "same inode");
    await link(original, linked);

    await expect(validatePath(config, { kind: "output", path: "linked.txt", overwrite: true })).rejects.toMatchObject({
      payload: { code: "SECURITY_HARDLINK_DENIED" }
    });
  });

  it("denies symlink outputs when the platform allows creating one", async () => {
    const { root, config } = await tempProject();
    const target = path.join(root, "target.txt");
    const linkPath = path.join(root, "link.txt");
    await writeFile(target, "target");
    try {
      await symlink(target, linkPath);
    } catch {
      return;
    }

    await expect(validatePath(config, { kind: "output", path: "link.txt", overwrite: true })).rejects.toMatchObject({
      payload: { code: "SECURITY_SYMLINK_DENIED" }
    });
  });

  it("creates run folders and updates manifest", async () => {
    const { config } = await tempProject();
    const folder = await createRunFolder(config, "2026-05-09T12-34-56Z_ab12cd");
    const manifestRaw = await readFile(folder.manifestPath, "utf8");

    expect(manifestRaw).toContain("officegen.manifest@1.2");
    await updateManifest(folder, (manifest) => {
      manifest.inputs.push({ path: "<project>/source.pptx", trusted: false });
    });
    expect(await readFile(folder.manifestPath, "utf8")).toContain("<project>/source.pptx");
  });
});
