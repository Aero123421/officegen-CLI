import { mkdir, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { mkdtemp } from "node:fs/promises";
import { describe, expect, it } from "vitest";
import { computeCapabilitiesHash, getCapabilities } from "./capabilities.js";
import { getBuiltinConfig, loadConfig } from "./config.js";

describe("config and capabilities", () => {
  it("uses substrate defaults with authoring features enabled and plugins disabled", () => {
    const config = getBuiltinConfig("substrate");
    const capabilities = getCapabilities(config, { agent: true });

    expect(config.features.inspect.enabled).toBe(true);
    expect(config.features.template.enabled).toBe(true);
    expect(config.features.design.visibleToAgents).toBe(true);
    expect(capabilities.visibleCommands).toContain("inspect");
    expect(capabilities.visibleCommands).toContain("template candidates");
    expect(capabilities.visibleCommands).toContain("renderer doctor");
    expect(capabilities.disabled).toEqual(expect.arrayContaining(["plugin"]));
  });

  it("provides authoring and enterprise built-in profiles", () => {
    const authoring = getBuiltinConfig("authoring");
    const enterprise = getBuiltinConfig("enterprise");

    expect(authoring.features.template.visibleToAgents).toBe(true);
    expect(authoring.features.plugin.enabled).toBe(false);
    expect(enterprise.features.plugin.enabled).toBe(true);
    expect(enterprise.features.renderer.enabled).toBe(true);
  });

  it("loads user then project config, with project taking precedence", async () => {
    const root = await mkdtemp(path.join(tmpdir(), "officegen-core-"));
    const userDir = path.join(root, "user");
    const projectDir = path.join(root, "project");
    await mkdir(path.join(projectDir, ".officegen"), { recursive: true });
    await mkdir(userDir, { recursive: true });
    const userConfigPath = path.join(userDir, "config.json");
    const projectConfigPath = path.join(projectDir, ".officegen", "config.json");
    await writeFile(userConfigPath, JSON.stringify({ profile: "authoring" }));
    await writeFile(
      projectConfigPath,
      JSON.stringify({
        features: {
          design: { enabled: true, visibleInHelp: true, visibleToAgents: false }
        }
      })
    );

    const config = await loadConfig({ cwd: projectDir, userConfigPath, projectConfigPath });

    expect(config.profile).toBe("authoring");
    expect(config.features.template.enabled).toBe(true);
    expect(config.features.design.visibleInHelp).toBe(true);
    expect(config.features.design.visibleToAgents).toBe(false);
  });

  it("changes capabilities hash when feature visibility changes", () => {
    const config = getBuiltinConfig("authoring");
    const before = computeCapabilitiesHash(config);
    config.features.design.visibleToAgents = false;

    expect(computeCapabilitiesHash(config)).not.toBe(before);
  });
});
