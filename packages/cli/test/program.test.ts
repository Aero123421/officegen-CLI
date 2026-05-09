import { mkdir, mkdtemp, readFile, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { afterEach, describe, expect, it } from "vitest";
import { runCli } from "../src/program.js";

interface Captured {
  stdout: string[];
  stderr: string[];
}

async function run(args: string[], cwd = process.cwd(), env: NodeJS.ProcessEnv = {}): Promise<Captured> {
  const captured: Captured = { stdout: [], stderr: [] };
  process.exitCode = undefined;
  await runCli(["node", "officegen", ...args], {
    cwd,
    now: new Date("2026-05-09T00:00:00.000Z"),
    stdout: (text) => captured.stdout.push(text),
    stderr: (text) => captured.stderr.push(text),
    env
  });
  return captured;
}

async function tempWorkspace(config?: unknown): Promise<string> {
  const unique = await mkdtemp(path.join(os.tmpdir(), "officegen-cli-test-"));
  if (config) {
    const configDir = path.join(unique, ".officegen");
    await mkdir(configDir, { recursive: true });
    await writeFile(path.join(configDir, "config.json"), `${JSON.stringify(config, null, 2)}\n`, "utf8");
  }
  return unique;
}

function parseEnvelope(captured: Captured): any {
  return JSON.parse(captured.stdout[0] ?? captured.stderr[0]);
}

afterEach(() => {
  process.exitCode = undefined;
});

describe("officegen CLI command surface", () => {
  it("wraps capabilities --agent --json in the v1.2 envelope and hides disabled optional commands", async () => {
    const captured = await run(["capabilities", "--agent", "--json"]);
    const envelope = parseEnvelope(captured);

    expect(envelope.schema).toBe("officegen.envelope@1.2");
    expect(envelope.ok).toBe(true);
    expect(envelope.result.schema).toBe("officegen.capabilities@1.2");
    expect(envelope.result.enabled).toContain("inspect");
    expect(envelope.result.disabled).toContain("template");
    expect(envelope.availableCommands).toContain("inspect");
    expect(envelope.availableCommands).not.toContain("template");
    expect(envelope.nextSuggestedCommands).toContain("officegen capabilities --agent --json");
  });

  it("returns availableCommands and nextSuggestedCommands for unknown commands", async () => {
    const captured = await run(["create", "--json"]);
    const envelope = parseEnvelope(captured);

    expect(envelope.ok).toBe(false);
    expect(envelope.error.code).toBe("UNKNOWN_COMMAND");
    expect(envelope.availableCommands).toContain("inspect");
    expect(envelope.nextSuggestedCommands).toContain("officegen capabilities --json");
    expect(process.exitCode).toBe(2);
  });

  it("distinguishes agent-hidden commands from disabled commands", async () => {
    const cwd = await tempWorkspace({
      profile: "authoring",
      features: {
        design: {
          enabled: true,
          visibleInHelp: true,
          visibleToAgents: false
        }
      }
    });

    const captured = await run(["design", "--agent", "--json"], cwd);
    const envelope = parseEnvelope(captured);

    expect(envelope.ok).toBe(false);
    expect(envelope.error.code).toBe("FEATURE_HIDDEN");
    expect(envelope.error.feature).toBe("design");
    expect(envelope.availableCommands).not.toContain("design");
    expect(envelope.nextSuggestedCommands).toContain("officegen capabilities --agent --json");
    expect(process.exitCode).toBe(5);
  });

  it("registers every v1.2 top-level command when enabled by enterprise profile", async () => {
    const cwd = await tempWorkspace({ profile: "enterprise" });
    const commands = [
      "capabilities",
      "help",
      "config",
      "doctor",
      "schema",
      "errors",
      "inspect",
      "view",
      "edit",
      "render",
      "scaffold",
      "export",
      "validate",
      "diagnose",
      "repair",
      "run",
      "asset",
      "chart",
      "diagram",
      "template",
      "design",
      "layout",
      "agent",
      "mcp",
      "renderer",
      "plugin"
    ];

    const captured = await run(["capabilities", "--json"], cwd);
    const envelope = parseEnvelope(captured);
    expect(envelope.ok).toBe(true);
    for (const command of commands) {
      expect(envelope.availableCommands, command).toContain(command);
    }
  });

  it("writes scaffold IR while still returning a JSON envelope", async () => {
    const cwd = await tempWorkspace();
    const captured = await run(["scaffold", "--kind", "pptx", "--title", "Proposal", "--out", "deck.ir.json", "--json"], cwd);
    const envelope = parseEnvelope(captured);
    const written = JSON.parse(await readFile(path.join(cwd, "deck.ir.json"), "utf8"));

    expect(envelope.ok).toBe(true);
    expect(envelope.result.document.schema).toBe("officegen.ir.document@1.2");
    expect(written.metadata.title).toBe("Proposal");
  });

  it("rejects unknown group subcommands", async () => {
    const captured = await run(["asset", "delete", "--json"]);
    const envelope = parseEnvelope(captured);

    expect(envelope.ok).toBe(false);
    expect(envelope.error.code).toBe("UNKNOWN_COMMAND");
    expect(process.exitCode).toBe(2);
  });

  it("denies out-of-project scaffold output by default", async () => {
    const cwd = await tempWorkspace();
    const captured = await run(["scaffold", "--out", "../outside.ir.json", "--json"], cwd);
    const envelope = parseEnvelope(captured);

    expect(envelope.ok).toBe(false);
    expect(envelope.error.code).toBe("SECURITY_PATH_OUTSIDE_ROOT");
    expect(process.exitCode).toBe(4);
  });

  it("hides optional schemas from agents", async () => {
    const captured = await run(["schema", "get", "officegen.template.map@1.2", "--agent", "--json"]);
    const envelope = parseEnvelope(captured);

    expect(envelope.ok).toBe(false);
    expect(envelope.error.code).toBe("FEATURE_HIDDEN_FROM_AGENT");
    expect(process.exitCode).toBe(5);
  });

  it("prints native help with profile-aware commands and exits successfully", async () => {
    const cwd = await tempWorkspace({ profile: "substrate" });
    const captured = await run(["--help"], cwd);

    expect(captured.stdout[0]).toContain("officegen - AI-friendly Office/PDF runtime");
    expect(captured.stdout[0]).toContain("inspect");
    expect(captured.stdout[0]).not.toContain("template");
    expect(process.exitCode).toBeUndefined();
  });

  it("warns when the supplied capabilities hash is stale", async () => {
    const captured = await run(["capabilities", "--json", "--capabilities-hash", "sha256:stale"]);
    const envelope = parseEnvelope(captured);

    expect(envelope.ok).toBe(true);
    expect(envelope.warnings).toEqual(expect.arrayContaining([
      expect.objectContaining({
        code: "AGENT_ADAPTER_STALE",
        expected: "sha256:stale"
      })
    ]));
  });

  it("also accepts the expected capabilities hash from the environment", async () => {
    const captured = await run(["capabilities", "--json"], process.cwd(), { OFFICEGEN_CAPABILITIES_HASH: "sha256:env-stale" });
    const envelope = parseEnvelope(captured);

    expect(envelope.warnings).toEqual(expect.arrayContaining([
      expect.objectContaining({
        code: "AGENT_ADAPTER_STALE",
        expected: "sha256:env-stale"
      })
    ]));
  });

  it("truncates agent JSON output when an explicit budget is exceeded", async () => {
    const captured = await run(["capabilities", "--agent", "--json", "--json-budget-bytes", "512"]);
    const envelope = parseEnvelope(captured);

    expect(envelope.ok).toBe(true);
    expect(envelope.result.schema).toBe("officegen.progressive-disclosure@1.2");
    expect(envelope.warnings).toEqual(expect.arrayContaining([
      expect.objectContaining({ code: "AGENT_JSON_BUDGET_EXCEEDED" })
    ]));
  });
});
