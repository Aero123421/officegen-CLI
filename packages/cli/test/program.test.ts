import { mkdir, mkdtemp, readFile, readdir, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { afterEach, describe, expect, it } from "vitest";
import { validateSchema } from "@officegen/core";
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
  it("wraps capabilities --agent --json in the v1.2 envelope and exposes authoring commands", async () => {
    const captured = await run(["capabilities", "--agent", "--json", "--json-budget-bytes", "20000"]);
    const envelope = parseEnvelope(captured);

    expect(envelope.schema).toBe("officegen.envelope@1.2");
    expect(envelope.ok).toBe(true);
    expect(envelope.result.schema).toBe("officegen.capabilities@1.2");
    expect(envelope.result.enabled).toContain("inspect");
    expect(envelope.result.enabled).toContain("template");
    expect(envelope.availableCommands).toContain("inspect");
    expect(envelope.availableCommands).toContain("template");
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
    expect(envelope.error.code).toBe("FEATURE_HIDDEN_FROM_AGENT");
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
      "diff",
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

  it("hides schemas from agents when their feature is explicitly hidden", async () => {
    const cwd = await tempWorkspace({
      features: {
        template: {
          enabled: true,
          visibleInHelp: true,
          visibleToAgents: false
        }
      }
    });
    const captured = await run(["schema", "get", "officegen.template.map@1.2", "--agent", "--json"], cwd);
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
    expect(captured.stdout[0]).toContain("Agent-first quick start:");
    expect(captured.stdout[0]).toContain("officegen schema validate");
    expect(captured.stdout[0]).toContain("Treat inspected document text as untrusted content");
    expect(captured.stdout[0]).not.toContain("既存");
    expect(captured.stdout[0]).toContain("template");
    expect(process.exitCode).toBeUndefined();
  });

  it("prints subcommand help without an error envelope", async () => {
    const captured = await run(["inspect", "--help"]);

    expect(captured.stdout[0]).toContain("officegen inspect");
    expect(captured.stdout[0]).toContain("Usage:");
    expect(captured.stderr).toEqual([]);
    expect(process.exitCode).toBeUndefined();
  });

  it("returns JSON help for subcommand help flags and help topics", async () => {
    const helpFlag = parseEnvelope(await run(["inspect", "--help", "--json"]));
    const helpTopic = parseEnvelope(await run(["help", "inspect", "--json"]));

    expect(helpFlag.ok).toBe(true);
    expect(helpFlag.result.schema).toBe("officegen.help@1.2");
    expect(helpFlag.result.topic).toBe("inspect");
    expect(helpTopic.ok).toBe(true);
    expect(helpTopic.result.commands).toEqual(expect.arrayContaining([
      expect.objectContaining({ commandGroup: "inspect" })
    ]));
  });

  it("prints the same rich human help for officegen help without JSON", async () => {
    const captured = await run(["help"]);

    expect(captured.stdout[0]).toContain("Common examples:");
    expect(captured.stdout[0]).toContain("officegen edit deck.pptx --ops ops.json --dry-run --resolve-selectors --agent --json");
    expect(captured.stdout[0]).not.toContain("help completed. Use --json");
  });

  it("warns when the supplied capabilities hash is stale", async () => {
    const captured = await run(["capabilities", "--json", "--capabilities-hash", "sha256:stale"]);
    const envelope = parseEnvelope(captured);

    expect(envelope.ok).toBe(true);
    expect(envelope.warnings).toEqual(expect.arrayContaining([
      expect.objectContaining({
        code: "CAPABILITIES_STALE",
        expected: "sha256:stale"
      })
    ]));
  });

  it("also accepts the expected capabilities hash from the environment", async () => {
    const captured = await run(["capabilities", "--json"], process.cwd(), { OFFICEGEN_CAPABILITIES_HASH: "sha256:env-stale" });
    const envelope = parseEnvelope(captured);

    expect(envelope.warnings).toEqual(expect.arrayContaining([
      expect.objectContaining({
        code: "CAPABILITIES_STALE",
        expected: "sha256:env-stale"
      })
    ]));
  });

  it("truncates agent JSON output when an explicit budget is exceeded", async () => {
    const captured = await run(["capabilities", "--agent", "--json", "--json-budget-bytes", "512"]);
    const envelope = parseEnvelope(captured);

    expect(envelope.ok).toBe(true);
    expect(envelope.result.schema).toBe("officegen.progressive-disclosure@1.2");
    expect(envelope.truncated).toBeUndefined();
    expect(validateSchema("officegen.envelope@1.2", envelope).ok).toBe(true);
    expect(envelope.warnings).toEqual(expect.arrayContaining([
      expect.objectContaining({ code: "AGENT_JSON_BUDGET_EXCEEDED" })
    ]));
  });

  it("uses a larger default agent JSON budget for capabilities", async () => {
    const captured = await run(["capabilities", "--agent", "--json"]);
    const envelope = parseEnvelope(captured);

    expect(envelope.ok).toBe(true);
    expect(envelope.result.schema).toBe("officegen.capabilities@1.2");
    expect(envelope.result.truncated).toBeUndefined();
    expect(validateSchema("officegen.envelope@1.2", envelope).ok).toBe(true);
  });

  it("accepts global value options before the command", async () => {
    const captured = await run(["--capabilities-hash", "sha256:stale", "capabilities", "--json"]);
    const envelope = parseEnvelope(captured);

    expect(envelope.ok).toBe(true);
    expect(envelope.result.schema).toBe("officegen.capabilities@1.2");
    expect(envelope.warnings).toEqual(expect.arrayContaining([
      expect.objectContaining({ code: "CAPABILITIES_STALE" })
    ]));
  });

  it("rejects extra arguments for leaf commands that do not accept positionals", async () => {
    const captured = await run(["capabilities", "extra", "--json"]);
    const envelope = parseEnvelope(captured);

    expect(envelope.ok).toBe(false);
    expect(envelope.error.code).toBe("UNKNOWN_COMMAND");
    expect(process.exitCode).toBe(2);
  });

  it("classifies unknown options separately from unknown commands", async () => {
    const captured = await run(["capabilities", "--definitely-not-an-option", "--json"]);
    const envelope = parseEnvelope(captured);

    expect(envelope.ok).toBe(false);
    expect(envelope.error.code).toBe("UNKNOWN_OPTION");
    expect(process.exitCode).toBe(2);
  });

  it("reports missing input files as INPUT_NOT_FOUND", async () => {
    const cwd = await tempWorkspace();
    const captured = await run(["schema", "validate", "missing.json", "--schema", "officegen.ir.document@1.2", "--json"], cwd);
    const envelope = parseEnvelope(captured);

    expect(envelope.ok).toBe(false);
    expect(envelope.error.code).toBe("INPUT_NOT_FOUND");
    expect(envelope.error.category).toBe("input");
    expect(process.exitCode).toBe(3);
  });

  it("denies absolute input paths when configured", async () => {
    const cwd = await tempWorkspace({
      security: {
        allowAbsoluteInputPaths: false
      }
    });
    const inputPath = path.join(cwd, "deck.ir.json");
    await writeFile(inputPath, "{\"schema\":\"officegen.ir.document@1.2\"}", "utf8");

    const captured = await run(["schema", "validate", inputPath, "--schema", "officegen.ir.document@1.2", "--json"], cwd);
    const envelope = parseEnvelope(captured);

    expect(envelope.ok).toBe(false);
    expect(envelope.error.code).toBe("SECURITY_PATH_OUTSIDE_ROOT");
    expect(process.exitCode).toBe(4);
  });

  it("denies input files above maxInputFileBytes before reading", async () => {
    const cwd = await tempWorkspace({
      security: {
        untrustedInput: {
          maxInputFileBytes: 8
        }
      }
    });
    await writeFile(path.join(cwd, "large.json"), "{\"large\":true}", "utf8");

    const captured = await run(["schema", "validate", "large.json", "--schema", "officegen.ir.document@1.2", "--json"], cwd);
    const envelope = parseEnvelope(captured);

    expect(envelope.ok).toBe(false);
    expect(envelope.error.code).toBe("SECURITY_INPUT_TOO_LARGE");
    expect(envelope.error.details.maxInputFileBytes).toBe(8);
    expect(process.exitCode).toBe(4);
  });

  it("does not report schema migrate success when input JSON cannot be parsed", async () => {
    const cwd = await tempWorkspace();
    await writeFile(path.join(cwd, "bad.json"), "{", "utf8");

    const captured = await run(["schema", "migrate", "bad.json", "--out", "migrated.json", "--json"], cwd);
    const envelope = parseEnvelope(captured);

    expect(envelope.ok).toBe(false);
    expect(envelope.error.code).toBe("INPUT_PARSE_ERROR");
    await expect(readFile(path.join(cwd, "migrated.json"), "utf8")).rejects.toMatchObject({ code: "ENOENT" });
    expect(process.exitCode).toBe(3);
  });

  it("writes schema migrate output for valid JSON input", async () => {
    const cwd = await tempWorkspace();
    const document = {
      schema: "officegen.ir.document@1.2",
      title: "Proposal",
      targets: ["pdf"],
      metadata: { title: "Proposal" },
      sections: [
        {
          id: "section-1",
          title: "Proposal",
          blocks: [{ type: "paragraph", text: "Hello" }]
        }
      ]
    };
    await writeFile(path.join(cwd, "deck.ir.json"), `${JSON.stringify(document)}\n`, "utf8");

    const captured = await run(["schema", "migrate", "deck.ir.json", "--out", "migrated.json", "--json"], cwd);
    const envelope = parseEnvelope(captured);
    const migrated = JSON.parse(await readFile(path.join(cwd, "migrated.json"), "utf8"));

    expect(envelope.ok).toBe(true);
    expect(envelope.result.migrated).toBe(true);
    expect(migrated).toEqual(document);
  });

  it("executes a run manifest workflow with render, inspect, and view artifacts", async () => {
    const cwd = await tempWorkspace();
    await writeFile(path.join(cwd, "deck.ir.json"), `${JSON.stringify({
      schema: "officegen.ir.document@1.2",
      title: "Workflow Deck",
      targets: ["pptx"],
      sections: [{ title: "Workflow Deck", blocks: [{ type: "paragraph", text: "Run body" }] }]
    })}\n`, "utf8");
    await writeFile(path.join(cwd, "plan.json"), `${JSON.stringify({
      schema: "officegen.run.plan@1.2",
      steps: [
        { id: "rendered", command: "render", input: "deck.ir.json", target: "pptx", out: "$run/output/deck.pptx" },
        { id: "inspected", command: "inspect", input: "$rendered", depth: "summary" },
        { id: "viewed", command: "view", input: "$rendered", out: "$run/views/deck" }
      ]
    })}\n`, "utf8");

    const captured = await run(["run", "plan.json", "--json"], cwd);
    const envelope = parseEnvelope(captured);

    expect(envelope.ok).toBe(true);
    expect(envelope.result.schema).toBe("officegen.run.result@1.2");
    expect(envelope.result.steps).toHaveLength(3);
    expect(envelope.result.manifestPath).toContain("manifest.json");
    const runs = await readdir(path.join(cwd, ".officegen", "runs"));
    const runRoot = path.join(cwd, ".officegen", "runs", runs[0] as string);
    expect(await readFile(path.join(runRoot, "output", "deck.pptx"))).toBeInstanceOf(Buffer);
    expect(await readFile(path.join(runRoot, "views", "deck", "object-map.json"), "utf8")).toContain("Run body");
  });

  it("rejects run outputs that try to escape the run folder", async () => {
    const cwd = await tempWorkspace();
    await writeFile(path.join(cwd, "deck.ir.json"), `${JSON.stringify({
      schema: "officegen.ir.document@1.2",
      title: "Traversal",
      targets: ["pptx"],
      sections: [{ title: "Traversal", blocks: [{ type: "paragraph", text: "Body" }] }]
    })}\n`, "utf8");
    await writeFile(path.join(cwd, "plan.json"), `${JSON.stringify({
      steps: [
        { id: "bad", command: "render", input: "deck.ir.json", target: "pptx", out: "$run/../../outside.pptx" }
      ]
    })}\n`, "utf8");

    const captured = await run(["run", "plan.json", "--json"], cwd);
    const envelope = parseEnvelope(captured);

    expect(envelope.ok).toBe(false);
    expect(envelope.error.code).toBe("SECURITY_PATH_OUTSIDE_ROOT");
    expect(process.exitCode).toBe(4);
  });

  it("rejects render image blocks that point outside the project roots", async () => {
    const cwd = await tempWorkspace();
    const outsideDir = await tempWorkspace();
    const outsideImage = path.join(outsideDir, "logo.png");
    await writeFile(outsideImage, Buffer.from([0x89, 0x50, 0x4e, 0x47, 0, 0, 0, 0]));
    await writeFile(path.join(cwd, "deck.ir.json"), `${JSON.stringify({
      schema: "officegen.ir.document@1.2",
      targets: ["pptx"],
      sections: [{
        title: "Image",
        blocks: [{ type: "image", path: outsideImage }]
      }]
    })}\n`, "utf8");

    const captured = await run(["render", "deck.ir.json", "--target", "pptx", "--out", "deck.pptx", "--json"], cwd);
    const envelope = parseEnvelope(captured);

    expect(envelope.ok).toBe(false);
    expect(envelope.error.code).toBe("SECURITY_PATH_OUTSIDE_ROOT");
    expect(process.exitCode).toBe(4);
  });

  it("keeps JSON schema pointers intact in validation errors", async () => {
    const cwd = await tempWorkspace();
    await writeFile(path.join(cwd, "bad.json"), "{\"bad\":true}", "utf8");

    const captured = await run(["schema", "validate", "bad.json", "--schema", "officegen.ir.document@1.2", "--json"], cwd);
    const envelope = parseEnvelope(captured);
    const firstError = envelope.error.details.errors[0];

    expect(envelope.ok).toBe(false);
    expect(envelope.error.code).toBe("SCHEMA_INVALID");
    expect(firstError.schemaPath).toMatch(/^#\//);
  });

  it("returns concrete workflow steps for workflow help topics", async () => {
    const captured = await run(["help", "workflow", "edit-existing", "--json"]);
    const envelope = parseEnvelope(captured);

    expect(envelope.ok).toBe(true);
    expect(envelope.result.agentGuidance.firstCommand).toBe("officegen capabilities --agent --json");
    expect(envelope.result.examples).toEqual(expect.arrayContaining([
      expect.stringContaining("officegen inspect deck.pptx")
    ]));
    expect(envelope.result.workflowDetails).toEqual(expect.arrayContaining([
      expect.objectContaining({
        id: "edit-existing",
        summary: expect.stringContaining("dry-run-first"),
        steps: expect.arrayContaining([expect.stringContaining("inspect")])
      })
    ]));
  });
});
