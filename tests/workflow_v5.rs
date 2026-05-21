use serde_json::{json, Value};
use std::fs;
use std::process::{Command, Output};
use tempfile::tempdir;

fn officegen(args: &[&str], cwd: &std::path::Path) -> Output {
    Command::new(env!("CARGO_BIN_EXE_officegen"))
        .args(args)
        .current_dir(cwd)
        .output()
        .expect("officegen runs")
}

fn parse_stdout(output: &Output) -> Value {
    serde_json::from_slice(&output.stdout).expect("stdout is JSON")
}

#[test]
fn workflow_runs_sequentially_and_writes_manifest_trace_summary() {
    let dir = tempdir().unwrap();
    fs::write(
        dir.path().join("workflow.json"),
        serde_json::to_vec_pretty(&json!({
            "schema": "officegen.workflow@2.0",
            "version": "2.0",
            "outputRoot": "out",
            "steps": [
                {
                    "id": "scaffold",
                    "command": "scaffold",
                    "args": ["--kind", "docx", "--title", "Workflow Smoke", "--out", "smoke.ir.json"]
                },
                {
                    "id": "render",
                    "command": "render",
                    "input": "out/smoke.ir.json",
                    "target": "docx",
                    "out": "smoke.docx"
                },
                {
                    "id": "inspect",
                    "command": "inspect",
                    "input": "out/smoke.docx"
                }
            ]
        }))
        .unwrap(),
    )
    .unwrap();

    let output = officegen(
        &["run", "workflow.json", "--agent", "--strict-json"],
        dir.path(),
    );
    let envelope = parse_stdout(&output);

    assert!(
        output.status.success(),
        "{}",
        String::from_utf8_lossy(&output.stderr)
    );
    assert_eq!(envelope["ok"], true);
    assert_eq!(
        envelope["result"]["schema"],
        "officegen.workflow.run.result@2.0"
    );
    assert_eq!(envelope["result"]["transport"], "cli-json-file");
    assert_eq!(envelope["result"]["mcp"], false);
    assert!(dir.path().join("out").join("smoke.ir.json").exists());
    assert!(dir.path().join("out").join("smoke.docx").exists());
    assert!(dir.path().join("out").join("manifest.json").exists());
    assert!(dir.path().join("out").join("trace.json").exists());
    assert!(dir.path().join("out").join("summary.json").exists());
    assert!(envelope["result"]["artifacts"]
        .as_array()
        .unwrap()
        .iter()
        .any(|artifact| artifact["path"] == "out/smoke.docx"));

    let trace: Value =
        serde_json::from_slice(&fs::read(dir.path().join("out").join("trace.json")).unwrap())
            .unwrap();
    assert_eq!(trace["ok"], true);
    assert_eq!(trace["steps"].as_array().unwrap().len(), 3);

    let manifest: Value =
        serde_json::from_slice(&fs::read(dir.path().join("out").join("manifest.json")).unwrap())
            .unwrap();
    assert!(manifest["artifacts"]
        .as_array()
        .unwrap()
        .iter()
        .any(|artifact| artifact["path"] == "out/smoke.docx"));
}

#[test]
fn workflow_stops_on_first_failure() {
    let dir = tempdir().unwrap();
    fs::write(
        dir.path().join("workflow.json"),
        serde_json::to_vec_pretty(&json!({
            "schema": "officegen.workflow@2.0",
            "version": "2.0",
            "outputRoot": "out",
            "steps": [
                {"id": "missing", "command": "inspect", "args": ["missing.docx"]},
                {"id": "after-failure", "command": "scaffold", "args": ["--kind", "docx", "--out", "should-not-exist.json"]}
            ]
        }))
        .unwrap(),
    )
    .unwrap();

    let output = officegen(
        &["run", "workflow.json", "--agent", "--strict-json"],
        dir.path(),
    );
    let envelope = parse_stdout(&output);

    assert!(!output.status.success());
    assert_eq!(envelope["ok"], false);
    assert_eq!(envelope["result"]["status"], "fail");
    assert_eq!(envelope["result"]["failedStep"], "missing");
    assert_eq!(envelope["result"]["completedSteps"], 0);
    assert_eq!(envelope["result"]["changed"], false);
    assert_eq!(envelope["result"]["workflowArtifactsWritten"], true);
    assert!(!dir
        .path()
        .join("out")
        .join("should-not-exist.json")
        .exists());

    let trace: Value =
        serde_json::from_slice(&fs::read(dir.path().join("out").join("trace.json")).unwrap())
            .unwrap();
    assert_eq!(trace["steps"].as_array().unwrap().len(), 1);
}

#[test]
fn workflow_edit_dry_run_does_not_require_out() {
    let dir = tempdir().unwrap();
    let scaffold = officegen(
        &[
            "scaffold",
            "--kind",
            "docx",
            "--title",
            "Workflow Dry Run",
            "--out",
            "input.ir.json",
            "--agent",
            "--strict-json",
        ],
        dir.path(),
    );
    assert!(scaffold.status.success());

    let render = officegen(
        &[
            "render",
            "input.ir.json",
            "--target",
            "docx",
            "--out",
            "input.docx",
            "--agent",
            "--strict-json",
        ],
        dir.path(),
    );
    assert!(render.status.success());

    fs::write(
        dir.path().join("ops.json"),
        serde_json::to_vec_pretty(&json!({
            "schema": "officegen.edit.ops@1.2",
            "ops": [
                {
                    "op": "docx.setText",
                    "selector": {"contains": "Workflow Dry Run"},
                    "text": "Preview Only"
                }
            ]
        }))
        .unwrap(),
    )
    .unwrap();
    fs::write(
        dir.path().join("workflow.json"),
        serde_json::to_vec_pretty(&json!({
            "schema": "officegen.workflow@2.0",
            "version": "2.0",
            "outputRoot": "out",
            "steps": [
                {
                    "id": "dry-edit",
                    "command": "edit",
                    "input": "input.docx",
                    "ops": "ops.json",
                    "dryRun": true
                }
            ]
        }))
        .unwrap(),
    )
    .unwrap();

    let output = officegen(
        &["run", "workflow.json", "--agent", "--strict-json"],
        dir.path(),
    );
    let envelope = parse_stdout(&output);

    assert!(
        output.status.success(),
        "{}",
        String::from_utf8_lossy(&output.stderr)
    );
    assert_eq!(envelope["ok"], true);
    assert_eq!(envelope["mutationStatus"], "not_applicable");
    assert_eq!(envelope["result"]["changed"], false);
    assert_eq!(envelope["result"]["documentChanged"], false);
    assert_eq!(envelope["result"]["completedSteps"], 1);
    assert!(!dir.path().join("out").join("input-edited.docx").exists());
}

#[test]
fn workflow_preparation_failures_use_workflow_step_failed_contract() {
    let dir = tempdir().unwrap();
    fs::write(
        dir.path().join("workflow.json"),
        serde_json::to_vec_pretty(&json!({
            "schema": "officegen.workflow@2.0",
            "version": "2.0",
            "outputRoot": "out",
            "steps": [
                {"id": "bad-render", "command": "render", "input": "missing.ir.json"}
            ]
        }))
        .unwrap(),
    )
    .unwrap();

    let output = officegen(
        &["run", "workflow.json", "--agent", "--strict-json"],
        dir.path(),
    );
    let envelope = parse_stdout(&output);

    assert!(!output.status.success());
    assert_eq!(envelope["ok"], false);
    assert_eq!(envelope["error"]["code"], "WORKFLOW_STEP_FAILED");
    assert_eq!(envelope["result"]["failedStep"], "bad-render");
    assert!(envelope["result"]["message"]
        .as_str()
        .unwrap()
        .contains("OUTPUT_REQUIRED"));
}

#[test]
fn workflow_denies_outputs_outside_output_root() {
    let dir = tempdir().unwrap();
    fs::write(
        dir.path().join("workflow.json"),
        serde_json::to_vec_pretty(&json!({
            "schema": "officegen.workflow@2.0",
            "version": "2.0",
            "outputRoot": "out",
            "steps": [
                {"id": "escape", "command": "scaffold", "args": ["--kind", "docx", "--out", "../escape.json"]}
            ]
        }))
        .unwrap(),
    )
    .unwrap();

    let output = officegen(
        &["run", "workflow.json", "--agent", "--strict-json"],
        dir.path(),
    );
    let envelope = parse_stdout(&output);

    assert!(!output.status.success());
    assert_eq!(envelope["ok"], false);
    assert_eq!(envelope["failureClass"], "workflow");
    assert_eq!(envelope["error"]["code"], "WORKFLOW_STEP_FAILED");
    assert!(envelope["result"]["message"]
        .as_str()
        .unwrap()
        .contains("SECURITY_PATH_OUTSIDE_ROOT"));
    assert!(!dir.path().join("escape.json").exists());
    assert!(dir.path().join("out").join("trace.json").exists());
}
