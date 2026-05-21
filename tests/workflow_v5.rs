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

    let trace: Value =
        serde_json::from_slice(&fs::read(dir.path().join("out").join("trace.json")).unwrap())
            .unwrap();
    assert_eq!(trace["ok"], true);
    assert_eq!(trace["steps"].as_array().unwrap().len(), 3);
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
    assert_eq!(envelope["failureClass"], "security");
    assert!(!dir.path().join("escape.json").exists());
    assert!(dir.path().join("out").join("trace.json").exists());
}
