use serde_json::Value;
use std::fs;
use std::process::Command;
use tempfile::tempdir;

fn raw_officegen(args: &[&str], cwd: &std::path::Path) -> std::process::Output {
    Command::new(env!("CARGO_BIN_EXE_officegen"))
        .args(args)
        .current_dir(cwd)
        .output()
        .expect("officegen command should run")
}

fn officegen(args: &[&str], cwd: &std::path::Path) -> Value {
    let output = raw_officegen(args, cwd);
    assert!(
        output.status.success(),
        "command failed\nstdout:\n{}\nstderr:\n{}",
        String::from_utf8_lossy(&output.stdout),
        String::from_utf8_lossy(&output.stderr)
    );
    serde_json::from_slice(&output.stdout).expect("stdout should be JSON")
}

#[test]
fn strict_json_errors_use_cataloged_codes_and_categories() {
    let dir = tempdir().unwrap();
    fs::write(dir.path().join("bad.json"), "{not-json").unwrap();
    let invalid = raw_officegen(
        &[
            "schema",
            "validate",
            "bad.json",
            "--schema",
            "edit-ops",
            "--agent",
            "--strict-json",
        ],
        dir.path(),
    );
    assert!(!invalid.status.success());
    let invalid_payload: Value = serde_json::from_slice(&invalid.stdout).unwrap();
    assert_eq!(invalid_payload["error"]["code"], "SCHEMA_INVALID");
    assert_eq!(invalid_payload["error"]["category"], "schema");

    fs::write(dir.path().join("input.txt"), "plain").unwrap();
    fs::write(
        dir.path().join("ops.json"),
        serde_json::to_vec_pretty(&serde_json::json!({
            "schema": "officegen.edit.ops@1.2",
            "ops": [{"op": "setText", "selector": {"contains": "plain"}, "text": "new"}]
        }))
        .unwrap(),
    )
    .unwrap();
    let unsupported = raw_officegen(
        &[
            "edit",
            "input.txt",
            "--ops",
            "ops.json",
            "--out",
            "out.txt",
            "--agent",
            "--strict-json",
        ],
        dir.path(),
    );
    assert!(!unsupported.status.success());
    let unsupported_payload: Value = serde_json::from_slice(&unsupported.stdout).unwrap();
    assert_eq!(unsupported_payload["error"]["code"], "UNSUPPORTED_FORMAT");
    assert_eq!(unsupported_payload["error"]["category"], "unsupported");

    let inspected = officegen(
        &[
            "errors",
            "inspect",
            "UNSUPPORTED_FORMAT",
            "--agent",
            "--strict-json",
        ],
        dir.path(),
    );
    assert_eq!(
        inspected["result"]["matches"][0]["category"],
        unsupported_payload["error"]["category"]
    );
}

#[test]
fn schema_fetch_out_writes_raw_schema_artifact() {
    let dir = tempdir().unwrap();
    let payload = officegen(
        &[
            "schema",
            "fetch",
            "edit-ops",
            "--out",
            "edit-ops.schema.json",
            "--agent",
            "--strict-json",
        ],
        dir.path(),
    );
    assert_eq!(payload["ok"], true);
    assert_eq!(payload["mutationStatus"], "changed");
    let schema: Value =
        serde_json::from_slice(&fs::read(dir.path().join("edit-ops.schema.json")).unwrap())
            .unwrap();
    assert_eq!(schema["$id"], "officegen.edit.ops@1.2");
}

#[test]
fn emitted_error_codes_are_cataloged_and_inspectable() {
    let dir = tempdir().unwrap();
    let list = officegen(&["errors", "list", "--agent", "--strict-json"], dir.path());
    let codes = list["result"]["errors"]
        .as_array()
        .unwrap()
        .iter()
        .filter_map(|entry| entry["code"].as_str())
        .collect::<Vec<_>>();
    for code in [
        "OUTPUT_REQUIRED",
        "SELECTOR_NOT_FOUND",
        "UNSUPPORTED_FORMAT",
        "OOXML_PARSE_FAILED",
        "WORKFLOW_STEP_FAILED",
        "WORKFLOW_RECURSION_DENIED",
    ] {
        assert!(codes.contains(&code), "missing {code}");
        let inspected = officegen(
            &["errors", "inspect", code, "--agent", "--strict-json"],
            dir.path(),
        );
        assert_eq!(inspected["result"]["matches"].as_array().unwrap().len(), 1);
    }
}

#[test]
fn topic_help_includes_operational_contracts() {
    let dir = tempdir().unwrap();
    let template = officegen(
        &["help", "template", "fill", "--agent", "--strict-json"],
        dir.path(),
    );
    assert!(template["result"]["details"]["usage"]
        .as_str()
        .unwrap()
        .contains("--data"));
    let run = officegen(&["help", "run", "--agent", "--strict-json"], dir.path());
    assert!(run["result"]["details"]["usage"]
        .as_str()
        .unwrap()
        .contains("workflow.json"));
}
