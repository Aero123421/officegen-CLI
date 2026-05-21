use serde_json::{json, Value};
use std::fs;
use std::io::Write;
use std::process::Command;
use tempfile::tempdir;
use zip::write::SimpleFileOptions;
use zip::ZipWriter;

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
fn schema_validate_honors_schema_id_alias() {
    let dir = tempdir().unwrap();
    fs::write(
        dir.path().join("doc.json"),
        serde_json::to_vec_pretty(&json!({
            "schema": "officegen.ir.document@2.0",
            "title": "Valid"
        }))
        .unwrap(),
    )
    .unwrap();
    let payload = officegen(
        &[
            "schema",
            "validate",
            "doc.json",
            "--schema-id",
            "officegen.ir.document@2.0",
            "--agent",
            "--strict-json",
        ],
        dir.path(),
    );
    assert_eq!(payload["ok"], true);
    assert_eq!(payload["result"]["ok"], true);
    assert_eq!(payload["result"]["schemaId"], "officegen.ir.document@2.0");
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

#[test]
fn json_budget_does_not_turn_errors_into_partial_success() {
    let dir = tempdir().unwrap();
    let output = raw_officegen(
        &[
            "not-a-command",
            "--json-budget-bytes",
            "512",
            "--agent",
            "--strict-json",
        ],
        dir.path(),
    );
    assert!(!output.status.success());
    let payload: Value = serde_json::from_slice(&output.stdout).unwrap();
    assert_eq!(payload["ok"], false);
    assert_eq!(payload["objectiveOk"], false);
    assert_eq!(payload["readiness"], "blocked");
}

#[test]
fn non_mutating_diff_does_not_report_mutation_changed() {
    let dir = tempdir().unwrap();
    fs::write(dir.path().join("a.txt"), "before").unwrap();
    fs::write(dir.path().join("b.txt"), "after").unwrap();
    let payload = officegen(
        &["diff", "a.txt", "b.txt", "--agent", "--strict-json"],
        dir.path(),
    );
    assert_eq!(payload["ok"], true);
    assert_eq!(payload["result"]["changed"], true);
    assert_eq!(payload["mutationStatus"], "not_applicable");
}

#[test]
fn artifact_paths_keep_output_directories() {
    let dir = tempdir().unwrap();
    fs::write(
        dir.path().join("chart.json"),
        serde_json::to_vec_pretty(&json!({
            "title": "Revenue",
            "labels": ["Q1", "Q2"],
            "values": [1, 2]
        }))
        .unwrap(),
    )
    .unwrap();
    let payload = officegen(
        &[
            "chart",
            "render",
            "chart.json",
            "--out",
            "out/chart.svg",
            "--agent",
            "--strict-json",
        ],
        dir.path(),
    );
    assert_eq!(payload["ok"], true);
    assert_eq!(payload["mutationStatus"], "changed");
    assert_eq!(payload["artifacts"][0]["path"], "out/chart.svg");
    assert!(dir.path().join("out").join("chart.svg").exists());
}

#[test]
fn chart_and_diagram_reject_or_parse_specs_honestly() {
    let dir = tempdir().unwrap();
    fs::write(
        dir.path().join("bad-chart.json"),
        serde_json::to_vec_pretty(&json!({
            "type": "qux",
            "labels": ["A"],
            "values": [1]
        }))
        .unwrap(),
    )
    .unwrap();
    let bad_chart = raw_officegen(
        &[
            "chart",
            "render",
            "bad-chart.json",
            "--out",
            "bad.svg",
            "--agent",
            "--strict-json",
        ],
        dir.path(),
    );
    assert!(!bad_chart.status.success());
    let payload: Value = serde_json::from_slice(&bad_chart.stdout).unwrap();
    assert_eq!(payload["error"]["code"], "SCHEMA_INVALID");
    assert!(!dir.path().join("bad.svg").exists());

    fs::write(
        dir.path().join("diagram.json"),
        serde_json::to_vec_pretty(&json!({
            "type": "flow",
            "nodes": [{"id": "A", "label": "Start"}, {"id": "B", "label": "End"}],
            "edges": [{"from": "A", "to": "B"}]
        }))
        .unwrap(),
    )
    .unwrap();
    let diagram = officegen(
        &[
            "diagram",
            "render",
            "diagram.json",
            "--out",
            "diagram.svg",
            "--agent",
            "--strict-json",
        ],
        dir.path(),
    );
    assert_eq!(diagram["ok"], true);
    assert_eq!(diagram["result"]["nodes"].as_array().unwrap().len(), 2);
    assert_eq!(diagram["result"]["edges"].as_array().unwrap().len(), 1);
    let svg = fs::read_to_string(dir.path().join("diagram.svg")).unwrap();
    assert!(svg.contains("Start"));
    assert!(svg.contains("End"));
    assert!(!svg.contains(">type<"));

    fs::write(dir.path().join("flow.mmd"), "graph TB\nA[Start] --> B[End]").unwrap();
    let mermaid = officegen(
        &[
            "diagram",
            "render",
            "flow.mmd",
            "--out",
            "flow.svg",
            "--agent",
            "--strict-json",
        ],
        dir.path(),
    );
    assert_eq!(mermaid["ok"], true);
    let svg = fs::read_to_string(dir.path().join("flow.svg")).unwrap();
    assert!(svg.contains("Start"));
    assert!(svg.contains("End"));
    assert!(!svg.contains(">TB<"));
}

#[test]
fn pdf_edit_is_fail_closed_until_portable_writer_exists() {
    let dir = tempdir().unwrap();
    fs::write(
        dir.path().join("pdf.json"),
        serde_json::to_vec_pretty(&json!({
            "schema": "officegen.ir.document@2.0",
            "title": "PDF",
            "blocks": [{"type": "paragraph", "text": "Visible text"}]
        }))
        .unwrap(),
    )
    .unwrap();
    officegen(
        &[
            "render",
            "pdf.json",
            "--target",
            "pdf",
            "--out",
            "input.pdf",
            "--agent",
            "--strict-json",
        ],
        dir.path(),
    );
    fs::write(
        dir.path().join("ops.json"),
        serde_json::to_vec_pretty(&json!({
            "schema": "officegen.edit.ops@1.2",
            "operations": [{"op": "pdf.annotate", "text": "note"}]
        }))
        .unwrap(),
    )
    .unwrap();
    let output = raw_officegen(
        &[
            "edit",
            "input.pdf",
            "--ops",
            "ops.json",
            "--out",
            "edited.pdf",
            "--agent",
            "--strict-json",
        ],
        dir.path(),
    );
    assert!(!output.status.success());
    let payload: Value = serde_json::from_slice(&output.stdout).unwrap();
    assert_eq!(payload["error"]["code"], "PDF_UNSUPPORTED_OPERATION");
    assert!(!dir.path().join("edited.pdf").exists());
}

#[test]
fn repair_out_writes_a_verified_copy_when_no_changes_needed() {
    let dir = tempdir().unwrap();
    fs::write(
        dir.path().join("doc.json"),
        serde_json::to_vec_pretty(&json!({
            "schema": "officegen.ir.document@2.0",
            "title": "Repair Copy"
        }))
        .unwrap(),
    )
    .unwrap();
    officegen(
        &[
            "render",
            "doc.json",
            "--target",
            "docx",
            "--out",
            "input.docx",
            "--agent",
            "--strict-json",
        ],
        dir.path(),
    );
    let payload = officegen(
        &[
            "repair",
            "input.docx",
            "--out",
            "repaired.docx",
            "--agent",
            "--strict-json",
        ],
        dir.path(),
    );
    assert_eq!(payload["ok"], true);
    assert_eq!(payload["mutationStatus"], "noop");
    assert_eq!(payload["artifactStatus"], "complete");
    assert!(dir.path().join("repaired.docx").exists());
}

#[test]
fn repair_out_rejects_unsafe_zip_instead_of_copying_it() {
    let dir = tempdir().unwrap();
    let file = fs::File::create(dir.path().join("unsafe.docx")).unwrap();
    let mut writer = ZipWriter::new(file);
    let options = SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    writer.start_file("../evil.txt", options).unwrap();
    writer.write_all(b"bad").unwrap();
    writer.finish().unwrap();

    let output = raw_officegen(
        &[
            "repair",
            "unsafe.docx",
            "--out",
            "copied.docx",
            "--agent",
            "--strict-json",
        ],
        dir.path(),
    );
    assert!(!output.status.success());
    let payload: Value = serde_json::from_slice(&output.stdout).unwrap();
    assert_eq!(payload["error"]["code"], "SECURITY_ZIP_UNSAFE");
    assert!(!dir.path().join("copied.docx").exists());
}

#[test]
fn rendered_pdf_inspect_exposes_visible_text_not_ir_metadata() {
    let dir = tempdir().unwrap();
    fs::write(
        dir.path().join("pdf.json"),
        serde_json::to_vec_pretty(&json!({
            "schema": "officegen.ir.document@2.0",
            "title": "Quarterly Review",
            "targets": ["pdf"],
            "blocks": [{"type": "paragraph", "text": "Human visible summary"}]
        }))
        .unwrap(),
    )
    .unwrap();
    officegen(
        &[
            "render",
            "pdf.json",
            "--target",
            "pdf",
            "--out",
            "input.pdf",
            "--agent",
            "--strict-json",
        ],
        dir.path(),
    );
    let inspected = officegen(
        &["inspect", "input.pdf", "--agent", "--strict-json"],
        dir.path(),
    );
    let preview = inspected["result"]["untrusted"]["textPreview"]
        .as_str()
        .unwrap();
    assert!(preview.contains("Quarterly Review"));
    assert!(preview.contains("Human visible summary"));
    assert!(!preview.contains("officegen.ir.document"));
    assert!(!preview.contains("paragraph"));
}
