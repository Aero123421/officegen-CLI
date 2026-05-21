use serde_json::{json, Value};
use std::fs;
use std::process::Command;
use tempfile::tempdir;

fn officegen(args: &[&str], cwd: &std::path::Path) -> Value {
    let output = Command::new(env!("CARGO_BIN_EXE_officegen"))
        .args(args)
        .current_dir(cwd)
        .output()
        .expect("officegen command should run");
    assert!(
        output.status.success(),
        "command failed\nstdout:\n{}\nstderr:\n{}",
        String::from_utf8_lossy(&output.stdout),
        String::from_utf8_lossy(&output.stderr)
    );
    serde_json::from_slice(&output.stdout).expect("stdout should be JSON")
}

#[test]
fn docx_v5_renders_blocks_and_reports_presence_in_verify_and_view() {
    let dir = tempdir().unwrap();
    fs::write(
        dir.path().join("doc.json"),
        serde_json::to_vec(&json!({
            "schema": "officegen.ir.document@2.0",
            "title": "V5 Authoring Brief",
            "targets": ["docx"],
            "sections": [{
                "title": "Authoring",
                "blocks": [
                    {"type": "paragraph", "text": "Portable DOCX blocks render as editable text."},
                    {"type": "table", "rows": [["Area", "Status"], ["Tables", "Present"]]},
                    {"type": "image", "alt": "Architecture snapshot"},
                    {"type": "chart", "title": "Presence chart"}
                ]
            }]
        }))
        .unwrap(),
    )
    .unwrap();

    let render = officegen(
        &[
            "render",
            "doc.json",
            "--target",
            "docx",
            "--out",
            "out.docx",
            "--agent",
            "--strict-json",
        ],
        dir.path(),
    );
    assert_eq!(render["ok"], true);

    let inspect = officegen(
        &["inspect", "out.docx", "--agent", "--strict-json"],
        dir.path(),
    );
    assert_eq!(inspect["result"]["trusted"]["summary"]["tables"], 1);
    assert_eq!(inspect["result"]["trusted"]["summary"]["images"], 1);
    assert_eq!(inspect["result"]["trusted"]["summary"]["charts"], 1);
    assert!(inspect["result"]["untrusted"]["textPreview"]
        .as_str()
        .unwrap()
        .contains("Portable DOCX blocks"));

    let verify = officegen(
        &["verify", "out.docx", "--agent", "--strict-json"],
        dir.path(),
    );
    assert_eq!(verify["result"]["status"], "pass");
    assert_eq!(verify["result"]["semanticPresence"]["tables"], 1);
    assert_eq!(verify["result"]["semanticPresence"]["images"], 1);
    assert_eq!(verify["result"]["semanticPresence"]["charts"], 1);

    let view = officegen(
        &[
            "view",
            "out.docx",
            "--format",
            "html",
            "--out",
            "view",
            "--agent",
            "--strict-json",
        ],
        dir.path(),
    );
    assert_eq!(view["result"]["artifactUsable"], true);
    let html = fs::read_to_string(dir.path().join("view").join("index.html")).unwrap();
    assert!(html.contains("OOXML Preview"));
    assert!(html.contains("<table>"));
}

#[test]
fn pptx_v5_table_edit_diff_and_svg_view_are_semantic() {
    let dir = tempdir().unwrap();
    fs::write(
        dir.path().join("deck.json"),
        serde_json::to_vec(&json!({
            "schema": "officegen.ir.document@2.0",
            "title": "V5 Deck",
            "targets": ["pptx"],
            "slides": [{
                "title": "Release Plan",
                "blocks": [
                    {"type": "paragraph", "text": "Keep edits scoped."},
                    {"type": "table", "rows": [["Work", "Owner"], ["PPTX", "Runtime"]]},
                    {"type": "chart", "title": "Readiness"}
                ]
            }]
        }))
        .unwrap(),
    )
    .unwrap();
    officegen(
        &[
            "render",
            "deck.json",
            "--target",
            "pptx",
            "--out",
            "before.pptx",
            "--agent",
            "--strict-json",
        ],
        dir.path(),
    );
    fs::write(
        dir.path().join("ops.json"),
        serde_json::to_vec(&json!({
            "schema": "officegen.edit.ops@2.0",
            "operations": [{
                "op": "pptx.setTableCell",
                "selector": {"sourcePath": "ppt/slides/slide1.xml", "tableIndex": 0, "rowIndex": 1, "columnIndex": 1},
                "text": "OOXML runtime"
            }]
        }))
        .unwrap(),
    )
    .unwrap();

    let edit = officegen(
        &[
            "edit",
            "before.pptx",
            "--ops",
            "ops.json",
            "--out",
            "after.pptx",
            "--agent",
            "--strict-json",
        ],
        dir.path(),
    );
    assert_eq!(edit["result"]["applied"], 1);

    let inspect = officegen(
        &["inspect", "after.pptx", "--agent", "--strict-json"],
        dir.path(),
    );
    assert_eq!(inspect["result"]["trusted"]["summary"]["tables"], 1);
    assert_eq!(inspect["result"]["trusted"]["summary"]["charts"], 1);
    assert!(inspect["result"]["untrusted"]["textPreview"]
        .as_str()
        .unwrap()
        .contains("OOXML runtime"));

    let diff = officegen(
        &[
            "diff",
            "before.pptx",
            "after.pptx",
            "--agent",
            "--strict-json",
        ],
        dir.path(),
    );
    assert_eq!(diff["result"]["changed"], true);
    assert_eq!(diff["result"]["semantic"]["tableChanged"], true);
    assert_eq!(diff["result"]["semantic"]["presenceAfter"]["charts"], 1);

    officegen(
        &[
            "view",
            "after.pptx",
            "--format",
            "svg",
            "--out",
            "view",
            "--agent",
            "--strict-json",
        ],
        dir.path(),
    );
    let svg = fs::read_to_string(dir.path().join("view").join("page-001.svg")).unwrap();
    assert!(svg.contains("semantic preview"));
    assert!(svg.contains("OOXML runtime"));
    assert!(svg.contains("table:"));
}

#[test]
fn inspect_controls_bound_object_map_and_json_budget() {
    let dir = tempdir().unwrap();
    fs::write(
        dir.path().join("deck.json"),
        serde_json::to_vec_pretty(&json!({
            "schema": "officegen.ir.document@2.0",
            "targets": ["pptx"],
            "slides": [{
                "title": "Budgeted Deck",
                "blocks": (0..20).map(|i| json!({"type": "paragraph", "text": format!("Line {i}")})).collect::<Vec<_>>()
            }]
        }))
        .unwrap(),
    )
    .unwrap();
    officegen(
        &[
            "render",
            "deck.json",
            "--target",
            "pptx",
            "--out",
            "deck.pptx",
            "--agent",
            "--strict-json",
        ],
        dir.path(),
    );

    let limited = officegen(
        &[
            "inspect",
            "deck.pptx",
            "--object-map-limit",
            "2",
            "--agent",
            "--strict-json",
        ],
        dir.path(),
    );
    assert_eq!(limited["result"]["objectMap"].as_array().unwrap().len(), 2);
    assert_eq!(
        limited["result"]["truncated"]["objectMap"]["returnedCount"],
        2
    );

    let budgeted = officegen(
        &[
            "inspect",
            "deck.pptx",
            "--json-budget-bytes",
            "4000",
            "--agent",
            "--strict-json",
        ],
        dir.path(),
    );
    assert_eq!(budgeted["partial"], true);
    assert_eq!(budgeted["readiness"], "partial");
    assert!(serde_json::to_string(&budgeted).unwrap().len() <= 4000);

    let tiny_budget = officegen(
        &[
            "inspect",
            "deck.pptx",
            "--json-budget-bytes",
            "50",
            "--agent",
            "--strict-json",
        ],
        dir.path(),
    );
    assert_eq!(tiny_budget["partial"], true);
    assert_eq!(tiny_budget["result"]["fullResultOmitted"], true);
    assert!(serde_json::to_string(&tiny_budget).unwrap().len() > 50);
}

#[test]
fn pdf_inspect_does_not_expose_raw_pdf_structure() {
    let dir = tempdir().unwrap();
    fs::write(
        dir.path().join("doc.json"),
        serde_json::to_vec_pretty(&json!({
            "schema": "officegen.ir.document@1.2",
            "title": "PDF Safe Preview",
            "targets": ["pdf"],
            "sections": [{"blocks": [{"type": "paragraph", "text": "Visible PDF text"}]}]
        }))
        .unwrap(),
    )
    .unwrap();
    officegen(
        &[
            "render",
            "doc.json",
            "--target",
            "pdf",
            "--out",
            "doc.pdf",
            "--agent",
            "--strict-json",
        ],
        dir.path(),
    );
    let inspect = officegen(
        &["inspect", "doc.pdf", "--agent", "--strict-json"],
        dir.path(),
    );
    let preview = inspect["result"]["untrusted"]["textPreview"]
        .as_str()
        .unwrap();
    assert!(preview.contains("Visible PDF text"));
    for raw in ["%PDF", "stream", "xref", "trailer"] {
        assert!(!preview.contains(raw), "preview leaked {raw}");
    }

    fs::write(
        dir.path().join("raw.pdf"),
        b"%PDF-1.7\n1 0 obj\n<</Length 4>>stream\n\x00\x01\x02\x03\nendstream\nxref\n%%EOF",
    )
    .unwrap();
    let raw = officegen(
        &["inspect", "raw.pdf", "--agent", "--strict-json"],
        dir.path(),
    );
    assert_eq!(raw["result"]["untrusted"]["textPreview"], "");
    assert_eq!(
        raw["result"]["trusted"]["summary"]["extractionConfidence"],
        "none"
    );
}
