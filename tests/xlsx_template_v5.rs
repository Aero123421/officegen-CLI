use serde_json::json;
use std::fs;
use std::io::Read;
use std::path::Path;
use std::process::Command;
use tempfile::tempdir;
use zip::ZipArchive;

fn officegen(dir: &Path, args: &[&str]) -> serde_json::Value {
    let output = Command::new(env!("CARGO_BIN_EXE_officegen"))
        .current_dir(dir)
        .args(args)
        .output()
        .expect("officegen command runs");
    assert!(
        output.status.success(),
        "officegen failed\nstdout:\n{}\nstderr:\n{}",
        String::from_utf8_lossy(&output.stdout),
        String::from_utf8_lossy(&output.stderr)
    );
    serde_json::from_slice(&output.stdout).expect("strict-json output")
}

fn officegen_raw(dir: &Path, args: &[&str]) -> std::process::Output {
    Command::new(env!("CARGO_BIN_EXE_officegen"))
        .current_dir(dir)
        .args(args)
        .output()
        .expect("officegen command runs")
}

fn zip_text(path: &Path, part: &str) -> String {
    let file = fs::File::open(path).expect("open xlsx");
    let mut zip = ZipArchive::new(file).expect("zip opens");
    let mut entry = zip.by_name(part).expect("zip part exists");
    let mut text = String::new();
    entry.read_to_string(&mut text).expect("part is text");
    text
}

#[test]
fn render_xlsx_ir_v2_sheets_tables_formulas_validations_and_charts() {
    let dir = tempdir().unwrap();
    fs::write(
        dir.path().join("report.json"),
        serde_json::to_vec_pretty(&json!({
            "schema": "officegen.ir.document@2.0",
            "title": "Sales",
            "targets": ["xlsx"],
            "sheets": [{
                "name": "Summary",
                "rows": [["Region", "Value"], ["North", 12], ["South", 18]],
                "formulas": [{"cell": "B4", "formula": "SUM(B2:B3)"}],
                "tables": [{"name": "SalesTable", "range": "A1:B3"}],
                "dataValidations": [{"range": "A2:A3", "type": "list", "formula1": "\"North,South\""}],
                "charts": [{"title": "By region", "labels": ["North", "South"], "values": [12, 18]}]
            }, {
                "name": "Detail",
                "rows": [["Metric", "Value"], ["Pipeline", 30]],
                "tables": [{"name": "DetailTable", "range": "A1:B2"}],
                "charts": [{"title": "Detail chart", "labels": ["Pipeline"], "values": [30]}]
            }]
        }))
        .unwrap(),
    )
    .unwrap();

    let payload = officegen(
        dir.path(),
        &[
            "render",
            "report.json",
            "--target",
            "xlsx",
            "--out",
            "report.xlsx",
            "--agent",
            "--strict-json",
        ],
    );

    assert_eq!(payload["ok"], true);
    let xlsx = dir.path().join("report.xlsx");
    let sheet = zip_text(&xlsx, "xl/worksheets/sheet1.xml");
    let chart = zip_text(&xlsx, "xl/charts/chart1.xml");
    assert!(sheet.contains("SUM(B2:B3)"));
    assert!(sheet.contains("<dataValidations"));
    assert!(sheet.contains("<tableParts"));
    assert!(sheet.contains("<drawing "));
    assert!(chart.contains("North"));
    assert!(chart.contains("18"));
    let sheet2_rels = zip_text(&xlsx, "xl/worksheets/_rels/sheet2.xml.rels");
    assert!(sheet2_rels.contains("Target=\"../tables/table2.xml\""));
    assert!(sheet2_rels.contains("Target=\"../drawings/drawing2.xml\""));
    assert!(zip_text(&xlsx, "xl/tables/table2.xml").contains("DetailTable"));
    assert!(zip_text(&xlsx, "xl/charts/chart2.xml").contains("30"));
}

#[test]
fn xlsx_formula_external_and_network_references_fail_closed() {
    let dir = tempdir().unwrap();
    fs::write(
        dir.path().join("report.json"),
        serde_json::to_vec_pretty(&json!({
            "schema": "officegen.ir.document@2.0",
            "targets": ["xlsx"],
            "sheets": [{
                "name": "Unsafe",
                "rows": [["Name", "Value"]],
                "formulas": [{"cell": "B2", "formula": "WEBSERVICE(\"https://example.test\")"}]
            }]
        }))
        .unwrap(),
    )
    .unwrap();

    let output = officegen_raw(
        dir.path(),
        &[
            "render",
            "report.json",
            "--target",
            "xlsx",
            "--out",
            "report.xlsx",
            "--agent",
            "--strict-json",
        ],
    );
    let payload: serde_json::Value =
        serde_json::from_slice(&output.stdout).expect("strict-json output");
    assert!(!output.status.success());
    assert_eq!(payload["ok"], false);
    assert_eq!(payload["error"]["code"], "SCHEMA_INVALID");
    assert!(!dir.path().join("report.xlsx").exists());
}

#[test]
fn edit_xlsx_package_ops_update_real_parts() {
    let dir = tempdir().unwrap();
    fs::write(
        dir.path().join("base.json"),
        serde_json::to_vec_pretty(&json!({
            "schema": "officegen.ir.document@2.0",
            "title": "Base",
            "targets": ["xlsx"],
            "sheets": [{"name": "Sheet1", "rows": [["A", "B"]]}]
        }))
        .unwrap(),
    )
    .unwrap();
    officegen(
        dir.path(),
        &[
            "render",
            "base.json",
            "--target",
            "xlsx",
            "--out",
            "base.xlsx",
            "--agent",
            "--strict-json",
        ],
    );
    fs::write(
        dir.path().join("ops.json"),
        serde_json::to_vec_pretty(&json!({
            "schema": "officegen.edit.ops@1.2",
            "operations": [
                {"op": "xlsx.addSheet", "name": "Extra"},
                {"op": "xlsx.setCell", "sheet": "Extra", "cell": "C1", "value": "Scoped"},
                {"op": "xlsx.setRange", "sheet": "Extra", "cell": "A1", "values": [["Name", "Score"], ["Ada", 42]]},
                {"op": "xlsx.setFormula", "sheet": "Extra", "cell": "B3", "formula": "SUM(B2:B2)"},
                {"op": "xlsx.renameSheet", "sheet": "Extra", "name": "Final"},
                {"op": "xlsx.setNamedRange", "name": "ScoreRange", "refersTo": "Final!$B$2:$B$2"},
                {"op": "xlsx.setDataValidation", "sheet": "Final", "range": "A2:A2", "type": "list", "formula1": "\"Ada\""},
                {"op": "xlsx.addTable", "sheet": "Final", "name": "ScoreTable", "range": "A1:B2"},
                {"op": "xlsx.addChart", "sheet": "Final", "title": "Scores", "labels": ["Ada"], "values": [42]}
            ]
        }))
        .unwrap(),
    )
    .unwrap();

    let payload = officegen(
        dir.path(),
        &[
            "edit",
            "base.xlsx",
            "--ops",
            "ops.json",
            "--out",
            "edited.xlsx",
            "--agent",
            "--strict-json",
        ],
    );

    assert_eq!(payload["ok"], true);
    assert_eq!(payload["result"]["applied"], 9);
    let xlsx = dir.path().join("edited.xlsx");
    assert!(zip_text(&xlsx, "xl/workbook.xml").contains("Final"));
    assert!(zip_text(&xlsx, "xl/workbook.xml").contains("ScoreRange"));
    let sheet2 = zip_text(&xlsx, "xl/worksheets/sheet2.xml");
    assert!(sheet2.contains("Ada"));
    assert!(sheet2.contains("Scoped"));
    assert!(sheet2.contains("SUM(B2:B2)"));
    assert!(sheet2.contains("<dataValidations"));
    assert!(!sheet2.contains("<dataValidations count=\"1\"><dataValidations"));
    assert!(zip_text(&xlsx, "xl/tables/table1.xml").contains("ScoreTable"));
    assert!(zip_text(&xlsx, "xl/charts/chart1.xml").contains("42"));
}

#[test]
fn template_inspect_and_fill_xlsx_placeholders() {
    let dir = tempdir().unwrap();
    fs::write(
        dir.path().join("template.json"),
        serde_json::to_vec_pretty(&json!({
            "schema": "officegen.ir.document@2.0",
            "targets": ["xlsx"],
            "sheets": [{"name": "Template", "rows": [["Name", "{{name}}"]]}]
        }))
        .unwrap(),
    )
    .unwrap();
    officegen(
        dir.path(),
        &[
            "render",
            "template.json",
            "--target",
            "xlsx",
            "--out",
            "template.xlsx",
            "--agent",
            "--strict-json",
        ],
    );
    let inspect = officegen(
        dir.path(),
        &[
            "template",
            "inspect",
            "template.xlsx",
            "--agent",
            "--strict-json",
        ],
    );
    assert_eq!(inspect["result"]["placeholderCount"], 1);

    fs::write(
        dir.path().join("data.json"),
        serde_json::to_vec_pretty(&json!({"name": "Nano"})).unwrap(),
    )
    .unwrap();
    let fill = officegen(
        dir.path(),
        &[
            "template",
            "fill",
            "template.xlsx",
            "--data",
            "data.json",
            "--out",
            "filled.xlsx",
            "--agent",
            "--strict-json",
        ],
    );

    assert_eq!(fill["ok"], true);
    let sheet = zip_text(&dir.path().join("filled.xlsx"), "xl/worksheets/sheet1.xml");
    assert!(sheet.contains("Nano"));
    assert!(!sheet.contains("{{name}}"));
}

#[test]
fn template_fill_missing_fields_fails_closed_without_artifact() {
    let dir = tempdir().unwrap();
    fs::write(
        dir.path().join("template.json"),
        serde_json::to_vec_pretty(&json!({
            "schema": "officegen.ir.document@2.0",
            "targets": ["xlsx"],
            "sheets": [{"name": "Template", "rows": [["Name", "{{name}}"]]}]
        }))
        .unwrap(),
    )
    .unwrap();
    officegen(
        dir.path(),
        &[
            "render",
            "template.json",
            "--target",
            "xlsx",
            "--out",
            "template.xlsx",
            "--agent",
            "--strict-json",
        ],
    );
    fs::write(dir.path().join("data.json"), "{}").unwrap();

    let output = officegen_raw(
        dir.path(),
        &[
            "template",
            "fill",
            "template.xlsx",
            "--data",
            "data.json",
            "--out",
            "filled.xlsx",
            "--agent",
            "--strict-json",
        ],
    );
    let payload: serde_json::Value =
        serde_json::from_slice(&output.stdout).expect("strict-json output");

    assert!(!output.status.success());
    assert_eq!(payload["ok"], false);
    assert_eq!(payload["readiness"], "blocked");
    assert_eq!(payload["result"]["missingFields"][0], "name");
    assert!(!dir.path().join("filled.xlsx").exists());
}

#[test]
fn template_fill_formula_placeholder_is_revalidated() {
    let dir = tempdir().unwrap();
    fs::write(
        dir.path().join("template.json"),
        serde_json::to_vec_pretty(&json!({
            "schema": "officegen.ir.document@2.0",
            "targets": ["xlsx"],
            "sheets": [{"name": "Template", "rows": [["Formula", {"formula": "{{formula}}"}]]}]
        }))
        .unwrap(),
    )
    .unwrap();
    officegen(
        dir.path(),
        &[
            "render",
            "template.json",
            "--target",
            "xlsx",
            "--out",
            "template.xlsx",
            "--agent",
            "--strict-json",
        ],
    );
    fs::write(
        dir.path().join("data.json"),
        serde_json::to_vec_pretty(&json!({"formula": "'[other.xlsx]Sheet1'!A1"})).unwrap(),
    )
    .unwrap();

    let output = officegen_raw(
        dir.path(),
        &[
            "template",
            "fill",
            "template.xlsx",
            "--data",
            "data.json",
            "--out",
            "filled.xlsx",
            "--agent",
            "--strict-json",
        ],
    );
    let payload: serde_json::Value =
        serde_json::from_slice(&output.stdout).expect("strict-json output");

    assert!(!output.status.success());
    assert_eq!(payload["ok"], false);
    assert_eq!(payload["error"]["code"], "SCHEMA_INVALID");
    assert!(!dir.path().join("filled.xlsx").exists());
}

#[test]
fn template_fill_formula_numeric_entity_references_are_revalidated() {
    let dir = tempdir().unwrap();
    fs::write(
        dir.path().join("template.json"),
        serde_json::to_vec_pretty(&json!({
            "schema": "officegen.ir.document@2.0",
            "targets": ["xlsx"],
            "sheets": [{"name": "Template", "rows": [["Formula", {"formula": "{{formula}}"}]]}]
        }))
        .unwrap(),
    )
    .unwrap();
    officegen(
        dir.path(),
        &[
            "render",
            "template.json",
            "--target",
            "xlsx",
            "--out",
            "template.xlsx",
            "--agent",
            "--strict-json",
        ],
    );
    fs::write(
        dir.path().join("data.json"),
        serde_json::to_vec_pretty(&json!({"formula": "cmd&#124;/C calc!A1"})).unwrap(),
    )
    .unwrap();

    let output = officegen_raw(
        dir.path(),
        &[
            "template",
            "fill",
            "template.xlsx",
            "--data",
            "data.json",
            "--out",
            "filled.xlsx",
            "--agent",
            "--strict-json",
        ],
    );
    let payload: serde_json::Value =
        serde_json::from_slice(&output.stdout).expect("strict-json output");

    assert!(!output.status.success());
    assert_eq!(payload["ok"], false);
    assert_eq!(payload["error"]["code"], "SCHEMA_INVALID");
    assert!(!dir.path().join("filled.xlsx").exists());
}
