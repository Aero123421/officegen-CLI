#[path = "../src/schemas.rs"]
mod schemas;

use schemas::{
    fetch_schema, get_schema, list_schemas, resolve_alias, validate_minimal_required_fields,
};
use serde_json::{json, Value};

#[test]
fn lists_v45_catalog_in_stable_order() {
    let ids = list_schemas()
        .iter()
        .map(|entry| entry.id)
        .collect::<Vec<_>>();

    assert_eq!(
        ids,
        vec![
            "officegen.envelope@1.2",
            "officegen.capabilities@1.2",
            "officegen.ir.document@1.2",
            "officegen.ir.document@2.0",
            "officegen.edit.ops@1.2",
            "officegen.manifest@1.2",
            "officegen.workflow@1.2",
            "officegen.workflow@2.0",
            "officegen.error.catalog@1.2",
        ]
    );
}

#[test]
fn fetch_resolves_aliases_and_embedded_schema_ids() {
    assert_eq!(resolve_alias("envelope"), Some("officegen.envelope@1.2"));
    assert_eq!(
        resolve_alias("officegen.ir.document"),
        Some("officegen.ir.document@2.0")
    );
    assert_eq!(resolve_alias("edit-ops"), Some("officegen.edit.ops@1.2"));
    assert_eq!(resolve_alias("errors"), Some("officegen.error.catalog@1.2"));

    let document = fetch_schema("workflow").unwrap();

    assert_eq!(document.id, "officegen.workflow@2.0");
    assert_eq!(document.path, "schemas/workflow-2.0.schema.json");
    assert_eq!(document.schema["$id"], "officegen.workflow@2.0");
    assert!(get_schema("missing").is_none());
}

#[test]
fn every_catalog_document_is_parseable_and_self_identifies() {
    for entry in list_schemas() {
        let document = fetch_schema(entry.id).unwrap();

        assert_eq!(
            document.schema["$schema"],
            "https://json-schema.org/draft/2020-12/schema"
        );
        assert_eq!(document.schema["$id"], entry.id);
        assert_eq!(document.schema["type"], "object");
        assert!(document.schema["required"]
            .as_array()
            .unwrap()
            .contains(&json!("schema")));
    }
}

#[test]
fn validates_minimal_required_fields_for_golden_payloads() {
    let golden_payloads: Vec<(&str, Value)> = vec![
        (
            "envelope",
            json!({
                "schema": "officegen.envelope@1.2",
                "ok": true,
                "cliVersion": "4.5.0",
                "pathsRedacted": true,
                "warnings": [],
                "diagnostics": [],
                "artifacts": [],
                "nextSuggestedCommands": []
            }),
        ),
        (
            "capabilities",
            json!({
                "schema": "officegen.capabilities@1.2",
                "ok": true,
                "profile": "substrate",
                "capabilitiesHash": "sha256:test",
                "visibleCommands": ["schema list"],
                "formatCapabilities": {},
                "featureContracts": [],
                "unsupportedNow": [],
                "nextSuggestedCommands": []
            }),
        ),
        (
            "officegen.ir.document@1.2",
            json!({
                "schema": "officegen.ir.document@1.2",
                "targets": ["docx"],
                "sections": []
            }),
        ),
        (
            "document",
            json!({
                "schema": "officegen.ir.document@2.0",
                "targets": ["docx"],
                "sections": []
            }),
        ),
        (
            "ops",
            json!({
                "schema": "officegen.edit.ops@1.2",
                "operations": [{"op": "docx.setText"}]
            }),
        ),
        (
            "manifest",
            json!({
                "schema": "officegen.manifest@1.2",
                "version": "4.5.0",
                "artifacts": []
            }),
        ),
        (
            "officegen.workflow@1.2",
            json!({
                "schema": "officegen.workflow@1.2",
                "version": "1.0",
                "steps": [{"id": "inspect"}]
            }),
        ),
        (
            "workflow",
            json!({
                "schema": "officegen.workflow@2.0",
                "version": "2.0",
                "steps": [{"id": "inspect", "command": "inspect"}]
            }),
        ),
        (
            "errors",
            json!({
                "schema": "officegen.error.catalog@1.2",
                "version": "1.0",
                "errors": [{"code": "UNKNOWN_COMMAND", "category": "usage", "severity": "error"}]
            }),
        ),
    ];

    for (schema, payload) in golden_payloads {
        let report = validate_minimal_required_fields(schema, &payload);

        assert!(report.ok, "{schema}: {:?}", report.errors);
    }
}

#[test]
fn reports_missing_schema_mismatch_and_type_errors() {
    let missing = validate_minimal_required_fields(
        "manifest",
        &json!({
            "schema": "officegen.manifest@1.2",
            "version": "4.5.0"
        }),
    );
    assert!(!missing.ok);
    assert_eq!(missing.errors[0].instance_path, "/artifacts");

    let mismatch = validate_minimal_required_fields(
        "officegen.workflow@1.2",
        &json!({
            "schema": "officegen.manifest@1.2",
            "version": "1.0",
            "steps": []
        }),
    );
    assert!(!mismatch.ok);
    assert!(mismatch.errors.iter().any(|error| {
        error.instance_path == "/schema" && error.message.contains("officegen.workflow@1.2")
    }));

    let wrong_type = validate_minimal_required_fields(
        "officegen.ir.document@1.2",
        &json!({
            "schema": "officegen.ir.document@1.2",
            "targets": "docx",
            "sections": []
        }),
    );
    assert!(!wrong_type.ok);
    assert!(wrong_type.errors.iter().any(|error| {
        error.instance_path == "/targets" && error.message.contains("expected type array")
    }));

    let optional_wrong_type = validate_minimal_required_fields(
        "officegen.ir.document@1.2",
        &json!({
            "schema": "officegen.ir.document@1.2",
            "title": 42,
            "targets": ["docx"],
            "sections": []
        }),
    );
    assert!(!optional_wrong_type.ok);
    assert!(optional_wrong_type.errors.iter().any(|error| {
        error.instance_path == "/title" && error.message.contains("expected type string")
    }));

    let missing_nested = validate_minimal_required_fields(
        "ops",
        &json!({
            "schema": "officegen.edit.ops@1.2",
            "operations": [{}]
        }),
    );
    assert!(!missing_nested.ok);
    assert!(missing_nested
        .errors
        .iter()
        .any(|error| error.instance_path == "/operations/0/op"));
}

#[test]
fn edit_ops_schema_accepts_ops_alias_but_rejects_empty_operation_payloads() {
    let missing_ops = validate_minimal_required_fields(
        "ops",
        &json!({
            "schema": "officegen.edit.ops@1.2"
        }),
    );
    assert!(!missing_ops.ok);
    assert!(missing_ops
        .errors
        .iter()
        .any(|error| error.message.contains("anyOf")));

    let alias_ops = validate_minimal_required_fields(
        "ops",
        &json!({
            "schema": "officegen.edit.ops@1.2",
            "ops": [{"op": "docx.setText"}]
        }),
    );
    assert!(alias_ops.ok, "{:?}", alias_ops.errors);

    let canonical_ops = validate_minimal_required_fields(
        "ops",
        &json!({
            "schema": "officegen.edit.ops@1.2",
            "operations": [{"op": "docx.setText"}]
        }),
    );
    assert!(canonical_ops.ok, "{:?}", canonical_ops.errors);

    let conflicting_aliases = validate_minimal_required_fields(
        "ops",
        &json!({
            "schema": "officegen.edit.ops@1.2",
            "operations": [{"op": "docx.setText"}],
            "ops": [{"op": "pptx.setText"}]
        }),
    );
    assert!(!conflicting_aliases.ok);
    assert!(conflicting_aliases
        .errors
        .iter()
        .any(|error| error.message.contains("operations and ops")));
}
