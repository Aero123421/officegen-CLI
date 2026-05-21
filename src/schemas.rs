use regex::Regex;
use serde_json::Value;
use std::fmt;

const ENVELOPE_SCHEMA_JSON: &str = include_str!("../schemas/envelope-1.2.schema.json");
const CAPABILITIES_SCHEMA_JSON: &str = include_str!("../schemas/capabilities-1.2.schema.json");
const IR_DOCUMENT_SCHEMA_JSON: &str = include_str!("../schemas/ir-document-1.2.schema.json");
const IR_DOCUMENT_V2_SCHEMA_JSON: &str = include_str!("../schemas/ir-document-2.0.schema.json");
const EDIT_OPS_SCHEMA_JSON: &str = include_str!("../schemas/edit-ops-1.2.schema.json");
const MANIFEST_SCHEMA_JSON: &str = include_str!("../schemas/manifest-1.2.schema.json");
const WORKFLOW_SCHEMA_JSON: &str = include_str!("../schemas/workflow-1.2.schema.json");
const WORKFLOW_V2_SCHEMA_JSON: &str = include_str!("../schemas/workflow-2.0.schema.json");
const ERROR_CATALOG_SCHEMA_JSON: &str = include_str!("../schemas/error-catalog-1.2.schema.json");

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SchemaCatalogEntry {
    pub id: &'static str,
    pub aliases: &'static [&'static str],
    pub path: &'static str,
}

#[derive(Clone, Debug, PartialEq)]
pub struct SchemaDocument {
    pub id: &'static str,
    pub path: &'static str,
    pub schema: Value,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ValidationError {
    pub instance_path: String,
    pub message: String,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ValidationReport {
    pub schema_id: &'static str,
    pub ok: bool,
    pub errors: Vec<ValidationError>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum SchemaCatalogError {
    UnknownSchema(String),
    InvalidCatalogSchema { id: &'static str, message: String },
}

#[derive(Clone, Copy, Debug)]
struct RawSchema {
    id: &'static str,
    aliases: &'static [&'static str],
    path: &'static str,
    json: &'static str,
}

const RAW_SCHEMAS: &[RawSchema] = &[
    RawSchema {
        id: "officegen.envelope@1.2",
        aliases: &["envelope", "envelope@1.2", "runtime-envelope"],
        path: "schemas/envelope-1.2.schema.json",
        json: ENVELOPE_SCHEMA_JSON,
    },
    RawSchema {
        id: "officegen.capabilities@1.2",
        aliases: &["capabilities", "capabilities@1.2"],
        path: "schemas/capabilities-1.2.schema.json",
        json: CAPABILITIES_SCHEMA_JSON,
    },
    RawSchema {
        id: "officegen.ir.document@1.2",
        aliases: &["document@1.2", "ir-document@1.2"],
        path: "schemas/ir-document-1.2.schema.json",
        json: IR_DOCUMENT_SCHEMA_JSON,
    },
    RawSchema {
        id: "officegen.ir.document@2.0",
        aliases: &[
            "ir.document",
            "document",
            "document@2.0",
            "ir-document",
            "ir-document@2.0",
        ],
        path: "schemas/ir-document-2.0.schema.json",
        json: IR_DOCUMENT_V2_SCHEMA_JSON,
    },
    RawSchema {
        id: "officegen.edit.ops@1.2",
        aliases: &["edit.ops", "edit-ops", "ops", "ops@1.2"],
        path: "schemas/edit-ops-1.2.schema.json",
        json: EDIT_OPS_SCHEMA_JSON,
    },
    RawSchema {
        id: "officegen.manifest@1.2",
        aliases: &["manifest", "manifest@1.2"],
        path: "schemas/manifest-1.2.schema.json",
        json: MANIFEST_SCHEMA_JSON,
    },
    RawSchema {
        id: "officegen.workflow@1.2",
        aliases: &["workflow@1.2"],
        path: "schemas/workflow-1.2.schema.json",
        json: WORKFLOW_SCHEMA_JSON,
    },
    RawSchema {
        id: "officegen.workflow@2.0",
        aliases: &["workflow", "workflow@2.0"],
        path: "schemas/workflow-2.0.schema.json",
        json: WORKFLOW_V2_SCHEMA_JSON,
    },
    RawSchema {
        id: "officegen.error.catalog@1.2",
        aliases: &["error.catalog", "error-catalog", "errors", "errors@1.2"],
        path: "schemas/error-catalog-1.2.schema.json",
        json: ERROR_CATALOG_SCHEMA_JSON,
    },
];

pub fn list_schemas() -> Vec<SchemaCatalogEntry> {
    RAW_SCHEMAS
        .iter()
        .map(|schema| SchemaCatalogEntry {
            id: schema.id,
            aliases: schema.aliases,
            path: schema.path,
        })
        .collect()
}

#[allow(dead_code)]
pub fn resolve_alias(id_or_alias: &str) -> Option<&'static str> {
    raw_schema(id_or_alias).map(|schema| schema.id)
}

#[allow(dead_code)]
pub fn get_schema(id_or_alias: &str) -> Option<Value> {
    fetch_schema(id_or_alias)
        .ok()
        .map(|document| document.schema)
}

pub fn fetch_schema(id_or_alias: &str) -> Result<SchemaDocument, SchemaCatalogError> {
    let schema = raw_schema(id_or_alias)
        .ok_or_else(|| SchemaCatalogError::UnknownSchema(id_or_alias.to_string()))?;
    let parsed = parse_schema(schema)?;
    Ok(SchemaDocument {
        id: schema.id,
        path: schema.path,
        schema: parsed,
    })
}

pub fn validate_minimal_required_fields(id_or_alias: &str, value: &Value) -> ValidationReport {
    let Some(schema) = raw_schema(id_or_alias) else {
        return ValidationReport {
            schema_id: "unknown",
            ok: false,
            errors: vec![ValidationError {
                instance_path: String::new(),
                message: format!("unknown schema: {id_or_alias}"),
            }],
        };
    };

    let schema_doc = match parse_schema(schema) {
        Ok(schema_doc) => schema_doc,
        Err(error) => {
            return ValidationReport {
                schema_id: schema.id,
                ok: false,
                errors: vec![ValidationError {
                    instance_path: String::new(),
                    message: error.to_string(),
                }],
            };
        }
    };

    if !value.is_object() {
        return ValidationReport {
            schema_id: schema.id,
            ok: false,
            errors: vec![ValidationError {
                instance_path: String::new(),
                message: "expected object".to_string(),
            }],
        };
    }

    let mut errors = Vec::new();
    validate_value("", value, &schema_doc, &mut errors);
    validate_schema_specific_contracts(schema.id, value, &mut errors);

    ValidationReport {
        schema_id: schema.id,
        ok: errors.is_empty(),
        errors,
    }
}

fn validate_schema_specific_contracts(
    schema_id: &str,
    value: &Value,
    errors: &mut Vec<ValidationError>,
) {
    if schema_id != "officegen.edit.ops@1.2" {
        return;
    }
    let operations = value.get("operations").and_then(Value::as_array);
    let ops_alias = value.get("ops").and_then(Value::as_array);
    if let (Some(left), Some(right)) = (operations, ops_alias) {
        if left != right {
            errors.push(ValidationError {
                instance_path: String::new(),
                message: "operations and ops must be identical when both are present".to_string(),
            });
        }
    }
}

fn raw_schema(id_or_alias: &str) -> Option<&'static RawSchema> {
    let requested = id_or_alias.trim();
    if let Some(schema) = RAW_SCHEMAS
        .iter()
        .find(|schema| schema.id == requested || schema.aliases.contains(&requested))
    {
        return Some(schema);
    }
    RAW_SCHEMAS
        .iter()
        .rev()
        .find(|schema| matches_schema_name(schema, requested))
}

fn matches_schema_name(schema: &RawSchema, requested: &str) -> bool {
    let id_without_version = schema.id.split('@').next().unwrap_or(schema.id);
    let short_id = id_without_version
        .strip_prefix("officegen.")
        .unwrap_or(id_without_version);
    requested == id_without_version
        || requested == short_id
        || requested == short_id.replace('.', "-")
}

fn parse_schema(schema: &RawSchema) -> Result<Value, SchemaCatalogError> {
    serde_json::from_str(schema.json).map_err(|error| SchemaCatalogError::InvalidCatalogSchema {
        id: schema.id,
        message: error.to_string(),
    })
}

fn validate_value(path: &str, value: &Value, schema: &Value, errors: &mut Vec<ValidationError>) {
    if let Some(any_of) = schema.get("anyOf").and_then(Value::as_array) {
        let matches_any_branch = any_of.iter().any(|branch| {
            let mut branch_errors = Vec::new();
            validate_value(path, value, branch, &mut branch_errors);
            branch_errors.is_empty()
        });
        if !matches_any_branch {
            errors.push(ValidationError {
                instance_path: path.to_string(),
                message: "expected value to match at least one anyOf schema".to_string(),
            });
        }
    }

    if let Some(expected) = schema.get("const") {
        if value != expected {
            errors.push(ValidationError {
                instance_path: path.to_string(),
                message: format!("expected const {}", json_literal(expected)),
            });
        }
    }

    if let Some(expected_type) = schema.get("type") {
        if !type_accepts(expected_type, value) {
            errors.push(ValidationError {
                instance_path: path.to_string(),
                message: format!(
                    "expected type {}, got {}",
                    expected_type_label(expected_type),
                    json_type(value)
                ),
            });
        }
    }

    if let Some(enum_values) = schema.get("enum").and_then(Value::as_array) {
        if !enum_values.iter().any(|allowed| allowed == value) {
            errors.push(ValidationError {
                instance_path: path.to_string(),
                message: format!(
                    "expected one of {}",
                    json_literal(schema.get("enum").unwrap())
                ),
            });
        }
    }

    if let Some(min_items) = schema.get("minItems").and_then(Value::as_u64) {
        if value
            .as_array()
            .map(|items| items.len() as u64)
            .unwrap_or(min_items)
            < min_items
        {
            errors.push(ValidationError {
                instance_path: path.to_string(),
                message: format!("expected at least {min_items} item(s)"),
            });
        }
    }

    if let (Some(text), Some(min_length)) = (
        value.as_str(),
        schema.get("minLength").and_then(Value::as_u64),
    ) {
        if text.chars().count() < min_length as usize {
            errors.push(ValidationError {
                instance_path: path.to_string(),
                message: format!("expected string length at least {min_length}"),
            });
        }
    }

    if let (Some(text), Some(max_length)) = (
        value.as_str(),
        schema.get("maxLength").and_then(Value::as_u64),
    ) {
        if text.chars().count() > max_length as usize {
            errors.push(ValidationError {
                instance_path: path.to_string(),
                message: format!("expected string length at most {max_length}"),
            });
        }
    }

    if let (Some(text), Some(pattern)) = (
        value.as_str(),
        schema.get("pattern").and_then(Value::as_str),
    ) {
        match Regex::new(pattern) {
            Ok(regex) if !regex.is_match(text) => errors.push(ValidationError {
                instance_path: path.to_string(),
                message: format!("expected string matching pattern {pattern}"),
            }),
            Err(error) => errors.push(ValidationError {
                instance_path: path.to_string(),
                message: format!("invalid schema pattern {pattern}: {error}"),
            }),
            _ => {}
        }
    }

    if let (Some(number), Some(minimum)) = (
        value.as_f64(),
        schema.get("minimum").and_then(Value::as_f64),
    ) {
        if number < minimum {
            errors.push(ValidationError {
                instance_path: path.to_string(),
                message: format!("expected number >= {minimum}"),
            });
        }
    }

    if let (Some(number), Some(maximum)) = (
        value.as_f64(),
        schema.get("maximum").and_then(Value::as_f64),
    ) {
        if number > maximum {
            errors.push(ValidationError {
                instance_path: path.to_string(),
                message: format!("expected number <= {maximum}"),
            });
        }
    }

    if let Some(object) = value.as_object() {
        let properties = schema
            .get("properties")
            .and_then(Value::as_object)
            .cloned()
            .unwrap_or_default();

        for (field, field_value) in object {
            let field_path = child_path(path, field);
            if let Some(field_schema) = properties.get(field) {
                validate_value(&field_path, field_value, field_schema, errors);
            } else if schema.get("additionalProperties").and_then(Value::as_bool) == Some(false) {
                errors.push(ValidationError {
                    instance_path: field_path,
                    message: format!("additional property is not allowed: {field}"),
                });
            }
        }

        if let Some(required) = schema.get("required").and_then(Value::as_array) {
            for field in required.iter().filter_map(Value::as_str) {
                let field_path = child_path(path, field);
                if !object.contains_key(field) {
                    errors.push(ValidationError {
                        instance_path: field_path,
                        message: format!("missing required field: {field}"),
                    });
                }
            }
        }
    }

    if let (Some(items), Some(item_schema)) = (value.as_array(), schema.get("items")) {
        for (index, item) in items.iter().enumerate() {
            validate_value(
                &child_path(path, &index.to_string()),
                item,
                item_schema,
                errors,
            );
        }
    }
}

fn child_path(parent: &str, child: &str) -> String {
    if parent.is_empty() {
        format!("/{child}")
    } else {
        format!("{parent}/{child}")
    }
}

fn type_accepts(expected_type: &Value, value: &Value) -> bool {
    match expected_type {
        Value::String(expected) => {
            json_type(value) == expected
                || (expected == "number" && matches!(value, Value::Number(_)))
                || (expected == "integer"
                    && value
                        .as_i64()
                        .or_else(|| value.as_u64().map(|number| number as i64))
                        .is_some())
        }
        Value::Array(types) => types.iter().any(|expected| type_accepts(expected, value)),
        _ => true,
    }
}

fn json_type(value: &Value) -> &'static str {
    match value {
        Value::Null => "null",
        Value::Bool(_) => "boolean",
        Value::Number(number) if number.is_i64() || number.is_u64() => "integer",
        Value::Number(_) => "number",
        Value::String(_) => "string",
        Value::Array(_) => "array",
        Value::Object(_) => "object",
    }
}

fn expected_type_label(expected_type: &Value) -> String {
    match expected_type {
        Value::String(value) => value.clone(),
        Value::Array(values) => values
            .iter()
            .filter_map(Value::as_str)
            .collect::<Vec<_>>()
            .join("|"),
        other => json_literal(other),
    }
}

fn json_literal(value: &Value) -> String {
    serde_json::to_string(value).unwrap_or_else(|_| "<json>".to_string())
}

impl fmt::Display for SchemaCatalogError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            SchemaCatalogError::UnknownSchema(id) => write!(formatter, "unknown schema: {id}"),
            SchemaCatalogError::InvalidCatalogSchema { id, message } => {
                write!(formatter, "invalid catalog schema {id}: {message}")
            }
        }
    }
}

impl std::error::Error for SchemaCatalogError {}
