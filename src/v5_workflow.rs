#![allow(dead_code)]

use crate::registry;
use crate::safety;
use anyhow::{anyhow, bail, Context as AnyhowContext, Result};
use serde_json::{json, Value};
use std::ffi::OsStr;
use std::fs;
use std::path::{Path, PathBuf};
use std::process::Command;
use std::time::{SystemTime, UNIX_EPOCH};

const VERSION: &str = env!("CARGO_PKG_VERSION");
const OUTPUT_FLAGS: &[&str] = &["--out"];

pub fn run_workflow(cwd: &Path, args: &[String]) -> Result<Value> {
    let workflow_arg = first_workflow_input(args)
        .ok_or_else(|| anyhow!("INPUT_REQUIRED: run requires workflow.json"))?;
    let workflow_path = safety::resolve_input_path(cwd, &workflow_arg)?;
    let workflow_text = fs::read_to_string(&workflow_path)
        .with_context(|| format!("failed to read workflow {}", redacted(&workflow_path)))?;
    let workflow: Value = serde_json::from_str(&workflow_text).map_err(|error| {
        anyhow!(
            "SCHEMA_INVALID: failed to parse workflow {}: {error}",
            redacted(&workflow_path)
        )
    })?;

    if workflow.get("schema").and_then(Value::as_str) != Some("officegen.workflow@2.0") {
        bail!("SCHEMA_INVALID: run requires officegen.workflow@2.0");
    }

    let output_root_arg = option_value(args, "--output-root")
        .or_else(|| option_value(args, "--out"))
        .or_else(|| {
            workflow
                .get("outputRoot")
                .and_then(Value::as_str)
                .map(str::to_string)
        })
        .unwrap_or_else(|| default_output_root(&workflow_path));
    let output_root = safety::resolve_output_path(cwd, &output_root_arg)?;

    let steps = workflow
        .get("steps")
        .and_then(Value::as_array)
        .ok_or_else(|| anyhow!("SCHEMA_INVALID: workflow steps must be an array"))?;
    if steps.is_empty() {
        bail!("SCHEMA_INVALID: workflow steps must not be empty");
    }

    fs::create_dir_all(&output_root).with_context(|| {
        format!(
            "failed to create workflow output root {}",
            redacted(&output_root)
        )
    })?;

    let mut trace_steps = Vec::new();
    let mut stopped = false;
    let mut failed_step = Value::Null;
    let mut failure_message = None;

    for (index, step) in steps.iter().enumerate() {
        let step_id = step
            .get("id")
            .and_then(Value::as_str)
            .ok_or_else(|| anyhow!("SCHEMA_INVALID: workflow step {index} is missing id"))?;
        let command = step
            .get("command")
            .and_then(Value::as_str)
            .ok_or_else(|| anyhow!("SCHEMA_INVALID: workflow step {step_id} is missing command"))?;
        let raw_args = step_args(step)?;
        let started_at = unix_millis();

        let step_result = prepare_step_args(cwd, &output_root, command, &raw_args)
            .and_then(|prepared| execute_step(cwd, prepared));
        let finished_at = unix_millis();

        match step_result {
            Ok(executed) => {
                let ok = executed.ok;
                trace_steps.push(json!({
                    "id": step_id,
                    "index": index,
                    "command": executed.command,
                    "args": executed.args,
                    "status": if ok { "pass" } else { "fail" },
                    "exitCode": executed.exit_code,
                    "startedAtUnixMs": started_at,
                    "finishedAtUnixMs": finished_at,
                    "durationMs": finished_at.saturating_sub(started_at),
                    "envelope": executed.envelope,
                    "stderr": executed.stderr
                }));
                if !ok {
                    stopped = true;
                    failed_step = json!(step_id);
                    failure_message = Some(format!("WORKFLOW_STEP_FAILED: step {step_id} failed"));
                    break;
                }
            }
            Err(error) => {
                stopped = true;
                failed_step = json!(step_id);
                let detail = error.to_string();
                failure_message = Some(format!(
                    "WORKFLOW_STEP_FAILED: step {step_id} failed before execution: {detail}"
                ));
                trace_steps.push(json!({
                    "id": step_id,
                    "index": index,
                    "command": command,
                    "args": raw_args,
                    "status": "fail",
                    "exitCode": Value::Null,
                    "startedAtUnixMs": started_at,
                    "finishedAtUnixMs": finished_at,
                    "durationMs": finished_at.saturating_sub(started_at),
                    "error": detail
                }));
                break;
            }
        }
    }

    let ok = !stopped;
    let trace_path = output_root.join("trace.json");
    let summary_path = output_root.join("summary.json");
    let manifest_path = output_root.join("manifest.json");

    let trace = json!({
        "schema": "officegen.workflow.trace@2.0",
        "workflowSchema": workflow["schema"],
        "workflow": redacted(&workflow_path),
        "ok": ok,
        "stopped": stopped,
        "failedStep": failed_step,
        "steps": trace_steps
    });
    let completed_steps = trace["steps"]
        .as_array()
        .map(|steps| steps.iter().filter(|step| step["status"] == "pass").count())
        .unwrap_or_default();
    let document_changed = workflow_document_changed(&trace);
    let summary = json!({
        "schema": "officegen.workflow.summary@2.0",
        "ok": ok,
        "status": if ok { "pass" } else { "fail" },
        "workflow": redacted(&workflow_path),
        "outputRoot": redacted(&output_root),
        "stepCount": steps.len(),
        "completedSteps": completed_steps,
        "failedStep": trace["failedStep"],
        "message": failure_message.clone().unwrap_or_else(|| "workflow completed".to_string())
    });
    let mut artifacts = vec![
        declared_artifact(&manifest_path, "manifest", "json"),
        declared_artifact(&trace_path, "trace", "json"),
        declared_artifact(&summary_path, "summary", "json"),
    ];
    artifacts.extend(step_artifacts(&trace));
    let manifest = json!({
        "schema": "officegen.manifest@1.2",
        "version": VERSION,
        "createdAtUnixMs": unix_millis(),
        "artifacts": artifacts,
        "metadata": {
            "kind": "workflow-run",
            "workflowSchema": workflow["schema"],
            "workflow": redacted(&workflow_path),
            "outputRoot": redacted(&output_root)
        }
    });

    write_json(&trace_path, &trace)?;
    write_json(&summary_path, &summary)?;
    write_json(&manifest_path, &manifest)?;

    Ok(json!({
        "schema": "officegen.workflow.run.result@2.0",
        "ok": ok,
        "status": if ok { "pass" } else { "fail" },
        "changed": document_changed,
        "workflowArtifactsWritten": true,
        "documentChanged": document_changed,
        "execution": "sequential",
        "transport": "cli-json-file",
        "mcp": false,
        "workflowSchema": workflow["schema"],
        "outputRoot": redacted(&output_root),
        "manifest": declared_artifact(&manifest_path, "manifest", "json"),
        "trace": declared_artifact(&trace_path, "trace", "json"),
        "summary": declared_artifact(&summary_path, "summary", "json"),
        "artifacts": artifacts,
        "stepCount": steps.len(),
        "completedSteps": completed_steps,
        "failedStep": trace["failedStep"],
        "message": failure_message.unwrap_or_else(|| "workflow completed".to_string())
    }))
}

struct ExecutedStep {
    command: String,
    args: Vec<String>,
    ok: bool,
    exit_code: Option<i32>,
    envelope: Value,
    stderr: String,
}

fn prepare_step_args(
    cwd: &Path,
    output_root: &Path,
    command: &str,
    raw_args: &[String],
) -> Result<Vec<String>> {
    let mut prepared = command
        .split_whitespace()
        .map(str::to_string)
        .chain(raw_args.iter().cloned())
        .collect::<Vec<_>>();
    if prepared.is_empty() {
        bail!("SCHEMA_INVALID: workflow step command must not be empty");
    }
    if prepared[0] == "run" {
        bail!("WORKFLOW_RECURSION_DENIED: workflow steps cannot invoke run");
    }
    if prepared[0] == "mcp" {
        bail!("FEATURE_REMOVED_FROM_SCOPE: MCP is intentionally outside workflow scope");
    }
    if registry::find_command(&command_text(&prepared)).is_none() {
        bail!("UNKNOWN_COMMAND: {}", command_text(&prepared));
    }
    if mutating_step_requires_output(&prepared)
        && !has_option(&prepared, "--out")
        && !dry_run_can_omit_output(&prepared)
    {
        bail!("OUTPUT_REQUIRED: workflow mutating steps require --out inside outputRoot");
    }

    let mut index = 0;
    while index < prepared.len() {
        if OUTPUT_FLAGS.contains(&prepared[index].as_str()) {
            let Some(value) = prepared.get(index + 1).cloned() else {
                bail!("OUTPUT_REQUIRED: {} requires a value", prepared[index]);
            };
            let scoped = scoped_output_arg(cwd, output_root, &value)?;
            prepared[index + 1] = scoped;
            index += 2;
        } else {
            index += 1;
        }
    }

    if !prepared.iter().any(|arg| arg == "--agent") {
        prepared.push("--agent".to_string());
    }
    if !prepared.iter().any(|arg| arg == "--strict-json") {
        prepared.push("--strict-json".to_string());
    }
    Ok(prepared)
}

fn execute_step(cwd: &Path, prepared: Vec<String>) -> Result<ExecutedStep> {
    let exe = std::env::current_exe().context("failed to resolve current executable")?;
    let result = Command::new(&exe)
        .args(&prepared)
        .current_dir(cwd)
        .output()
        .with_context(|| format!("failed to execute {}", display_exe(&exe)))?;
    let stdout = String::from_utf8_lossy(&result.stdout);
    let stderr = String::from_utf8_lossy(&result.stderr).to_string();
    let envelope = serde_json::from_str::<Value>(&stdout)
        .with_context(|| "workflow step did not emit a JSON envelope")?;
    let envelope_ok = envelope.get("ok").and_then(Value::as_bool).unwrap_or(false);
    Ok(ExecutedStep {
        command: command_text(&prepared),
        args: prepared,
        ok: result.status.success() && envelope_ok,
        exit_code: result.status.code(),
        envelope,
        stderr,
    })
}

fn step_args(step: &Value) -> Result<Vec<String>> {
    let command = step.get("command").and_then(Value::as_str).unwrap_or("");
    let mut out = Vec::new();

    if command_text_parts(command).first().map(String::as_str) == Some("diff") {
        push_string_field(step, "before", &mut out)?;
        push_string_field(step, "after", &mut out)?;
    } else {
        push_string_field(step, "input", &mut out)?;
    }

    push_flag_value(step, "ops", "--ops", &mut out)?;
    push_flag_value(step, "data", "--data", &mut out)?;
    push_flag_value(step, "out", "--out", &mut out)?;
    push_flag_value(step, "target", "--target", &mut out)?;
    push_flag_value(step, "format", "--format", &mut out)?;
    push_flag_value(step, "sheet", "--sheet", &mut out)?;
    push_flag_value(step, "range", "--range", &mut out)?;
    push_flag_value(step, "slides", "--slides", &mut out)?;
    push_flag_value(step, "pages", "--pages", &mut out)?;
    push_flag_value(step, "schema", "--schema", &mut out)?;
    push_flag_value(step, "manifest", "--manifest", &mut out)?;
    push_flag_value(step, "trace", "--trace", &mut out)?;
    push_flag_value(step, "summary", "--summary", &mut out)?;

    push_bool_flag(step, "dryRun", "--dry-run", &mut out)?;
    push_bool_flag(step, "resolveSelectors", "--resolve-selectors", &mut out)?;
    push_bool_flag(step, "visual", "--visual", &mut out)?;

    if let Some(args) = step.get("args") {
        out.extend(
            args.as_array()
                .ok_or_else(|| anyhow!("SCHEMA_INVALID: workflow step args must be an array"))?
                .iter()
                .map(|value| {
                    value.as_str().map(str::to_string).ok_or_else(|| {
                        anyhow!("SCHEMA_INVALID: workflow step args must be strings")
                    })
                })
                .collect::<Result<Vec<_>>>()?,
        );
    }

    Ok(out)
}

fn command_text_parts(command: &str) -> Vec<String> {
    command.split_whitespace().map(str::to_string).collect()
}

fn push_string_field(step: &Value, field: &str, out: &mut Vec<String>) -> Result<()> {
    if let Some(value) = step.get(field) {
        let value = value
            .as_str()
            .ok_or_else(|| anyhow!("SCHEMA_INVALID: workflow step {field} must be a string"))?;
        out.push(value.to_string());
    }
    Ok(())
}

fn push_flag_value(step: &Value, field: &str, flag: &str, out: &mut Vec<String>) -> Result<()> {
    if let Some(value) = step.get(field) {
        let value = value
            .as_str()
            .ok_or_else(|| anyhow!("SCHEMA_INVALID: workflow step {field} must be a string"))?;
        out.push(flag.to_string());
        out.push(value.to_string());
    }
    Ok(())
}

fn push_bool_flag(step: &Value, field: &str, flag: &str, out: &mut Vec<String>) -> Result<()> {
    if let Some(value) = step.get(field) {
        let enabled = value
            .as_bool()
            .ok_or_else(|| anyhow!("SCHEMA_INVALID: workflow step {field} must be a boolean"))?;
        if enabled {
            out.push(flag.to_string());
        }
    }
    Ok(())
}

fn scoped_output_arg(cwd: &Path, output_root: &Path, value: &str) -> Result<String> {
    let scoped_path = safety::resolve_output_path(output_root, value)?;
    let cwd = cwd.canonicalize().context("failed to canonicalize cwd")?;
    if !scoped_path.starts_with(output_root) {
        bail!("SECURITY_PATH_OUTSIDE_ROOT: workflow output resolves outside outputRoot");
    }
    let relative = scoped_path.strip_prefix(&cwd).map_err(|_| {
        anyhow!("SECURITY_PATH_OUTSIDE_ROOT: workflow output root must be inside cwd")
    })?;
    Ok(relative.to_string_lossy().replace('\\', "/"))
}

fn mutating_step_requires_output(args: &[String]) -> bool {
    let command = command_text(args);
    matches!(
        command.as_str(),
        "scaffold"
            | "render"
            | "view"
            | "edit"
            | "export"
            | "prepare"
            | "asset extract"
            | "chart render"
            | "diagram render"
    )
}

fn dry_run_can_omit_output(args: &[String]) -> bool {
    command_text(args) == "edit" && has_option(args, "--dry-run")
}

fn step_artifacts(trace: &Value) -> Vec<Value> {
    trace
        .get("steps")
        .and_then(Value::as_array)
        .into_iter()
        .flatten()
        .flat_map(|step| {
            let out_arg = output_arg_from_step(step);
            step.pointer("/envelope/artifacts")
                .and_then(Value::as_array)
                .or_else(|| {
                    step.pointer("/envelope/result/artifacts")
                        .and_then(Value::as_array)
                })
                .into_iter()
                .flatten()
                .cloned()
                .map(move |artifact| normalize_step_artifact_path(artifact, out_arg.as_deref()))
        })
        .collect()
}

fn output_arg_from_step(step: &Value) -> Option<String> {
    let args = step.get("args")?.as_array()?;
    args.windows(2).find_map(|pair| {
        if pair[0].as_str() == Some("--out") {
            pair[1].as_str().map(str::to_string)
        } else {
            None
        }
    })
}

fn normalize_step_artifact_path(mut artifact: Value, out_arg: Option<&str>) -> Value {
    let Some(out_arg) = out_arg else {
        return artifact;
    };
    let Some(path) = artifact.get("path").and_then(Value::as_str) else {
        return artifact;
    };
    let out_name = Path::new(out_arg)
        .file_name()
        .and_then(OsStr::to_str)
        .unwrap_or(out_arg);
    if path == out_name {
        artifact["path"] = json!(out_arg.replace('\\', "/"));
    }
    artifact
}

fn workflow_document_changed(trace: &Value) -> bool {
    trace
        .get("steps")
        .and_then(Value::as_array)
        .into_iter()
        .flatten()
        .filter(|step| step["status"] == "pass")
        .any(|step| {
            let envelope = &step["envelope"];
            envelope["mutationStatus"] == "changed"
                || envelope["result"]["changed"].as_bool().unwrap_or(false)
                || envelope["result"]["documentChanged"]
                    .as_bool()
                    .unwrap_or(false)
        })
}

fn first_workflow_input(args: &[String]) -> Option<String> {
    positionals(args).get(1).cloned()
}

fn positionals(args: &[String]) -> Vec<String> {
    let mut out = Vec::new();
    let mut iter = args.iter().skip(1).peekable();
    while let Some(arg) = iter.next() {
        if arg.starts_with("--") {
            if option_takes_value(arg) {
                iter.next();
            }
            continue;
        }
        if arg.starts_with('-') {
            continue;
        }
        out.push(arg.clone());
    }
    out
}

fn option_value(args: &[String], flag: &str) -> Option<String> {
    args.windows(2).find_map(|pair| {
        if pair[0] == flag {
            Some(pair[1].clone())
        } else {
            None
        }
    })
}

fn has_option(args: &[String], flag: &str) -> bool {
    args.iter().any(|arg| arg == flag)
}

fn option_takes_value(arg: &str) -> bool {
    !matches!(
        arg,
        "--json"
            | "--agent"
            | "--strict-json"
            | "--help"
            | "--version"
            | "--dry-run"
            | "--resolve-selectors"
            | "--overwrite"
            | "--visual"
            | "--native"
            | "--embedded"
            | "--images"
            | "--summary-only"
            | "--source-only"
            | "--plan"
            | "--no-object-map"
            | "--compact"
            | "--supported-only"
            | "--full"
    )
}

fn command_text(args: &[String]) -> String {
    if args.is_empty() {
        return String::new();
    }
    let first = args[0].as_str();
    let second = args.get(1).map(String::as_str).unwrap_or("");
    let two = format!("{first} {second}");
    if registry::find_command(&two).is_some() {
        two
    } else {
        first.to_string()
    }
}

fn default_output_root(workflow_path: &Path) -> String {
    let stem = workflow_path
        .file_stem()
        .and_then(OsStr::to_str)
        .unwrap_or("workflow");
    format!(".officegen/workflows/{stem}")
}

fn write_json(path: &Path, value: &Value) -> Result<()> {
    safety::atomic_write(path, &serde_json::to_vec_pretty(value)?)?;
    Ok(())
}

fn artifact(path: &Path, kind: &str, format: &str) -> Value {
    json!({"path": redacted(path), "kind": kind, "format": format, "exists": path.exists()})
}

fn declared_artifact(path: &Path, kind: &str, format: &str) -> Value {
    json!({"path": redacted(path), "kind": kind, "format": format, "exists": true})
}

fn unix_millis() -> u128 {
    SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|duration| duration.as_millis())
        .unwrap_or_default()
}

fn redacted(path: &Path) -> String {
    path.file_name()
        .and_then(OsStr::to_str)
        .unwrap_or("<path>")
        .to_string()
}

fn display_exe(path: &PathBuf) -> String {
    path.file_name()
        .and_then(OsStr::to_str)
        .unwrap_or("officegen")
        .to_string()
}
