#[path = "../src/registry.rs"]
mod registry;

use registry::{
    command_registry, compact_agent_visible_commands, find_command, human_visible_commands,
    CommandStatus, OfficeFormat,
};

#[test]
fn compact_agent_visible_commands_omit_mcp_and_deferred_commands() {
    let compact = compact_agent_visible_commands();
    let commands = compact
        .iter()
        .map(|entry| entry.command)
        .collect::<Vec<_>>();

    assert!(commands.contains(&"inspect"));
    assert!(commands.contains(&"edit"));
    assert!(commands.contains(&"schema list"));
    assert!(commands.contains(&"errors list"));
    assert!(!commands.contains(&"plan"));
    assert!(!commands.contains(&"template candidates"));
    assert!(!commands.contains(&"mcp serve"));
    assert!(!commands.contains(&"template create"));
    assert!(!commands.contains(&"design edit"));
    assert!(!commands.contains(&"agent install"));
    assert!(!commands.contains(&"renderer trust"));
}

#[test]
fn compact_agent_visible_commands_only_expose_supported_limited_and_plan_only() {
    let compact = compact_agent_visible_commands();

    assert!(!compact.is_empty());
    assert!(compact.iter().all(|entry| matches!(
        entry.status,
        CommandStatus::Supported | CommandStatus::Limited | CommandStatus::PlanOnly
    )));
    assert!(compact
        .iter()
        .any(|entry| entry.status == CommandStatus::Supported));
    assert!(compact
        .iter()
        .any(|entry| entry.status == CommandStatus::Limited));
    assert!(compact.iter().all(|entry| !entry.summary.trim().is_empty()));
}

#[test]
fn mcp_is_removed_from_scope_and_not_visible_to_agents_or_humans() {
    let mcp = find_command("mcp serve").expect("mcp serve registry entry");

    assert_eq!(mcp.status, CommandStatus::RemovedFromScope);
    assert!(!mcp.agent_visible);
    assert!(!mcp.human_visible);
    assert!(!mcp.mutates_files);
    assert!(mcp.supported_formats.is_empty());
}

#[test]
fn registry_captures_mutation_dry_run_and_format_metadata() {
    let edit = find_command("edit").expect("edit registry entry");
    let plan = find_command("plan").expect("plan registry entry");
    let inspect = find_command("inspect").expect("inspect registry entry");

    assert_eq!(edit.status, CommandStatus::Limited);
    assert!(edit.mutates_files);
    assert!(edit.supports_dry_run);
    assert!(edit.supported_formats.contains(&OfficeFormat::Pptx));
    assert!(edit.supported_formats.contains(&OfficeFormat::Docx));
    assert!(edit.supported_formats.contains(&OfficeFormat::Xlsx));

    assert_eq!(plan.status, CommandStatus::PlanOnly);
    assert!(!plan.mutates_files);
    assert!(!plan.supports_dry_run);
    assert_eq!(plan.supported_formats, &[OfficeFormat::Json]);

    assert_eq!(inspect.status, CommandStatus::Supported);
    assert!(!inspect.mutates_files);
    assert!(!inspect.supports_dry_run);
    assert!(inspect.supported_formats.contains(&OfficeFormat::Pdf));
}

#[test]
fn public_registry_views_are_available_for_runtime_integration() {
    let all = command_registry();
    let human_visible = human_visible_commands();

    assert!(all.len() > human_visible.len());
    assert!(all.iter().any(|entry| entry.command == "mcp serve"));
    assert!(human_visible.iter().any(|entry| entry.command == "help"));
    assert!(!human_visible
        .iter()
        .any(|entry| entry.status == CommandStatus::RemovedFromScope));
}
