import type { CapabilityEntry, FeatureKey } from "./types.js";

export const COMMAND_METADATA: CapabilityEntry[] = [
  meta("capabilities", "有効機能とAgent可視機能", ["capabilities"]),
  meta("help", "動的help", ["help", "help workflow", "help error"]),
  meta("config", "config確認・設定", ["config show", "config set"]),
  meta("doctor", "環境確認", ["doctor"]),
  meta("schema", "schema取得・検証・migration", ["schema list", "schema get", "schema validate", "schema migrate"]),
  meta("errors", "エラーカタログ", ["errors list", "errors inspect"]),
  core("inspect", "既存ファイル解析", ["inspect"]),
  core("view", "SVG/HTMLプレビューとobject map", ["view"]),
  core("edit", "EditOpsで既存ファイル編集", ["edit"]),
  core("render", "IR/Specから新規ファイル生成", ["render"]),
  core("scaffold", "LLMなしの雛形IR/ops/data生成", ["scaffold"]),
  core("export", "形式変換", ["export"]),
  core("validate", "スキーマ・構造・品質検証", ["validate"]),
  core("diagnose", "問題検出", ["diagnose"]),
  core("repair", "修復または修復案生成", ["repair"]),
  core("run", "複合workflow実行", ["run"]),
  core("asset", "画像・添付物・メディア", ["asset add", "asset inspect", "asset extract", "asset replace"]),
  core("chart", "グラフ", ["chart render"]),
  core("diagram", "図解", ["diagram render"]),
  optional("template", "テンプレート作成・充填", [
    "template list",
    "template inspect",
    "template candidates",
    "template create",
    "template apply-map",
    "template validate",
    "template fill"
  ]),
  optional("design", "デザイン知識抽出・保存・適用", [
    "design list",
    "design inspect",
    "design init",
    "design edit",
    "design update",
    "design validate",
    "design capture",
    "design apply"
  ]),
  optional("layout", "自動レイアウト", ["layout apply"]),
  optional("agent", "Agent adapter生成", ["agent install", "agent refresh"]),
  optional("mcp", "MCP server", ["mcp serve"], true),
  optional("renderer", "外部レンダラー管理", ["renderer list", "renderer inspect", "renderer trust"], true),
  optional("plugin", "plugin管理", ["plugin list", "plugin inspect", "plugin install", "plugin trust"], true)
];

export function metadataFor(feature: FeatureKey): CapabilityEntry | undefined {
  return COMMAND_METADATA.find((entry) => entry.feature === feature);
}

function meta(feature: FeatureKey, description: string, commands: string[]): CapabilityEntry {
  return entry(feature, description, commands, false, false);
}

function core(feature: FeatureKey, description: string, commands: string[]): CapabilityEntry {
  return entry(feature, description, commands, false, false);
}

function optional(feature: FeatureKey, description: string, commands: string[], externalProcess = false): CapabilityEntry {
  return entry(feature, description, commands, false, externalProcess);
}

function entry(
  feature: FeatureKey,
  description: string,
  commands: string[],
  network: boolean,
  externalProcess: boolean
): CapabilityEntry {
  return {
    feature,
    moduleId: `officegen.core.${feature}`,
    commandGroup: feature,
    description,
    stability: "stable",
    commands,
    requires: [],
    security: {
      network,
      externalProcess
    }
  };
}
