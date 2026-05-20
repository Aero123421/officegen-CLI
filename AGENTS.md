# AGENTS.md instructions for Officegen-CLI

主に日本語を使用する。
オーケストレーターに徹して複雑で量の多い作業は複数サブエージェントに任せる。
基本的にサブエージェントはGPT-5.5 high を使用するが、タスクの難易度に応じて適切なモデルを割り当てる。
簡単なタスクはGPT-5.5 low を割り当てる。
エージェントの最大起動数は15。

## bigfaster-worker 併用

大規模タスクや実装速度を優先したい作業では、Codex subagent として `bigfaster-worker` を使ってよい。
定義は `.codex/agents/bigfaster-worker.toml` に置き、元リポジトリは https://github.com/Aero123421/bigfaster-worker 。
小さな修正や即座に終わる確認には使わず、Cursor CLI に実装試行を任せる価値がある場合に限定する。
関連する複数修正が同じ package や機能領域に触れる場合は、原則として1つの `bigfaster-worker` 呼び出しにまとめ、明確な ownership boundary と既知の検証コマンドを渡す。
複数の `bigfaster-worker` を並列に呼ぶのは、ユーザーが並列実装・比較案・明確に分離された担当範囲を求めた場合に限定する。
worker は Cursor CLI の headless 既定形 `agent --model auto --print --trust -- "<TASK_PROMPT>"` を使い、dirty tree・リスクの高い作業・並列作業では `--worktree <WORKTREE_NAME>` を使う。
Cursor の worktree 差分は盲目的にマージせず、意図したファイルだけ確認して元リポジトリへ反映し、生成物・lockfile・version・docs は明示スコープ外なら変更しない。

## Version / Release 運用

version は手作業で複数ファイルを直接編集しない。

version を上げるときは必ず以下を使う。

```bash
npm run version:bump -- patch
npm run version:bump -- minor
npm run version:bump -- major
npm run version:bump -- 1.2.7
```

このスクリプトは root/workspace package manifests、`package-lock.json`、`OFFICEGEN_CLI_VERSION`、README の release URL 例をまとめて同期する。

リリース前、タグ作成前、version まわりを触った後は必ず以下を実行する。

```bash
npm run version:check
npm run typecheck
npm test
npm run build
npm run pack:smoke
```

`npm run version:check` が失敗する状態で release tag を切らない。

GitHub Release は `vX.Y.Z` tag または手動 workflow で作るが、workflow は manifest version と release version の一致を検査する。ズレた場合は workflow が失敗するため、先に `npm run version:bump -- <version>` を実行してから commit/push/tag する。
