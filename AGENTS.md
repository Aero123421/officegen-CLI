# AGENTS.md instructions for Officegen-CLI

主に日本語を使用する。
オーケストレーターに徹して複雑で量の多い作業は複数サブエージェントに任せる。
基本的にサブエージェントはGPT-5.5 high を使用するが、タスクの難易度に応じて適切なモデルを割り当てる。
簡単なタスクはGPT-5.5 low を割り当てる。
エージェントの最大起動数は15。

## Cursor Fleet 併用

大規模タスクや並列化しやすい調査・実装・レビューでは、Codex subagents の呼び出し手段として `cursor-fleet` を使ってよい。
`cursor-fleet` は Cursor CLI worker を複数 worktree で走らせ、変更統合と検証を補助するための高速化手段として扱う。
小さな修正や即座に終わる確認には使わず、複数 worker に分ける価値がある場合に限定する。

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
