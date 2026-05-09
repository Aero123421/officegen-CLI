# AGENTS.md instructions for Officegen-CLI

主に日本語を使用する。

簡単なタスクは gpt5.3spark を割り当てる。

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
