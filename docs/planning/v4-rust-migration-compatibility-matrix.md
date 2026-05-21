# Officegen-CLI v4 Rust Migration Compatibility Matrix

Baseline: v3.1.8.

Scope: map the v3.1.8 command surface, compatibility features, and regression tests to the planned Rust module ownership and release gates. This is a planning matrix, not an implementation status report.

v4.5.0 update: this older parity matrix is superseded by the v4.5 contract-freeze plan for release decisions. MCP and plugin runtime support are removed from the built-in CLI scope, not policy-enabled optional surfaces. Keep this file only as a historical v3.1.8 parity reference.

Status values: `not-started`, `ported`, `parity-tested`, `deferred`, `removed-by-policy`.

## Gate Vocabulary

| Gate | Required evidence |
| --- | --- |
| RG-0 Surface parity | Rust CLI exposes the same enabled command groups, subcommands, global flags, command options, help JSON, capabilities hash behavior, and profile visibility as v3.1.8. |
| RG-1 Contract parity | Golden JSON envelope tests pass for success, warning, partial, blocked, unsupported, and unknown-command outcomes. |
| RG-2 Read-only format parity | Inspect, view, verify, diff, critique, asset inspect, chart render, diagram render, and benchmark read paths match v3.1.8 fixtures within documented tolerances. |
| RG-3 Mutation safety parity | Edit, render, export, repair, asset replace/extract, template/design/layout mutations preserve output policy, rollback, OOXML validation, risky-package blocking, and post-mutation verify warnings. |
| RG-4 Optional/external boundary | Native renderer and external-process commands remain disabled unless policy enables them; MCP/plugin runtime support is outside the built-in CLI scope. |
| RG-5 Release packaging | Version check, type/build equivalent, tests, pack smoke, schema coverage, contract check, remediation check, security fixtures, benchmark local fixtures, perfect-spec evidence, and release tarball smoke all pass for the v4 artifact. |

Before a v4 tag, run the v3 baseline gates as compatibility controls until Rust replacements exist: `npm run version:check`, `npm run typecheck`, `npm test`, `npm run build`, `npm run pack:smoke`, `npm run remediation:check`, `npm run contract:check`, `npm run schema:coverage`, `npm run security:fixtures`, `npm run benchmark:local-fixtures`, `npm run perfect-spec:evidence`, `npm run perfect-spec:check -- --gate=publish`, and release tarball smoke for the target artifact.

## Planned Rust Module Ownership

| Rust area | Owns | v3.1.8 source reference |
| --- | --- | --- |
| `officegen_cli::commands` | Command registration, option validation, help output, JSON stdout routing | `packages/cli/src/commands/register.ts`, `packages/cli/src/shared/metadata.ts` |
| `officegen_cli::payloads` | Command dispatch, envelope projection, field projection, report writes | `packages/cli/src/commands/payloads.ts`, `packages/cli/src/shared/envelope.ts` |
| `officegen_core` | Capabilities, config, schemas, errors, paths, redaction, run trace, zip safety | `packages/core/src/*.ts` |
| `officegen_ooxml` | OPC/ZIP graph, relationships, XML token index, patch engine, transaction, validator, PPTX/DOCX/XLSX semantic readers | `packages/formats/src/ooxml/*.ts` |
| `officegen_formats` | Inspect, view, edit, render, export, verify, diagnose, repair, diff, assets, charts, diagrams | `packages/formats/src/*.ts` |
| `officegen_pdf` | PDF object graph, PDF preview, PDF overlay/annotation safety, CJK font handling | `packages/formats/src/pdf/*`, `packages/formats/src/pdfFonts.ts` |
| `officegen_optional` | Historical v3 template/design/layout/renderer/agent references; plugin and MCP are removed from built-in v4.5 scope | `packages/optional/src/*.ts` |
| `officegen_benchmark` | Benchmark manifest policy, corpus run/compare reports | `scripts/benchmark-*.mjs`, `benchmarks/office-corpus/manifest.json` |
| `officegen_release` | Pack, smoke, install, perfect-spec evidence adapters for Rust artifact | `scripts/*smoke*.mjs`, `scripts/check-*.mjs` |

## Command Compatibility Matrix

| v3.1.8 command/feature | Rust owner | v3 parity tests | Release gates |
| --- | --- | --- | --- |
| `capabilities` | `officegen_core::capabilities`, `officegen_cli::payloads` | `packages/cli/test/program.test.ts`, `packages/core/src/config.test.ts` | RG-0, RG-1, RG-5 |
| `help`, `help workflow`, `help error` | `officegen_cli::commands` | `packages/cli/test/program.test.ts` help and workflow cases | RG-0, RG-1 |
| `config show`, `config set` | `officegen_core::config`, `officegen_cli::payloads` | `packages/cli/test/program.test.ts`, `packages/core/src/config.test.ts` | RG-0, RG-1, RG-4 |
| `doctor` | `officegen_cli::payloads`, `officegen_optional::renderer` discovery | `packages/cli/test/program.test.ts` doctor/native proof cases | RG-1, RG-4 |
| `schema list/get/fetch/validate/migrate` | `officegen_core::schemas` | `packages/cli/test/program.test.ts`, `packages/core/src/schemas.test.ts` | RG-0, RG-1, RG-5 |
| `errors list/inspect` | `officegen_core::errors` | `packages/cli/test/program.test.ts`, `packages/core/src/envelope-errors.test.ts` | RG-1 |
| `inspect` PPTX/DOCX/XLSX/PDF | `officegen_formats::inspect`, `officegen_ooxml`, `officegen_pdf` | `packages/formats/tests/formats.test.ts`, `inspect-workbook-map.test.ts`, `pptx-object-map.test.ts`, `docx-story-run-graph.test.ts`, `packages/cli/test/program.test.ts` | RG-1, RG-2 |
| `view` SVG/HTML/PNG/object crop | `officegen_formats::view`, `officegen_pdf` | `packages/cli/test/program.test.ts`, `packages/formats/tests/formats.test.ts`, `visual-diff.test.ts` | RG-2, RG-4 |
| `edit` JSON EditOps | `officegen_formats::edit`, `officegen_ooxml::{patch_engine,transaction,operations}` | `edit-office-xml-integration.test.ts`, `edit-transaction.test.ts`, `operation-registry.test.ts`, `packages/cli/test/program.test.ts` | RG-1, RG-3 |
| `render` IR to PPTX/DOCX/XLSX/PDF | `officegen_formats::render`, `officegen_pdf::fonts` | `packages/cli/test/program.test.ts`, `packages/formats/tests/formats.test.ts` | RG-2, RG-3 |
| `scaffold` | `officegen_cli::payloads`, `officegen_formats::render` IR helpers | `packages/cli/test/program.test.ts` scaffold cases | RG-1, RG-3 |
| `export` | `officegen_formats::export`, `officegen_optional::renderer` | `packages/cli/test/program.test.ts`, `packages/formats/tests/formats.test.ts` | RG-3, RG-4 |
| `validate` | `officegen_core::schemas`, `officegen_ooxml::validator` | `packages/core/src/schemas.test.ts`, `ooxml-validator.test.ts` | RG-1, RG-5 |
| `verify` | `officegen_formats::verify`, `officegen_ooxml::validator`, `officegen_optional::renderer` | `packages/cli/test/program.test.ts`, `ooxml-validator.test.ts`, `visual-diff.test.ts` | RG-2, RG-3, RG-4 |
| `diagnose` | `officegen_formats::diagnose`, `officegen_formats::verify` | `packages/cli/test/program.test.ts`, `packages/formats/tests/formats.test.ts` | RG-2 |
| `repair` | `officegen_formats::repair`, `officegen_formats::edit` | `packages/cli/test/program.test.ts`, `edit-office-xml-integration.test.ts` | RG-1, RG-3 |
| `diff` semantic/visual | `officegen_formats::diff`, `officegen_formats::visual_diff`, `officegen_ooxml` | `visual-diff.test.ts`, `packages/formats/tests/formats.test.ts`, `packages/cli/test/program.test.ts` | RG-2, RG-4 |
| `prepare`, `prepare reference` | `officegen_cli::payloads`, `officegen_formats::{inspect,view}` | `packages/cli/test/program.test.ts` prepare cases | RG-1, RG-2 |
| `manifest`, `manifest inspect`, `manifest verify` | `officegen_core::run`, `officegen_cli::payloads` | `packages/cli/test/program.test.ts` manifest/lock cases | RG-1, RG-5 |
| `select` | `officegen_formats::selector_graph`, `officegen_cli::payloads` | `packages/cli/test/program.test.ts` select cases | RG-1, RG-2 |
| `plan` | `officegen_cli::payloads`, `officegen_formats::edit` plan builder | `packages/cli/test/program.test.ts` Japanese plan/EditOps case | RG-1, RG-3 |
| `rollback` | `officegen_ooxml::transaction`, `officegen_cli::payloads` | `edit-transaction.test.ts`, `packages/cli/test/program.test.ts` | RG-3 |
| `lock` | `officegen_core::paths`, `officegen_cli::payloads` | `packages/cli/test/program.test.ts` scoped lock cases | RG-1 |
| `merge` | `officegen_cli::payloads`, future `officegen_formats::merge` | `packages/cli/test/program.test.ts` command-surface coverage | RG-0, RG-1 |
| `run`, `run prepare-reference`, `run office-edit`, `run office-agent` | `officegen_core::run`, `officegen_cli::payloads` | `packages/cli/test/program.test.ts`, `packages/core/src/run-trace.test.ts`, `paths-run.test.ts` | RG-1, RG-3, RG-5 |
| `critique` | `officegen_formats::diagnose`, `officegen_cli::payloads` | `packages/cli/test/program.test.ts`, `packages/formats/tests/formats.test.ts` | RG-2 |
| `improve` | `officegen_formats::diagnose`, `officegen_cli::payloads` | `packages/cli/test/program.test.ts` improve cases | RG-1, RG-2 |
| `benchmark run/compare` | `officegen_benchmark` | `packages/cli/test/program.test.ts`, `scripts/benchmark-local-fixtures.mjs` | RG-2, RG-5 |
| `asset inspect/extract/replace` | `officegen_formats::assets`, `officegen_ooxml` | `packages/cli/test/program.test.ts`, `packages/formats/tests/formats.test.ts` | RG-2, RG-3 |
| `chart render` | `officegen_formats::charts` | `packages/cli/test/program.test.ts`, `packages/formats/tests/formats.test.ts` | RG-2 |
| `diagram render` | `officegen_formats::diagrams` | `diagrams-render.test.ts`, `packages/cli/test/program.test.ts` | RG-2 |
| `template list/inspect/candidates/create/apply-map/validate/fill` | `officegen_optional::template` | `packages/optional/src/template-design.test.ts`, `packages/cli/test/program.test.ts` | RG-1, RG-3 |
| `design list/inspect/init/edit/update/validate/capture/apply` | `officegen_optional::design` | `packages/optional/src/template-design.test.ts`, `packages/cli/test/program.test.ts` | RG-1, RG-3 |
| `layout apply` | `officegen_optional::layout`, `officegen_formats::diff` | `packages/optional/src/layout.test.ts`, `packages/cli/test/program.test.ts` | RG-3 |
| `agent install/refresh` | `officegen_optional::agent` | `packages/cli/test/program.test.ts` command-surface coverage | RG-0, RG-4 |
| `renderer list/inspect/trust/doctor` | `officegen_optional::renderer` | `packages/cli/test/program.test.ts` renderer policy cases | RG-4 |
| `plugin list/inspect/install/trust` | removed from built-in runtime scope | Removal/negative smoke; historical v3 command registration only | RG-4 |
| `mcp serve` | removed from built-in runtime scope | Removal/negative smoke; historical v3 command registration only | RG-4 |

## v3.1.8 Compatibility Fix Matrix

| v3.1.8 ID | Compatibility requirement for v4 Rust | Rust owner | Required test/evidence | Gate |
| --- | --- | --- | --- | --- |
| V31-P001 | PPTX chart backing workbooks classify as chart workbook info, not OLE embedded objects. | `officegen_core::zip_safety`, `officegen_ooxml::package_graph` | Zip safety fixture plus asset/verify parity fixture | RG-2, RG-3 |
| V31-P002 | `asset inspect --embedded` and `verify` agree on chart workbook vs true embedded package counts. | `officegen_formats::assets`, `officegen_formats::verify` | Embedded PPTX asset fixture | RG-2 |
| V31-P003 | Post-edit package diffs report part counts, changed/removed/added parts, byte delta, and compression note. | `officegen_formats::edit`, `officegen_ooxml::transaction` | Edit integration fixture with package diff schema | RG-3, RG-5 |
| V31-P004 | Unknown/custom OOXML package parts survive supported PPTX text edits. | `officegen_ooxml::transaction`, `officegen_formats::edit` | Custom XML preservation fixture | RG-3 |
| V31-X001 | Scoped XLSX `--sheet/--range` summary matches scoped objectMap cell coverage. | `officegen_ooxml::xlsx`, `officegen_formats::inspect` | `inspect-workbook-map.test.ts` parity | RG-2 |
| V31-X002 | `xlsx.setFormula` resolves workbook relationship-targeted worksheets and existing formula text selectors. | `officegen_ooxml::xlsx`, `officegen_formats::edit` | XLSX formula selector fixture | RG-3 |
| V31-X003 | Older workbook path variants remain covered for formula selector guards. | `officegen_ooxml::relationships`, `officegen_ooxml::xlsx` | Relationship-targeted worksheet fixture | RG-3 |
| V31-D001 | DOCX render does not duplicate document title as repeated section headings. | `officegen_formats::render`, `officegen_ooxml::docx` | DOCX render fixture | RG-2, RG-3 |
| V31-D002 | DOCX comments count actual comment entries separately from `comments.xml` presence. | `officegen_ooxml::docx`, `officegen_formats::inspect` | DOCX comments fixture | RG-2 |
| V31-U001 | Recommended profile commands remain shell-safe on Windows PowerShell. | `officegen_cli::commands`, `officegen_cli::payloads` | Help/error/renderer guidance golden output on Windows | RG-0, RG-1 |
| V31-U002 | Error catalog suggestions keep `--agent --strict-json` in agent/strict contexts. | `officegen_core::errors`, `officegen_cli::payloads` | Error inspect golden output | RG-1 |
| V31-U003 | `doctor` reports warning readiness when native proof is unavailable by policy. | `officegen_cli::payloads`, `officegen_optional::renderer` | Doctor native proof policy fixture | RG-1, RG-4 |
| V31-U004 | Native renderer blocked messages include concrete next commands. | `officegen_optional::renderer`, `officegen_formats::export` | Renderer blocked guidance fixture | RG-4 |
| V31-U005 | Native proof remains optional-gated and never claimed universal. | `officegen_optional::renderer`, `officegen_formats::verify` | Native proof unavailable fixture | RG-4 |
| V31-R001 | `TEXT_OVERFLOW_RISK` repairs and top risks are de-duplicated. | `officegen_formats::verify`, `officegen_formats::diagnose` | Overflow risk fixture | RG-2 |
| V31-R002 | Repair suggestions emit directly usable edit ops with `op: "setText"` and `officegen.edit.ops@1.2`. | `officegen_formats::diagnose`, `officegen_formats::repair` | Repair plan schema fixture | RG-1, RG-3 |
| V31-R003 | Repair execution updates `wouldWrite` and `planOnly` according to actual writes. | `officegen_formats::repair` | Repair write vs plan-only fixture | RG-3 |
| V31-C001 | `chart render` rejects unsupported chart data instead of silently using sample data. | `officegen_formats::charts` | Chart invalid data fixture | RG-2 |
| V31-G001 | Diagram parser separates Mermaid-like node ids and labels. | `officegen_formats::diagrams` | `diagrams-render.test.ts` label/id fixture | RG-2 |
| V31-T001 | `template candidates --source-only` excludes stale registry candidates. | `officegen_optional::template` | Template source-only fixture | RG-1, RG-3 |
| V31-A001 | `officegen.packageDiff@1` remains registered and schema-covered. | `officegen_core::schemas`, `officegen_formats::edit` | Schema coverage plus package-diff fixture | RG-5 |
| V31-A002 | Broad nested inspect/verify schemas remain explicitly documented until hardened. | `officegen_core::schemas` | Schema coverage with deferred-hardening note | RG-5 |
| V31-A003 | npm publish remains blocked unless release policy changes explicitly. | `officegen_release` | Pack/publish policy smoke | RG-5 |

## Test Parity Matrix

| v3 test suite | Rust test target | Compatibility surface |
| --- | --- | --- |
| `packages/cli/test/program.test.ts` | `crates/officegen-cli/tests/cli_parity.rs` | Command surface, help, JSON envelope, option validation, policy gates, run workflows, path policies, output artifacts. |
| `packages/core/src/config.test.ts` | `crates/officegen-core/tests/config.rs` | Profile defaults, feature visibility, capabilities hash. |
| `packages/core/src/envelope-errors.test.ts` | `crates/officegen-core/tests/errors.rs` | Error catalog and envelope success/failure classification. |
| `packages/core/src/paths-run.test.ts` | `crates/officegen-core/tests/paths_run.rs` | Run path policy, output root enforcement, redaction. |
| `packages/core/src/redaction.test.ts` | `crates/officegen-core/tests/redaction.rs` | Home/project path redaction. |
| `packages/core/src/run-trace.test.ts` | `crates/officegen-core/tests/run_trace.rs` | Run manifest, JSONL trace, artifact evidence. |
| `packages/core/src/schemas.test.ts` | `crates/officegen-core/tests/schemas.rs` | Schema registry, edit op schema diagnostics, package-diff schema. |
| `packages/core/src/zipSafety.test.ts` | `crates/officegen-core/tests/zip_safety.rs` | ZIP risk classification, chart workbook classification, risky package gates. |
| `packages/formats/src/shared.test.ts` | `crates/officegen-formats/tests/zip_loading.rs` | Unsafe ZIP metadata rejection before parse. |
| `packages/formats/tests/ooxml-validator.test.ts` | `crates/officegen-ooxml/tests/validator.rs` | OPC relationship, malformed XML, risky part detection. |
| `packages/formats/tests/package-graph.test.ts` | `crates/officegen-ooxml/tests/package_graph.rs` | Part graph, rel graph, external link and embedded package classification. |
| `packages/formats/tests/ooxml-token-index.test.ts` | `crates/officegen-ooxml/tests/token_index.rs` | XML token spans and source mapping. |
| `packages/formats/tests/ooxml-patch-engine.test.ts` | `crates/officegen-ooxml/tests/patch_engine.rs` | Byte-preserving patches, malformed XML rejection, stale fingerprint rejection. |
| `packages/formats/tests/edit-transaction.test.ts` | `crates/officegen-ooxml/tests/transaction.rs` | Atomic rollback, best-effort behavior, cloned snapshots. |
| `packages/formats/tests/edit-office-xml-integration.test.ts` | `crates/officegen-formats/tests/edit_integration.rs` | EditOps, selector resolution, no partial output unless explicit, dry-run behavior. |
| `packages/formats/tests/formats.test.ts` | `crates/officegen-formats/tests/formats_parity.rs` | PPTX/DOCX/XLSX/PDF inspect/render/view/verify/diff fixtures. |
| `packages/formats/tests/inspect-workbook-map.test.ts` | `crates/officegen-formats/tests/xlsx_inspect.rs` | Formula attribution and scoped summary parity. |
| `packages/formats/tests/xlsx-formula-graph.test.ts` | `crates/officegen-ooxml/tests/xlsx_formula_graph.rs` | Shared formulas, structured references, unsafe formulas, chart/pivot related ranges. |
| `packages/formats/tests/docx-story-run-graph.test.ts` | `crates/officegen-ooxml/tests/docx_story_graph.rs` | DOCX stories, runs, hyperlinks, bookmarks, comments, revisions, fields. |
| `packages/formats/tests/pptx-object-map.test.ts` | `crates/officegen-ooxml/tests/pptx_object_map.rs` | Stable PPTX object ids, paragraph/run semantics, shapes, charts, SmartArt signals. |
| `packages/formats/tests/object-graph.test.ts` | `crates/officegen-formats/tests/object_graph.rs` | Object graph projection and selector graph inputs. |
| `packages/formats/tests/visual-diff.test.ts` | `crates/officegen-formats/tests/visual_diff.rs` | Approximate visual diff, blocked native diff classification. |
| `packages/formats/tests/diagrams-render.test.ts` | `crates/officegen-formats/tests/diagrams.rs` | Mermaid-like parsing, SVG size, small-width layout. |
| `packages/formats/tests/pdf-redaction.test.ts` | `crates/officegen-pdf/tests/redaction.rs` | PDF overlay/annotation boundaries and non-redaction disclosure. |
| `packages/optional/src/template-design.test.ts` | `crates/officegen-optional/tests/template_design.rs` | Template candidates, source-only behavior, design capture/apply/update/edit. |
| `packages/optional/src/layout.test.ts` | `crates/officegen-optional/tests/layout.rs` | PPTX geometry mutation and diff evidence. |

## Release Gate Checklist

| Stage | Must pass before advancing |
| --- | --- |
| Porting branch | Rust equivalents for target modules compile; no command exposed without RG-0 metadata and help coverage. |
| Feature parity PR | Relevant Rust unit tests plus v3 fixture replay for the rows touched in the command and fix matrices. |
| v4 alpha | RG-0 and RG-1 for all command groups; RG-4 for renderer/plugin/MCP/agent surfaces; no package/source manifest edits outside the approved release workflow. |
| v4 beta | RG-2 for all read-only commands; RG-3 for edit/render/export/repair/asset/template/design/layout; benchmark local fixtures pass. |
| v4 release candidate | Full RG-5, perfect-spec evidence regenerated, release tarball smoke passes, and post-tag smoke plan names the exact `vX.Y.Z` artifact. |
