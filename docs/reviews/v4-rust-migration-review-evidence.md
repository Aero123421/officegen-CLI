# Officegen-CLI v4 Rust Migration Review Evidence

Date: 2026-05-21.

Baseline: v3.1.8 compatibility matrix in `docs/planning/v4-rust-migration-compatibility-matrix.md`.

Scope: docs/test evidence only. No Rust implementation, package manifest, release workflow, installer, or version script files were edited for this review.

v4.5.0 update: this is historical v4.0.0 evidence. It is superseded by `docs/reviews/v4.5.0-review-checklist.md` for release decisions. MCP and plugin runtime support are now removed from the built-in CLI scope rather than optional policy-gated surfaces.

Status values: `pass`, `warning`, `blocked`, `not-run`, `planned`.

## Compatibility Evidence

| Gate | Evidence | Status | Notes |
| --- | --- | --- | --- |
| RG-0 Surface parity | `target\debug\officegen.exe capabilities --agent --strict-json` | pass | Returned `officegen.envelope@1.2`, `runtime.kind: rust-native`, `nodeRequired: false`, version `4.0.0`, stable capabilities hash, and the planned visible command surface including core, format, optional, renderer, plugin, MCP, and benchmark commands. |
| RG-1 Contract parity | `target\debug\officegen.exe frobnicate --agent --strict-json` | pass | Returned a structured error envelope with `ok: false`, `objectiveOk: false`, `readiness: blocked`, `code: UNKNOWN_COMMAND`, and agent-safe next actions. |
| RG-1 Contract parity | `target\debug\officegen.exe errors inspect UNKNOWN_COMMAND --agent --strict-json` | pass | Returned `officegen.error.inspect@1.2` with `officegen help --agent --strict-json` as the recommended recovery command. |
| RG-1/RG-4 Optional gating | `target\debug\officegen.exe doctor --agent --strict-json` | warning | Rust runtime check passed and native Office/LibreOffice proof remained explicitly opt-in via a warning check. The envelope readiness was `pass`; release review should decide whether warning readiness should propagate to the envelope before v4 RC. |
| RG-2 Read-only format parity | Temporary PPTX smoke: `scaffold`, `render`, `inspect`, `verify` through `target\debug\officegen.exe` | pass | Rendered a PPTX in `%TEMP%`, inspected it as `format: pptx`, found text objects, and verified status `pass`. This is a smoke check only, not fixture parity. |
| RG-2 Chart guard | Temporary invalid chart JSON with `chart render --agent --strict-json` | pass | Unsupported chart data returned `ok: false` with `SCHEMA_INVALID`, preserving the v3.1.8 no-sample-data compatibility requirement. |
| RG-2 Diagram parsing | Temporary Mermaid-like `graph TD; A[Alpha Label] --> B[Beta Label]` with `diagram render --out diagram.svg --agent --strict-json` | pass | Command produced an SVG artifact successfully. Full id/label golden parity remains a Rust fixture candidate. |
| RG-5 Rust unit/integration tests | `cargo test` | blocked | `cargo` and `rustc` were not available in this Windows shell, so Rust tests could not be executed here. |

## Targeted Smoke Script

The following PowerShell smoke was executed successfully against the existing built binary and wrote only to `%TEMP%`:

```powershell
$exe = (Resolve-Path 'target\debug\officegen.exe').Path
$tmp = Join-Path $env:TEMP ('officegen-v4-smoke-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tmp | Out-Null
Push-Location $tmp

$cap = (& $exe capabilities --agent --strict-json | ConvertFrom-Json)
$scaffold = (& $exe scaffold --kind pptx --title 'Rust Smoke' --out ir.json --agent --strict-json | ConvertFrom-Json)
$render = (& $exe render ir.json --target pptx --out smoke.pptx --agent --strict-json | ConvertFrom-Json)
$inspect = (& $exe inspect smoke.pptx --agent --strict-json | ConvertFrom-Json)
$verify = (& $exe verify smoke.pptx --agent --strict-json | ConvertFrom-Json)

'{"title":"Bad Chart","data":{"foo":1}}' | Set-Content -Encoding UTF8 chart-invalid.json
$chart = (& $exe chart render chart-invalid.json --agent --strict-json | ConvertFrom-Json)

'graph TD; A[Alpha Label] --> B[Beta Label]' | Set-Content -Encoding UTF8 diagram.mmd
$diagram = (& $exe diagram render diagram.mmd --out diagram.svg --agent --strict-json | ConvertFrom-Json)

Pop-Location
```

Assertions checked:

- capabilities returns `ok: true`, `runtime: rust-native`, `nodeRequired: false`, and includes `edit` plus `schema validate`.
- scaffold writes `ir.json`.
- render writes `smoke.pptx` and reports `mutationStatus: changed`.
- inspect returns `format: pptx` with at least one text object.
- verify returns `status: pass`.
- invalid chart data returns `ok: false` and `SCHEMA_INVALID`.
- diagram render writes `diagram.svg`.

## Rust Test Additions

No Rust tests were added in this pass because there is no root `tests/` directory in the current Rust workspace. To stay inside the requested ownership boundary, this review records test candidates instead of creating a new Rust test tree.

Recommended first integration tests when the Rust tests directory is introduced:

| Candidate file | Coverage |
| --- | --- |
| `tests/cli_contract.rs` | `capabilities`, unknown command, `errors inspect`, envelope fields, `--agent --strict-json` next actions. |
| `tests/pptx_smoke.rs` | `scaffold` -> `render` -> `inspect` -> `verify` temporary PPTX workflow. |
| `tests/chart_diagram_guards.rs` | Invalid chart input rejects with `SCHEMA_INVALID`; Mermaid-like node ids and labels remain separated. |
| `tests/native_policy.rs` | `doctor`, renderer, plugin, MCP, and native proof surfaces remain opt-in or discovery-only. |

## Review Notes

- The current Rust smoke is sufficient for early compatibility evidence, but it does not replace the v3 fixture replay matrix.
- `doctor` currently carries a warning check while the outer envelope still reports readiness `pass`; this is acceptable as documented warning evidence for now, but should be reviewed before release candidate gating.
- `cargo test` remains a required local/CI gate once a Rust toolchain is available.
