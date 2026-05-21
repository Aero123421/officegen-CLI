# v2.6.0 Product Completion Checklist

Use this checklist as the human sign-off view of
`office-operation-os-v2.6.0.goal.json`.

## Agent Contract

- [ ] `officegen capabilities --agent --json` exposes command groups, profiles, gates, and known limitations.
- [ ] `officegen help --agent --json` and workflow help are enough for an agent to pick the next command.
- [ ] `npm run contract:check` passes and no agent-visible command is a placeholder.
- [ ] `npm run schema:coverage` passes and `schema list/get/validate` work for representative schemas.
- [ ] Negative commands produce structured errors with stable codes, hints, and redacted context.

## Office Runtime

- [ ] PPTX, DOCX, XLSX, and PDF fixtures inspect into stable object maps with trust markers.
- [ ] View artifacts exist for representative Office/PDF fixtures and link previews to object-map context.
- [ ] Scaffold/render/verify works for PPTX, DOCX, XLSX, and PDF structured IR.
- [ ] Chart and diagram SVG rendering succeeds for valid specs and fails cleanly for invalid specs.
- [ ] Template/design/layout behavior records reusable signals and discloses best-effort limits.

## Mutation Runtime

- [ ] Edit dry-run resolves selectors, reports target counts, and leaves source hashes unchanged.
- [ ] Representative PPTX, DOCX, XLSX, and PDF edits apply to new outputs.
- [ ] Diff and verify evidence exists for every edited output.
- [ ] Embedded assets can be inspected, extracted, and replaced with trust annotations and sha256 evidence.

## Verification and Workflow OS

- [ ] Verify distinguishes clean outputs from damaged or repair-risk fixtures.
- [ ] Diagnose, repair, and diff form a documented closure loop.
- [ ] Critique/improve are plan-first and do not mutate source files unless explicitly requested.
- [ ] `officegen run` produces strict JSON, JSONL events, manifest metadata, sha256 values, and a Markdown summary.
- [ ] Output-root escape attempts fail with no escaped writes.

## Safety

- [ ] Document-derived strings, metadata, hyperlinks, macros, and embedded files are treated as untrusted data.
- [ ] Zip traversal, absolute path, oversized entry, and outside-root fixtures fail safely.
- [ ] Native renderer, plugin, MCP, and external-process capabilities are gated by policy.
- [ ] Redaction tests cover secret-like values in errors and logs.

## Release

- [ ] `npm run goal:check` passes.
- [ ] `npm run version:check` passes after the release version is bumped by the version script.
- [ ] `npm run remediation:check` passes for `docs/reviews/v2.6.0-remediation-matrix.md`.
- [ ] `npm run typecheck`, `npm test`, `npm run build`, `npm run binary:smoke`, and `npm run pack:smoke` pass.
- [ ] GitHub install smoke or equivalent release-candidate install smoke passes.
