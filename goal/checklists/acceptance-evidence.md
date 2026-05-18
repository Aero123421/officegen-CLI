# v2.6.0 Acceptance Evidence Checklist

The acceptance evidence bundle should be reproducible enough for CI and readable
enough for a release reviewer.

## Required Evidence Root

- [ ] `acceptance/index.json` lists every acceptance test ID and status.
- [ ] `acceptance/sha256.txt` includes every generated artifact and transcript.
- [ ] `acceptance/reviewer-notes.md` records manual decisions, platform constraints, and accepted limitations.
- [ ] `runs/latest/manifest.json` records `exists`, `bytes`, and `sha256` for expected outputs.
- [ ] `runs/latest/events.jsonl` is parseable line by line.
- [ ] `runs/latest/summary.md` summarizes pass/fail state without hiding limitations.

## Blocking Logs

- [ ] `goal-check.txt`
- [ ] `contract-check.txt`
- [ ] `schema-coverage.txt`
- [ ] `remediation-check.txt`
- [ ] `typecheck.txt`
- [ ] `test.txt`
- [ ] `build.txt`
- [ ] `binary-smoke.txt`
- [ ] `pack-smoke.txt`
- [ ] `install-smoke.txt`

## Product Artifacts

- [ ] Inspect JSON for PPTX, DOCX, XLSX, PDF, and hostile-content fixtures.
- [ ] View preview artifacts and object-map overlay JSON.
- [ ] Rendered PPTX, DOCX, XLSX, PDF outputs plus verify JSON.
- [ ] Chart and diagram SVG outputs plus invalid-input failure JSON.
- [ ] Edit dry-run JSON, applied edit JSON, diff JSON, and verify JSON per format.
- [ ] Asset inspect, extract, replace, and verification outputs.
- [ ] Diagnose, repair, diff, critique, and improve outputs.
- [ ] Security negative-case outputs for output-root escape, unsafe archive/path input, and gated external capabilities.

## Sign-off Rules

- [ ] Every blocking acceptance test in `acceptance-suite.v2.6.0.json` is `pass` or has an approved release-blocking waiver.
- [ ] Every waiver cites a remediation row and a user-visible known limitation.
- [ ] No acceptance result is marked pass when its output contains blocking warnings.
- [ ] Manual native-renderer evidence states the platform, backend, and configuration used.
