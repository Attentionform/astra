# ASTra benchmarks

Reproducible measurements of the **published** ASTra npm packages, through the same MCP interface an
AI assistant uses. Work in progress: ASTra changes every day, so every result belongs to one named release
(`results/<version>/`). Numbers will change; the method is meant to stay.

* `PROTOCOL.md`: what is measured, how, and what we expected **before** measuring. Read this first.
* `results/<version>/REPORT.md`: the report. Raw per-trial data (`*.jsonl`) sits beside it. Nothing is dropped,
  failures included.
* `corpus/`: the generator and manifest (SHA-256 of every document). Documents are regenerated, not stored.
* `lib/`: the harness and the independent checks (they contain no ASTra code).
* `run/`: one script per experiment.

## Run it

Requirements: Node 22.19 or newer, Python 3.10+ with `python-docx`, `lxml` and `Pillow`, LibreOffice (`soffice`) and
Poppler (`pdfinfo`) on the `PATH`.

```bash
cd benchmarks
npm ci
python3 corpus/generate.py          # corpus/files/*.docx + corpus/manifest.json
./run/all.sh                        # every experiment, then the report
```

Single experiments: `node run/r1_representation.mjs`, `node run/e1_edit_locality.mjs`, and so on; build the report with
`node run/report.mjs`. The package under test is the one pinned in `package.json`; change that pin to measure another release.

## Run it on GitHub Actions

`CI.md` explains how to run the suite on a hosted runner from a `bench/**` branch, against the published package
and a build of any branch, one experiment at a time.

## What this is not

* Not a comparison with other products.
* Not a statement about how well a particular AI model uses the tools (see PROTOCOL section 10).
* Not a Word-compatibility proof; LibreOffice and a Python reader are the independent readers here.

## Licence

Source-available under the repository licence (BUSL-1.1). The benchmark scripts generate their own documents; no
third-party or customer content is included.
