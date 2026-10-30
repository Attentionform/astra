# Adversarial fidelity corpus

Legal-but-pathological Word output: merged tables, mixed-script revision runs,
anchored drawings with negative offsets, nested tables, numbering restarts,
revision soup. Each fixture proves the engine can ingest, mutate two sentences,
and save without touching the weirdness.

## Layout

- `fixtures/<name>.docx` + `fixtures/<name>.manifest.json` — one pathology per
  fixture; every fixture has a manifest with a recorded disposition.
- `hostile/<family>.twin.docx` + `hostile/<family>.twin.json` — negative
  control per family: same shape, exactly one corruption; the gate must refuse
  it. Green without teeth proves nothing.
- `manifest.schema.json` — single source of truth for manifest fields. Required:
  `fixture`, `family`, `motivatingPathology`, `knownLimitations`,
  `scopeExclusions`, `ecma376Difficulty`, `licenseProvenance`,
  `disposition{verdict, reason}`. Optional: `sourceDocument`. Unknown fields
  fail (misnamed fields are never silently accepted); the Safe Docx alias names
  (`limitation`, `prose`, `description`, `discussion`, …) fail explicitly.
- `invariants.json` — full invariant registry: claim + evidence tier + caveats
  + named falsifier per family, one entry per family across the whole corpus.
- `generate-fixtures.ts` — deterministic fixture/twin builder. Runner of record
  for binaries: `npx tsx tests/harness/fidelity/generate-fixtures.ts`.
- `fidelity-corpus.test.ts` — the runner (runs under the existing
  `tests/harness` vitest include; no config changes needed).

## Pipeline (per fixture)

Ingest (`DocxParser` + `getAST`) → normalize (engine-owned allowlist pass
inside the save path; log captured per run) → mutate exactly two sentences via
one `executeBatch` replace pair → rezip → `validateOOXML`.

## Gates

- HARD: `validateOOXML` passes; reloaded node checksums show ONLY the two
  targets changed; `word/document.xml` tag histogram is identical (no valid tag
  stripped or added); every other package part is preserved (XML parts
  semantically identical modulo namespace/whitespace serialization, binary
  parts byte-identical); node id set is unchanged; a second pass from the same
  source bytes is byte-identical.
- Mutation targets are the first two paragraphs whose runs carry no direct
  formatting. Replace carries the first run's properties onto rewritten text,
  but paragraphs with distinct per-run formatting would still conflate
  replace merge semantics with the preservation under test, so formatted
  targets stay excluded by rule. The exclusion is stated here and in the
  gallery manifests — not a silent skip.
- Manifest gate: `node scripts/check-fidelity-manifests.mjs` passes (also
  asserted inside the runner against the same schema file).
- Conform-or-document: every manifest carries `disposition.verdict` —
  `conform` (Word-legal shape the engine preserves) or `deliberately-differ`
  with a written reason. No silent third option.
- Render similarity is advisory only (LibreOffice is not a hard gate); it runs
  in `render-similarity.test.ts`, never here. See "Render similarity" below.
- Variance study verdict lives in `variance-report.json`: currently
  `do-not-promote`. Render similarity may gate merges only after the study
  records self-variance under budget for three consecutive campaigns with no
  open converter crashers; the record test enforces that rule on the file.

## First run (fresh clone)

`fixtures/*.docx` (71) and `hostile/*.docx` (19) are gitignored (`*.docx`,
except `test-gallery/`) — a fresh clone holds only `.manifest.json` /
`.twin.json`. The first `npx vitest run
tests/harness/fidelity/fidelity-corpus.test.ts` rebuilds them at import via
`npx --yes tsx tests/harness/fidelity/generate-fixtures.ts` (about a second;
the `fixture …` / `gallery … <- …` / `twin … (from …)` lines are the
generator, then the 163 tests run). Binaries present: two-readdir no-op, no
generator output. Generator failure throws naming the generator step — never
an empty suite. Same command regenerates by hand:
`npx tsx tests/harness/fidelity/generate-fixtures.ts`.

## Regenerate / verify

```
npx tsx tests/harness/fidelity/generate-fixtures.ts
node scripts/check-fidelity-manifests.mjs
npx vitest run tests/harness/fidelity/fidelity-corpus.test.ts
```

Budget: whole file green in well under 5 minutes wall time (measured and
asserted by the `budget` case).

## Render similarity (advisory only, never a merge gate)

`render-similarity.ts` + `render-similarity.test.ts` + `variance-study.test.ts`
+ `variance-report.json`. Per fixture the advisory runner replays the corpus
pipeline (mutate two sentences), renders before/after with the local
LibreOffice, and records one entry: `measured` (pixel, page, text scores plus
an advisory band) or `unavailable` with the reason (missing binary, timeout,
renderer crash). No magnitude ever fails the suite; a missing renderer skips
the render cases cleanly.

Metric: same-binary `soffice --headless --convert-to pdf` for both sides
(PDF bytes never compared — the exporter stamps timestamps), then per-page
`pdftoppm -gray -r 36` rasters scored as mean(|a-b|)/255 over shared pages
(mean reported, worst page alongside), page-count equality from `pdfinfo`,
and whitespace-token Jaccard of `pdftotext -layout` output (informational:
it is dominated by the two-sentence mutation itself, median 0.33 on the
calibration run, so it never bands). Bands: pixel <= 0.02 `consistent`,
<= 0.10 `localized-change`, above that or on page-count change `needs-eyes`.

Calibration (2026-10-10, LibreOffice 25.2.7.2): 69 of 71 measured —
64 `consistent`, 4 `localized-change` (all gallery baselines), 1 `needs-eyes`
(`gallery-coffee-onepager`: the shorter replacement text reflows 2 pages to
1; verified by eye that only the two target paragraphs changed). Hand-built
fixtures all score <= 0.011, so the bands sit where the data puts them.

Oracle trust boundary: structural projections only (page count, extracted
text, pixel score); no engine internals read, nothing asserted about engine
output. Binary-gated on renderer presence. Caveats: 36 dpi grayscale is
blind to sub-point shifts and color-only changes; `pdftotext` order follows
the content stream, not reading order (bidi/complex-script similarity is
approximate); the exporter crashes on endnote shapes (2 fixtures record
`unavailable`; see FID-005 in `scratch/FIDELITY_FINDINGS.md`); one bulk
convert shares fate across files, so the runner converts per file.

Harness notes: one isolated renderer profile per test worker (sequential
conversions share it; a fresh profile costs one ~20 s warm-up, later
conversions are sub-second; a crashed conversion does not poison it). Run:

```
npx vitest run tests/harness/fidelity/render-similarity.test.ts
npx vitest run tests/harness/fidelity/variance-study.test.ts
ASTRA_ADVISORY_OUT=/tmp/advisory.json npx vitest run tests/harness/fidelity/render-similarity.test.ts
```

## Adding a fixture (builder contract)

Numbered recipe; a newcomer following steps 1-7 lands a green fixture.

1. Add a builder in `generate-fixtures.ts` (`builders` record, keyed by
   fixture stem) or a gallery copy in `galleryCopies`. Binaries are
   generated: the generator is the runner of record for every `.docx` in
   `fixtures/` and `hostile/`. Never hand-edit a `.docx` in place; the next
   regen overwrites it. Paths resolve relative to the generator file, so cwd
   does not matter.
2. Run the generator: `npx tsx tests/harness/fidelity/generate-fixtures.ts`.
   Regen rewrites every builder `.docx`, re-copies gallery baselines, and
   re-derives every hostile twin. It never writes manifests or
   `invariants.json`; a `.docx` with no builder survives regen but breaks the
   runner-of-record rule, so every hand-built fixture needs a builder entry.
3. Hand-write `fixtures/<stem>.manifest.json`. Required fields: `fixture`
   (equals file stem), `family`, `motivatingPathology`, `knownLimitations`,
   `scopeExclusions`, `ecma376Difficulty`, `licenseProvenance`,
   `disposition{verdict, reason}`. `fixture`/`family` match
   `^[a-z0-9-]+$`. `disposition.verdict` is `conform` or
   `deliberately-differ`. Word-count ranges (enforced by the check script and
   the runner against `manifest.schema.json`): `motivatingPathology` 40-150,
   `knownLimitations` 40-150, `scopeExclusions` 40-150,
   `ecma376Difficulty` 40-150, `licenseProvenance` 5-80,
   `disposition.reason` 20-150. A sibling `<stem>.docx` must exist. Unknown
   fields fail; the Safe Docx alias names (`limitation`, `prose`,
   `description`, `discussion`, ...) fail explicitly.
4. `sourceDocument` convention: hand-built fixtures omit it; gallery
   baselines set it to the gallery origin (e.g.
   `test-gallery/T1_coffee_onepager.docx`). Optional string only.
5. Add or extend the hostile twin in `generate-fixtures.ts` (`twins`
   array): one corruption per family, derived from a `base` fixture stem.
   Twin JSON contract (`hostile/<family>.twin.json`): `family`, `twin`
   (`<family>.twin`), `base` (fixture stem), `corruption` (one-sentence
   description), `expectedRefusal` (`open`, `ingest`, `rezip`, or
   `validate`). The runner refuses each twin at its documented stage.
6. Update `invariants.json` when adding a family or changing preservation
   behavior: one entry per family with `id` (`FID-<FAMILY>-NNN`), `claim`,
   `tier` (`tested-only` until a differential oracle lands;
   `empirically-validated` needs a deterministic differential harness or
   reference oracle), `evidence` (runner path plus fixture stems),
   `caveats`, and `falsifier` (hostile twin path). Keep one claim per family;
   every tier stays `tested-only` until a differential oracle with a variance
   study promotes it — the honest default, recorded in the registry scope note.
7. Verify: `node scripts/check-fidelity-manifests.mjs`, then
   `npx vitest run tests/harness/fidelity/fidelity-corpus.test.ts`.
   Mutation targets are the first two unformatted paragraphs by rule (replace
   carries the first run's properties onto rewritten text, so the rule is now
   conservative: it keeps multi-run formatting distinctions out of the
   preservation gate rather than working around dropped properties).
