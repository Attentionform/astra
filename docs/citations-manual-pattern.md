# Citations: manual references pattern

No CSL processor. No bibliography command. `astra batch --schema` exposes
`clearDoc, comment, delete, insert, list, move, namedStyle, replace, style,
table`; no `citation`/`bibliography` action. References are plain paragraphs.

## Pattern

1. Insert a heading plus one paragraph per entry (T8 pattern: `References` + `[1] …`).
2. Optional in-text notes via engine `footnote`/`endnote` ops (pass through `astra batch`).

```sh
cp test-gallery/T8_footnotes_endnotes.docx /tmp/citedemo.docx
astra batch /tmp/citedemo.docx ops.json --out /tmp/citedemo.docx
# ops.json: {"operations": [{"action": "insert", "type": "paragraph",
#   "anchor": "25B763D2", "position": "after",
#   "text": "[3] Vega and Okafor, Field Methods for Archive Verification, 2023."}]}
# => BATCH OK /tmp/citedemo.docx -> /tmp/citedemo.docx (applied=1)

astra validate /tmp/citedemo.docx
# => VALID /tmp/citedemo.docx

astra export /tmp/citedemo.docx --format markdown -o /tmp/citedemo.md
# => EXPORT OK ... (MARKDOWN, 786 bytes)
# md carries manual entries verbatim plus [^1]/[^en-1] defs for real notes:
# [^1]: T8 footnote body FN-ALPHA: ...
# [^en-1]: T8 endnote body EN-BETA: ...
```

Verified 2026-09-24 against `node packages/cli/dist/cli.js` (exits 0; bad ops exit 1/2 per `docs/cli-guide.md`).
Live example: `test-gallery/T8_footnotes_endnotes.docx` (`References` heading + entries `[1]`, `[2]`).

## Automation

Any CSL-style auto-formatting is future work. Until then, manual pattern
only: no automation track exists yet, so there is nothing to wait on.
