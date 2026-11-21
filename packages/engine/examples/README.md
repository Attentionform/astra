# Examples

Small scripts that use `@attentionform/astra-engine` as a library. CI builds the engine, runs every script and fails when a
script throws or when a file it writes does not validate. The documentation copies its code from these files.

Run one from the repository root, after `npm run build:packages`:

```bash
node packages/engine/examples/01-create-from-markdown.mjs
```

Scripts 01 and 03 are self-contained and write into the folder that you run them from. The others write into a new folder under the system temp directory and print the path; set `ASTRA_EXAMPLE_OUT` to choose it. `run-all.mjs` runs all of them.

| Script | Shows |
|--------|-------|
| `01-create-from-markdown.mjs` | A new document from a Markdown string |
| `02-create-from-operations.mjs` | A new document built only from operations, with a heading |
| `03-open-find-edit-save.mjs` | Open a saved file, find text, edit it, validate, save |
| `04-table-and-chart.mjs` | A table and a chart |
| `05-comments-and-tracked-changes.mjs` | A comment, tracked edits, accept and reject |
| `06-history-and-recovery.mjs` | Saves with history, undo, and recovery of the original |
| `07-templates.mjs` | Every built-in template, in Letter and A4, and a compose onto one |
| `08-validate.mjs` | Check a document against the schema, see a failure, and handle its error code |
| `09-history-older-version.mjs` | List recorded states, preview a revert and read an older version back (uses `@attentionform/astra-engine/history`) |
| `10-render.mjs` | Find out which renderers are installed, make a PDF and a page image when they are |
