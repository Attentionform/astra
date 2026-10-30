# Changelog

All notable changes to ASTra are recorded here. The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and the packages use [Semantic Versioning](https://semver.org/). While the major version is 0, minor releases may change behavior; each such change is listed.

Packages released together: `@attentionform/astra-engine`, `@attentionform/astra-mcp-server`, `@attentionform/astra` (command line) and `@attentionform/ooxml-validator`, all at the same version.

## [Unreleased]

## [0.1.0] - 2026-10-31

First public release.

### Added

- **Engine** (`@attentionform/astra-engine`): reads a `.docx` into a structured document with stable node ids; 27 typed operations (insert, replace, delete, style, tables, sections and page setup, headers and footers, fields, equations, lists, charts, images, callouts, footnotes and endnotes, watermarks, page borders, cover pages, hyperlinks, bookmarks, comments, tracked changes, metadata); batches that apply all or nothing; revision history.
- **Validation on every write.** A file is checked against the Office Open XML schema before it is written. An invalid result is refused with `E_VALIDATION`, an explanation and a suggested fix, and the file on disk is left as it was. This covers `save`, `sanitize` and PDF export.
- **Schema validator** (`@attentionform/ooxml-validator`): the .NET 10 Open XML validator behind the check above, with a persistent worker process, in-memory validation without temporary files, and prebuilt binaries for Linux, macOS and Windows on x64 and arm64.
- **MCP server** (`@attentionform/astra-mcp-server`): 19 tools over stdio, including open, inspect, search, execute a batch, save, compose from Markdown, import, export (HTML, Markdown, PDF), visual inspection, history, preview and revert, validate, statistics, sanitize and normalize. Every call is confined to the folder given with `--allowed-dir`.
- **Command line** (`@attentionform/astra`): 15 commands for scripts and pipelines, with a machine-readable `--json` output and documented exit codes.
- **Equations.** LaTeX in documents becomes native Word equations, including accents (`\dot`, `\ddot`, `\hat`, `\bar`, `\vec`, `\tilde`), and round-trips back to LaTeX. In Markdown, `\( \)` and `$$ $$` are always math; `$ $` is math only when you turn it on, so prices such as "$5 and $10" stay text.
- **License and contributions.** ASTra is source-available under the Business Source License 1.1 (`BUSL-1.1`). Free production use covers individuals (including freelancers and consultants), schools and academic research, non-profits, public bodies and small businesses; other production use needs a commercial license. Each version becomes Apache-2.0 four years after it is first made public. Outside contributions require a one-time contributor license agreement, signed by adding a signature file to the first pull request and checked by the `cla` workflow (see `CONTRIBUTING.md`).
- **Privacy tools.** `sanitize` removes document metadata by default; `--redact-pii` and `--client-safe` also redact personal data in the body text. Footnotes and endnotes are preserved.
- **Numbers count from 1.** Every number you type or read starts at 1: table rows and columns (`rowIndex`, `colIndex`, `index`, `afterIndex`, merge ranges), chart series (`trendline.series`), sections (`sectionIndex`), list levels (`level`, shown as `[L1]` in `inspect`) and pages. `inspect` prints the same numbers it accepts. `0` is refused with the valid range; for list `changeLevel`, `0` takes a paragraph out of its list. `offset` (how many items to skip) and the raw file values stay 0-based.
- **Cell addresses.** A table cell is addressed as `<tableId>:r<row>c<col>` (for example `t_001:r2c3`), or `…:p<n>` for one paragraph inside it. `replace`, `style`, `namedStyle`, `clearFormat`, `hyperlink`, `bookmark` and `comment` accept it as `target`; any other action refuses it and names the fix. `search_content` can return cells, and `compose_document` returns a `persistentIds` map from its temporary `md_*` ids to the ids the saved file carries.
- **Unknown settings are refused.** An operation that carries a setting it does not read (for example `bold` inside `params` of a table `updateCell`) fails with `E_VALIDATION` before anything is applied, naming the setting and the ones that are accepted.
- **Real dates and your own author.** Output carries real timestamps by default. For byte-identical output pass `--reproducible` (or set `ASTRA_REPRODUCIBLE=1` or `SOURCE_DATE_EPOCH`); `ASTRA_NO_REPRODUCIBLE=1` wins over both. `--author <name>` sets the default author for tracked changes and comments that do not name one.
- **Better defaults when composing Markdown.** Composed tables get a navy, bold header row and striped body rows (a row added later keeps the stripes, and deleting a row re-alternates them); a subtitle under the title is kept. The old plain look is still available as the `legacy` style. `headerBold` (default on) turns the bold header off, flat or in `params`, or with `- Header Bold: false` under `## Tables` in a style guide. Spacing and indents accept twips or a point string such as `"12pt"`.
- **Complete new files and a safe history location.** New documents carry the full package (document properties, settings, font table and theme) and pass the Open XML SDK validator. `ASTRA_HISTORY_DIR` keeps revision history out of the document's folder.
- **Visual output.** `visual_inspect` takes `outputDir` to choose where page images are written (a symlink at the image name is refused). The text wireframe shows table columns, the header and first rows, list markers and `[CALLOUT]` blocks in about 380 tokens.
- **Lists have nine levels.** A new list definition carries all nine levels, so an item at any level has its marker and indent: bullets cycle `•` `◦` `▪`, numbers cycle `1.` `a.` `i.`. Before, only level 1 was defined and a deeper item had no marker or indent of its own. A list is reused only when its top level is the requested kind (a Word list that is numbered at level 1 and bulleted at level 2 used to be handed out for a bullet list), never when it is heading numbering or a style link, and a list that stops short of the level an item is placed at gets the missing levels in place.
- **`pageSetup` turns the page.** `orientation: "landscape"` (or `"portrait"`) now swaps the section's width and height, so the page really turns in Word and LibreOffice; before, only the orientation flag changed unless a `paperSize` was given too. Margins are written as given, never rotated. A `paperSize` given alone keeps the section's orientation.
- **`validate` lists the repairs a save would make.** `astra validate` and `validate_document` report `pendingNormalizations` (for example property children that would be put back in schema order) next to the verdict, so a file that is valid only once ASTra has saved it no longer reads as plainly valid.
- **Sections are read from the file.** A document opens with the sections it has: one per section break, with its own page setup, ids `sec_root`, `sec_2`, `sec_3`, … by position, and `inspect` reports the real count (it always said 1 before). Section operations, headers, footers, page borders and columns work on a file you opened, not only on sections made in the same session. A header kept only for a later section's first page is no longer typed as the document header.
- **New content goes where the file puts it.** With no anchor, an insert, table, callout, table of contents, field or chart goes to the end of the document (it went into the first section before, and a callout or table of contents with no anchor crashed on a document with several sections). Content placed right after a section break paragraph starts the next section. `equation` and `move` follow the same rule.
- **A failed batch changes nothing you can see.** The ids and versions returned by earlier batches (`_bk_new_…`, `md_…`, section ids) still resolve after a failed batch, a dry run or a failed single operation, and an `expectedVersion` taken before the failure is still valid. If the document cannot be matched block for block, the error says the ids were re-read.
- **Sections follow the section breaks.** After every batch, if a section break was deleted, moved or added, the sections are rebuilt from the document and the result carries a warning. Deleting a break merges the section into the next, as in Word; the text takes the page setup of the section after it. The header and footer of a deleted section are removed from the file too, with any picture only they used, so their text does not stay inside the saved document; nothing is removed while anything still refers to it, and if the package would be left with a broken reference the removal is skipped.
- **Headers, footers, page borders and columns fail loudly.** When a section has no section properties to hold the change, the batch fails and rolls back; before, it did nothing and a header or footer part could be left behind. A header or footer can now be set on an empty trailing section.
- **One Node version for every package.** The validator wrapper needs Node 22.19 or newer, like the engine, the command-line tool and the MCP server (it said 20 before).
- **Validator binaries keep working when one platform's build fails.** The native binaries have their own version, separate from the wrapper's, and the wrapper accepts a range of them, so a platform whose newest binary could not be published keeps its last one.

### Security

- The model can no longer approve its own shell commands; approval always comes from the person at the keyboard or from an explicit flag.
- Archive entries, XML external entities, decompression bombs and active content are rejected or removed (see `docs/security.md`).
- Dependency advisories resolved before release: `zod`, `undici` and `@vitest/mocker`.

### Requirements and validation

- PDF export and page images need LibreOffice and Poppler installed on your machine.
- Every output file is validated with the Open XML SDK before it is saved, so files are valid, uncorrupted and open in Microsoft Word.

[Unreleased]: https://github.com/Attentionform/astra/compare/v0.1.0...HEAD
[0.1.0]: https://github.com/Attentionform/astra/releases/tag/v0.1.0
