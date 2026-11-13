# Reading a document cheaply

An agent that reads a whole document to find one paragraph pays for every word. ASTra has several ways to read less. This page says which one to use. Part 1 holds the choices and the rules, and has no numbers. Part 2 holds sizes measured on one generated document, so you can see the order of magnitude and repeat the measurement.

## Part 1. Choosing a control

The tier of `inspect_document` is set with `fidelity`: `outline`, `compact`, `standard` (the default) or `full`. Every other control below narrows or pages the rows of that tier.

| You want | Use | You get | It does not |
|---|---|---|---|
| The shape of the document | `fidelity: "outline"` (MCP) or `astra inspect --outline` (CLI) | One row per heading: id, level, style, text, and the headings above it | Show body text, tables or images |
| Its size | `get_stats` or `astra stats` | Words, paragraphs, tables, images and a page estimate | Show structure or ids |
| Where a phrase is | `search_content` or `astra search` | The ids of the nodes that match | Return the rows around a match |
| The rows around a node you know | `aroundNode` with `window` (MCP) | The rows from `window` before to `window` after the node | Combine with `offset` or `limit` (the call is refused) |
| A range of rows | `offset` and `limit` (MCP) | A page of rows, with `nextOffset` while more remain | Combine with `aroundNode` |
| Shorter rows | `maxTextLength` (MCP) or `--max-text` (CLI) | Row text cut at that many characters, with `...` | Change which rows are returned |
| A reply that stays under a size | `maxTokens` (MCP) | Whole rows only; the reply says `truncated` when it cut | Promise an exact size (see below) |
| One part of the file | `sectionId` (MCP) or `--section` (CLI) | One page section of the file, or one node by its id | Select by heading: see "What a section is" |
| Formatting detail | `fidelity: "full"` with a `sectionId` | The run and paragraph properties of those rows | Run on a whole document |
| The text as prose | `astra export --format markdown` | The document as Markdown, without ids | Give you ids to edit with |

A cheap way to read a long document is to go from the outline to the part you need:

1. Read the outline (`fidelity: "outline"`).
2. Pick the heading you need and zoom in on it with `aroundNode` and `window`, or search for a phrase and zoom in on the node it names.
3. Edit with the node ids the rows gave you.

### What counts as a heading

A heading is a paragraph in the body of the document whose own style id contains `heading`, `title` or `subtitle` (in any case). The level comes from the style: `Heading1` to `Heading9` give 1 to 9, `Title` gives 1, `Subtitle` gives 2, and any other style that matches by name (`TOCHeading`, `BookTitle`) gives 1. These limits follow from that rule:

- A heading style with a custom name, such as `Kapitel`, is not listed. ASTra does not read the outline level that Word stores in a style.
- A paragraph inside a table cell or a textbox is not listed.
- `h1` to `h9` are not heading styles.

The outline names the headings above a row in its `Section` column, nested by level, joined with ` > `. A top-level heading has an empty `Section`.

A heading can contain the characters that separate cells, so the outline escapes them in `Text`, `Style` and `Section`: a backslash is written `\\`, a pipe `\|`, a line break `\n` (and `\r`, `\t`), and a `>` inside a heading in `Section` is written `\>`. Split a row on ` | `, then decode each cell. The other tiers do not escape: their table cells already use ` | ` inside the `Text` column.

### Counts in the outline

For the outline, `rowCount` and `totalNodes` count heading rows, and `offset`, `limit`, `hasMore` and `nextOffset` page over heading rows. The reply also carries `outline`, with counts for the whole document: `headings`, `nodes`, `tables`, `images` and `words`. The CLI prints the same counts in its footer and in `--json`.

With `aroundNode`, a node that is not a heading shows the heading above it, and the reply gives that heading's id as `aroundResolvedTo`. Before the first heading, or in a document with no headings, `aroundResolvedTo` is `null` and a `note` says why.

### What a section is

A section in the sense of `sectionId`, the `sections=` footer and the `sec_root` id is a page section of the file: the stretch between two section breaks. Most documents have one. A section is not a heading and does not nest. To read the part of the document under a heading, use the outline to find the heading, then `aroundNode` on its id.

### Paging rules

- With no paging argument, a small reply comes back whole. A large one starts at the first rows and tells you how to continue.
- `aroundNode` and `offset` or `limit` cannot be used together.
- `limit` is at most 200 per call; follow `nextOffset` to read more.
- `maxTokens` drops rows from the end of the window. It counts tokens with an estimate of characters divided by 3.8, which is not a tokenizer.

### The page estimate

`pages~` in `astra stats`, `astra inspect` and `get_stats` is an estimate: `max(1, ceil(words/450 + 0.45 per table or callout + 0.5 per image + 1 per page break))`. A callout is a one-cell filled table; `astra stats` counts it apart from the tables, and the estimate counts both. Tables, callouts, images and page breaks are counted at fixed weights because the estimate cannot size them, and the output says how much of the estimate comes from them. For the renderer's page count, use `astra visual`.

## Part 2. Measured sizes

This part can be removed without touching part 1.

**What was measured.** One document generated by `scripts/measure-read-tiers.mjs`: 12 chapters, each with a Heading 1, three Heading 2 sections of two Heading 3 parts, body paragraphs and one table. It has 11,100 words, 12 tables and 324 nodes, of which 120 are headings. The text is generated, so the document and its outputs can be shared.

**How.** `node scripts/measure-read-tiers.mjs --keep <dir>` builds the document with `astra create --from-md`, reads it with `astra inspect` at each tier, and prints rows, characters and the estimate of characters divided by 3.8. The token counts below were taken once, on the files that command keeps, with the `o200k_base` encoding from the `gpt-tokenizer` package (version 4.1.2). That is a proxy: it is not the tokenizer of any Claude model, and counts for a Claude model differ.

| Read | Rows | Characters | Tokens (o200k_base) |
|---|---:|---:|---:|
| `astra inspect --outline` | 120 | 6,939 | 3,134 |
| `astra inspect --fidelity compact` | 324 | 86,942 | 19,094 |
| `astra inspect --fidelity standard` | 324 | 87,122 | 19,190 |
| `astra inspect --fidelity compact --max-text 40` | 324 | 23,108 | 8,632 |
| `astra inspect --fidelity compact --section <one node id>` | 1 | 136 | 51 |
| the Markdown source the document was made from | | 74,724 | 12,890 |

What these show, for this document only:

- The outline is about one sixth of the compact tier in tokens. The saving depends on how much body text sits under each heading.
- Cutting the text of every row with `--max-text 40` removes more than half of the tokens of the compact tier and keeps every row and id.
- The estimate of characters divided by 3.8 is far from the o200k count in both directions: it gives 1,826 tokens for the outline (counted: 3,134) and 22,879 for the compact tier (counted: 19,094). Short rows made of ids and short headings use more tokens per character than prose. Treat `maxTokens` as a rough cap, and leave room under it.
- The standard tier adds little over the compact tier here because the generated paragraphs carry no formatting. A document with formatted runs adds more.

No claim is made about other documents or other tools.
