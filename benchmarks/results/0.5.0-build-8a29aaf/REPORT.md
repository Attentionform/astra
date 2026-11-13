# ASTra benchmark report: release 0.5.0-build-8a29aaf

> **Work in progress.** ASTra is improved every day; these numbers describe one published release on one machine and will change. The protocol (`PROTOCOL.md`) was committed before the first measurement. Raw per-trial data sits next to this file as `*.jsonl`. Nothing was dropped.

## Pre-registered expectations versus results

| id | expectation | measured | verdict |
| --- | --- | --- | --- |
| H1 | median standard-outline ratio >= 3x | 2.89x | NOT met |
| H2 | 0 non-target paragraph content changes (N vs E) | 1047/1047 | met |
| H3 | all produced files pass O1-O5 (E1) | 1132/1132 | met |
| H8 | C1 content fidelity >= 98% of blocks | 100.00% | met |
| H4 | gate catches >= 95% of XSD/package faults, 0 false alarms | 86.3% caught (95.0% caught or repaired), 0 false alarms | NOT met |
| H5 | 100% atomic outcomes among rejected batches | 108/108 | met |
| H6 | 100% byte-exact restores | document.xml 300/300; every part 300/300; reachable parts 300/300; whole file 0/300 | NOT met |

## Environment

```json
{
 "release_under_test": "0.5.0-build-8a29aaf",
 "measured_on": "2026-11-13",
 "os": "Linux 7.0.0-1016-aws x64",
 "wsl2": false,
 "cpu": "Intel(R) Xeon(R) 6975P-C",
 "cpu_threads": 4,
 "ram_gb": 15.3,
 "node": "v22.23.4",
 "libreoffice": "LibreOffice 26.2.6.3 620(Build:3)",
 "python": "Python 3.14.4",
 "python_docx": "1.2.0",
 "lxml": "(6, 1, 3, 0)",
 "pdfinfo": "poppler pdfinfo",
 "tokenizer": "gpt-tokenizer (o200k_base), a proxy for any particular model",
 "packages_under_test": {
  "@attentionform/astra-engine": {
   "version": "0.4.2",
   "integrity": "sha512-jTFeFvgn1O4YwNkSRfcCG3ImhHvsQoR0HkSsxxutPnmUnZhb75JVV+UlUKW7D+DrKloc9rwXUcHmNIYm7V1Ehg=="
  },
  "@attentionform/astra-engine/node_modules/adm-zip": {
   "version": "0.6.1",
   "integrity": "sha512-Xwrja8nx9e5o2N1my4DsKCeKpdrnACyr1wtbPxBDgGzKzKyE9kRtBFA8mWldI+RVlD7CBZNWY/wQ2+ydwOR6kQ=="
  },
  "@attentionform/astra-mcp-server": {
   "version": "0.4.2",
   "integrity": "sha512-C71NCNaNDBrF3eFolJTWCHK9jKT50ljU7SmoUHOnnoJmMAPib31+TKVHM1OyuJnw1HoWuPiGHOmVj7+w4Z8qDw=="
  },
  "@attentionform/astra-mcp-server/node_modules/adm-zip": {
   "version": "0.6.1",
   "integrity": "sha512-Xwrja8nx9e5o2N1my4DsKCeKpdrnACyr1wtbPxBDgGzKzKyE9kRtBFA8mWldI+RVlD7CBZNWY/wQ2+ydwOR6kQ=="
  },
  "@attentionform/ooxml-validator": {
   "version": "0.4.2",
   "integrity": "sha512-xBDRp3uLk6vKl4EDEDTsWPk3ytsqG8Ba4H1Obkvn8Gmx9QClAM1qZyIDr1b8YTPixfJ82yW/m67tcBiudc+eoQ=="
  },
  "@attentionform/ooxml-validator-darwin-arm64": {
   "version": "0.4.2",
   "integrity": "sha512-thZeIAVScr6by5SjqmBBktP0IpTtUODacXB7Mi81r9TLYxsrKEBRvwPUpngivnXrjTpk74ZJOjbSXaVbki+b+g=="
  },
  "@attentionform/ooxml-validator-darwin-x64": {
   "version": "0.4.2",
   "integrity": "sha512-SX5FDROP9WQUUQCZ3YUVKTBDUtfStY4+0/JkQHdoHwM9Jnr+KzYTyqxTCf6fpWKOMHiLMuNQ1WC1vjJLpnACTQ=="
  },
  "@attentionform/ooxml-validator-linux-arm64": {
   "version": "0.4.2",
   "integrity": "sha512-sdz7/311AIESDTFSoizOGmIEQscAA6PZEM5v2BJXKR49uuqjI5gEQyM2eGlHK/B2Yi+hP+ZTHTWwi485XfTvDw=="
  },
  "@attentionform/ooxml-validator-linux-x64": {
   "version": "0.4.2",
   "integrity": "sha512-UkJus9xYSpo9KD/pKobLAhC++gpnHg4Yz5ps+3aTZA709KTg5s0tBku3krMCfOzzlael9NMciGagJHaUbnbejA=="
  },
  "@attentionform/ooxml-validator-win32-arm64": {
   "version": "0.4.2",
   "integrity": "sha512-g3MSRHnFjKKz5KIPraU0w1WB6qpH4UU75qKEk2lqamSTERfph1eqdLYh7/Oop4LJOX7mLT5ad9FnrrvqQ2e6Fg=="
  },
  "@attentionform/ooxml-validator-win32-x64": {
   "version": "0.4.2",
   "integrity": "sha512-sYr6X5WO8YQdkg9YWlYUdchA8toQEKIXn9UswkKbY+z4JZY+OxPF+CMMiY9ZP5iPE80x3EVOsEi0UncXM/78gw=="
  },
  "@modelcontextprotocol/sdk": {
   "version": "1.32.2",
   "integrity": "sha512-SWsVtH/0vULSQStoZ6xxH2K09Li8fXYTMC64ZeogCdpZ3XSi1eCPOByV1OpuJxuI3OcmHhosDWYyii07tpPsLw=="
  },
  "adm-zip": {
   "version": "0.5.16",
   "integrity": "sha512-TGw5yVi4saajsSEgz25grObGHEUaDrniwvA2qwSC060KfqGPdglhvPMA2lPIoxs3PQIItj2iag35fONcQqgUaQ=="
  },
  "gpt-tokenizer": {
   "version": "3.0.1",
   "integrity": "sha512-5jdaspBq/w4sWw322SvQj1Fku+CN4OAfYZeeEg8U7CWtxBz+zkxZ3h0YOHD43ee+nZYZ5Ud70HRN0ANcdIj4qg=="
  }
 },
 "local_build": {
  "server_bin": "/home/runner/work/astra/astra/target/packages/mcp-server/dist/index.js",
  "git_head": "8a29aaf4914c70fa7a0bb0e007a6d61837d76ae4",
  "github_run_id": "43466748250",
  "runner": "GitHub Actions 1000001365"
 }
}
```

## R1. Reading cost: document.xml versus the ASTra outline

Tokenizer: o200k_base (proxy, see PROTOCOL section 8). Ratio = tokens of `word/document.xml` divided by tokens of the outline (standard fidelity, every page of results). "vs text floor" is outline tokens divided by the tokens of the bare document text, i.e. the overhead the outline adds on top of the words themselves.

**All 40 documents: median ratio 2.89x (bootstrap 95% CI 2.61 to 4.30), interquartile range 2.51 to 5.10, range 1.70 to 25.91.**

| document kind | n | median ratio | min | max | outline / text floor (median) |
| --- | --- | --- | --- | --- | --- |
| comments | 2 | 2.64 | 2.52 | 2.75 | 1.40 |
| heavy-runs | 1 | 12.83 | 12.83 | 12.83 | 1.65 |
| links-bookmarks | 1 | 2.84 | 2.84 | 2.84 | 1.44 |
| lists | 6 | 2.22 | 1.70 | 4.79 | 2.27 |
| mixed | 6 | 4.30 | 4.24 | 11.04 | 1.87 |
| prose | 14 | 2.52 | 2.34 | 3.32 | 1.39 |
| tables | 7 | 6.25 | 6.03 | 25.91 | 1.99 |
| tables-wide | 1 | 6.38 | 6.38 | 6.38 | 1.85 |
| tracked | 2 | 2.74 | 2.55 | 2.93 | 1.38 |

By writer: python-docx median 2.87x (n=27); python-docx+lxml median 2.84x (n=3); libreoffice median 4.65x (n=10)

By size: 1-2 pages median 4.05x (n=6); 3-12 pages median 2.89x (n=16); 13-60 pages median 2.55x (n=11); >60 pages median 2.49x (n=7)

Compact fidelity median ratio: 2.98x. Standard with full paragraph text (maxTextLength 10000): 2.89x.

## E1. Edit correctness and locality

1092 edit trials (40 documents x up to 8 edit types x 4 positions; types that need a table are skipped where a document has none: 188 skipped). Locality compares N (document saved with no edit) and E (document saved with the edit): content view, identifier attributes stripped.

- Edit applied and saved: 1092/1092 (100.0%, 95% CI 99.6% to 100.0%)
- Targets that could be located uniquely and analysed: 1047
- **Zero non-target paragraphs changed:** 1047/1047 (100.0%, 95% CI 99.6% to 100.0%)
- Text of every non-target paragraph preserved: 1047/1047 (100.0%, 95% CI 99.6% to 100.0%)
- Intended change present in the target: 1047/1047 (100.0%, 95% CI 99.6% to 100.0%)

| edit type | trials | saved | zero non-target changes | intended change present | median parts changed |
| --- | --- | --- | --- | --- | --- |
| replace | 148 | 148 | 142/142 (100.0%, 95% CI 97.4% to 100.0%) | 142/142 (100.0%, 95% CI 97.4% to 100.0%) | 1.0 |
| style-word | 148 | 148 | 141/141 (100.0%, 95% CI 97.3% to 100.0%) | 141/141 (100.0%, 95% CI 97.3% to 100.0%) | 1.0 |
| insert | 148 | 148 | 141/141 (100.0%, 95% CI 97.3% to 100.0%) | 141/141 (100.0%, 95% CI 97.3% to 100.0%) | 1.0 |
| delete | 148 | 148 | 142/142 (100.0%, 95% CI 97.4% to 100.0%) | 142/142 (100.0%, 95% CI 97.4% to 100.0%) | 1.0 |
| comment | 148 | 148 | 142/142 (100.0%, 95% CI 97.4% to 100.0%) | 142/142 (100.0%, 95% CI 97.4% to 100.0%) | 6.0 |
| tracked-replace | 148 | 148 | 142/142 (100.0%, 95% CI 97.4% to 100.0%) | 142/142 (100.0%, 95% CI 97.4% to 100.0%) | 1.0 |
| named-style | 148 | 148 | 141/141 (100.0%, 95% CI 97.3% to 100.0%) | 141/141 (100.0%, 95% CI 97.3% to 100.0%) | 1.0 |
| table-cell | 56 | 56 | 56/56 (100.0%, 95% CI 93.6% to 100.0%) | 56/56 (100.0%, 95% CI 93.6% to 100.0%) | 1.0 |

### First-save normalization (A versus N, 40 documents)

Opening and saving a document written by another program, with no edit, changes it as follows. Median share of paragraphs whose raw XML changes: 100.0%; share whose content changes once identifier attributes are ignored: 0.0%. Median number of parts with different bytes: 2.0; with different XML content: 2.0. Text identical after save in 40/40 documents.


### E1 control: plain parse-and-reserialize of every XML part (no product, 40 documents)

Median parts with different bytes: 8.0 of 17.0; with different XML content: 0.0. Median share of paragraphs with different raw XML: 0.0%; content-level: 0.0%.

## Independent oracles on E1 output

1132/1132 (100.0%, 95% CI 99.7% to 100.0%) pass O1 to O5 (O1 1132/1132, O2 1132/1132, O3 1132/1132, O4 1132/1132, O5 1132/1132).

## Independent oracles on C1 output

26/26 (100.0%, 95% CI 87.1% to 100.0%) pass O1 to O5 (O1 26/26, O2 26/26, O3 26/26, O4 26/26, O5 26/26).

## Independent oracles on C2 output

15/15 (100.0%, 95% CI 79.6% to 100.0%) pass O1 to O5 (O1 15/15, O2 15/15, O3 15/15, O4 15/15, O5 15/15).

## C2. Building with typed operations (15 recipes)

15/15 (100.0%, 95% CI 79.6% to 100.0%) recipes built, saved and contain the intended structure.

| recipe | batch applied | saved | gate verdict | intended structure |
| --- | --- | --- | --- | --- |
| styled-headings | yes | yes | valid | present |
| paragraph-format | yes | yes | valid | present |
| bullet-list | yes | yes | valid | present |
| numbered-list | yes | yes | valid | present |
| table-merged | yes | yes | valid | present |
| table-styled | yes | yes | valid | present |
| inline-image | yes | yes | valid | present |
| callout | yes | yes | valid | present |
| chart | yes | yes | valid | present |
| equation | yes | yes | valid | present |
| footnote | yes | yes | valid | present |
| header-footer | yes | yes | valid | present |
| page-setup-columns | yes | yes | valid | present |
| section-break | yes | yes | valid | present |
| hyperlink-comment | yes | yes | valid | present |



## C1. Creating documents from Markdown (26 trials)

Compose, validate and save succeeded in 26/26 (100.0%, 95% CI 87.1% to 100.0%). Content fidelity (independent reader, blocks found in order): **6602/6602 (100.0%, 95% CI 99.9% to 100.0%)**; heading level / list numbering / table shape / quote style correct on found blocks: 4713/4713 (100.0%, 95% CI 99.9% to 100.0%). Writing cost: the Markdown an agent writes is a median 12.8x fewer tokens than the `word/document.xml` that results (range 9.9 to 17.8).

| trial | template | md tokens | xml tokens | xml/md | compose | save | gate | fidelity |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| md-00-1p | blank | 273 | 4333 | 15.9 | 216 ms | 140 ms | valid | 12/12 |
| md-01-1p | blank | 392 | 6961 | 17.8 | 81 ms | 148 ms | valid | 6/6 |
| md-02-2p | blank | 303 | 3000 | 9.9 | 86 ms | 132 ms | valid | 27/27 |
| md-03-2p | blank | 471 | 6942 | 14.7 | 76 ms | 140 ms | valid | 13/13 |
| md-04-3p | blank | 640 | 7214 | 11.3 | 81 ms | 160 ms | valid | 29/29 |
| md-05-5p | blank | 1179 | 17812 | 15.1 | 151 ms | 192 ms | valid | 47/47 |
| md-06-5p | blank | 1031 | 12521 | 12.1 | 98 ms | 154 ms | valid | 49/49 |
| md-07-8p | blank | 1917 | 24689 | 12.9 | 190 ms | 225 ms | valid | 68/68 |
| md-08-10p | blank | 2644 | 37492 | 14.2 | 252 ms | 274 ms | valid | 97/97 |
| md-09-10p | blank | 2161 | 25990 | 12.0 | 173 ms | 221 ms | valid | 101/101 |
| md-10-15p | blank | 3085 | 35569 | 11.5 | 287 ms | 285 ms | valid | 146/146 |
| md-11-20p | blank | 4931 | 62240 | 12.6 | 450 ms | 397 ms | valid | 224/224 |
| md-12-25p | blank | 5879 | 65372 | 11.1 | 480 ms | 435 ms | valid | 283/283 |
| md-13-25p | blank | 6549 | 92320 | 14.1 | 661 ms | 556 ms | valid | 286/286 |
| md-14-30p | blank | 6858 | 87454 | 12.8 | 629 ms | 519 ms | valid | 283/283 |
| md-15-40p | blank | 10407 | 133263 | 12.8 | 1.14 s | 776 ms | valid | 417/417 |
| md-16-50p | blank | 11969 | 136392 | 11.4 | 1.23 s | 810 ms | valid | 523/523 |
| md-17-50p | blank | 12669 | 174041 | 13.7 | 1.55 s | 966 ms | valid | 496/496 |
| md-18-60p | blank | 14323 | 176053 | 12.3 | 1.81 s | 1.00 s | valid | 626/626 |
| md-19-75p | blank | 19304 | 246614 | 12.8 | 2.58 s | 1.37 s | valid | 760/760 |
| md-20-100p | blank | 23396 | 287827 | 12.3 | 3.60 s | 1.66 s | valid | 978/978 |
| md-21-100p | blank | 24545 | 305328 | 12.4 | 3.96 s | 1.79 s | valid | 1030/1030 |
| md-22-2p | blank | 417 | 5315 | 12.7 | 52 ms | 103 ms | valid | 17/17 |
| md-23-5p | blank | 1033 | 13490 | 13.1 | 91 ms | 147 ms | valid | 50/50 |
| tpl-report | report | 585 | 8681 | 14.8 | 71 ms | 116 ms | valid | 17/17 |
| tpl-memo | memo | 585 | 8680 | 14.8 | 63 ms | 118 ms | valid | 17/17 |

### C1b. Out-of-scope Markdown must fail closed (12 inputs)

| input | result | code | document unchanged | note |
| --- | --- | --- | --- | --- |
| heading-level-4 | refused | E_VALIDATION | yes | [E_VALIDATION] line 3: heading level 4 is out of scope (levels 1-3 only) Remedy: use #, ## |
| nested-list | refused | E_VALIDATION | yes | [E_VALIDATION] line 2: nested lists are out of scope Remedy: flatten nested items into the |
| task-list | refused | E_VALIDATION | yes | [E_VALIDATION] line 1: task-list items are out of scope Remedy: use a plain bullet or numb |
| raw-html | refused | E_VALIDATION | yes | [E_VALIDATION] line 3: raw HTML blocks are out of scope Remedy: use plain Markdown blocks  |
| nested-table-in-quote | refused | E_VALIDATION | yes | [E_VALIDATION] line 1: tables inside a quote are out of scope Remedy: Word cannot indent a |
| heading-in-quote | refused | E_VALIDATION | yes | [E_VALIDATION] line 1: headings inside a quote are out of scope Remedy: put the heading be |
| remote-image | refused | E_VALIDATION | yes | [E_VALIDATION] line 1: image reference points outside the allowed directory Remedy: keep r |
| image-without-allowedDir | refused | E_VALIDATION | yes | [E_VALIDATION] line 1: image needs the folder holding the picture file Remedy: pass allowe |
| table-without-separator | accepted (documented) | - | - | no separator row means plain text, as in Markdown |
| html-comment-script | refused | E_VALIDATION | yes | [E_VALIDATION] line 1: raw HTML blocks are out of scope Remedy: use plain Markdown blocks  |
| deep-nested-quote | accepted (documented) | - | - | documented: quote levels beyond 3 are written at level 3 |
| custom-style-block | refused | E_VALIDATION | yes | [E_VALIDATION] line 1: container fences (:::) are out of scope as custom styles Remedy: us |

Refused with nothing written: 10 of 10 inputs that are outside the documented subset.

## E3. Schema gate: injected faults

158 fault files (60 fault classes x 3 base documents; a class that does not apply to a base is skipped). Each file was opened, validated, then saved after a trivial edit in strict mode and in default mode. "Flagged invalid" means `validate_document` reports more errors than the same base document had before injection.

- **Caught (refused at open or flagged invalid):** 129/158 (81.6%, 95% CI 74.9% to 86.9%); excluding the semantic cross-reference class: 120/139 (86.3%, 95% CI 79.6% to 91.1%)
- **Caught or repaired by the save-time normalizer** (element-order faults are put back in schema order and the file is written valid): 143/158 (90.5%, 95% CI 84.9% to 94.2%); excluding the semantic class: 132/139 (95.0%, 95% CI 90.0% to 97.5%)
- Strict save refused the caught faults (bases with no pre-existing errors, 155 files): 126/126 (100.0%, 95% CI 97.0% to 100.0%); strict saves that changed the file on disk: 0
- Default mode wrote the file for 112 of 126 caught faults: default mode is documented to accept a save that adds no new error signature, and these faults were already in the input file; it refuses edits that add new errors (phase B below).
- **False alarms** on valid-mutation controls (more errors than their base, or a strict refusal on an error-free base): 0/58 (0.0%, 95% CI 0.0% to 6.2%)
- Clean corpus documents that already carry a schema error before any injection: 0 of 40

| category | files | refused at open | flagged invalid | repaired at save | not flagged |
| --- | --- | --- | --- | --- | --- |
| attribute | 28 | 0 | 28 | 0 | 0 |
| identifier | 26 | 0 | 15 | 3 | 8 |
| missing | 27 | 0 | 21 | 2 | 4 |
| order | 25 | 0 | 16 | 9 | 0 |
| package | 24 | 3 | 18 | 0 | 3 |
| unknown | 28 | 0 | 28 | 0 | 0 |

### Not flagged (15)

These are published, not hidden. 15 of the 15 are also accepted by the standalone upstream validator run directly on the injected file (no ASTra code in the path), so they are limits of the upstream rules (or of how I labelled a fault as invalid), not something ASTra added. The semantic class (cross-references) is outside what an XSD validator checks.

What default mode writes: 112 of 141 files written by default-mode saves of these fault files still contain the injected error, because default mode is documented to tolerate errors that were already in the input and to refuse only edits that add new ones. Use strict mode to refuse such inputs (strict refused 100% of the caught faults above).

| id | base | kind | fault | rule |
| --- | --- | --- | --- | --- |
| D01 | pd-mixed-010p | schema | tc without any paragraph | CT_Tc requires a block-level element |
| D03 | pd-mixed-010p | schema | tbl without rows | CT_Tbl requires at least one tr |
| E04 | pd-mixed-010p | semantic | bookmarkEnd without start | bookmark pairing |
| E06 | pd-mixed-010p | semantic | numId refers to missing num | numbering reference |
| E07 | pd-mixed-010p | semantic | pStyle refers to missing style | style reference |
| F10 | pd-mixed-010p | package | document part has a wrong content type | content type of main part |
| D01 | lo-mixed-010p | schema | tc without any paragraph | CT_Tc requires a block-level element |
| D03 | lo-mixed-010p | schema | tbl without rows | CT_Tbl requires at least one tr |
| E04 | lo-mixed-010p | semantic | bookmarkEnd without start | bookmark pairing |
| E07 | lo-mixed-010p | semantic | pStyle refers to missing style | style reference |
| F10 | lo-mixed-010p | package | document part has a wrong content type | content type of main part |
| E04 | pd-comments-005p | semantic | bookmarkEnd without start | bookmark pairing |
| E06 | pd-comments-005p | semantic | numId refers to missing num | numbering reference |
| E07 | pd-comments-005p | semantic | pStyle refers to missing style | style reference |
| F10 | pd-comments-005p | package | document part has a wrong content type | content type of main part |

### E3 phase B. Edits that try to introduce invalid content (36 edits on a valid document, default mode)

| route | edit | outcome | code | file |
| --- | --- | --- | --- | --- |
| raw-xml | x01 jc=middle | blocked at save | E_VALIDATION | nothing written |
| raw-xml | x02 unknown element | blocked at save | E_VALIDATION | nothing written |
| raw-xml | x03 sz=abc | blocked at save | E_VALIDATION | nothing written |
| raw-xml | x04 table without rows | rejected at edit | E_VALIDATION | nothing written |
| raw-xml | x05 malformed xml | rejected at edit | E_VALIDATION | nothing written |
| raw-xml | x06 empty drawing | rejected at edit | E_VALIDATION | nothing written |
| raw-xml | x07 bookmark without id | blocked at save | E_VALIDATION | nothing written |
| raw-xml | x08 pPr order | accepted and saved | - | written, passes O1-O5 |
| raw-xml | x09 color=ZZZZZZ | blocked at save | E_VALIDATION | nothing written |
| raw-xml | x10 tc without p | rejected at edit | E_VALIDATION | nothing written |
| raw-xml | x11 ins without author | blocked at save | E_VALIDATION | nothing written |
| raw-xml | x12 foreign namespace | rejected at edit | E_VALIDATION | nothing written |
| raw-xml | x13 xml:space=bad | blocked at save | E_VALIDATION | nothing written |
| raw-xml | x14 text outside p | blocked at save | E_VALIDATION | nothing written |
| raw-xml | x15 VALID paragraph | accepted and saved | - | written, passes O1-O5 |
| raw-xml | x16 VALID centred | accepted and saved | - | written, passes O1-O5 |
| typed | t01 color=banana | blocked at save | E_VALIDATION | nothing written |
| typed | t02 fontSize=-1 | rejected at edit | TOOL_INPUT_SCHEMA | nothing written |
| typed | t03 fontSize=0 | accepted and saved | - | written, passes O1-O5 |
| typed | t04 fontSize=99999 | accepted and saved | - | written, passes O1-O5 |
| typed | t05 alignment=middle | rejected at edit | TOOL_INPUT_SCHEMA | nothing written |
| typed | t06 highlight=purple | blocked at save | E_VALIDATION | nothing written |
| typed | t07 mergeCells out of range | rejected at edit | E_VALIDATION | nothing written |
| typed | t08 image not base64 | rejected at edit | E_VALIDATION | nothing written |
| typed | t09 image bytes not an image | rejected at edit | E_VALIDATION | nothing written |
| typed | t10 chart non-numeric values | rejected at edit | TOOL_INPUT_SCHEMA | nothing written |
| typed | t11 chart no series | accepted and saved | - | written, passes O1-O5 |
| typed | t12 hyperlink javascript: | rejected at edit | E_VALIDATION | nothing written |
| typed | t13 bookmark name with spaces | accepted and saved | - | written, passes O1-O5 |
| typed | t14 unknown named style | accepted and saved | - | written, passes O1-O5 |
| typed | t15 NUL in text | blocked at save | E_VALIDATION | nothing written |
| typed | t16 negative margin | accepted and saved | - | written, passes O1-O5 |
| typed | t17 columns=99 | rejected at edit | TOOL_INPUT_SCHEMA | nothing written |
| typed | t18 empty comment | rejected at edit | TOOL_INPUT_SCHEMA | nothing written |
| typed | t19 ragged table rows | accepted and saved | - | written, passes O1-O5 |
| typed | t20 VALID bold | accepted and saved | - | written, passes O1-O5 |

Outcome counts: blocked-at-save 11, rejected-at-edit 14, saved 11. Every file that was written passed O1 to O5 (11/11).

Edits that were accepted and saved are schema-valid even when they are unwise (font size 0 or 99999, a negative margin, a chart with no series, a bookmark name with spaces, a style name the document does not define, a table with ragged rows). The gate answers "is this valid OOXML", not "is this sensible".

## E4. Atomicity

144 trials: batches of 2, 5, 20 or 100 operations in which one operation (first, middle or last) is bad, in four ways. 108 batches were rejected and 36 were accepted (the "unknown style" operation is not an error: ASTra applies a style name that the document does not define; see limits). Of the 108 rejected batches, **the session was unchanged afterwards (outline, revision, dirty flag, history length) in 108/108 (100.0%, 95% CI 96.6% to 100.0%).** The same batch without the bad operation applied afterwards in 144/144 (100.0%, 95% CI 97.4% to 100.0%).

| failure kind | trials | rejected | rejected by | session unchanged after rejection |
| --- | --- | --- | --- | --- |
| unknown-target | 36 | 36 | E_VALIDATION | 36/36 |
| schema-invalid-param | 36 | 36 | TOOL_INPUT_SCHEMA | 36/36 |
| unknown-style | 36 | 0 | not rejected | n/a |
| absent-match | 36 | 36 | E_NOT_FOUND | 36/36 |

Note: "TOOL_INPUT_SCHEMA" means the MCP tool's own input schema refused the call before the engine ran, so those trials show the contract, not the engine's rollback. The engine-level rollback trials are the 72 rejected with E_VALIDATION or E_NOT_FOUND.

## E5. History, revert and redo

30 sequences of 10 saved edits on 10 documents x 3 seeds.

**Revert to a saved commit (by commit id), every depth 1 to 10, each followed by redo: 300 reverts.**

- Restored package parts identical to the state saved at that commit (SHA-256 of every part): 300/300 (100.0%, 95% CI 98.7% to 100.0%)
- Restored `word/document.xml` identical: 300/300 (100.0%, 95% CI 98.7% to 100.0%)
- **Parts reachable through relationships identical** (what Word or any reader loads): 300/300 (100.0%, 95% CI 98.7% to 100.0%)
- Parts left behind or missing after the revert (name: reverts): none
- Whole file byte-identical, including the ZIP container: 0/300 (0.0%, 95% CI 0.0% to 1.3%)
- Redo returns every part of the head state identically: 300/300 (100.0%, 95% CI 98.7% to 100.0%)
- `preview_revert` file list equals the parts that actually changed: 300/300 (100.0%, 95% CI 98.7% to 100.0%)
- Every save appears in the history listing: 30/30; listing order matches save order: 30/30

**Revert by step count (`steps`), fresh history per depth 1, 2, 5, 9 (first seed of each document): 40 reverts.** Every part identical: 40/40 (100.0%, 95% CI 91.2% to 100.0%); reachable parts identical: 40/40 (100.0%, 95% CI 91.2% to 100.0%); preview names the intended commit: 40/40 (100.0%, 95% CI 91.2% to 100.0%).


### E5 follow-up: files reverted past a comment (8 documents)

After the revert, `word/people.xml` is still in the ZIP in 0/8, `word/commentsExtended.xml` in 0/8 and `word/comments.xml` in 1/8; the revert target is the state saved before the comment was added. Independent checks on the reverted files: O1 to O5 pass 8/8; standalone validator error count equals the count before the comment was added in 8/8.

### History store growth

| document | saves | docx bytes | history bytes | history / docx |
| --- | --- | --- | --- | --- |
| pd-prose-005p | 0 | 41513 | 935335 | 22.53 |
| pd-prose-005p | 10 | 39836 | 977834 | 24.55 |
| pd-prose-005p | 20 | 38360 | 1006629 | 26.24 |
| pd-prose-005p | 30 | 37068 | 1021434 | 27.56 |
| pd-prose-005p | 40 | 37072 | 1036352 | 27.96 |
| pd-prose-005p | 50 | 37068 | 1051260 | 28.36 |
| pd-prose-050p | 0 | 79624 | 1223621 | 15.37 |
| pd-prose-050p | 10 | 78200 | 1649030 | 21.09 |
| pd-prose-050p | 20 | 76723 | 2060392 | 26.85 |
| pd-prose-050p | 30 | 75224 | 2456255 | 32.65 |
| pd-prose-050p | 40 | 73879 | 2839707 | 38.44 |
| pd-prose-050p | 50 | 72334 | 3207001 | 44.34 |
| pd-mixed-100p | 0 | 157745 | 1908112 | 12.10 |
| pd-mixed-100p | 10 | 157119 | 2921977 | 18.60 |
| pd-mixed-100p | 20 | 156599 | 3930216 | 25.10 |
| pd-mixed-100p | 30 | 155823 | 4930783 | 31.64 |
| pd-mixed-100p | 40 | 155321 | 5926150 | 38.15 |
| pd-mixed-100p | 50 | 154682 | 6915533 | 44.71 |
