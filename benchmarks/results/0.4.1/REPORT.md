# ASTra benchmark report: release 0.4.1

> **Work in progress.** ASTra is improved every day; these numbers describe one published release on one machine and will change. The protocol (`PROTOCOL.md`) was committed before the first measurement. Raw per-trial data sits next to this file as `*.jsonl`. Nothing was dropped.

## Pre-registered expectations versus results

| id | expectation | measured | verdict |
| --- | --- | --- | --- |
| H1 | median standard-outline ratio >= 3x | 2.89x | NOT met |
| H2 | 0 non-target paragraph content changes (N vs E) | 1047/1047 | met |
| H3 | all produced files pass O1-O5 (E1) | 1132/1132 | met |
| H8 | C1 content fidelity >= 98% of blocks | 100.00% | met |
| H4 | gate catches >= 95% of XSD/package faults, 0 false alarms | 83.5% caught (92.1% caught or repaired), 0 false alarms | NOT met |
| H5 | 100% atomic outcomes among rejected batches | 108/108 | met |
| H6 | 100% byte-exact restores | document.xml 300/300; every part 189/300; reachable parts 300/300; whole file 2/300 | NOT met |
| H7 | median save < 5 s at 100 pages | 712 ms | met |

## Environment

```json
{
 "release_under_test": "0.4.1",
 "measured_on": "2026-11-11",
 "os": "Linux 6.6.114.1-microsoft-standard-WSL2 x64",
 "wsl2": true,
 "cpu": "13th Gen Intel(R) Core(TM) i7-13700HX",
 "cpu_threads": 24,
 "ram_gb": 7.6,
 "node": "v26.8.1",
 "libreoffice": "LibreOffice 25.2.7.2 520(Build:2)",
 "python": "Python 3.13.13",
 "python_docx": "1.2.0",
 "lxml": "(5, 3, 2, 0)",
 "pdfinfo": "poppler pdfinfo",
 "tokenizer": "gpt-tokenizer (o200k_base), a proxy for any particular model",
 "packages_under_test": {
  "@attentionform/astra-engine": {
   "version": "0.4.1",
   "integrity": "sha512-QBLjRb4BxClGwxm3PIc8vrcPIy9bOdyU5Wk3qWQ1GtsJouz8s9mi9K2Jb4SGH9bZuPa3i/sHLviUtdm4YKhB5w=="
  },
  "@attentionform/astra-engine/node_modules/adm-zip": {
   "version": "0.6.1",
   "integrity": "sha512-Xwrja8nx9e5o2N1my4DsKCeKpdrnACyr1wtbPxBDgGzKzKyE9kRtBFA8mWldI+RVlD7CBZNWY/wQ2+ydwOR6kQ=="
  },
  "@attentionform/astra-mcp-server": {
   "version": "0.4.1",
   "integrity": "sha512-cYYvZMVRyqSfIXTIwI7L+J0yvSZch0dFi06SknEBotxU9kkUzqT451PJOxUR8JvaHD+efn6NnORT6eqY/03n4Q=="
  },
  "@attentionform/astra-mcp-server/node_modules/adm-zip": {
   "version": "0.6.1",
   "integrity": "sha512-Xwrja8nx9e5o2N1my4DsKCeKpdrnACyr1wtbPxBDgGzKzKyE9kRtBFA8mWldI+RVlD7CBZNWY/wQ2+ydwOR6kQ=="
  },
  "@attentionform/ooxml-validator": {
   "version": "0.4.1",
   "integrity": "sha512-I0tPJWNM6g/VdEZfnuBFcaibIgBV9z7I2v2Er26j4QBUO5yR82dyBJwMzU1qSIKNBwwtzFOMXoH99tt1JlJYWQ=="
  },
  "@attentionform/ooxml-validator-darwin-arm64": {
   "version": "0.4.1",
   "integrity": "sha512-lbOlpocHu+ivsQzYm+u7y7n/V6eS4O6YQ2GZ4FXoVlHgG7rzPg+dSfKyArLHvspXRJE3y/Br+KTF3069ASBdFw=="
  },
  "@attentionform/ooxml-validator-darwin-x64": {
   "version": "0.4.1",
   "integrity": "sha512-SUPjRVs0H5F52w2YOKX4USCA7B6x2xzyCuXTY4yaPSUc9J+xpHXTzGwCupeE/+39cF93W3zxDnjNNIf2vE4+ug=="
  },
  "@attentionform/ooxml-validator-linux-arm64": {
   "version": "0.4.1",
   "integrity": "sha512-JnMpdaRJmlYX7FwrVUC4EPxKcjGDP8YSHotV1FrSOg9DQmAvE0CKwk9CyN8gxiY5qms3OM/loEvFUCq+aPpuIA=="
  },
  "@attentionform/ooxml-validator-linux-x64": {
   "version": "0.4.1",
   "integrity": "sha512-Rt2r+o3CmAg7S+jMq8635l4oH+vKr04dUFPXuzkwboTOEjaxj5NCOzFt+DYXN1SEh9OG2R8NVGGTXpoVBgYklw=="
  },
  "@attentionform/ooxml-validator-win32-arm64": {
   "version": "0.4.1",
   "integrity": "sha512-fC7pUvO3UoiRYMCaXkFkLBZCjJ2JOvongU13Nw3Dcg7ayDPVwV6oBxuMYDJIWjWNzXn38fl5NzNLSpsxYz6I9A=="
  },
  "@attentionform/ooxml-validator-win32-x64": {
   "version": "0.4.1",
   "integrity": "sha512-p9fG8dqVTgL5xeVHaxEh1z9AuEqFVLHNbAUODQOd8X9K0NcE/rchS8UcuXzYveh9MNjkfx+3ZdPLQzfmI9Teqg=="
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

By size: 1-2 pages median 3.32x (n=7); 3-12 pages median 2.90x (n=15); 13-60 pages median 2.55x (n=11); >60 pages median 2.49x (n=7)

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
| md-00-1p | blank | 273 | 4333 | 15.9 | 133 ms | 240 ms | valid | 12/12 |
| md-01-1p | blank | 392 | 6961 | 17.8 | 49 ms | 81 ms | valid | 6/6 |
| md-02-2p | blank | 303 | 3000 | 9.9 | 55 ms | 80 ms | valid | 27/27 |
| md-03-2p | blank | 471 | 6942 | 14.7 | 48 ms | 83 ms | valid | 13/13 |
| md-04-3p | blank | 640 | 7214 | 11.3 | 54 ms | 88 ms | valid | 29/29 |
| md-05-5p | blank | 1179 | 17812 | 15.1 | 92 ms | 112 ms | valid | 47/47 |
| md-06-5p | blank | 1031 | 12521 | 12.1 | 57 ms | 90 ms | valid | 49/49 |
| md-07-8p | blank | 1917 | 24689 | 12.9 | 93 ms | 107 ms | valid | 68/68 |
| md-08-10p | blank | 2644 | 37492 | 14.2 | 137 ms | 149 ms | valid | 97/97 |
| md-09-10p | blank | 2161 | 25990 | 12.0 | 102 ms | 125 ms | valid | 101/101 |
| md-10-15p | blank | 3085 | 35569 | 11.5 | 129 ms | 149 ms | valid | 146/146 |
| md-11-20p | blank | 4931 | 62240 | 12.6 | 223 ms | 211 ms | valid | 224/224 |
| md-12-25p | blank | 5879 | 65372 | 11.1 | 251 ms | 235 ms | valid | 283/283 |
| md-13-25p | blank | 6549 | 92320 | 14.1 | 344 ms | 290 ms | valid | 286/286 |
| md-14-30p | blank | 6858 | 87454 | 12.8 | 313 ms | 280 ms | valid | 283/283 |
| md-15-40p | blank | 10407 | 133263 | 12.8 | 572 ms | 409 ms | valid | 417/417 |
| md-16-50p | blank | 11969 | 136392 | 11.4 | 612 ms | 439 ms | valid | 523/523 |
| md-17-50p | blank | 12669 | 174041 | 13.7 | 778 ms | 528 ms | valid | 496/496 |
| md-18-60p | blank | 14323 | 176053 | 12.3 | 910 ms | 568 ms | valid | 626/626 |
| md-19-75p | blank | 19304 | 246614 | 12.8 | 1.32 s | 740 ms | valid | 760/760 |
| md-20-100p | blank | 23396 | 287827 | 12.3 | 1.83 s | 859 ms | valid | 978/978 |
| md-21-100p | blank | 24545 | 305328 | 12.4 | 2.03 s | 942 ms | valid | 1030/1030 |
| md-22-2p | blank | 417 | 5315 | 12.7 | 34 ms | 56 ms | valid | 17/17 |
| md-23-5p | blank | 1033 | 13490 | 13.1 | 51 ms | 76 ms | valid | 50/50 |
| tpl-report | report | 585 | 8681 | 14.8 | 38 ms | 67 ms | valid | 17/17 |
| tpl-memo | memo | 585 | 8680 | 14.8 | 37 ms | 63 ms | valid | 17/17 |

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

- **Caught (refused at open or flagged invalid):** 125/158 (79.1%, 95% CI 72.1% to 84.7%); excluding the semantic cross-reference class: 116/139 (83.5%, 95% CI 76.4% to 88.7%)
- **Caught or repaired by the save-time normalizer** (element-order faults are put back in schema order and the file is written valid): 139/158 (88.0%, 95% CI 82.0% to 92.2%); excluding the semantic class: 128/139 (92.1%, 95% CI 86.4% to 95.5%)
- Strict save refused the caught faults (bases with no pre-existing errors, 102 files): 82/82 (100.0%, 95% CI 95.5% to 100.0%); strict saves that changed the file on disk: 0
- Default mode wrote the file for 72 of 82 caught faults: default mode is documented to accept a save that adds no new error signature, and these faults were already in the input file; it refuses edits that add new errors (phase B below).
- **False alarms** on valid-mutation controls (more errors than their base, or a strict refusal on an error-free base): 0/58 (0.0%, 95% CI 0.0% to 6.2%)
- Clean corpus documents that already carry a schema error before any injection: 3 of 40 (lo-lists-001p.docx, lo-lists-005p.docx, lo-mixed-010p.docx; the validator reports The element has unexpected child element 'w:'. in the writer's numbering definitions; strict mode refuses to save them, default mode saves them and reports the error)

| category | files | refused at open | flagged invalid | repaired at save | not flagged |
| --- | --- | --- | --- | --- | --- |
| attribute | 28 | 0 | 28 | 0 | 0 |
| identifier | 26 | 0 | 15 | 3 | 8 |
| missing | 27 | 0 | 21 | 2 | 4 |
| order | 25 | 0 | 16 | 9 | 0 |
| package | 24 | 3 | 14 | 0 | 7 |
| unknown | 28 | 0 | 28 | 0 | 0 |

### Not flagged (19)

These are published, not hidden. 19 of the 19 are also accepted by the standalone upstream validator run directly on the injected file (no ASTra code in the path), so they are limits of the upstream rules (or of how I labelled a fault as invalid), not something ASTra added. The semantic class (cross-references) is outside what an XSD validator checks.

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
| F03 | lo-mixed-010p | package | styles override removed from content types | content type for every part |
| F04 | lo-mixed-010p | package | styles part deleted, relationship kept | OPC relationship target exists |
| F05 | lo-mixed-010p | package | duplicate relationship id | unique relationship ids |
| F08 | lo-mixed-010p | package | relationship of unknown type with bad target | relationship integrity |
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

- Restored package parts identical to the state saved at that commit (SHA-256 of every part): 189/300 (63.0%, 95% CI 57.4% to 68.3%)
- Restored `word/document.xml` identical: 300/300 (100.0%, 95% CI 98.7% to 100.0%)
- **Parts reachable through relationships identical** (what Word or any reader loads): 300/300 (100.0%, 95% CI 98.7% to 100.0%)
- Parts left behind or missing after the revert (name: reverts): left behind word/commentsExtended.xml: 111; left behind word/people.xml: 111
- Whole file byte-identical, including the ZIP container: 2/300 (0.7%, 95% CI 0.2% to 2.4%)
- Redo returns every part of the head state identically: 300/300 (100.0%, 95% CI 98.7% to 100.0%)
- `preview_revert` file list equals the parts that actually changed: 300/300 (100.0%, 95% CI 98.7% to 100.0%)
- Every save appears in the history listing: 30/30; listing order matches save order: 30/30

**Revert by step count (`steps`), fresh history per depth 1, 2, 5, 9 (first seed of each document): 40 reverts.** Every part identical: 23/40 (57.5%, 95% CI 42.2% to 71.5%); reachable parts identical: 40/40 (100.0%, 95% CI 91.2% to 100.0%); preview names the intended commit: 40/40 (100.0%, 95% CI 91.2% to 100.0%).


### E5 follow-up: files reverted past a comment (8 documents)

The unreferenced `commentsExtended.xml` and `people.xml` stay in the ZIP after a revert past a comment. Independent checks on the reverted files: O1 to O5 pass 8/8; standalone validator error count equals the count before the comment was added in 8/8.

### History store growth

| document | saves | docx bytes | history bytes | history / docx |
| --- | --- | --- | --- | --- |
| pd-prose-005p | 0 | 41384 | 420447 | 10.16 |
| pd-prose-005p | 10 | 39707 | 461553 | 11.62 |
| pd-prose-005p | 20 | 38231 | 488964 | 12.79 |
| pd-prose-005p | 30 | 36939 | 502370 | 13.60 |
| pd-prose-005p | 40 | 36943 | 515894 | 13.96 |
| pd-prose-005p | 50 | 36939 | 529415 | 14.33 |
| pd-prose-050p | 0 | 79495 | 673219 | 8.47 |
| pd-prose-050p | 10 | 78071 | 1097235 | 14.05 |
| pd-prose-050p | 20 | 76594 | 1507197 | 19.68 |
| pd-prose-050p | 30 | 75095 | 1901658 | 25.32 |
| pd-prose-050p | 40 | 73750 | 2283710 | 30.97 |
| pd-prose-050p | 50 | 72205 | 2649606 | 36.70 |
| pd-mixed-100p | 0 | 157568 | 1302177 | 8.26 |
| pd-mixed-100p | 10 | 156942 | 2314623 | 14.75 |
| pd-mixed-100p | 20 | 156422 | 3321459 | 21.23 |
| pd-mixed-100p | 30 | 155646 | 4320622 | 27.76 |
| pd-mixed-100p | 40 | 155144 | 5314574 | 34.26 |
| pd-mixed-100p | 50 | 154505 | 6302548 | 40.79 |

## E6. Latency and memory by document size

15 timed runs after 2 warm-up runs per size, median (p95). One machine, otherwise idle. Memory is the peak resident size of the server process.

| document | n | open | inspect | search | 1 edit | 50-op batch | validate | save | validator share of save | peak RSS |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| pd-prose-001p | 15 | 103 ms (p95 110 ms) | 6 ms (p95 7 ms) | 0 ms (p95 1 ms) | 116 ms (p95 123 ms) | 118 ms (p95 130 ms) | 313 ms (p95 336 ms) | 316 ms (p95 335 ms) | 98.9% | 387 MB |
| pd-prose-010p | 15 | 118 ms (p95 128 ms) | 10 ms (p95 13 ms) | 1 ms (p95 1 ms) | 129 ms (p95 133 ms) | 132 ms (p95 137 ms) | 327 ms (p95 341 ms) | 332 ms (p95 343 ms) | 98.5% | 384 MB |
| pd-prose-050p | 15 | 211 ms (p95 236 ms) | 30 ms (p95 32 ms) | 1 ms (p95 2 ms) | 204 ms (p95 212 ms) | 207 ms (p95 217 ms) | 492 ms (p95 507 ms) | 500 ms (p95 527 ms) | 97.8% | 463 MB |
| pd-prose-100p | 15 | 313 ms (p95 345 ms) | 59 ms (p95 62 ms) | 2 ms (p95 3 ms) | 309 ms (p95 317 ms) | 310 ms (p95 322 ms) | 692 ms (p95 718 ms) | 712 ms (p95 749 ms) | 98.8% | 436 MB |
| pd-prose-200p | 15 | 549 ms (p95 561 ms) | 115 ms (p95 126 ms) | 5 ms (p95 7 ms) | 524 ms (p95 539 ms) | 522 ms (p95 542 ms) | 1.13 s (p95 1.16 s) | 1.12 s (p95 1.15 s) | 100.6% | 659 MB |

"Validator share of save" is the separately timed validate call divided by the save time. It is an estimate: a value near or above 100% means the schema check is essentially the whole cost of a save and the two timings differ only by noise.

Cold first run (new server process): pd-prose-001p open 814 ms, save 489 ms; pd-prose-010p open 795 ms, save 525 ms; pd-prose-050p open 998 ms, save 815 ms; pd-prose-100p open 1.10 s, save 1.09 s; pd-prose-200p open 1.40 s, save 1.60 s.

### E7. Batch scaling (100-page document, replace operations, 10 runs each)

| operations | runs | batch median | batch p95 | ms per operation | save median |
| --- | --- | --- | --- | --- | --- |
| 1 | 10 | 307 ms | 343 ms | 307.1 | 701 ms |
| 10 | 10 | 302 ms | 320 ms | 30.2 | 696 ms |
| 50 | 10 | 306 ms | 314 ms | 6.1 | 681 ms |
| 100 | 10 | 304 ms | 316 ms | 3.0 | 655 ms |
| 250 | 10 | 299 ms | 309 ms | 1.2 | 585 ms |
| 500 | 10 | 295 ms | 304 ms | 0.6 | 483 ms |
