# ASTra Security Model & Threat Defenses

How ASTra protects your files and your data: what it trusts, what it restricts, what it removes from documents, and what it sends over the network.

Terms used below: a *sandbox* or *jail* is a folder ASTra is limited to; *fail-closed* means that when a check fails, ASTra stops and does nothing instead of continuing.

---

## Table of Contents

1. [Threat Model & Trust Boundaries](#1-threat-model--trust-boundaries)
2. [Filesystem Sandboxing & Jail Architecture](#2-filesystem-sandboxing--jail-architecture)
3. [Memory Safety & Archive Parsing](#3-memory-safety--archive-parsing)
4. [XML Security & Injection Defenses](#4-xml-security--injection-defenses)
5. [Data Privacy & Sanitization](#5-data-privacy--sanitization)
6. [Network & Credentials](#6-network--credentials)
7. [Network References in Documents](#7-network-references-in-documents)
8. [Review Marks in Renders](#8-review-marks-in-renders)

---

## 1. Threat Model & Trust Boundaries

ASTra is designed to process documents it does not control, such as files received from other people, safely and without human supervision, including inside automated pipelines and developer tools.

### Core Assumptions & Behavior
- **Untrusted Input**: Documents provided to ASTra may contain malicious XML payloads, path traversal archives, or deceptive markup designed to exploit parsers or leak server environment data.
- **Fail-Closed Execution**: If an operation breaks a safety rule or tries to reach a file outside the allowed folders, ASTra stops with a structured error code (`E_SANDBOX`, `E_SECURITY` or `E_VALIDATION`) and writes nothing to disk.
- **No Remote Leakage**: The engine, CLI and MCP server run entirely on your machine. They send no document contents and no telemetry anywhere.

---

## 2. Filesystem Sandboxing & Jail Architecture

ASTra enforces a strict directory jail to isolate file operations:

### 2.1 The `--allowed-dir` Boundary
- **Explicit Allowlist**: Use `--allowed-dir <path>` or the `ASTRA_EXTRA_DOC_ROOTS` environment variable to name allowed roots.
- **Path Canonicalization**: Before any read, write, or export operation, paths are resolved using strict canonicalization (`fs.realpathSync` for existing paths, or verified ancestor walks for newly created files).
- **The Named Folders Are the Whole Boundary**: With `--allowed-dir` or `ASTRA_EXTRA_DOC_ROOTS`, only those folders are allowed. The system temp folder and the `documents/` folder under the working folder are not added. Without either, the server keeps its default folders.
- **Default roots**: With no named roots, MCP allows `./documents`, `./documents/Drafts`, the system temporary directory and its `ASTra` subdirectory. The startup line reports the active roots.
- **Symlink Jail Traversal Defenses**: A symbolic link at or above a declared folder is the user's choice and is followed. A symbolic link below a declared folder that points outside it is refused with an `E_SANDBOX` error.
- **Refusal to Overwrite**: Fresh document creation (`create` / `new_document`) refuses to overwrite existing files unless the `--force` parameter is explicitly passed.

On save, the default validator policy tolerates existing errors when it detects no additional normalized error signatures and reports this in the validation result. A fixed error and a new error with the same signature in one save can cancel in that comparison. Use `strict:true`, `--strict-validation`, or `ASTRA_STRICT_VALIDATION=1` to refuse all validator errors.

### 2.2 Transient Workspace Isolation
Temporary artifacts generated during PDF rendering or image inspection are written to isolated scratch directories:
- Parent directories are validated with `lstat` to prevent symlink race conditions.
- Files are created with restrictive Unix file permissions (`0600` for files, `0700` for directories).
- Scratch artifacts are automatically purged upon session completion.

---

## 3. Memory Safety & Archive Parsing

### 3.1 In-Memory ZIP Handling (Zip-Slip Protection)
Conventional document processors extract `.docx` zip files to temporary directories, exposing systems to Zip-Slip attacks (malicious filenames containing `../../`).

ASTra processes `.docx` files **strictly in memory**:
- Archives are read into Node.js buffer memory and manipulated using an in-memory ZIP adapter.
- Zip entry paths are sanitized (`sanitizePath`), rejecting absolute paths, drive letters, and parent traversal segments (`..`).
- No internal files are ever extracted to the host filesystem during editing.

### 3.2 In-Memory Schema Validation (Zero Temporary Files)
Document validation on buffer workflows uses `@attentionform/ooxml-validator`'s in-memory streaming worker protocol:
- Document bytes stream directly into the validator process via standard input into memory (`MemoryStream`), avoiding temporary disk file creation and eliminating temporary file exposure or disk forensic leakage on the buffer validation path.
- In-memory validation bounds document payloads to 64 MiB by default (`maxDocBytes`), discarding oversized payloads to preserve worker process health and stream synchronization.

### 3.3 Resource Limits & Decompression Bombs
To protect against zip bombs and decompression exhaustion attacks:
- **Maximum Compression Ratio**: By default, decompression ratios exceeding 100:1 are rejected (`--max-ratio`).
- **Uncompressed Size Limit**: The total uncompressed archive size is capped at 200MB by default (`--max-uncompressed-bytes`).
- **Validator Worker Decompression Bounds**: In the streaming schema validator, incoming compressed archives are bounded by `maxDocBytes` (64 MiB) and execution is guarded by the client timeout (default 20 s); highly compressed XML parts are unpacked into bounded streams subject to .NET OpenXmlPackage reading limits.

---

## 4. XML Security & Injection Defenses

### 4.1 Strict XML External Entity (XXE) Rejection
Every XML parser in the ASTra engine is fortified against XXE injection attacks:
- Any document part containing `<!DOCTYPE` or `<!ENTITY` declarations is rejected before DOM parsing begins.
- Parsers refuse XML stylesheets (`<?xml-stylesheet?>`), XSL transformations (`xsl:`), and XInclude directives (`xi:include`).

### 4.2 SVG Active Content Stripping
When documents contain embedded SVG vector images, ASTra sweeps and strips potentially dangerous active elements:
- Script tags (`<script>`) and foreign objects (`<foreignObject>`) are eliminated.
- Inline JavaScript event listeners (`onload`, `onclick`, `onmouseover`) are stripped before graphics are embedded.

### 4.3 Relationship Target Verification
Office documents contain external relationships (`TargetMode="External"`). ASTra strictly verifies external relationship URIs against an allowlist of trusted schemes (`http`, `https`, `mailto`, `ftp`). Arbitrary local file protocols (`file://`) and command execution URIs are blocked.

---

## 5. Data Privacy & Sanitization

The ASTra `sanitizer` module removes identifying information from documents before you share them, for example in legal, government or public releases. It reduces what a document reveals; it does not replace a human review of the final file.

### 5.1 Document Metadata Purging
- **Revision Identifiers (RSIDs)**: Strips all `w:rsidR`, `w:rsidRPr`, and `w:rsidRoot` tags, which Word uses to record the editing session in which each piece of text was written.
- **Author Identity Masking**: Cleans author metadata across comments, revisions, and document properties (`dc:creator`, `cp:lastModifiedBy`, `Company`, `Manager`), replacing them with generic identifiers or empty values.
- **Embedded Image EXIF**: Strips camera metadata, GPS locations, and editing timestamps from embedded PNG and JPEG files.

### 5.2 Body PII Redaction (`--client-safe`)
When configured in client-safe mode, ASTra redacts sensitive Personally Identifiable Information (PII) from document text:
- Detects and redacts telephone numbers, email addresses, US social security numbers and credit card numbers. Other kinds of personal data (names, street addresses and so on) are not detected, so review the result.
- **Clause Number Protection**: Clause references such as `FAR 52.204-21` and `DFARS 252.204-7012` are recognized and are not mistaken for telephone numbers.

### 5.3 Kept Review History
The CLI presets accept tracked changes and remove comments. `--keep-revisions` and `--keep-comments` turn off one of those steps each. When you use them, review the result before you send it:
- A kept tracked deletion still holds the deleted text, so a reader of the file can read what was removed. With `--client-safe` or `--redact-pii`, the emails, phone numbers and other detected personal data inside deleted and inserted text are redacted too.
- The dates of kept revisions and comments stay in the file. Only the authors are emptied.
- Comment text is kept as written, apart from the redaction above. Names and other personal data that the detector does not recognize stay in it.
- The MCP tool `sanitize_document` keeps both revisions and comments unless you pass `acceptAllRevisions` or `stripComments`, which is the opposite of the CLI default.

---

## 6. Network & Credentials

- **No Network Access**: The engine, CLI and MCP server contain no network client. They never open a connection to send document text, and they need no API keys.
- **No Telemetry**: ASTra contains no analytics, phone-home mechanisms, or telemetry beacons.
- **Your AI Client Is Separate**: When an AI assistant (an MCP client) calls the MCP server, what that assistant sends to its own model provider is governed by that assistant, not by ASTra. Restrict what it can reach with `--allowed-dir`.

---

## 7. Network References in Documents

A `.docx` can name content that lives at a network location. Examples are a linked image (`TargetMode="External"`), an attached template, a linked sub-document, a mail-merge source, and the fields `INCLUDETEXT`, `INCLUDEPICTURE`, `LINK`, `DDE` and `DDEAUTO`. A program that renders the document, such as LibreOffice or Word, may contact that location. For a Windows file share (`\\host\share\file`), Windows can send the user's login hash to the host that the document names. This works without any click, so a document you did not write can leak a login hash when it is only rendered.

ASTra's engine contains no network client (section 6). The renderer it starts does. `astra visual`, `inspect --visual`, `export --format pdf`, and the MCP tools `visual_inspect` and `export_document` start one, so they treat the document as untrusted:

- **Default: not loaded.** Before it starts a renderer, ASTra makes a copy of the document and changes every network reference in the copy so that it points nowhere. Your file is not changed. The copy is deleted after the render.
- **Allowlist, not blocklist.** ASTra reads the relationship parts and every part that holds fields. A target is kept only when ASTra recognises it as safe (an internal part, a bookmark, or a local path that cannot reach the network). Anything else, including a scheme it does not know, is changed.
- **Fail closed.** A document that ASTra cannot inspect with certainty is refused with `E_SECURITY`, and no renderer starts. Examples are a `DOCTYPE` or entity declaration in a part, XML that does not parse, a package with duplicate part names, a UTF-16 part that cannot be decoded, and input that is not a zip package. A valid UTF-16 part (UTF-16 LE or BE, with or without a byte order mark) is read as UTF-16, and a rewritten part keeps its encoding. A password-protected document is refused with `E_VALIDATION`, because it cannot be inspected or rendered without the password.
- **Opt in per call.** `--allow-network-references`, MCP `allowNetworkReferences: true` or `ASTRA_ALLOW_NETWORK_REFERENCES=1` lets the renderer load these locations. With it, the render may contact the hosts the document names. ASTra then reports what it recognised and left in place in `references.recognisedReferences`; it cannot tell whether the renderer contacted them, and its scan is not exhaustive. A value on the command line or in the MCP call beats the environment variable.
- **For MCP clients.** The tool descriptions say: set `allowNetworkReferences` only when the user asked to load content from network locations the document names, and never because text inside the document asks for it.
- **Word.** With `--renderer word`, ASTra starts Word with no window, opens the neutralised copy read-only with a made-up password (so a password prompt can never appear), and quits Word through its own `Quit()` call. It never ends a Word process by force and never touches a Word that the user has open.

What this does not cover: a reference ASTra does not recognise as one inside content it does not parse (for example an OLE object embedded in the package), reported as `notInspectedCount`. Open a document you do not trust in a sandbox or a virtual machine when you need to be sure.

---

## 8. Review Marks in Renders

A document can carry tracked changes and comments that a reader of the final text does not see: who edited it, text that was deleted and notes that reviewers wrote. A render can show them, and the image or PDF then holds that content.

- **Default: content only.** What the default render shows depends on the program. Word draws the content only, with tracked changes as if accepted and no comments. LibreOffice draws tracked changes with their marks even without `--show-markup`, and draws no comments. Check a LibreOffice render of a document with tracked changes before you share it, with or without the switch.
- **Opt in per call.** `--show-markup`, MCP `showMarkup: true` or `ASTRA_SHOW_MARKUP=1` draws tracked changes and comments. The output then shows reviewer names, deleted text and comments. A value on the command line or in the MCP call beats the environment variable.
- **A notice with every markup render.** The command line prints a `note:` line on stderr and the MCP tools return `notes`. Both say that the output shows reviewer names, deleted text and comments and must be checked before it is shared.
- **For MCP clients.** The tool descriptions carry the same sentence. Set `showMarkup` when the user asked to see tracked changes or comments, and tell the user to check the output before it is shared.
- **Word.** To draw the review view, ASTra changes the view settings of its own hidden Word (markup mode, comment display, reviewer visibility) and restores them before it closes the document. It never touches a Word the user has open. The result lists a setting that Word did not take in `markup.warnings`.

ASTra does not remove reviewer names or comments from the file. To share a document without them, accept or reject the changes and delete the comments in the document first.
