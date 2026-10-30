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
- **Explicit Allowlist**: The MCP server and CLI require one or more root directories declared via `--allowed-dir <path>` or the `ASTRA_EXTRA_DOC_ROOTS` environment variable.
- **Path Canonicalization**: Before any read, write, or export operation, paths are resolved using strict canonicalization (`fs.realpathSync` for existing paths, or verified ancestor walks for newly created files).
- **Symlink Jail Traversal Defenses**: If a symlink points outside the declared allowed directories, ASTra detects the boundary escape and refuses access with an `E_SANDBOX` error.
- **Refusal to Overwrite**: Fresh document creation (`create` / `new_document`) refuses to overwrite existing files unless the `--force` parameter is explicitly passed.

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

---

## 6. Network & Credentials

- **No Network Access**: The engine, CLI and MCP server contain no network client. They never open a connection to send document text, and they need no API keys.
- **No Telemetry**: ASTra contains no analytics, phone-home mechanisms, or telemetry beacons.
- **Your AI Client Is Separate**: When an AI assistant (an MCP client) calls the MCP server, what that assistant sends to its own model provider is governed by that assistant, not by ASTra. Restrict what it can reach with `--allowed-dir`.
