# Contributing to ASTra

Thank you for your interest in contributing to ASTra! ASTra is a source-available OOXML engine and document-editing toolkit (Business Source License 1.1; see [LICENSING.md](../LICENSING.md)), and we welcome contributions from developers, technical writers, and document engineers of all skill levels.

This guide outlines our development workflow, coding standards, testing procedures, and the process for submitting issues and Pull Requests.

---

## Table of Contents

- [Code of Conduct](#code-of-conduct)
- [Prerequisites & Toolchain](#prerequisites--toolchain)
- [Local Development Setup](#local-development-setup)
- [Monorepo Architecture](#monorepo-architecture)
- [Testing & Quality Gates](#testing--quality-gates)
- [Coding Standards & Invariants](#coding-standards--invariants)
- [Git Workflow & Branching](#git-workflow--branching)
- [Submitting Issues & Pull Requests](#submitting-issues--pull-requests)
- [Releasing](#releasing)

---

## Code of Conduct

We are committed to a welcoming, harassment-free environment for everyone. The full text is in [CODE_OF_CONDUCT.md](../CODE_OF_CONDUCT.md). The short rules for commits, sign-off and the use of AI tools are in [CONTRIBUTING.md](../CONTRIBUTING.md); they are enforced by an automatic check on every pull request.

---

## Prerequisites & Toolchain

Before contributing, ensure your development environment meets the following requirements:

- **Node.js**: `>= 22.19.0` (matches the `engines` field in `package.json`; CI pins Node 22).
- **npm**: `npm@^11` (CI runs `npm install -g npm@^11`; the lockfile is v3; run that upgrade command first if your npm is older).
- **Git**: any maintained release.
- **Operating System**: Linux, macOS, or Windows (WSL2 recommended for Windows).
- **Optional: LibreOffice** (`soffice` in `PATH`): needed for PDF export (`export_document` with `format: 'pdf'`) and image renders. Without it those calls fail with `E_DEPENDENCY` and nothing is written.
  - Ubuntu / Debian / WSL: `sudo apt-get update && sudo apt-get install -y libreoffice-writer libreoffice-calc` (same packages CI installs in `.github/workflows/ci.yml`).
  - macOS: `brew install --cask libreoffice`
  - Check: `which soffice` should print a path.
- **Optional: Poppler** (`pdftoppm` in `PATH`): needed for PDF-to-image rasterization in visual renders. Without it image renders fail with `E_DEPENDENCY` while text-only telemetry still works.
  - Ubuntu / Debian / WSL: `sudo apt-get update && sudo apt-get install -y poppler-utils` (same package CI installs).
  - macOS: `brew install poppler`
  - Check: `which pdftoppm` should print a path.
  - Combined check: `which soffice pdftoppm`
- **.NET 10 SDK** (required for building from source): The OOXML schema validator uses an in-house native binary (`@attentionform/ooxml-validator`). When working from a source checkout, install the .NET 10 SDK (`10.0.x`) and compile the host binary using `npm run build:native --workspace=@attentionform/ooxml-validator`. When installed as a package dependency from npm, platform-specific binaries are downloaded automatically as optional dependencies.

---

## Local Development Setup

1. **Clone the repository:**
   - Maintainers and direct contributors (SSH):
     ```bash
     git clone git@github.com:attentionform/astra.git
     cd astra
     ```
     SSH needs a key on your GitHub account (GitHub Settings → SSH keys; see the GitHub docs page "Connecting to GitHub with SSH" if you see `Permission denied (publickey)`).
   - HTTPS alternative (no SSH key needed):
     ```bash
     git clone https://github.com/Attentionform/astra.git
     cd astra
     ```
   - Outside contributors: fork `attentionform/astra` on GitHub first, then clone your fork and open pull requests from a branch on the fork. Clone the main repo directly only when you have write access.

2. **Clean install dependencies (mandatory `npm ci`):**
   ```bash
   npm ci
   ```
   *Note: Always use `npm ci` rather than `npm install` to ensure lockfile fidelity.*

3. **Optional: API keys for chat**: copy the template and fill only the providers you use:
   ```bash
   cp .env.example .env
   ```
   `.env.example` lists OpenAI, Anthropic, Groq, Google, Azure, OpenRouter/OpenCode, plus custom OpenAI-compatible endpoints; Ollama needs no key, only a local server on port 11434. Never commit `.env`.

4. **Build monorepo workspace packages:**
   ```bash
   npm run build:packages
   ```
   *Note: Packages depend on `@attentionform/astra-engine` compilation output. Run this build step before running tests.*

5. **Verify typecheck:**
   ```bash
   npx tsc --noEmit
   ```

6. **Run the full test suite:**
   ```bash
   npm test
   ```
   New here? Run one fast file first (see [Testing & Quality Gates](#testing--quality-gates)) before the full run.

---

## Monorepo Architecture

ASTra is structured as an npm workspaces monorepo containing four core packages:

```
packages/
├── engine/       # @attentionform/astra-engine: Core OOXML parser, AST, 27 operations, TOON, validator gate
├── cli/          # @attentionform/astra: Pure POSIX command-line interface (astra bin)
├── mcp-server/   # @attentionform/astra-mcp-server: Model Context Protocol 19-tool stdio server
└── ooxml-validator/  # @attentionform/ooxml-validator: .NET Open XML schema validator and its platform binaries
```

---

## Testing & Quality Gates

Producing valid documents is ASTra's core promise, so every change must pass the test suite before it is merged.

### Running Tests
- **First run? Start with one fast file** (~8 seconds, 47 checks, verified on this checkout):
  ```bash
  npx vitest run packages/cli/tests/cli-smoke.test.ts
  ```
  Success looks like:
  ```text
  Test Files  1 passed (1)
        Tests  47 passed (47)
  Duration  ~8s
  ```
  Use this for a quick single-file check. Only then run the full suite.
- **Run all unit & integration tests** (several minutes; hundreds of files, thousands of checks; budget time and record your own counts):
  ```bash
  npm test
  ```
  Success looks like `Test Files  N passed (N)` plus `Tests  N passed (N)` with zero failures; copy the exact numbers from your run into the PR, never reuse old counts.
- **Run tests in watch mode during development:**
  ```bash
  npx vitest packages/cli/tests/cli-smoke.test.ts
  ```

### Document Gallery Schema Gate
ASTra enforces an automated ECMA-376 XML schema validation gate across our enterprise template gallery:
```bash
npm run validate:docs
```
This script tests all templates in `test-gallery/*.docx` through the validator gate (`validateDocxFile`). All 9 gallery templates must pass schema validation without errors.

### Docs truth rule
Docs ship with the code that proves them. When you change a CLI command, a tool, a flag, or a count, update the matching README, usage text, and registry entry in the same pull request. Run `npm run lint:docs` plus the docs-truth golden (`npx vitest run tests/unit/docs-truth.test.ts`) before you ask for a second look. Copy-paste every new example into a clean terminal and confirm it works verbatim. Numbers always come from a live run or a cited source line, never from memory. If a new reader cannot reach first success with only the docs, that confusion is a defect; fix the docs, not the reader.

### User-facing docs gate
`npm run lint:docs` (`npx tsx scripts/lint-docs.ts`) scans `README.md`, `docs/*.md`, `packages/*/README.md`, registry JSON, and integration README plus JSON for unfinished markers, the legacy env literal, internal terms, and stale labels. Prose only is checked for internal terms (fenced blocks and inline code are skipped); markers and the legacy literal are checked everywhere. Prints `file:line` on hit and exits non-zero. Grandfathered hits live in `scripts/lint-docs-allowlist.json` with a reason each. There is no commit hook in this repo, so run the script by hand and keep it green before you open a pull request.

## Coding Standards & Invariants

To maintain enterprise document fidelity, all code contributions must adhere to the following invariants:

1. **Fail-Closed Validation**:
   - Never write a `.docx` file to disk without passing the ECMA-376 schema validation gate (`validateDocxBuffer` / `validateDocxFile`).
   - If an operation creates invalid OOXML structure, it must throw an `E_VALIDATION` error with a clear, actionable remedy and leave the document uncorrupted.
2. **Strict TypeScript & Schema Discipline**:
   - Monorepo runs under `strict: true`. Avoid ungrounded `any` casts.
   - Use Zod schemas for all public CLI arguments, MCP tool inputs, and operation payloads.
3. **Atomic Operations**:
   - Operations executed via `OperationEngine.executeBatch` must maintain snapshot rollback. If one operation fails in a multi-operation batch, the document state must rewind cleanly to before the batch.
   - Shape example: good (id comes from `inspect_document`, single edit = 1-item batch):
     ```json
     {"action": "replace", "target": "3F2A9C1B", "text": "Hello"}
     ```
   - Shape example: good insert (anchor + position required together):
     ```json
     {"action": "insert", "type": "paragraph", "anchor": "3F2A9C1B", "position": "after", "text": "Hello"}
     ```
   - Shape example: bad (rejected with `E_VALIDATION`, nothing written):
     ```json
     {"action": "replace", "text": "Hello"}
     ```
     Missing `target`, and any invented id that never came from `inspect_document` is rejected the same way.
4. **OOXML Namespaces**:
   - Always reference XML elements and attributes using the canonical namespace constants from `@attentionform/astra-engine` (`src/ooxml/ns.ts`), never raw hardcoded namespace strings.
5. **Linting**:
   - Lint (`npm run lint`, same as `npx eslint .`) must report no errors before you open a pull request. Warnings are tolerated, errors are not.

---

## Git Workflow & Branching

We follow standard GitHub Flow:

1. **Branch Naming**:
   - Feature branches: `feat/<descriptive-name>` (e.g., `feat/table-border-styling`)
   - Bug fixes: `fix/<issue-description>` (e.g., `fix/vmerge-restart-offset`)
   - Documentation: `docs/<topic>` (e.g., `docs/mcp-setup-guide`)
2. **Commit Messages**:
   - We follow the [Conventional Commits](https://www.conventionalcommits.org/) specification:
     - `feat(engine): add support for cell vertical alignment`
     - `fix(cli): correct exit code on invalid json input`
     - `docs(mcp): document smart-windowing parameters`
3. **Clean Commits**:
   - Only stage files relevant to your change.
   - Never commit `.env` files, local scratch notes, or temporary build outputs.

---

## Submitting Issues & Pull Requests

### Reporting Issues
- Check existing GitHub Issues before opening a new one.
- New contributors: look for open Issues marked `good first issue` for small scoped tasks.
- Stuck? Ask on the repo's GitHub Issues and include Node.js version, operating system, and the exact error code (`E_VALIDATION`, `E_IO`, `E_SANDBOX`).
- Provide a minimal reproducible example (including a sample `.docx` or operation batch JSON where applicable).

### Pre-PR Checklist
Before opening a Pull Request, verify that all checkboxes pass locally:

- [ ] `npx tsc --noEmit` runs with 0 errors.
- [ ] `npm test` passes (all unit and integration suites green).
- [ ] `npm run validate:docs` exits with code 0.
- [ ] `npm run lint:docs` exits with code 0 (user-facing docs carry no banned strings).
- [ ] `npm run build:packages` builds all workspace packages cleanly.
- [ ] New features include corresponding automated tests in `tests/` or `packages/*/tests/`.
- [ ] Documentation is updated to reflect any new CLI flags, tool parameters, or exported functions.

### Contributor license agreement

Sign the one-time [CLA](../CLA.md) by adding a signature file to your first pull request. One command writes it: `npm run cla:sign -- <your-github-login> "Your Full Name"`; the `cla` check also shows you the exact file if you forget (see [CONTRIBUTING.md](../CONTRIBUTING.md), rule 4). Company contributors use the [corporate CLA](../CLA-CORPORATE.md).

Once opened, our GitHub Actions CI pipeline will run typechecking, package build, tests, and document validation gates. Maintainers will review your PR and provide constructive feedback!

---

## Releasing

Maintainers publish packages with the manual **Release** workflow: a dry run first, then a real run that waits for approval. Nothing is published by a push or a tag. See [Releasing](releasing.md) for the choices, the checks a run makes, how the validator binaries are versioned, and the one-time setup.
