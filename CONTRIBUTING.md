# Contributing to ASTra

Thank you for helping. ASTra edits Word documents for people and for AI agents, and it makes one promise: it never writes a broken file. Most of the rules below exist to protect that promise.

The complete developer guide, with the toolchain, architecture and test details, is [`docs/contributing.md`](docs/contributing.md). This page is the short version and contains the rules we enforce.

## Before you start

- Read the [Code of Conduct](CODE_OF_CONDUCT.md).
- Search the existing issues first.
- To report a bug, open an issue with the exact error code (for example `E_VALIDATION`), your Node.js version and operating system, and the smallest `.docx` file or operation batch that shows the problem. Do not attach documents that contain real personal or confidential data.
- To report a security problem, do **not** open an issue. Follow [SECURITY.md](SECURITY.md).
- For a larger change, open an issue or a discussion first, so we can agree on the approach before you spend time on it.

## Set up

You need Node.js 22.19 or newer and npm 11.

```bash
git clone https://github.com/Attentionform/astra.git
cd astra
npm install
npm run build:packages
npx vitest run packages/cli/tests/cli-smoke.test.ts   # a quick check, a few seconds
```

LibreOffice (`soffice`) and Poppler (`pdftoppm`) are optional. They are needed only for PDF export and page images. Without them those calls fail with `E_DEPENDENCY` and write nothing.

## Rules for code

1. **Never write a `.docx` that has not passed validation.** If your change can produce invalid OOXML, it must fail with `E_VALIDATION`, explain why, and leave the document as it was.
2. **A batch is all or nothing.** If one operation fails, the document returns to its state before the batch.
3. **Every public input has a schema.** Command-line arguments, MCP tool inputs and operation payloads are described with zod schemas. Do not add an input that bypasses them.
4. **Use the namespace constants** from the engine (`src/ooxml/ns.ts`) instead of typing namespace strings.
5. **Ship tests with the change.** A bug fix includes a test that fails without the fix. Name tests and test files after the behavior they check, not after a ticket or plan number.
6. **Ship documentation with the change.** If you add or change a flag, a tool parameter or an exported function, update the documentation in the same pull request. The `docs-truth` test checks counts and names against the code.

## Before you open a pull request

Run these. All of them should finish without errors.

```bash
npx tsc --noEmit
npm run lint
npm run build:packages
npm test
npm run validate:docs
npm run lint:docs
```

Copy your own test counts into the pull request description. Do not reuse old numbers.

## Commits

**Branches:** `feat/<name>`, `fix/<name>` or `docs/<name>`.

**Messages** follow [Conventional Commits](https://www.conventionalcommits.org/), for example `fix(cli): correct exit code on invalid json input`.

**Stage only what belongs to the change.** Never commit `.env` files, keys, personal notes or build output.

### Rule 1: use your real email address

Every commit must be authored with a real email address that you can receive mail at.

- GitHub's hidden addresses (`...@users.noreply.github.com`) are **not accepted**, and neither are placeholder addresses such as `you@example.com`.
- In GitHub, turn off **Settings > Emails > Keep my email addresses private** and **Block command line pushes that expose my email**, then set the same address in git:

  ```bash
  git config user.name  "Your Name"
  git config user.email "you@your-domain.com"
  ```

- Please understand that this puts your email address in the public history of the project. If you are not comfortable with that, please do not contribute.

### Rule 2: sign off every commit

Add a sign-off line with the same email address:

```bash
git commit -s -m "fix(engine): ..."
```

This adds `Signed-off-by: Your Name <you@your-domain.com>`. By signing off you confirm the [Developer Certificate of Origin](https://developercertificate.org/): you wrote the change, or you have the right to submit it under the project license. This sign-off is separate from the contributor license agreement, which you sign once (see rule 4).

An automatic check on every pull request enforces rules 1 and 2. A second check, `cla`, enforces rule 4. If it fails, fix the author with `git commit --amend --reset-author -s` (last commit) or `git rebase --exec` for several, and push again.

### Rule 3: you are responsible for what you commit, including AI-assisted work

You may use AI tools as much as you like. These conditions always apply:

- The commit must come from you, a person, under your own name and email address. Commits made under a bot or an AI tool's own identity are not accepted.
- You are solely responsible for everything you commit. Code, tests or text produced with the help of AI is treated as your own work. Read it, understand it, run it, and make sure you have the right to contribute it.
- "The AI wrote it" is not a reason a change is acceptable, and not a reason a defect is excusable.
- A `Co-Authored-By` line naming an AI tool is welcome, but it does not replace the human author.

### Rule 4: sign the contributor license agreement once

ASTra is source-available under the [Business Source License 1.1](LICENSE), and the owner also offers commercial licenses. To accept outside code, the owner needs your permission to use your contribution under any license terms, including commercial ones. You keep the copyright in your contribution.

1. Read the [individual CLA](CLA.md). If you contribute for your employer, ask your employer to sign the [corporate CLA](CLA-CORPORATE.md) instead (write to astra@attentionform.com).
2. Create your signature file. From a clone of the repository, run one command (it needs `npm install` once, and your GitHub login and full name):

   ```bash
   npm run cla:sign -- <your-github-login> "Your Full Name"
   ```

   It looks up your numeric GitHub id, writes `cla/signatures/<your-numeric-github-id>.json` with the current fingerprint and statement filled in, and checks the file with the same rules as the `cla` check before it saves it. It never overwrites an existing file. Add `--print` to see the file without writing it. Add the file to your first pull request, in a commit you wrote yourself.

   You do not have to use the command. If you open your pull request without the file, the `cla` check posts the exact file for you, with your id and login already filled in; you only type your full name. The file looks like this:

   ```json
   {
     "github_id": 1234567,
     "github_login": "janedoe",
     "name": "Jane Doe",
     "cla_version": 1,
     "cla_sha256": "<fingerprint of the CLA text, filled in by the check>",
     "statement": "I have read the ASTra CLA v1 and I agree to it."
   }
   ```

   Writing it by hand? Your id is the number at `https://api.github.com/users/<your-login>`.
3. Push. The check runs again by itself and turns green. You sign once, for all future pull requests.

The check is strict on purpose. The file must be named after your id, hold exactly these six keys, carry the exact statement and the fingerprint of the current CLA text, and be added in a commit written by your own account. A pull request may add only its own authors' files, may not change anyone else's signature, and may not touch the CLA files or `cla/corporate/`. If the check fails it says which rule was broken. It never runs code from your pull request and it cannot write to the repository.

Your signature is public: the repository then shows your GitHub id and login, the name you typed, the CLA version and the fingerprint of the text you accepted, and Git records the date. Nothing else is collected. The file is kept to prove that you agreed, on the basis of the agreement itself and the project's legitimate interest in proving its right to license your work, and so it cannot be erased while ASTra is licensed with your contribution in it; write to astra@attentionform.com with any question. If the CLA text changes, the version number goes up and you are asked to raise `cla_version` in your own file in a later pull request.

Dependabot and the repository owner do not need to sign.

## License and ownership

ASTra is licensed under the [Business Source License 1.1](LICENSE), a source-available license that is not an open-source license. [LICENSING.md](LICENSING.md) explains in plain words who may use it for free. Each version becomes Apache-2.0 four years after it is first made public. You keep the copyright in your contribution and give the owner the license described in the [CLA](CLA.md).

The names and logos "ASTra" and "Attentionform" are not covered by the license; see [TRADEMARKS.md](TRADEMARKS.md).
