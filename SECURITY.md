# Security Policy

ASTra reads and writes Word documents for people and for AI agents, so a flaw can mean a file written outside its folder, or a document that opens wrongly. We take reports seriously.

## Reporting a vulnerability

**Please do not open a public issue for a security problem.**

Use one of these:

1. **GitHub private report (preferred):** go to the repository's **Security** tab and choose **Report a vulnerability**.
2. **Email:** astra@attentionform.com
3. **Email, second address:** attentionform@protonmail.com reaches the same person. Mail from a Proton Mail account to it is encrypted end to end.

Please include:

- what you found and what it lets an attacker do,
- the smallest document, command or tool call that shows it,
- the package and version (`astra --version`, or the version in `package.json`), your operating system and Node.js version.

You do not need a working exploit. A clear description is enough.

## What to expect

These are targets, not guarantees; ASTra is early-access software maintained by a small team.

- We aim to acknowledge a report within **5 working days**.
- We aim to tell you whether we can reproduce it within **14 days**.
- We will keep you informed, agree a disclosure date with you, and credit you in the release notes if you wish.
- We ask for **90 days** from your report before public disclosure, or less once a fix is released.

## Supported versions

Only the latest published release of each package receives security fixes while ASTra is in the 0.x series.

| Package | Supported |
| --- | --- |
| `@attentionform/astra-engine`, `@attentionform/astra-mcp-server`, `@attentionform/astra`, `@attentionform/ooxml-validator` | latest release |

## What is in scope

- Escaping the `--allowed-dir` folder (path traversal, symlinks, archive entries such as zip-slip).
- Reading or writing files, or reaching the network, in a way the documented behavior forbids.
- Malicious documents: XML external entities, decompression bombs, macros or active content that survive where they should be refused or removed.
- Getting ASTra to write a package that fails validation, or to skip the save-time validator.

## What is out of scope

- Documents you open that are simply wrong or ugly but valid.
- Problems that need an attacker who already controls your machine, your npm account or your model provider account.
- Denial of service by an input far larger than the documented limits.

## Safe harbor

If you act in good faith, stay within what is needed to show the problem, do not read or change other people's data, and give us a reasonable time to fix it, we will not take legal action against you for your research.

## Privacy note for reporters

Do not send us documents that contain real personal or confidential information. Make a small test file instead.
