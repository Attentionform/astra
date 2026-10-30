# Third-Party Notice: Native OOXML Validator (.NET OpenXmlValidator)

- **Upstream Package:** `ooxml-validator` (.NET subsystem)
- **Upstream Repository:** https://github.com/xarsh/ooxml-validator
- **Version:** 0.4.0
- **Upstream Commit:** `11e3e8c5430d054c8d9102e3e5f5e15c6697c7b2`
- **Vendoring Date:** 2026-10-27
- **License:** MIT (see `./LICENSE-MIT-xarsh` and root `/THIRD_PARTY_NOTICES.md`)
- **Copyright:** Copyright (c) 2025 Takao Naito

## Provenance and Adaptations

The C# project files and validation logic in this directory are adapted from `xarsh/ooxml-validator` version 0.4.0 (commit `11e3e8c5430d054c8d9102e3e5f5e15c6697c7b2`).

### Adaptations and Enhancements:
1. Target .NET 10.0 (`net10.0`) with locked package dependencies (`packages.lock.json`).
2. Configured invariant globalization (`InvariantGlobalization=true`) for smaller self-contained builds without ICU dependencies.
3. Added persistent streaming worker mode (`--worker`) with stdin/stdout newline-delimited framing and raw memory stream validation.
4. Added structured error attribution (reporting target node names, namespaces, and error classification).
5. Retained 100% byte-for-byte JSON output compatibility in one-shot mode (`ooxml-validator <file> [OfficeVersion]`).

For complete license terms, see `LICENSE-MIT-xarsh` in this directory and the root repository `THIRD_PARTY_NOTICES.md`.
