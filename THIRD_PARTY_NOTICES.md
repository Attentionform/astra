# Third-Party Software Notices and Information

ASTra includes third-party software. The notices below are provided as required by their licenses.

---

## 1. xarsh/ooxml-validator

- **Component:** OOXML validator wrapper (runner, binary resolution, one-shot native program)
- **Upstream Project:** ooxml-validator
- **Version:** 0.4.0
- **Upstream Commit:** `11e3e8c5430d054c8d9102e3e5f5e15c6697c7b2`
- **Source URL:** https://github.com/xarsh/ooxml-validator
- **Copyright:** Copyright (c) 2025 Takao Naito
- **License:** MIT
- **Where it appears in ASTra:**
  - `packages/ooxml-validator/src/**`
  - `packages/ooxml-validator/native/**`
  - `@attentionform/ooxml-validator-<platform>` packages
- **Modifications:** Modified by Attention Form for use in ASTra.

---

## 2. Open XML SDK (DocumentFormat.OpenXml)

- **Component:** OOXML schema validation
- **Source URL:** https://github.com/dotnet/Open-XML-SDK
- **Version:** 3.5.1
- **Copyright:** Copyright (c) .NET Foundation and Contributors
- **License:** MIT
- **Where it appears:** linked into each platform binary.

---

## 3. .NET runtime

- **Component:** self-contained runtime bundled in each platform binary
- **Source URL:** https://github.com/dotnet/runtime
- **Copyright:** Copyright (c) .NET Foundation and Contributors
- **License:** MIT
- **Where it appears:** bundled in each platform binary. The runtime's licence (`DOTNET-LICENSE.TXT`) and its own third-party notices (`THIRD-PARTY-NOTICES.TXT`) ship in every platform package, next to the binary.

---

# License Texts

## MIT License (xarsh/ooxml-validator)

MIT License

Copyright (c) 2025 Takao Naito

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

---

## MIT License (Open XML SDK and .NET runtime)

The MIT License (MIT)

Copyright (c) .NET Foundation and Contributors

All rights reserved.

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

