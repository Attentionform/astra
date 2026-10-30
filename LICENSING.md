# Licensing

ASTra is licensed under the [Business Source License 1.1](LICENSE) (BSL). This page explains it in plain words. If this page and the [LICENSE](LICENSE) disagree, the LICENSE wins.

## In one minute

- You can read, run, modify and redistribute the source code.
- **Free production use** if you are: an individual (including a freelancer or consultant), a school or university (or doing academic research), a non-profit using it for its mission, a public body doing governmental work, or a small business (under USD 1,000,000 revenue and fewer than 10 people).
- **Everyone else needs a commercial license**: write to astra@attentionform.com.
- **Trying it out is always free**, for anyone: evaluation, development, testing and demos.
- **Your documents are yours.** The license puts no restriction on documents you create with ASTra.
- **Each version becomes Apache-2.0** four years after it is first made public.

ASTra is **source-available**. It is not "open source" in the sense of the Open Source Definition, because the free use is limited.

## Who can use it for free in production

| You are | Free? | Notes |
| --- | --- | --- |
| A freelancer or consultant, alone or through a company with no employees | Yes | Preparing documents for clients is covered. Running ASTra on a client's systems for the client's own continuing use is the client's use, so the client needs its own license (or must qualify itself). |
| A hobbyist or student | Yes | |
| A school, college or university, its students and staff | Yes | Teaching, learning and non-commercial academic research. |
| A registered charity or tax-exempt non-profit | Yes | For its mission. Not for selling products or services. |
| A government department, agency or municipality | Yes | For governmental functions. A state-owned company, or a body that sells products or services commercially, is treated as a company. |
| A company or other business under USD 1,000,000 revenue **and** fewer than 10 people (together with affiliates, last completed fiscal year; a new company that has not finished one uses what it reasonably expects for its first) | Yes | If you grow past either limit, you may continue until the end of your next fiscal year, then you need a license. |
| A company at or above those limits | No | Commercial license. |

## Common questions

**I am a freelancer. Can I use ASTra for client work?** Yes. Using it to produce documents for clients is covered. You pay nothing, and you do not need to tell anyone.

**My startup is small. Do I need to do anything?** No. If you are under both limits you may use ASTra in production and in your own product. Keep an eye on the limits.

**Can I build ASTra into my own product or service?** Yes, if you are an individual or a small business and your product has substantial functionality of its own. Your customers use ASTra only as part of your product. A larger company needs a commercial license for this.

**Can I make videos, stream, review or teach with ASTra?** Yes. Recording a tutorial, running a live stream, writing a review or teaching a course is fine, and it does not matter whether you earn money from the video or the course. Individuals are free whatever they earn, and demonstrating the software is non-production use, which is free for everyone. Hosting ASTra as a service for your audience, or reselling it, is a competing offering and always needs a commercial license.

**Can I build an open-source project on ASTra?** Yes, if your project has substantial functionality of its own. Your own code can be under any license, such as MIT or Apache-2.0. You list ASTra as a dependency and you do not relicense it. People who use your project use ASTra only as part of it, which is covered when you are an individual or a small business, even if your project is free of charge. A company at or above the small-business limits needs a commercial license for its own production use of ASTra, and the people who use a project published by such a company are not covered by the rule above and need their own right to use ASTra. A thin wrapper whose value is mainly ASTra itself is a competing offering and needs a commercial license. If you copy ASTra source into your repository, keep ASTra's LICENSE with it: ASTra code cannot be relicensed as MIT or Apache-2.0 until its own Change Date. Some license scanners and software distributions do not treat the BSL as open source, so describe your project as "depends on ASTra (source-available)".

**What is not allowed without a license?** A "competing offering": a product or service whose value consists mainly of ASTra itself, for example hosting ASTra as a service for others or reselling it. This always needs a commercial license.

**Do the limits include my parent company and affiliates?** Yes, they are counted together.

**What if I stop qualifying?** You can continue your existing use for 90 days (for small businesses, until the end of your next fiscal year), then you need a commercial license or you stop.

**Do I have to publish my own code?** No. The BSL has no copyleft and no obligation to share your code.

**Is the license going to get stricter?** The terms of a released version never become stricter. A later version may carry different terms.

**Can I send a pull request?** Yes. Sign the [contributor license agreement](CLA.md) once, by adding one small signature file to your first pull request. See [CONTRIBUTING.md](CONTRIBUTING.md).

**What about the names and logos?** They are covered separately: see [TRADEMARKS.md](TRADEMARKS.md).

**What about third-party code inside ASTra?** It keeps its own license. See [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md). A folder that has its own license file is under that license.

## Commercial licenses

Write to astra@attentionform.com with a short description of your use. We offer licenses for companies that embed ASTra in a product, run it internally at scale, or need support.

## For maintainers

- The `LICENSE` file describes the Licensed Work as every version the Licensor makes publicly available, so it needs no edit per release. Each version's Change Date is four years after the day it was first made public: write that day next to the version in `CHANGELOG.md` when you release.
- The `LICENSE` in the root and in each package must stay identical. `tests/unit/license-consistency.test.ts` checks this.
- Use the SPDX id `BUSL-1.1` in `package.json`.
- Never describe ASTra as "open source". Say "source-available".
