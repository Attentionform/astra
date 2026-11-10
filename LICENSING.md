# Licensing

ASTra is licensed under the [Business Source License 1.1](LICENSE) (BSL). This page explains it in plain words. If this page and the [LICENSE](LICENSE) disagree, the LICENSE wins.

This page describes version 0.4.0 and later. Every version released before 0.4.0 keeps the terms it was released with.

## In one minute

- You can read, run, modify and redistribute the source code.
- **Free production use** if you are: a person who works alone (a freelancer, consultant or sole proprietor, in any legal form, at any revenue) for their own work and their client deliverables, a school or university (or doing academic research), a non-profit using it for its mission, a public body doing governmental work, a small organization (up to 10 people and under USD 250,000 revenue) for its own work, or the maker of a product or service with under USD 250,000 revenue that has ASTra inside it.
- **Internal use** means your own people, systems and agents using any part of ASTra (the MCP server, the CLI, the engine as a library, the validator) for your own operations, including scripts, CI/CD, batch jobs and internal pipelines. Above the small-organization limits it needs a commercial license, priced by the number of people.
- **Embedding or SaaS** means ASTra is part of a product or service of yours that has substantial features of its own. It is free until that product's revenue reaches USD 250,000, whatever the size of your company. After that it needs an Embedded license, priced by the product's revenue.
- **Everyone else needs a commercial license**: see the [pricing page](https://attentionform.com/pricing), or write to astra@attentionform.com.
- **Trying it out is always free**, for anyone: evaluation, development, testing and demos.
- **Your documents are yours.** The license puts no restriction on documents you create with ASTra.
- **Each version becomes Apache-2.0** four years after it is first made public.

ASTra is **source-available**. It is not "open source" in the sense of the Open Source Definition, because the free use is limited.

## Which question applies to you

1. **Do you work alone?** Free, at any revenue, for your own work and your client deliverables. This covers freelancers, consultants, sole proprietors and one-person companies. It does not cover a product or service of yours that has ASTra inside it and that you offer to others (question 2 applies, with its USD 250,000 revenue limit, even if you work alone), and it never covers a competing offering (see below).
2. **Is ASTra part of a product or service that you ship or host for others, and that has substantial features of its own?** Free until that product's revenue reaches USD 250,000. After that, an Embedded license, priced by the product's revenue.
3. **Otherwise, do your own people use ASTra for your own work?** Free up to 10 people and under USD 250,000 revenue. Above either limit, an internal license, priced by the number of people.

Schools, non-profits and public bodies are free for their own mission. Hosting ASTra for others, a thin API or MCP wrapper, and reselling it are competing offerings and always need their own license.

## Who can use it for free in production

| You are | Free? | Notes |
| --- | --- | --- |
| A freelancer or consultant who works alone, personally or through your own one-person company | Yes | At any revenue. Preparing documents for clients, one job at a time, is covered. A product or service of your own with ASTra inside, offered to others, is judged by the product's revenue (USD 250,000), not by this row. Running ASTra as a continuing service or pipeline for a client's own operations is the client's use, so the client needs its own license (or must qualify itself). |
| A sole proprietor with no employees | Yes | A sole proprietorship is the person who owns it, so its use is Individual Use at any income. Personal projects, your own reports and documentation, planning, and deliverables for clients are all free. |
| A hobbyist or student | Yes | |
| A school, college or university, its students and staff | Yes | Teaching, learning and non-commercial academic research. |
| A registered charity or tax-exempt non-profit | Yes | For its mission. Not for selling products or services. |
| A government department, agency or municipality | Yes | For governmental functions. A state-owned company, or a body that sells products or services commercially, is treated as a company. |
| A company or other business with up to 10 people **and** under USD 250,000 revenue, using ASTra for its own work (together with affiliates, last completed fiscal year; a new company that has not finished one uses what it reasonably expects for its first) | Yes | If you pass either limit, you may continue for 90 days after the end of that fiscal year, then you need a license. |
| A company of any size whose product or service that has ASTra inside it has under USD 250,000 revenue | Yes, for the product | Headcount and the company's other revenue do not matter. The company's own internal use of ASTra is judged separately, in the row above. |
| A company above those limits using ASTra for its own work | No | Internal license. |
| A product or service with USD 250,000 revenue or more that has ASTra inside it | No | Embedded license. |

## What counts as internal use

Internal use is use of any part of ASTra by an organization's own people, systems and agents for its own operations. That includes:

- the MCP server, the CLI, the engine imported as a library into your own code, and the validator;
- scripts, automation, CI/CD jobs, batch jobs, internal pipelines and internal tools;
- AI agents that run for your organization;
- writing code against the engine for those purposes.

One license covers every part of ASTra together, on any number of computers, servers, agents and pipelines. Nothing is counted per machine, per server or per installation. Headcount counts people: employees, owners and partners who work in the business, and regular contractors. Servers, bots and AI agents do not count.

A pipeline that processes documents **for your own operations** is internal use. A pipeline or service that processes documents **for your customers**, or that lets third parties reach ASTra's operations, is not internal use. It is an embedded product, or a competing offering if its value is mainly ASTra.

Running your own tests and CI to build and check your code is free for everyone, because testing is non-production use.

## Common questions

**I am a freelancer. Can I use ASTra for client work?** Yes. Using it to produce documents for clients is covered. You pay nothing, and you do not need to tell anyone.

**I work alone and I am building a product with ASTra inside. Is it free at any revenue?** No. Working alone makes your own use free at any revenue, but a product or service of yours that has ASTra inside it and that you offer to others is judged by that product's revenue. It is free while the product's revenue is under USD 250,000, and after that it needs an Embedded license, whether you work alone or not. Your own use of ASTra for your own work stays free at any revenue.

**I work alone. Can I host ASTra or sell access to it?** No. A competing offering (a product or service whose value is mainly ASTra itself, such as hosting it for others, an API or MCP endpoint, or a thin editor on top of it) always needs a commercial license, for one person as for a large company.

**I work through my own company with no employees. Is that still free?** Yes. If you are the only person who works in the business, it is Individual Use at any revenue. If you hire an employee or take on a partner or a regular contractor, you are no longer working alone, and the small-organization limits apply to you.

**My startup is small. Do I need to do anything?** No. If you have up to 10 people and under USD 250,000 revenue, you may use ASTra for your own work. If you ship a product with ASTra inside, it is free until that product reaches USD 250,000 revenue. Keep an eye on the limits.

**Can I build ASTra into my own product or service?** Yes, if your product has substantial functionality of its own. It is free until that product's revenue reaches USD 250,000, whatever the size or headcount of your company. Your customers use ASTra only as part of your product. After that you need an Embedded license. Revenue means the revenue of the product or service, counting only the plans, modules or editions that include or rely on ASTra, not the rest of a larger suite. A free product has no revenue, unless it is a free edition of a product you also sell. Your own staff's internal use of ASTra is a separate matter (see the next question).

**My product is free, and my company is large. Do my own staff need a license?** If your own people use ASTra for your own work, yes: internal use above 10 people or USD 250,000 revenue needs an internal license. Building and testing the product does not count as internal use. It is non-production use and is free for everyone.

**Can I make videos, stream, review or teach with ASTra?** Yes. Recording a tutorial, running a live stream, writing a review or teaching a course is fine, and it does not matter whether you earn money from the video or the course. A person working on their own account is free whatever they earn, and demonstrating the software is non-production use, which is free for everyone. Hosting ASTra as a service for your audience, or reselling it, is a competing offering and always needs a commercial license.

**Can I build an open-source project on ASTra?** Yes, if your project has substantial functionality of its own. Your own code can be under any license, such as MIT or Apache-2.0. You list ASTra as a dependency and you do not relicense it. People who use your project use ASTra only as part of it, even if your project is free of charge. If your project is the product of a company, the company's own internal use of ASTra needs its own right to use it. A thin wrapper whose value is mainly ASTra itself is a competing offering and needs a commercial license. If you copy ASTra source into your repository, keep ASTra's LICENSE with it: ASTra code cannot be relicensed as MIT or Apache-2.0 until its own Change Date. Some license scanners and software distributions do not treat the BSL as open source, so describe your project as "depends on ASTra (source-available)".

**What is not allowed without a license?** A "competing offering": a product or service offered to others whose value consists mainly of ASTra itself. It includes hosting ASTra as a service for others; making ASTra's operations, commands, API or tools available to third parties as a service in itself, directly or through a thin interface (for example a public API or MCP endpoint whose main purpose is access to ASTra); a desktop, web or other editor whose value is mainly ASTra's document functionality; and selling ASTra, or a product or service whose value is mainly ASTra, with or without another name. A product that uses ASTra to implement substantial features of its own is not a competing offering. That includes an AI agent or an agent platform that uses ASTra as one of its tools: it is a product, free while its revenue is under USD 250,000, and it needs an Embedded license after that. You may copy, modify and redistribute ASTra free of charge, modified or not, provided every copy keeps the LICENSE and the copyright and license notices. Three tests decide it. It is a competing offering if most of what customers pay for or use it for is ASTra's document functionality; if little would remain without ASTra (an interface, login, billing or prompts around ASTra's operations are not features of your own); or if its main purpose is to let third parties reach ASTra's operations, commands, API or tools. A competing offering always needs a commercial license, whatever your size.

**Can I ask before I build?** Yes. Describe the product to astra@attentionform.com and we aim to tell you in writing within five working days whether it is a competing offering. This is a target, not a promise of a deadline.

**Do the limits include my parent company and affiliates?** Yes. Revenue and headcount are counted for your whole group: companies that control you, companies you control, and companies under the same control. A separate company set up to stay under the limits does not count as separate. An investor that holds 50% or less and has only customary minority rights, such as a board seat or consent over a financing or a sale, does not make its other portfolio companies count together with you. For an embedded product, only the revenue of that product counts, not the revenue of the whole group.

**How many people count?** Employees, owners and partners who work in the business, and regular contractors. Founders count. Contractors hired for one specific project, servers, bots and AI agents do not count. Headcount is counted at the end of the last completed fiscal year. Revenue in another currency is converted into US dollars at the exchange rate on the last day of your fiscal year.

**Can my company try ASTra for as long as it likes?** Yes, on sample documents, with no time limit. Building and testing an integration and demonstrating it are always free. Running your company's real work through it, even in a pilot, is production use and needs a license if your company is above the small-organization limits. If a pilot is going well, write to us: we answer quickly.

**Does my open-source project's CI need a license?** No. Running your project's tests is non-production use and is free for everyone. Using ASTra to process real documents in an organization's regular operations is production use. People who run your project in production need their own right to use ASTra.

**If I publish a free product that includes ASTra, do my users need a license?** If your product has substantial functionality of its own, your users may use ASTra as part of your product, whatever their size, while the product's revenue is under USD 250,000 or you hold an Embedded license. If the product's value is mainly ASTra, it is a competing offering and needs a commercial license.

**What if I stop qualifying?** You can continue your existing use for 90 days, then you need a commercial license or you stop. You stop qualifying on the first day after the fiscal year in which you passed a limit. Example: a company passes USD 250,000 in July 2027 with a calendar fiscal year. It stays free through 2027, stops qualifying on 1 January 2028, and has until about 31 March 2028 to buy.

**Do I have to publish my own code?** No. The BSL has no copyleft and no obligation to share your code.

**Is the license going to get stricter?** The terms of a released version never become stricter. A later version may carry different terms.

**Can I send a pull request?** Yes. Sign the [contributor license agreement](CLA.md) once, by adding one small signature file to your first pull request. See [CONTRIBUTING.md](CONTRIBUTING.md).

**What about the names and logos?** They are covered separately: see [TRADEMARKS.md](TRADEMARKS.md).

**What about third-party code inside ASTra?** It keeps its own license. See [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md). A folder that has its own license file is under that license.

## Other Attentionform products

Other Attentionform products have their own licenses, unless a license says that they are part of the ASTra suite.

## Examples

| # | Case | Result |
| --- | --- | --- |
| 1 | Freelancer, working personally, prepares documents for clients | Free |
| 2 | Freelancer through their own company, no employees, USD 300,000 revenue | Free (works alone; Individual Use at any revenue) |
| 3 | Company with no employees, 40 regular contractors, USD 20M revenue | Needs a license (the contractors count as people) |
| 4 | Shell company with no employees, owned by a USD 2B group | Needs a license (the group counts) |
| 5 | Freelancer runs an always-on pipeline on their own cloud for a bank's operations | The bank's use. The bank needs its own right |
| 6 | Five-person agency sells a finished app with its own features to a bank, product revenue under USD 250,000 | Free (a product with its own functionality) |
| 7 | Startup of 6 people with a product that earns USD 200,000 and uses ASTra inside it | Free, for the product and for its own internal use (under both limits) |
| 8 | The same startup grows to 12 people | The product stays free. Its internal use needs a license after the end of that fiscal year and 90 days |
| 9 | Company of 200 people tests ASTra for six months on sample documents | Free (non-production) |
| 10 | The same company runs its real contract workflow through ASTra "as a pilot" | Production use; needs a license |
| 11 | Company of 200 people, internal pipeline that processes real documents | Needs an internal license for its size |
| 12 | Eight-person startup offers a "document API" that calls ASTra's operations | Competing offering; needs a license |
| 13 | The same startup builds a contract-review app that edits documents with ASTra, with product revenue under USD 250,000 | Free |
| 14 | A desktop or web editor whose value is mainly ASTra's document functionality | Competing offering; needs a license |
| 15 | Individual publishes an open-source project that depends on ASTra | Free |
| 16 | Large company publishes a free open-source project that depends on ASTra | The project is free (no revenue). The company's own staff need an internal license if they use ASTra for the company's work |
| 17 | University lab | Free |
| 18 | Cloud provider sells managed ASTra | Competing offering; needs a license |
| 19 | Two-person partnership with no employees, USD 200,000 revenue | Free (partners count as people; within both limits) |
| 20 | Company with 9 employees and 2 working founders | 11 people: needs a license after the grace period |
| 21 | Company with 4 employees and INR 100 million revenue | Converted to USD at the fiscal year-end rate; above USD 250,000 means a license |
| 22 | Anyone, using a version more than four years after its first release | Apache-2.0 |
| 23 | Open-source project runs its own tests with ASTra in CI | Free (testing is non-production use, for anyone) |
| 24 | Small startup's contract-review app grows past USD 250,000 product revenue | Free until the end of that fiscal year, then 90 days to buy an Embedded license |
| 25 | Student uses ASTra for assignments, reports, a thesis or a course project | Free (Educational Use, and also Individual Use) |
| 26 | Person uses ASTra for hobby projects, personal documents, CVs, letters, family or club paperwork | Free (Individual Use) |
| 27 | Individual contributes to an open-source project, on their own account | Free (Individual Use) |
| 28 | Individual freelancer prepares reports for clients, one job at a time | Free (Individual Use) |
| 29 | Employee uses ASTra for their employer's work, or contributes to open source as part of their job | The employer's use. Free only if the employer is within the free categories (small organization, non-profit, public sector, education) |
| 30 | Student or hobbyist sells a product built on ASTra | Free while the product has substantial functionality of its own and its revenue stays under USD 250,000. A product whose value is mainly ASTra is a competing offering |
| 31 | A "document assistant" app: a chat window and a login around ASTra's edit operations, sold as the product | Competing offering; needs a license |
| 32 | A contract-review product: ASTra edits the files, but its own clause analysis, workflow and reports are what customers pay for | Not a competing offering. ASTra is a component |
| 33 | A browser extension whose main feature is a button that runs ASTra's operations on the user's documents | Competing offering; needs a license |
| 34 | An internal tool that a large company built on ASTra for its own staff | Not a competing offering. The company needs an internal license for its size |
| 35 | An AI agent or agent platform that uses ASTra as one of its tools | Not competing. It is a product: free while its revenue is under USD 250,000, then an Embedded license. A service whose main purpose is access to ASTra's tools is competing |
| 36 | Sole proprietor, no employees, with a large web business, uses ASTra for planning and documents | Free (Individual Use), at any revenue |
| 37 | The same person hires one employee, and the business has USD 5M revenue | No longer works alone. Needs a license for its own use (above USD 250,000 revenue) |
| 38 | Sole proprietor sells a contract-review app with its own features, built on ASTra, with under USD 250,000 product revenue | Free |
| 39 | The same product reaches USD 250,000 revenue | Free until the end of that fiscal year, then 90 days to buy an Embedded license |
| 40 | Anyone, even a solo person, hosts ASTra and sells API access, or sells a thin tool whose value is mainly editing Word files | Competing offering; needs a license at any size |
| 41 | A three-person company publishes an MIT project with ASTra bundled in | Free while the product's revenue is under USD 250,000. Users may use ASTra only as part of the project |
| 42 | The same project, but users install ASTra themselves instead of receiving it with the project | Free for the maker. ASTra is then not part of the product, so each user's use is governed by that user's own category |
| 43 | A registered charity or foundation hosts a project built on ASTra | Free (Non-Profit Use), if it is not selling products or services |
| 44 | A company pays a maintainer to build a project for the company's own use | The company's use. It needs a license for its size |
| 45 | A company tries ASTra on a few of its real files to see whether it works, and does not rely on the results in its regular operations | Free (evaluation) |
| 46 | Funded startup of 200 people with no revenue ships an app built on ASTra | The product is free (no revenue). Building and testing it is free. The staff's own internal use of ASTra needs an internal license |
| 47 | Company of 6 people with USD 400,000 revenue runs internal reports through ASTra | Needs an internal license (revenue is above USD 250,000) |
| 48 | A developer imports the ASTra engine into the company's CI pipeline that builds reports for the company's own operations | Internal use. The company's size decides. One license covers the whole suite, on any number of machines |
| 49 | The same pipeline produces documents for the company's customers as part of a paid service | Not internal use. An embedded product, free under USD 250,000 revenue, or a competing offering if the service's value is mainly ASTra |
| 50 | Person who works alone sells an app with its own features, built on ASTra, with USD 600,000 product revenue | The app needs an Embedded license (above USD 250,000 product revenue). The person's own use of ASTra for their own work stays free |
| 51 | Person who works alone hosts ASTra and sells access to its operations through an API | Competing offering; needs a license, even for one person |

## Open-source projects

The MIT, Apache or other license on your project's own code does not change ASTra's license. ASTra stays under its own license, and each person's use of it is judged by that person's own category.

- **An individual maintainer, a qualifying non-profit or foundation, a school:** free.
- **A company or business behind the project:** the project is free while its revenue is under USD 250,000 (a free project has none). The company's own internal use of ASTra is judged separately, by its size. A project that earns USD 250,000 or more needs an Embedded license.
- **ASTra bundled in the project:** users may use ASTra only as part of the project, for as long as the maker qualifies or holds the license.
- **Users install ASTra themselves:** ASTra is not part of the project, so each user's use is governed by that user's own category.
- **Forks and redistribution:** you may copy, modify and redistribute ASTra free of charge, provided every copy keeps the LICENSE and the copyright and license notices. Selling ASTra itself, or a product or service that is mainly ASTra, is a competing offering.

## Trying ASTra on real files

Trying ASTra on a few of your real files, to see whether it works for you, is free for everyone. It becomes production use when you rely on the results in your regular operations.

## Commercial licenses

Plans, limits and prices are on the [pricing page](https://attentionform.com/pricing). This repository does not list prices, so that plans can change without a new release. A commercial license is available at the prices published there on the date of purchase. In short:

- **Internal license:** for an organization above the free limits whose own people use ASTra for its own work. The plan is chosen by the number of people, counted with affiliates. One license covers the whole ASTra suite. If you pass the limit of your plan, upgrade to the next plan promptly, and within 30 days. The pricing page explains how.
- **Embedded license:** for a product or service of yours with substantial value of its own that has ASTra inside it, whether you are one person or a large company. It is chosen by the product's revenue. One Embedded license covers one named product. It also covers your organization's own internal use of ASTra up to 49 people. An organization with 50 or more people either buys an internal license as well, or asks for a Custom license.
- **Custom license:** for anyone, of any size. It can combine an internal license, an Embedded license and support in one contract, cover needs that no plan covers (your own contract paper, invoicing, procurement or security review), or answer questions before you buy.
- **Competing offerings** (a thin wrapper, hosting, an API or MCP endpoint, a mostly-ASTra editor, resale): always need their own license, at any size. The terms and the price are fully negotiable.
- **Support plan:** every license includes community support (GitHub and Discord, best effort, no response time promised). Anyone who lawfully uses ASTra, free or licensed, an individual or an organization, can buy a paid support plan, written down so you have a record of what we promise. It adds a named contact, a private email channel, priority handling of your bugs and an allowance of tickets and hours, including help getting started. Response times in a plan are targets. A support plan is not a service level agreement: no uptime commitment, no promised fix dates, no custom development.

Not sure which applies? Ask before you build or buy. Write to astra@attentionform.com with your use case. We aim to reply within 5 working days. To get a quote, share the details: your product's revenue band, how ASTra is used in your product, how many people use it, and how you deploy it (cloud, on-premises or desktop). If our written answer says your use does not need a license, you have our permission for that use for one year from the date of our answer, without paying anything. Ask again before the year ends if the use continues. We will renew a written confirmation on request for the same use. We may decline only if the use has changed materially so that it no longer qualifies, or our terms for new versions have changed materially. If your use does need a license, the terms and the price are negotiable.

## For maintainers

- The `LICENSE` file describes the Licensed Work as every version the Licensor makes publicly available, so it needs no edit per release. Each version's Change Date is four years after the day it was first made public: write that day next to the version in `CHANGELOG.md` when you release.
- The `LICENSE` in the root and in each package must stay identical. `tests/unit/license-consistency.test.ts` checks this.
- Use the SPDX id `BUSL-1.1` in `package.json`.
- Never describe ASTra as "open source". Say "source-available".
- Do not put prices or plan names in this repository. They live on the pricing page, so that they can change without a release.
