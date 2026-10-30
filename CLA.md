# ASTra Individual Contributor License Agreement, version 1

This agreement is between you and Attentionform, the publisher of ASTra, operated by the holder of the GitHub account github.com/Userfrom1995 and also identified by the GitHub organization github.com/attentionform and the domain attentionform.com, together with its successors and assigns, including any company or other person to which it transfers ASTra ("Licensor"). It applies to every Contribution you submit after you accept it. "Contribution" means any work you intentionally submit to Licensor for inclusion in ASTra, in any form (code, documentation, tests, images, text), including changes to existing work, for example in a pull request. Work you clearly mark "Not a Contribution" is excluded. "You" means the individual who accepts this agreement. You are at least 18 years old, or of the age of majority where you live, and you are free to enter into this agreement.

1. **Copyright license.** In exchange for Licensor considering your Contributions for inclusion in ASTra, you grant Licensor a perpetual, worldwide, non-exclusive, no-charge, royalty-free, irrevocable copyright license to reproduce, prepare derivative works of, publicly display, publicly perform, **sublicense and distribute** your Contributions and such derivative works, in source and object form, **under any license terms, including commercial and proprietary terms**. This includes changing the license of any version of ASTra, converting it under a change license, and offering it under more than one license at once. Anyone else who receives your Contribution as part of ASTra receives it only under the license terms Licensor chooses to offer; this agreement gives them no rights in it.
2. **Patent license.** You grant Licensor and recipients of software distributed by Licensor a perpetual, worldwide, non-exclusive, no-charge, royalty-free, irrevocable patent license to make, have made, use, offer to sell, sell, import and otherwise transfer the work, under the patent claims you can license that are necessarily infringed by your Contribution alone or in combination with ASTra. If any entity starts patent litigation (including a cross-claim or counterclaim) alleging that your Contribution, or the ASTra work it was submitted to, infringes a patent, the patent licenses you granted to that entity under this agreement end on the date it starts.
3. **Your rights.** You confirm that: (a) except for material you identify under clause 4, each Contribution is your original work; (b) you have the right to grant the licenses above; (c) if your employer or anyone else has rights in your work, you have their written permission or waiver, or your employer has signed the corporate agreement ([CLA-CORPORATE.md](CLA-CORPORATE.md)); (d) if you used an AI tool, you reviewed what it produced, you followed the tool's terms, and to your knowledge it does not copy third-party material that Licensor could not license under clause 1.
4. **Third-party material.** If you include material you did not create, you will say so in the pull request and identify its source and license. Submit it only if its license allows Licensor to distribute it with ASTra under the Business Source License 1.1 and, after its change date, under the Apache License 2.0.
5. **No warranty and no duty.** Your Contributions are provided "as is", without warranty of any kind. You are not required to provide support. Licensor is not obliged to use, keep, credit or pay for any Contribution.
6. **Notice.** You will tell Licensor if any statement above becomes inaccurate.
7. **Ownership.** You keep the copyright in your Contributions. This agreement does not transfer ownership.
8. **Moral rights.** To the extent the law allows, you waive, and otherwise agree not to assert against Licensor or anyone who receives ASTra from Licensor, any moral rights (such as the right to be named or to object to changes) in your Contributions. The repository history keeps recording your name; Licensor is not required to credit you anywhere else.
9. **Versions.** If Licensor publishes a new version of this agreement, it applies to Contributions you submit after you accept the new version. Licenses you already granted are not affected and stay irrevocable. Once people have signed a version, its text does not change: the repository records a SHA-256 fingerprint of each version, and your signature file names the fingerprint of the text you accepted.
10. **Successors and assigns.** Licensor may assign or transfer this agreement, in whole or in part, to a successor, an affiliate (including a company formed to hold ASTra) or anyone that takes over ASTra. You may not assign it.
11. **General.** If a part of this agreement is invalid, the rest stays in force. This is the whole agreement about Contributions. The English text prevails over any translation.
12. **Acceptance and records.** You accept this agreement by adding your signature file, as described below. You agree that the file, the Git history, and GitHub's records of your account and pull request are evidence of your acceptance and are your electronic signature.

## How to sign

Add one file to your pull request, once: `cla/signatures/<your-numeric-github-id>.json`. If you open the pull request without it, the `cla` check posts the exact file for you, with the fingerprint filled in. The file must be added in commits you wrote yourself.

```json
{
  "github_id": 1234567,
  "github_login": "janedoe",
  "name": "Jane Doe",
  "cla_version": 1,
  "cla_sha256": "<fingerprint of this text, filled in by the check>",
  "statement": "I have read the ASTra CLA v1 and I agree to it."
}
```

Adding this file in your own commit is your electronic signature. The date of acceptance is the time GitHub recorded for your commit and pull request. The file is public, and it records your GitHub account number and name, the CLA version, the fingerprint of the text you accepted and nothing else. You do not need to sign again for later pull requests. If this agreement changes, the version number changes and you will be asked to raise `cla_version` (and the fingerprint) in your own file.

Privacy. The signature file is kept to show that you agreed to this agreement and so that Licensor can prove it may license your Contributions. Licensor processes it to perform this agreement and for its legitimate interest in proving that right. Because the record has to exist for as long as ASTra is licensed with your Contribution in it, a signature cannot be erased while that is so. Questions or requests: astra@attentionform.com.
