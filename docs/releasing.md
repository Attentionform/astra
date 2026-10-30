# Releasing

Packages reach npm in one way: the **Release** workflow (`.github/workflows/release.yml`), started by hand from the Actions tab. Nothing starts it on its own, and pushing a tag does not publish.

## What can be released

Tick any mix, or `all`. `all` goes in this order.

| Choice | Publishes |
|---|---|
| `validator_platforms` | the six native validator binaries, one package per platform (`@attentionform/ooxml-validator-<platform>`) |
| `validator_wrapper` | `@attentionform/ooxml-validator` |
| `engine` | `@attentionform/astra-engine` |
| `cli` | `@attentionform/astra` |
| `mcp_server` | `@attentionform/astra-mcp-server` |

A package whose version is already on npm is skipped. A run that would publish nothing stops with an error.

## Every release is two runs

1. **Dry run** (the default). It plans, tests, builds the platform packages on their own operating systems, installs the packed packages into an empty project and uses them, and runs `npm publish --dry-run`. Nothing is published or tagged.
2. **The real run.** The same choices with `dry_run` turned off. It waits for your approval in the `npm` environment, plans again (npm may have changed), publishes, and tags.

## What the plan stops

- Nothing selected, or every selected version already on npm.
- A package that needs another of ours (the CLI needs the engine and the validator) when that one is neither on npm nor in the same run.
- A wrapper release when some platform has no usable binary on npm (see below).
- Packages that disagree about the Node version they need (`engines.node`, now `>=22.19.0` for all of them).

## The validator binary has its own version

The platform packages carry the **binary version**; the wrapper has its own version. The wrapper's `package.json` says which binaries it works with:

```json
"validatorBinary": { "version": "0.1.0", "range": ">=0.1.0 <0.2.0" }
```

- `version` is what the next `validator_platforms` run publishes.
- `range` is every binary the wrapper accepts. It is what the wrapper asks npm for.

If one platform fails to build in a release, the others are published and the wrapper can still go out, as long as that platform has an older binary inside the range on npm. Users on that platform keep the older binary. If a platform has no binary in the range, the wrapper is held back. Tick `require_all_platforms` when every platform must be on the new binary (the first release, or a change in how the wrapper and the binary talk to each other).

Raise the binary `version` when the native validator changes. Move the `range` up only when the wrapper and an older binary stop working together; that release needs all six platforms.

## Versions and tags

Raise versions in a normal commit before releasing; the workflow publishes what the `package.json` files say. After a real publish, each package published in that run is tagged on the released commit: `validator-binary-v<version>` (only when all six platforms are on npm at it), `validator-v<version>`, `engine-v<version>`, `cli-v<version>`, `mcp-server-v<version>`. A tag that already exists is left alone.

## A partial release

Every publish skips a version already on npm, so running the same choices again finishes a release that stopped half way. A platform that failed shows as a failed job; fix the cause and run `validator_platforms` again.

## One-time setup

1. An npm organisation `attentionform` and an automation token that can publish to it.
2. A GitHub environment named **`npm`** with required reviewers (you), and the token as its secret `NPM_TOKEN`. Only the publish job uses this environment, so build and test steps never see the token.
3. Optional: a GitHub environment **`npm-dry-run`** with the same `NPM_TOKEN`, so the dry run also checks the npm login. Without it the dry run notes that the login was not checked.
4. Provenance (the signed "built from this repository" mark) is added automatically when the repository is public.

A bad release cannot be replaced under the same version. Publish a new patch version, and use `npm deprecate` on the bad one.
