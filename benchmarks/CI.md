# Running the benchmarks on GitHub Actions

The suite can run on a GitHub-hosted runner, so a developer machine is not loaded for hours and the
environment is the same every time. The workflow is `.github/workflows/benchmarks.yml`; the step runner is
`.github/scripts/run-benchmarks.sh`.

## Branch rules

* The runner lives on a **`bench/**` branch** (this one is `bench/runner`). Only pushes to `bench/**` and manual
  dispatches start it. It never runs on `main`, never publishes, never tags.
* The harness (`benchmarks/`) always comes from the bench branch. The code under test comes from `target_ref`,
  so any branch, tag or commit of this repository can be measured without copying it here.
* To pick up harness changes made on `benchmarks`, merge or cherry-pick them into the bench branch. Do not edit
  harness behaviour on a bench branch without recording why in `PROTOCOL.md`.

## Run it

1. Edit `benchmarks/bench.config.json`:
   * `target_ref`: branch, tag or commit to build and measure.
   * `baseline`: `published` also runs the package pinned in `benchmarks/package.json`; `none` skips it.
   * `experiments`: space-separated steps from `e1 e1_oracles e1_control e4 e5 e5_orphans e6` (`e6` also runs E7).
2. Commit and push the bench branch. The run starts by itself.
3. Or dispatch by hand and override any field:
   `gh workflow run benchmarks.yml --ref bench/runner -f target_ref=<ref> -f experiments="e6"`
   (a dispatch needs the workflow file on the default branch; the push route always works).

## What runs

Two jobs, one after the other, never in parallel (parallel jobs would disturb each other's timings):

| Job | Measures | Result label |
|---|---|---|
| `baseline` | the published package pinned in `benchmarks/package.json` | `<pinned version>-hosted-runner` |
| `build` | a local build of `target_ref` (`packages/mcp-server/dist`) | `<package version>-build-<sha7>` |

Inside a job the experiments run one at a time, in the order given. A failed step is recorded in
`sequence.log` and the next step still runs; a failure is a result, not something to hide.
E1 uses 2 workers and the oracles 4 LibreOffice workers (the runner has 4 vCPUs). Worker counts change run time,
not the E1/oracle results.

## Read the results

Each job uploads one artifact: `results-baseline-<version>` and `results-build-<label>`.

```bash
gh run download <run-id> --dir /tmp/bench-results
```

Inside: the raw `*.jsonl` per experiment, `env.json` (runner, CPU, versions, and for the build the exact git
commit), `sequence.log` (a timestamp for every step) and one `log-<step>.txt` per step.

## Read the numbers correctly

* Compare `baseline` with `build` from the **same run**. Both used the same kind of runner, one after the other.
* Do not compare hosted-runner timings with timings from another machine (for example the stored `results/0.4.1`).
  Counts and pass/fail results (E1 locality, E4 atomicity, E5 history) do not depend on the machine; timings
  (E6, E7) do. Shared runners are noisy: the protocol's median, p95 and n are reported, and a run with a very
  wide spread should be repeated before it is quoted.
* Every result names its build. A local build is not a release; say so wherever the numbers are quoted.

## Not covered here

C1, C2, E3 and R1 are not in the default list. Add them to `run-benchmarks.sh` (and `experiments`) when a
change can affect them.
