#!/usr/bin/env bash
# E6/E7 timing experiments, alone on one VM. Baseline (published package) and build alternate for PERF_ROUNDS
# rounds, so a slow or fast moment on the VM hits both sides. Results: results/perf-<label>-r<round>/.
set -u
ROUNDS=${PERF_ROUNDS:-2}
LOG="results/perf-sequence.log"
mkdir -p results
stamp() { echo "$(date -Is) $*" | tee -a "$LOG"; }
measure() { # label, server-bin ("" = published)
  local label=$1 bin=$2
  stamp "START $label"
  (
    export BENCH_LABEL="perf-${label}"
    if [ -n "$bin" ]; then export ASTRA_SERVER_BIN="$bin"; else unset ASTRA_SERVER_BIN; fi
    mkdir -p "results/${BENCH_LABEL}"
    node run/env.mjs > /dev/null
    node run/e6_performance.mjs > "results/${BENCH_LABEL}/log-e6_e7_performance.txt" 2>&1
  ) && stamp "END $label ok" || stamp "END $label FAILED"
}
for r in $(seq 1 "$ROUNDS"); do
  if [ "${WITH_BASELINE:-false}" = "true" ]; then measure "${BASELINE_VERSION}-published-r${r}" ""; fi
  measure "${BUILD_LABEL}-r${r}" "${TARGET_SERVER_BIN}"
done
stamp "ALL DONE"
exit 0
