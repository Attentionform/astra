#!/usr/bin/env bash
# Runs the requested benchmark steps one at a time, with a timestamp line per step.
# A failing step is recorded and the next step still runs (the failure is part of the result).
set -u
LOG="results/${BENCH_LABEL}/sequence.log"
mkdir -p "results/${BENCH_LABEL}"
stamp() { echo "$(date -Is) $*" | tee -a "$LOG"; }
failed=0
run() {
  local name=$1; shift
  stamp "START $name"
  if "$@" > "results/${BENCH_LABEL}/log-${name}.txt" 2>&1; then stamp "END $name ok"; else stamp "END $name FAILED"; failed=1; tail -n 20 "results/${BENCH_LABEL}/log-${name}.txt"; fi
}
for step in ${EXPERIMENTS}; do
  case "$step" in
    e1)         run e1_edit_locality env E1_WORKERS=2 node run/e1_edit_locality.mjs ;;
    e1_oracles) run e1_oracles node run/oracle_pass.mjs e1 ;;
    e1_control) run e1_control node run/e1_control.mjs ;;
    e4)         run e4_atomicity node run/e4_atomicity.mjs ;;
    e5)         run e5_history env E5_WORKER=0 E5_WORKERS=1 node run/e5_history.mjs ;;
    e5_orphans) run e5_orphans node run/e5_orphans.mjs ;;
    e6)         run e6_e7_performance node run/e6_performance.mjs ;;
    *) stamp "unknown step $step"; failed=1 ;;
  esac
done
stamp "ALL DONE failed=$failed"
exit 0
