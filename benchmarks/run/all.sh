#!/usr/bin/env bash
# Run every experiment in order. Performance (E6/E7) runs last so nothing else loads the machine.
set -euo pipefail
cd "$(dirname "$0")/.."
node run/env.mjs >/dev/null
[ -f corpus/files/manifest.json ] || [ -d corpus/files ] || npm run corpus
node run/r1_representation.mjs
E1_WORKERS=${E1_WORKERS:-6} node run/e1_edit_locality.mjs
node run/oracle_pass.mjs e1
node run/e1_control.mjs
node run/c2_build.mjs
node run/oracle_pass.mjs c2
node run/c1_create.mjs
node run/oracle_pass.mjs c1
python3 lib/faults.py corpus/files corpus/faults
node run/e3_gate.mjs
node run/e4_atomicity.mjs
node run/e5_history.mjs
node run/e6_performance.mjs
node run/report.mjs
