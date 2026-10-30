#!/usr/bin/env bash
# Installs the packages the way a user would get them, from tarballs, into an empty project, and uses
# them: the validator on a sample file, the CLI to make a document, the validator on that document.
# Catches a missing file, a wrong bin or a broken binary lookup before anything reaches npm.
#
#   scripts/release-rehearsal.sh [platform-out-dir] [work-dir]
#
# The validator binary comes from <platform-out-dir>/<this platform> when it was built in this run,
# otherwise from npm, at the wrapper's binary range.
set -euo pipefail

root="$(cd "$(dirname "$0")/.." && pwd)"

# Both directories are used after a `cd`, and npm reads a bare "dir/name" as a GitHub repository, so
# they must be absolute paths (the release workflow passes a relative one).
abs() { case "$1" in /*) printf '%s' "$1" ;; *) printf '%s/%s' "$PWD" "$1" ;; esac; }
platform_dir="$(abs "${1:-$root/platform-out}")"
work="$(abs "${2:-$(mktemp -d)}")"
mkdir -p "$work/pkgs" "$work/project"

for pkg in ooxml-validator engine cli mcp-server; do
  npm pack "$root/packages/$pkg" --pack-destination "$work/pkgs" > /dev/null
done

host="$(node -p 'process.platform + "-" + process.arch')"
extra=()
if [ -d "$platform_dir/$host" ]; then
  npm pack "$platform_dir/$host" --pack-destination "$work/pkgs" > /dev/null
  echo "rehearsal: using the $host binary built in this run"
else
  range="$(node -p "require('$root/packages/ooxml-validator/package.json').validatorBinary.range")"
  extra+=("@attentionform/ooxml-validator-$host@$range")
  echo "rehearsal: using the $host binary from npm ($range)"
fi

cd "$work/project"
npm init -y > /dev/null
npm install --no-audit --no-fund "$work"/pkgs/*.tgz ${extra[@]+"${extra[@]}"}

npx --no-install ooxml-validator "$root/packages/ooxml-validator/tests/fixtures/valid.docx"
npx --no-install astra create "$work/rehearsal.docx" --template memo
npx --no-install ooxml-validator "$work/rehearsal.docx"
node -e "require.resolve('@attentionform/astra-mcp-server/package.json'); console.log('mcp server package resolves')"
echo "rehearsal: ok"
