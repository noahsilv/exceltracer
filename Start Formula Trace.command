#!/bin/bash
set -eu
cd "$(dirname "$0")"
# Finder may start Terminal without the user's interactive shell configuration.
export PATH="$PATH:/opt/homebrew/bin:/usr/local/bin"
if ! command -v node >/dev/null || ! node -e 'process.exit(Number(process.versions.node.split(".")[0]) >= 22 ? 0 : 1)'; then
  echo 'Formula Trace needs Node.js 22 or newer. If using a version manager, activate it in Terminal and run this launcher there.' >&2
  exit 1
fi
exec node scripts/mac.mjs
