#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "$0")"

cp .clasp.personal.json .clasp.json
echo "→ Pushing Code/ to PERSONAL script (your live portfolio)..."
clasp push
