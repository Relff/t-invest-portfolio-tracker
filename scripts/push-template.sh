#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "$0")/.."

cp .clasp.template.json .clasp.json
echo "→ Pushing Code/ to TEMPLATE script (public GitHub template)..."
clasp push
