#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "$0")"

./push-template.sh
./push-personal.sh
