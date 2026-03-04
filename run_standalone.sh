#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PORT="${1:-8091}"

cd "$SCRIPT_DIR"

echo "Starting standalone OCR Picking Ticket server..."
echo "Folder: $SCRIPT_DIR"
echo "URL: http://127.0.0.1:${PORT}/"
echo "(redirects to /ocr_picking_ticket_standalone.html)"
echo

python3 -m http.server "$PORT"
