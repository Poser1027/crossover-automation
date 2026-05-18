#!/usr/bin/env bash
# ─── Hengli Crossover launcher (macOS / Linux) ───
set -e
cd "$(dirname "$0")"

if ! command -v python3 >/dev/null 2>&1; then
    echo "Python 3 not found. Please install Python 3.9+."
    exit 1
fi

if ! python3 -c "import fastapi, uvicorn, openpyxl" 2>/dev/null; then
    echo "Installing dependencies..."
    python3 -m pip install --quiet fastapi uvicorn openpyxl
fi

echo ""
echo "============================================================"
echo " Hengli Crossover - running at http://127.0.0.1:8000"
echo " Press Ctrl+C to stop."
echo "============================================================"
echo ""

# Open browser automatically
(sleep 2 && (command -v open >/dev/null && open http://127.0.0.1:8000 || \
             command -v xdg-open >/dev/null && xdg-open http://127.0.0.1:8000)) &

python3 -m uvicorn api:app --host 127.0.0.1 --port 8000
