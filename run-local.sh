#!/bin/bash
set -euo pipefail

cd "$(dirname "$0")"

PORT=8501
URL="http://localhost:${PORT}"

if lsof -ti :"${PORT}" >/dev/null 2>&1; then
  echo "Stopping previous Streamlit on port ${PORT}..."
  kill "$(lsof -ti :"${PORT}")" 2>/dev/null || true
  sleep 1
fi

echo ""
echo "Starting Junior Analyst Copilot"
echo "Open in browser: ${URL}"
echo "Press Ctrl+C in this window to stop."
echo ""

streamlit run app.py --server.port "${PORT}"
