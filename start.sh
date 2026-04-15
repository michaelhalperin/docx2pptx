#!/bin/bash
# ─────────────────────────────────────────────────────────
#  docx2pptx — launcher (Mac / Linux)
#  Double-click this file (or run: bash start.sh)
# ─────────────────────────────────────────────────────────

DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$DIR"

# Install dependencies if node_modules doesn't exist
if [ ! -d "node_modules" ]; then
  echo "📦 Installing dependencies (first run only)…"
  npm install pptxgenjs react react-dom react-icons sharp
fi

echo ""
echo "🚀 Starting docx2pptx server…"
echo "   Open your browser at: http://localhost:4242"
echo "   Press Ctrl+C to stop."
echo ""

node server.js
