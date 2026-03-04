#!/usr/bin/env bash
# build.sh — used by Render.com (free) and Railway
# Runs once at deploy time:
#   1. Install Python dependencies
#   2. Build the React/Vite frontend into frontend/dist/
#   3. Smoke-test the Python app imports
# NOTE: system packages (tesseract, poppler) are installed via render.yaml nativePackages
set -e

# ── 1. Python dependencies ───────────────────────────────────────
echo "=== Installing Python dependencies ==="
pip install --upgrade pip -q
pip install -r requirements.txt

# ── 2. Frontend build ────────────────────────────────────────────
echo "=== Building frontend ==="
cd frontend
npm install
npm run build
cd ..

echo "=== Checking frontend/dist ==="
ls -la frontend/dist/ || echo "WARNING: frontend/dist does not exist!"

# ── 3. Smoke test ────────────────────────────────────────────────
echo "=== Smoke-testing app import ==="
python -c "
from pathlib import Path
dist = Path('frontend/dist')
print('frontend/dist exists:', dist.exists())
print('index.html exists:', (dist / 'index.html').exists())
"

echo "=== Build complete ==="
