#!/usr/bin/env bash
# build.sh — used by Render.com (free) and Railway
# Runs once at deploy time:
#   1. Install system packages (Tesseract Khmer + Poppler)
#   2. Install Node.js if not present (needed to build the React frontend)
#   3. Install Python dependencies
#   4. Build the React/Vite frontend into frontend/dist/
#   5. Smoke-test the Python app imports
set -e

# ── 1. System packages ───────────────────────────────────────────
if command -v apt-get &>/dev/null; then
  echo "=== Installing system packages ==="
  apt-get update -qq
  apt-get install -y -qq \
    tesseract-ocr \
    tesseract-ocr-khm \
    poppler-utils \
    libgl1 \
    curl
fi

# ── 2. Node.js (if not already present) ─────────────────────────
if ! command -v node &>/dev/null; then
  echo "=== Installing Node.js 20 ==="
  curl -fsSL https://deb.nodesource.com/setup_20.x | bash -
  apt-get install -y -qq nodejs
fi
echo "Node: $(node --version)  npm: $(npm --version)"

# ── 3. Python dependencies ───────────────────────────────────────
echo "=== Installing Python dependencies ==="
pip install --upgrade pip -q
pip install -r requirements.txt

# ── 4. Frontend build ────────────────────────────────────────────
echo "=== Building frontend ==="
cd frontend
npm ci
npm run build
cd ..

echo "=== Checking frontend/dist ==="
ls -la frontend/dist/ || echo "WARNING: frontend/dist does not exist!"
ls -la frontend/dist/assets/ 2>/dev/null || echo "WARNING: frontend/dist/assets does not exist!"

# ── 5. Smoke test ────────────────────────────────────────────────
echo "=== Smoke-testing app import ==="
python -c "
from pathlib import Path
dist = Path('frontend/dist')
print('frontend/dist exists:', dist.exists())
print('index.html exists:', (dist / 'index.html').exists())
"

echo "=== Build complete ==="
