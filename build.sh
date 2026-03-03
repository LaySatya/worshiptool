#!/usr/bin/env bash
# build.sh — Railway build script
# Runs once at deploy time: installs Python deps + builds the React frontend.
set -e

echo "=== Installing Python dependencies ==="
pip install -r requirements.txt

echo "=== Building frontend ==="
cd frontend
npm ci
npm run build
cd ..

echo "=== Build complete ==="
