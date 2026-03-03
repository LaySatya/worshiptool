# ChurchTool — Khmer Worship Slide Generator

Convert Khmer worship song PDFs / scanned images into PowerPoint slides automatically.  
Built with **FastAPI + React (Vite) + Tesseract OCR**.

---

## 🚀 Deploy to Railway (recommended)

Railway handles Tesseract + Poppler system packages automatically via `railway.toml`.

### One-time setup

1. Push this repo to GitHub
2. Go to [railway.app](https://railway.app) → **New Project → Deploy from GitHub repo**
3. Select the repository — Railway will detect `railway.toml` and run `build.sh` automatically
4. Once deployed, click **Generate Domain** to get your public URL

That's it. The React frontend is built at deploy time and served by the same FastAPI process.

---

## 💻 Local development

### Prerequisites
```bash
# macOS
brew install tesseract tesseract-lang poppler

# Ubuntu/Debian
sudo apt install tesseract-ocr tesseract-ocr-khm poppler-utils
```

### Backend
```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
uvicorn main:app --reload --port 8000
```

### Frontend (separate terminal)
```bash
cd frontend
npm install
npm run dev        # http://localhost:5173
```

The Vite dev server proxies `/api/*` to `localhost:8000` automatically.

---

## Project structure
```
song_slide_app/
├── main.py             ← FastAPI backend (Bible slides + Song OCR)
├── requirements.txt    ← Python dependencies (pinned)
├── Procfile            ← gunicorn start command
├── railway.toml        ← Railway build + system package config
├── build.sh            ← Build script (pip install + npm build)
├── fonts/              ← Bundled Khmer OS fonts
├── frontend/
│   ├── src/            ← React + Vite source
│   └── dist/           ← Built output (generated, not committed)
├── output/             ← Generated PPTX files (local only)
└── uploads/            ← Temp upload storage (local only)
```
