# Song Slide Generator
Convert Khmer worship song PDFs (notation + lyrics) into PowerPoint slides automatically.

## Setup

### 1. Install Poppler (required for PDF conversion)
**Mac:**
```bash
brew install poppler
```
**Ubuntu/Linux:**
```bash
sudo apt install poppler-utils
```

### 2. Create virtual environment
```bash
python3 -m venv venv
source venv/bin/activate
```

### 3. Install Python packages
```bash
pip install -r requirements.txt
```

### 4. Run the app
```bash
python3 app.py
```

Open your browser at: **http://127.0.0.1:5000**

---

## How It Works
1. Upload your scanned song PDF
2. The app converts each page to a high-res image
3. It scans for horizontal white-space gaps to find each music block
4. Each block (notation staff + Khmer lyrics) becomes one slide
5. Download the ready `.pptx` file

## Project Structure
```
song_slide_app/
├── app.py              ← Flask backend
├── requirements.txt    ← Python dependencies
├── templates/
│   └── index.html      ← Upload UI
├── uploads/            ← Temp PDF storage
└── output/             ← Generated files
```

## Options
- **Auto (Smart detect)** — detects white gaps between music systems
- **Merge pairs** — merges adjacent small blocks together
- **DPI** — higher = better quality but slower (200 is recommended)
