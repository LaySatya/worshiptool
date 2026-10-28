"""
ChurchTool — FastAPI Backend
  • Bible Slide Generator
  • Song / Music-Sheet OCR → Lyric Slide Generator
"""
from __future__ import annotations

import io
import os
import re
import uuid
import zipfile
import base64
from collections import Counter
from pathlib import Path
from typing import Optional

import requests
from fastapi import FastAPI, File, Form, UploadFile, Request, Response, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, StreamingResponse

"""
Bible Reader: an embedded iframe of wordproject.org (Khmer Bible) loaded via a
same-origin proxy so text selections can be captured and piped into the slide
editor automatically.
"""
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from lxml import etree

app = FastAPI(title="ChurchTool — Bible & Song Slide Generator")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# ─────────────────────────────────────────────
# Models
# ─────────────────────────────────────────────

class SlideSettings(BaseModel):
    font_family: str = "Arial"
    font_size: int = Field(36, ge=8, le=120)
    bold: bool = False
    align: str = "center"
    line_spacing: float = Field(1.2, ge=0.5, le=4.0)
    bg_color: str = "#000000"
    text_color: str = "#ffffff"
    ref_color: str = "#aaaaaa"
    ref_font_size: int = Field(20, ge=8, le=72)
    verse_num_color: str = "#ffcc44"   # verse number colour
    padding: float = Field(0.55, ge=0.1, le=2.0)
    show_ref_each: bool = True


class GenerateRequest(BaseModel):
    raw_text: str
    per_slide: int = Field(1, ge=1, le=20)
    settings: SlideSettings = SlideSettings()
    reference: str = ""


class SlidePart(BaseModel):
    """A single rendered slide."""
    verse_num: str        # e.g. "12" — rendered in verse_num_color
    lines: list[str]
    ref: str


class PreviewRequest(BaseModel):
    raw_text: str
    per_slide: int = Field(1, ge=1, le=20)
    settings: SlideSettings = SlideSettings()
    reference: str = ""


# ─────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────

def _hex_to_rgb(hex_color: str) -> tuple[int, int, int]:
    h = hex_color.lstrip("#")
    if len(h) == 3:
        h = "".join(c * 2 for c in h)
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)


ALIGN_MAP = {
    "left": PP_ALIGN.LEFT,
    "center": PP_ALIGN.CENTER,
    "right": PP_ALIGN.RIGHT,
}

# ── Verse number pattern: Arabic digits OR Khmer digits at line start
# Matches: "12 ...", "12. ...", "12) ...", "[12] ...", "១២ ..."
_NUM_PAT = r"[0-9០-៩]+"
KHMER_VERSE_RE = re.compile(rf"^\[?({_NUM_PAT})\]?(?:[\.\:\)\-]\s*|\s+)(.+)$")

KHMER_DIGITS = "០១២៣៤៥៦៧៨៩"
ARABIC_TO_KHMER = str.maketrans("0123456789", KHMER_DIGITS)


def to_khmer_number(num_str: str) -> str:
    """Convert any Arabic digits in a verse number string to Khmer digits (e.g. '12' -> '១២')."""
    if not num_str:
        return ""
    return str(num_str).translate(ARABIC_TO_KHMER)


def get_slide_char_limit(settings: SlideSettings | None) -> int:
    """
    Calculate the character capacity to make a slide genuinely full before
    spilling over to a continuation slide.
    A full widescreen slide holds ~195-220 characters at 60pt font,
    and ~260-320 characters at 36pt font.
    """
    if not settings:
        return 200
    base = 200.0 * (60.0 / max(settings.font_size, 16)) ** 0.5
    base = base * (1.5 / max(settings.line_spacing, 0.8)) ** 0.3
    pad_factor = (13.33 - 2 * min(settings.padding, 1.5)) / (13.33 - 2 * 0.1)
    base = base * pad_factor
    return max(150, int(base))


def split_khmer_text_for_slides(text: str, max_chars: int = 200) -> list[str]:
    """
    Split a long Khmer text across slides.
    Ensures that the first slide is filled FULL first (taking as much text as fits),
    and only the remaining overflow text is cut into the next slide.
    """
    text = text.strip()
    if not text or len(text) <= max_chars:
        return [text] if text else []

    boundaries = []
    # 1. Punctuation (highest priority for clean breaks)
    for m in re.finditer(r'[។៕?!;,\n]', text):
        boundaries.append((m.end(), 10))
    # 2. Whitespace / spaces between clauses
    for m in re.finditer(r'\s+', text):
        boundaries.append((m.start(), 8))
    # 3. Zero-width space (\u200b)
    for m in re.finditer(r'\u200b', text):
        boundaries.append((m.end(), 5))

    boundaries.sort(key=lambda x: x[0])

    chunks = []
    start = 0
    while start < len(text):
        remaining = len(text) - start
        if remaining <= max_chars:
            chunks.append(text[start:].strip(' \t\n\u200b'))
            break

        max_p = start + max_chars
        # To make the first slide full first, search for a natural boundary
        # in the upper range [start + 65% of max_chars, max_p]
        upper_min = start + int(max_chars * 0.65)
        cands = [b for b in boundaries if upper_min <= b[0] <= max_p]

        if cands:
            # Pick candidates near the maximum boundary to fill the slide as much as possible
            furthest_pos = max(b[0] for b in cands)
            near_top = [b for b in cands if b[0] >= furthest_pos - 35]
            near_top.sort(key=lambda b: (b[1], b[0]), reverse=True)
            split_at = near_top[0][0]
        else:
            # If no boundary in upper 35%, search in [start + 45%, max_p]
            mid_min = start + int(max_chars * 0.45)
            mid_cands = [b for b in boundaries if mid_min <= b[0] <= max_p]
            if mid_cands:
                furthest_pos = max(b[0] for b in mid_cands)
                near_top = [b for b in mid_cands if b[0] >= furthest_pos - 35]
                near_top.sort(key=lambda b: (b[1], b[0]), reverse=True)
                split_at = near_top[0][0]
            else:
                # Look slightly ahead if a space/punct is just past max_chars (within 15 chars)
                ahead = [b for b in boundaries if max_p < b[0] <= max_p + 15 and b[1] >= 8]
                if ahead:
                    split_at = ahead[0][0]
                else:
                    fallback = [b for b in boundaries if start + 20 <= b[0] <= max_p]
                    split_at = fallback[-1][0] if fallback else max_p

        chunk = text[start:split_at].strip(' \t\n\u200b')
        if chunk:
            chunks.append(chunk)
        start = split_at

    return [c for c in chunks if c]


# ── English reference blocks: "John 3:16\nbody…"
VERSE_REF_RE = re.compile(
    r"""
    (?:^|\n)\s*
    (?P<ref>
        (?:\d+\s+)?[A-Z][a-z]+(?:\s+[A-Za-z]+)*
        \s+\d+:\d+(?:[–\-]\d+)?
    )\s*\n
    (?P<body>.*?)
    (?=(?:\d+\s+)?[A-Z][a-z]+(?:\s+[A-Za-z]+)*\s+\d+:\d+|$)
    """,
    re.DOTALL | re.VERBOSE,
)


def parse_verses(raw_text: str) -> list[dict]:
    """
    Returns list of {ref, verse_num, lines}.
    Verse numbers are always converted to Khmer digits (e.g. '12' -> '១២').

    Detection priority:
      1. Khmer-style leading verse numbers ("12 ដំ...", "12. ...", "12) ...", "១២ ...")
      2. English reference headers         ("John 3:16\\nbody")
      3. Blank-line block fallback
    """
    raw_text = raw_text.strip()
    if not raw_text:
        return []

    # Strategy 1 — Khmer / numbered lines
    # Detects lines starting with verse numbers (Arabic or Khmer digits)
    lines_with_numbers = [l for l in raw_text.splitlines() if l.strip() and KHMER_VERSE_RE.match(l.strip())]
    non_blank = [l for l in raw_text.splitlines() if l.strip()]

    if lines_with_numbers and len(lines_with_numbers) / max(len(non_blank), 1) >= 0.20:
        verses: list[dict] = []
        curr: dict | None = None
        for line in raw_text.splitlines():
            line = line.strip()
            if not line:
                continue
            m = KHMER_VERSE_RE.match(line)
            if m:
                curr = {
                    "ref": "",
                    "verse_num": to_khmer_number(m.group(1)),
                    "lines": [m.group(2).strip()],
                }
                verses.append(curr)
            elif curr:
                # Continuation of the current verse: join with space so it's a unified verse text
                curr["lines"][-1] = (curr["lines"][-1] + " " + line).strip()
            else:
                curr = {"ref": "", "verse_num": "", "lines": [line]}
                verses.append(curr)
        if verses:
            return verses

    # Strategy 2 — English reference blocks
    eng_matches = list(VERSE_REF_RE.finditer(raw_text))
    if eng_matches:
        verses = []
        for m in eng_matches:
            lines = [l.strip() for l in m.group("body").splitlines() if l.strip()]
            if lines:
                verses.append({"ref": m.group("ref").strip(), "verse_num": "", "lines": lines})
        if verses:
            return verses

    # Strategy 3 — blank-line blocks
    verses = []
    for block in re.split(r"\n\s*\n", raw_text):
        block = block.strip()
        if not block:
            continue
        b_lines = [l.strip() for l in block.splitlines() if l.strip()]
        if not b_lines:
            continue
        # Check if first line in block has a verse number
        m = KHMER_VERSE_RE.match(b_lines[0])
        if m:
            v_num = to_khmer_number(m.group(1))
            first_line = m.group(2).strip()
            # If there are subsequent lines in this block, join them with space
            rest_text = " ".join(b_lines[1:])
            full_text = (first_line + " " + rest_text).strip() if rest_text else first_line
            verses.append({"ref": "", "verse_num": v_num, "lines": [full_text]})
        else:
            verses.append({"ref": "", "verse_num": "", "lines": b_lines})
    return verses


def _chunk(items: list, size: int) -> list[list]:
    return [items[i: i + size] for i in range(0, len(items), size)]


def build_slides(
    verses: list[dict],
    per_slide: int,
    global_ref: str,
    show_ref_each: bool,
    settings: SlideSettings | None = None,
) -> list[SlidePart]:
    """
    Build slides from verses.
    - Standard verses each fit cleanly on their own slide.
    - Verse numbers are always rendered in Khmer numerals ('១', '២', '១២').
    - Only truly long verses that genuinely overflow the whole slide page (> char_limit)
      are cut into multiple slides representing that verse.
    """
    char_limit = get_slide_char_limit(settings)
    slides: list[SlidePart] = []

    for v in verses:
        ref = v.get("ref") or global_ref
        verse_num = to_khmer_number(v.get("verse_num", ""))
        raw_lines = [l.strip() for l in v.get("lines", []) if l.strip()]
        if not raw_lines:
            continue

        # Combine lines of this verse to measure total verse length
        full_verse_text = " ".join(raw_lines)

        # Check if the verse genuinely overflows the whole slide
        if len(full_verse_text) > char_limit:
            # Genuinely overflows the slide page -> cut into balanced slide chunks
            chunks = split_khmer_text_for_slides(full_verse_text, max_chars=char_limit)
            total_chunks = len(chunks)
            for i, chunk in enumerate(chunks):
                show = show_ref_each or (i == total_chunks - 1)
                # First slide gets the verse_num; subsequent continuation slides (e.g. second slide)
                # have no verse number (empty string), putting only the remaining text.
                slide_vnum = verse_num if i == 0 else ""
                slides.append(SlidePart(
                    verse_num=slide_vnum,
                    lines=[chunk],
                    ref=ref if show else "",
                ))
        else:
            # Fits on a single slide cleanly!
            slides.append(SlidePart(
                verse_num=verse_num,
                lines=[full_verse_text],
                ref=ref,
            ))

    return slides


# ─────────────────────────────────────────────
# PPTX builder
# ─────────────────────────────────────────────

def _fix_font(run, font_name: str):
    """
    python-pptx sets font name via theme inheritance which PowerPoint may override.
    Directly writing <a:latin typeface="..."/> on the run's rPr guarantees the font
    is respected without needing a manual click in PowerPoint.
    """
    rPr = run._r.get_or_add_rPr()
    # Remove any existing <a:latin> elements first
    for el in rPr.findall(qn("a:latin")):
        rPr.remove(el)
    latin = etree.SubElement(rPr, qn("a:latin"))
    latin.set("typeface", font_name)
    # For complex/East-Asian scripts (Khmer uses complex script path in pptx)
    for el in rPr.findall(qn("a:cs")):
        rPr.remove(el)
    cs = etree.SubElement(rPr, qn("a:cs"))
    cs.set("typeface", font_name)


def _set_line_spacing(p, line_spacing: float):
    """
    Apply line spacing as lnSpc/spcPct.
    PowerPoint value: 1.0 = 100000, 1.5 = 150000, 2.0 = 200000.
    We also clear any spcBef/spcAft so there is no extra gap between lines.
    """
    pPr = p._p.get_or_add_pPr()
    # Remove any existing line-spacing / space-before / space-after elements
    for tag in ("a:lnSpc", "a:spcBef", "a:spcAft"):
        for el in pPr.findall(qn(tag)):
            pPr.remove(el)
    # Insert lnSpc as the first child of pPr (required order by schema)
    lnSpc = etree.SubElement(pPr, qn("a:lnSpc"))
    spcPct = etree.SubElement(lnSpc, qn("a:spcPct"))
    spcPct.set("val", str(int(line_spacing * 100000)))
    # Zero out space before/after paragraphs
    for tag in ("a:spcBef", "a:spcAft"):
        spc = etree.SubElement(pPr, qn(tag))
        pts = etree.SubElement(spc, qn("a:spcPts"))
        pts.set("val", "0")


def _add_textbox(
    slide,
    lines: list[str],
    left: float, top: float, width: float, height: float,
    font_name: str, font_size: int, bold: bool,
    color_rgb: tuple, align,
    line_spacing: float,
    first_line_prefix: str = "",
    prefix_color_rgb: tuple | None = None,
):
    """
    Add a text box. If first_line_prefix is set (e.g. verse number), it is
    prepended as a separate run on the first paragraph with prefix_color_rgb.
    """
    txBox = slide.shapes.add_textbox(int(left), int(top), int(width), int(height))
    tf = txBox.text_frame
    tf.word_wrap = True

    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        _set_line_spacing(p, line_spacing)

        # Verse number prefix — only on first line, coloured differently
        if i == 0 and first_line_prefix:
            pre_run = p.add_run()
            pre_run.text = first_line_prefix + "  "   # two spaces as separator
            pre_run.font.size = Pt(font_size)
            pre_run.font.bold = True
            pre_run.font.color.rgb = RGBColor(*(prefix_color_rgb or color_rgb))
            _fix_font(pre_run, font_name)

        run = p.add_run()
        run.text = line
        run.font.name = font_name
        run.font.size = Pt(font_size)
        run.font.bold = bold
        run.font.color.rgb = RGBColor(*color_rgb)
        _fix_font(run, font_name)


def build_pptx(slides_data: list[SlidePart], s: SlideSettings) -> bytes:
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)

    align     = ALIGN_MAP.get(s.align, PP_ALIGN.CENTER)
    text_rgb  = _hex_to_rgb(s.text_color)
    ref_rgb   = _hex_to_rgb(s.ref_color)
    vnum_rgb  = _hex_to_rgb(s.verse_num_color)
    bg_rgb    = _hex_to_rgb(s.bg_color)

    pad = Inches(s.padding)
    sw  = prs.slide_width
    sh  = prs.slide_height

    for sd in slides_data:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        bg_fill = slide.background.fill
        bg_fill.solid()
        bg_fill.fore_color.rgb = RGBColor(*bg_rgb)

        # Body text — verse number is prepended inline on the first line
        body_w = sw - pad * 2
        body_h = sh - pad * 2 - Inches(0.85)
        _add_textbox(
            slide, sd.lines,
            pad, pad, body_w, body_h,
            s.font_family, s.font_size, s.bold,
            text_rgb, align, s.line_spacing,
            first_line_prefix=sd.verse_num,
            prefix_color_rgb=vnum_rgb,
        )

        # Reference (bottom-right)
        if sd.ref:
            _add_textbox(
                slide, [sd.ref],
                pad, sh - Inches(0.75),
                sw - pad * 2, Inches(0.6),
                s.font_family, s.ref_font_size, False,
                ref_rgb, PP_ALIGN.RIGHT, 1.0,
            )

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────────
# Routes
# ─────────────────────────────────────────────

@app.get("/api/health")
def health():
    return {"status": "ok"}


@app.post("/api/preview")
def preview(req: PreviewRequest):
    verses = parse_verses(req.raw_text)
    slides = build_slides(
        verses,
        req.per_slide,
        req.reference,
        req.settings.show_ref_each,
        settings=req.settings,
    )
    return {"slides": [s.model_dump() for s in slides], "total": len(slides)}


@app.post("/api/generate")
def generate(req: GenerateRequest):
    verses = parse_verses(req.raw_text)
    if not verses:
        verses = [{"ref": req.reference or "", "verse_num": "", "lines": ["(No content)"]}]
    slides = build_slides(
        verses,
        req.per_slide,
        req.reference,
        req.settings.show_ref_each,
        settings=req.settings,
    )
    pptx_bytes = build_pptx(slides, req.settings)
    filename = f"bible_slides_{uuid.uuid4().hex[:8]}.pptx"
    return StreamingResponse(
        io.BytesIO(pptx_bytes),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


# ═══════════════════════════════════════════════════════════════════
#  Song / OCR Slide Generator
# ═══════════════════════════════════════════════════════════════════

# ── Fonts ────────────────────────────────────────────────────────
FONT_DIR = Path(__file__).parent / "fonts"
FONT_FILE_MAP: dict[str, str] = {
    "Khmer OS Battambang": "KhmerOSBattambang-Regular.ttf",
    "Khmer OS":            "KhmerOS_.ttf",
    "Khmer OS Muol":       "KhmerOS_muol.ttf",
    "Khmer OS Moul Light": "KhmerOSMoulLight.ttf",
}

# ── Text-filtering helpers ───────────────────────────────────────
_HAS_KHMER_RE = re.compile(r'[\u1780-\u17FF\u19E0-\u19FF]')

# Khmer Unicode ranges
_KHMER_CHAR_RE = re.compile(r'[\u1780-\u17FF\u19E0-\u19FF]')

# Chord token (letter root + quality + optional slash bass)
_CHORD_TOKEN_RE = re.compile(
    r'^[A-G][#b]?(?:m|maj|min|aug|dim|sus|add)?(?:\d+)?(?:/[A-G][#b]?)?$',
    re.IGNORECASE,
)


def _line_is_khmer(s: str) -> bool:
    return bool(_HAS_KHMER_RE.search(s))


def _khmer_char_count(s: str) -> int:
    return len(_KHMER_CHAR_RE.findall(s))


def _is_repetitive(s: str, thr: float = 0.45) -> bool:
    """
    Return True if top-1 Khmer char >= thr fraction, OR top-3 chars >= 72 %
    of all Khmer chars.  Catches both per-char smear and solfege notation.
    """
    khmer_chars = _KHMER_CHAR_RE.findall(s)
    if len(khmer_chars) < 4:
        return False
    counts = Counter(khmer_chars)
    top1_frac = counts.most_common(1)[0][1] / len(khmer_chars)
    if top1_frac >= thr:
        return True
    top3_total = sum(c for _, c in counts.most_common(3))
    if top3_total / len(khmer_chars) >= 0.72:
        return True
    return False


def _has_repeated_bigram(tokens: list, max_frac: float = 0.40) -> bool:
    """Return True if any consecutive 2-token pair repeats >= max_frac of all bigrams."""
    if len(tokens) < 4:
        return False
    bigrams = [(tokens[i], tokens[i + 1]) for i in range(len(tokens) - 1)]
    top_count = Counter(bigrams).most_common(1)[0][1]
    return top_count / len(bigrams) >= max_frac


# OCR noise characters to strip from lyric lines before filtering
# These are common misreads: slashes, angles, brackets, punctuation artifacts
_OCR_NOISE_RE = re.compile(r'[/\\<>«»›‹|(){}\[\]#@&*^~`]+')

# Token-level smear detection regexes (compiled once at module level)
_LATIN_RE      = re.compile(r'[a-zA-Z]')
_REPEAT_SUB_RE = re.compile(r'(.{2,3})\1{2,}')  # e.g. "ោយោយោយ"


def _clean_line(s: str) -> str:
    """
    Strip common OCR noise characters from a lyric line while preserving
    Khmer text and natural punctuation (dash, comma, period, space).
    Also collapses multiple spaces.
    """
    s = _OCR_NOISE_RE.sub(' ', s)
    s = re.sub(r' {2,}', ' ', s)
    return s.strip()


def _is_lyric_line(s: str) -> bool:
    """
    Return True only if this line looks like a genuine Khmer lyric line.

    Rejection criteria (any one -> reject):
      1. No Khmer characters at all.
      2. Fewer than 6 Khmer characters.
      3. Top-1 Khmer char >= 45 % (repetitive OCR smear).
      4. Top-3 Khmer chars >= 72 % (solfege / rhythm notation).
      5. Fewer than 2 whitespace-separated Khmer word tokens.
      6. 2-3 tokens AND avg token length < 4.5 chars
         (short instrument-label / tab marker).
      7. Average Khmer token length < 2.5 chars (pure notation markers).
      8. Two consecutive identical Khmer tokens (rhythm: "i i", "pi pi").
      9. Repeated bigrams across the line (rhythm notation rows).
     10. >= 60 % of tokens end with Khmer vowel sign ii (\u17B8 = \u17b8 = i-vowel)
         -> solfege syllable row (pii, kii, rii, sii...).
     11. >= 50 % of tokens are exactly 2 chars (solfege syllables).
     12. Khmer chars < 55 % of all non-space chars (mixed notation/ascii).
    """
    if not _line_is_khmer(s):
        return False

    k_count = _khmer_char_count(s)
    if k_count < 6:
        return False

    if _is_repetitive(s):
        return False

    khmer_word_tokens = [t for t in s.split() if _line_is_khmer(t)]
    n_tokens = len(khmer_word_tokens)

    if n_tokens < 2:
        return False

    avg_word_len = sum(len(t) for t in khmer_word_tokens) / n_tokens

    # Short instrument-label gate: <= 3 tokens with short avg length
    if n_tokens <= 3 and avg_word_len < 4.5:
        return False

    # General avg-length floor
    if avg_word_len < 2.5:
        return False

    # Consecutive identical tokens
    for i in range(n_tokens - 1):
        if khmer_word_tokens[i] == khmer_word_tokens[i + 1]:
            return False

    # Repeated bigrams
    if _has_repeated_bigram(khmer_word_tokens):
        return False

    # Solfege gate 1: >= 60% of tokens end with ii-vowel sign (U+17B8)
    # Real Khmer words rarely end with ii; solfege syllables always do
    ii_vowel = '\u17B8'
    ii_end_frac = sum(1 for t in khmer_word_tokens if ii_vowel in t[-2:]) / n_tokens
    if ii_end_frac >= 0.60:
        return False

    # Solfege gate 2: >= 50% of tokens are exactly 2 chars
    short2_frac = sum(1 for t in khmer_word_tokens if len(t) == 2) / n_tokens
    if short2_frac >= 0.50:
        return False

    # Token-level smear gates: reject if ANY token looks like notation smear
    for tok in khmer_word_tokens:
        # Mixed Khmer+Latin in one token -> OCR smear of staff lines
        if _LATIN_RE.search(tok) and _KHMER_CHAR_RE.search(tok):
            return False
        # Repeated 2–3 char substring pattern -> notation run (e.g. ោយោយោយ)
        if len(tok) >= 6 and _REPEAT_SUB_RE.search(tok):
            return False
        # Long token with highly repetitive character bigrams -> notation run
        if len(tok) >= 7:
            chars = list(tok)
            bigrams = [chars[i] + chars[i+1] for i in range(len(chars)-1)]
            if bigrams:
                top_bg_count = Counter(bigrams).most_common(1)[0][1]
                if top_bg_count / len(bigrams) >= 0.22:
                    return False

    # Khmer density gate
    total_chars = len(s.replace(' ', ''))
    if total_chars > 0 and k_count / total_chars < 0.55:
        return False

    return True


def filter_lyrics(raw: str) -> str:
    """
    Pass 1: for each line, strip OCR noise chars then test _is_lyric_line().
    Pass 2: deduplicate adjacent identical lines (repeated verse detection).
    Blank lines are preserved as verse separators.
    """
    kept: list[str] = []
    prev_blank = True
    prev_line = None
    for line in raw.splitlines():
        s = line.strip()
        if not s:
            if not prev_blank:
                kept.append("")
            prev_blank = True
            prev_line = None
            continue
        prev_blank = False
        # Clean OCR noise first, then test
        cleaned = _clean_line(s)
        if _is_lyric_line(cleaned):
            if cleaned != prev_line:   # skip exact duplicate consecutive lines
                kept.append(cleaned)
                prev_line = cleaned
    while kept and kept[-1] == "":
        kept.pop()
    return "\n".join(kept)


# ── Numpy / PIL imports ─────────────────────────────────────────
import numpy as np
from PIL import Image, ImageOps, ImageFilter


def _find_tesseract() -> str:
    import shutil
    for candidate in (
        "/opt/homebrew/bin/tesseract",
        "/usr/local/bin/tesseract",
        "/usr/bin/tesseract",
    ):
        if os.path.exists(candidate):
            return candidate
    return shutil.which("tesseract") or "tesseract"


_TESSERACT_CMD: Optional[str] = None


def _get_tesseract() -> str:
    global _TESSERACT_CMD
    if _TESSERACT_CMD is None:
        _TESSERACT_CMD = _find_tesseract()
    return _TESSERACT_CMD


# ── Image pre-processing ────────────────────────────────────────

def _preprocess_for_ocr(img: Image.Image, target_width: int = 2800) -> Image.Image:
    """
    Prepare a lyric-strip image for Tesseract (Khmer-optimised):

      1. Greyscale
      2. Upscale so the longest dimension ≥ target_width px
         (Khmer vowel signs need glyph height ≥ 50 px to be read correctly)
      3. Auto-contrast (stretch histogram to full 0-255 range)
      4. Unsharp-mask (crisper edges than double-SHARPEN, doesn't over-thicken)
      5. Adaptive binarise via numpy:
           - Compute local mean in a 41-px window
           - Threshold = local_mean - 18  (handles uneven lighting / shadows)
           - This preserves faint vowel signs that hard-threshold at 150 erases
      6. Light despeckle: remove isolated dark pixels smaller than 2×2
      7. White border (Tesseract hates glyphs at the very edge)
    """
    img = img.convert("L")
    w, h = img.size

    # Step 2 — upscale
    if w < target_width:
        scale = target_width / w
        new_w = int(w * scale)
        new_h = max(int(h * scale), 80)   # ensure min height for vowel signs
        img = img.resize((new_w, new_h), Image.LANCZOS)

    # Step 3 — auto-contrast
    img = ImageOps.autocontrast(img, cutoff=2)

    # Step 4 — unsharp mask (radius 2, percent 150, threshold 3)
    img = img.filter(ImageFilter.UnsharpMask(radius=2, percent=150, threshold=3))

    # Step 5 — adaptive binarise
    arr = np.array(img, dtype=np.float32)
    from scipy.ndimage import uniform_filter
    try:
        local_mean = uniform_filter(arr, size=41)
        binary = (arr > (local_mean - 18)).astype(np.uint8) * 255
    except ImportError:
        # scipy not available → fall back to simple threshold at 140
        binary = (arr > 140).astype(np.uint8) * 255
    img = Image.fromarray(binary.astype(np.uint8), mode="L")

    # Step 6 — light despeckle: erode then dilate (opening) to remove 1-px noise
    img = img.filter(ImageFilter.MinFilter(3))   # erode (dark = ink)
    img = img.filter(ImageFilter.MaxFilter(3))   # dilate back

    # Step 7 — white border
    img = ImageOps.expand(img, border=20, fill=255)
    return img


# ── Staff-line detection ────────────────────────────────────────

def _row_dark_density(gray: np.ndarray, thr: int = 180) -> np.ndarray:
    """Fraction of pixels darker than thr in each row."""
    return (gray < thr).mean(axis=1).astype(float)


def _moving_avg(arr: np.ndarray, w: int) -> np.ndarray:
    if w < 2:
        return arr.copy()
    return np.convolve(arr, np.ones(w) / w, mode="same")


def _gap_spans(mask: np.ndarray, min_px: int) -> list[tuple[int, int]]:
    spans, in_g, g0 = [], False, 0
    for y, v in enumerate(mask):
        if v and not in_g:
            in_g, g0 = True, y
        elif not v and in_g:
            in_g = False
            if y - g0 >= min_px:
                spans.append((g0, y))
    if in_g and len(mask) - g0 >= min_px:
        spans.append((g0, len(mask)))
    return spans


def _row_horizontal_runs(gray_row: np.ndarray, thr: int = 180) -> float:
    """
    Return the mean length of dark horizontal runs in a single row.
    Staff lines have very long uniform runs (spanning most of the page width).
    Lyric rows have short, irregular runs (individual glyph strokes).
    """
    dark = gray_row < thr
    runs, in_run, run_len = [], False, 0
    for v in dark:
        if v:
            in_run = True
            run_len += 1
        elif in_run:
            runs.append(run_len)
            in_run = False
            run_len = 0
    if in_run:
        runs.append(run_len)
    return float(np.mean(runs)) if runs else 0.0


def _is_staff_strip(gray: np.ndarray, top: int, bot: int,
                    page_width: int) -> bool:
    """
    Decide whether the horizontal band gray[top:bot] is a music notation
    (staff) strip using TWO independent signals:

    Signal 1 — Long horizontal runs:
      Staff lines are thin, page-spanning dark lines.
      A strip with ≥ 3 rows where mean run length > 15 % of page width
      is almost certainly staff+notation.

    Signal 2 — Row-density profile:
      Staff rows have moderate but very uniform density across the height.
      The standard deviation of per-row densities is small for staff bands.

    Either signal alone is sufficient to label a strip as notation.
    """
    band       = gray[top:bot]
    h_band     = band.shape[0]
    if h_band < 3:
        return False

    run_threshold  = page_width * 0.12   # run > 12 % of page width → staff line
    long_run_rows  = 0
    row_densities  = []

    for row_idx in range(h_band):
        row = band[row_idx]
        mean_run = _row_horizontal_runs(row, thr=180)
        if mean_run > run_threshold:
            long_run_rows += 1
        row_densities.append(float((row < 180).mean()))

    # Signal 1: multiple rows with very long horizontal runs
    if long_run_rows >= max(2, h_band * 0.15):
        return True

    # Signal 2: high overall density (lots of ink — notes + beams) AND
    #           uniform density variance (regular pattern of staff lines)
    arr_d = np.array(row_densities)
    mean_d = float(arr_d.mean())
    std_d  = float(arr_d.std())
    if mean_d > 0.08 and std_d < 0.06:
        return True

    return False


def _detect_content_strips(gray: np.ndarray,
                            blank_thresh: float = 0.008,
                            min_gap_px:   int   = 6,
                            min_strip_px: int   = 15,
                            smooth_w:     int   = 3,
                            ) -> list[tuple[int, int]]:
    """
    Split a greyscale page into horizontal content strips by finding
    near-blank row gaps.  Uses a narrow smoothing window so stave/lyric
    boundaries stay sharp.
    """
    density = _moving_avg(_row_dark_density(gray, thr=180), w=smooth_w)
    blank   = density < blank_thresh
    strips, prev = [], 0
    for g0, g1 in _gap_spans(blank, min_gap_px):
        if g0 - prev >= min_strip_px:
            strips.append((prev, g0))
        prev = g1
    h = gray.shape[0]
    if h - prev >= min_strip_px:
        strips.append((prev, h))
    return strips


def _extract_lyric_strips(page_rgb: Image.Image) -> list[Image.Image]:
    """
    Return cropped PIL images of ONLY the lyric (Khmer text) rows from
    a music sheet page, discarding all notation/staff rows.

    Algorithm:
      1. Detect all horizontal content strips (separated by white gaps).
      2. For each strip, run _is_staff_strip() to detect notation rows
         using long-horizontal-run analysis.
      3. Any strip NOT identified as staff is a candidate lyric row.
      4. Safety: if zero lyric strips found (flat/photocopied page),
         return [] so the caller falls back to full-page OCR + text filter.
    """
    gray  = np.array(page_rgb.convert("L"))
    h, w  = gray.shape

    strips = _detect_content_strips(gray)
    if len(strips) < 2:
        return []

    lyric_crops: list[Image.Image] = []
    for top, bot in strips:
        if _is_staff_strip(gray, top, bot, page_width=w):
            continue   # skip notation / staff row entirely

        # This strip is NOT staff → it's lyrics (or chord symbols which
        # we'll filter at the text level after OCR).
        # Use generous vertical padding so vowel signs above/below baseline
        # are never clipped (Khmer stacks go well above and below the base row).
        pad = 24
        t = max(0, top - pad)
        b = min(h, bot + pad)
        lyric_crops.append(page_rgb.crop((0, t, w, b)))

    return lyric_crops


# ── OCR helpers ──────────────────────────────────────────────────

# Common Tesseract Khmer OCR mis-substitutions — map wrong → correct codepoint
# These are single-character swaps confirmed across many Khmer OCR test runs.
_KHMER_FIXES: list[tuple[str, str]] = [
    # Visually similar consonant pairs that Tesseract confuses
    ('\u1780', '\u1781'),   # ក ↔ ខ  (ka / kha) — only swap if context demands;
                            # actually we leave consonants alone (too risky without context)
    # Vowel signs commonly swapped by Tesseract
    ('\u17BE', '\u17C1'),   # ើ → េ  (oe → e) — Tesseract over-uses ើ
    # Subscript coeng (​្) sometimes dropped → we can't restore it, but
    # we CAN normalise Unicode composition
]

# Characters Tesseract inserts that are not valid Khmer (e.g. Cyrillic leaked
# from the model, Latin letters surrounded by Khmer)
_STRIP_NON_KHMER_IN_WORD_RE = re.compile(
    r'(?<=[\u1780-\u17FF\u19E0-\u19FF])[a-zA-Z\u0400-\u04FF]+'
    r'(?=[\u1780-\u17FF\u19E0-\u19FF])'
)


def _fix_ocr_text(text: str) -> str:
    """
    Post-process raw Tesseract output to fix the most common Khmer OCR errors:
      1. Unicode normalisation (NFC) — ensures combining marks are properly attached
      2. Remove Cyrillic/Latin characters that sneak in between Khmer characters
      3. Strip lines that are pure ASCII/numbers after Khmer-only filtering
      4. Collapse multiple spaces
    """
    import unicodedata
    text = unicodedata.normalize("NFC", text)
    # Remove stray non-Khmer chars embedded inside Khmer words
    text = _STRIP_NON_KHMER_IN_WORD_RE.sub('', text)
    # Collapse spaces
    text = re.sub(r' {2,}', ' ', text)
    return text


def _khmer_char_count_in(text: str) -> int:
    """Count total Khmer characters in a multi-line string."""
    return len(_KHMER_CHAR_RE.findall(text))


def _ocr_attempt(processed: 'Image.Image', lang: str, psm: str,
                 tessdata_dir: str | None = None) -> str:
    """Run one Tesseract attempt; return raw text or ''."""
    try:
        import pytesseract
    except ImportError:
        return ""
    cfg = f"--psm {psm} --oem 1"
    if tessdata_dir:
        cfg += f" --tessdata-dir {tessdata_dir}"
    try:
        return pytesseract.image_to_string(processed, lang=lang, config=cfg)
    except Exception:
        return ""


def _ocr_strip(img: Image.Image) -> str:
    """
    OCR a single lyric strip with multiple strategy attempts.
    Picks the result with the most Khmer characters (best signal).

    Strategies tried (in parallel order, best wins):
      A. lang=khm,       psm=7  (single line  — narrow strips)
      B. lang=khm,       psm=6  (uniform block — taller strips)
      C. lang=script/Khmer, psm=6  (script-level model, better glyph accuracy)
      D. lang=khm+eng,   psm=6  (allow some Latin — chord row guard)

    After picking the winner:
      • Keep only lines containing Khmer characters
      • Apply _fix_ocr_text() Unicode normalisation + noise removal
    """
    try:
        import pytesseract
    except ImportError:
        return ""

    pytesseract.pytesseract.tesseract_cmd = _get_tesseract()
    processed = _preprocess_for_ocr(img)
    _, ph = processed.size
    psm_primary = "7" if ph < 100 else "6"

    # Build script/Khmer tessdata path if available
    tess_prefix = os.path.dirname(os.path.dirname(_get_tesseract()))  # e.g. /opt/homebrew
    script_dir  = os.path.join(tess_prefix, "share", "tessdata")
    script_lang = "script/Khmer" if os.path.exists(
        os.path.join(script_dir, "script", "Khmer.traineddata")
    ) else None

    candidates: list[str] = []

    # Strategy A/B: khm with both psm modes
    for psm in (psm_primary, "6" if psm_primary == "7" else "7"):
        raw = _ocr_attempt(processed, "khm", psm)
        if raw:
            candidates.append(raw)

    # Strategy C: script/Khmer model (often better at individual glyph shapes)
    if script_lang:
        raw = _ocr_attempt(processed, script_lang, "6", tessdata_dir=script_dir)
        if raw:
            candidates.append(raw)

    # Strategy D: khm+eng fallback
    raw = _ocr_attempt(processed, "khm+eng", "6")
    if raw:
        candidates.append(raw)

    if not candidates:
        return ""

    # Pick the candidate with the most Khmer characters
    best = max(candidates, key=_khmer_char_count_in)

    # Post-process: keep Khmer-containing lines, fix encoding, normalise
    lines = [
        _fix_ocr_text(ln.strip())
        for ln in best.splitlines()
        if ln.strip() and _line_is_khmer(ln)
    ]
    return "\n".join(lines)


def _ocr_full_page_khmer(page_rgb: Image.Image) -> str:
    """
    Fallback: OCR the whole page with auto-layout detection (psm 3).
    Tries both khm and script/Khmer models; picks best (most Khmer chars).
    """
    try:
        import pytesseract
    except ImportError:
        return ""

    pytesseract.pytesseract.tesseract_cmd = _get_tesseract()
    processed = _preprocess_for_ocr(page_rgb)

    tess_prefix = os.path.dirname(os.path.dirname(_get_tesseract()))
    script_dir  = os.path.join(tess_prefix, "share", "tessdata")
    script_lang = "script/Khmer" if os.path.exists(
        os.path.join(script_dir, "script", "Khmer.traineddata")
    ) else None

    candidates: list[str] = []
    for lang in ("khm", "khm+eng"):
        raw = _ocr_attempt(processed, lang, "3")
        if raw:
            candidates.append(raw)
    if script_lang:
        raw = _ocr_attempt(processed, script_lang, "3", tessdata_dir=script_dir)
        if raw:
            candidates.append(raw)

    if not candidates:
        return ""

    best = max(candidates, key=_khmer_char_count_in)
    lines = [
        _fix_ocr_text(ln.strip())
        for ln in best.splitlines()
        if ln.strip() and _line_is_khmer(ln)
    ]
    return "\n".join(lines)


# ── Main extraction entry point ─────────────────────────────────

def extract_lyrics_from_bytes(data: bytes, ext: str) -> str:
    """
    Accept raw file bytes + extension (.pdf / .png / .jpg / .jpeg).
    Returns ONLY the Khmer lyric lines, with blank lines as verse separators.

    Pipeline per page:
      1. Convert PDF → RGB at 300 dpi (larger = better OCR)
      2. Detect & isolate lyric strips by discarding staff/notation rows
      3. OCR each lyric strip with lang=khm, keep only Khmer lines
      4. If strip isolation yields nothing, OCR full page and filter to Khmer
      5. Run filter_lyrics() as a final pass to ensure no garbage remains
    """
    all_lines: list[str] = []

    def process_page(page_rgb: Image.Image) -> None:
        strips = _extract_lyric_strips(page_rgb)

        if strips:
            for strip_img in strips:
                text = _ocr_strip(strip_img)
                lines = [ln for ln in text.splitlines() if ln.strip()]
                if lines:
                    all_lines.extend(lines)
                    all_lines.append("")   # verse separator after each strip
        else:
            # Fallback: full-page OCR filtered to Khmer only
            text = _ocr_full_page_khmer(page_rgb)
            lines = [ln for ln in text.splitlines() if ln.strip()]
            all_lines.extend(lines)

    if ext == ".pdf":
        try:
            from pdf2image import convert_from_bytes
            pages = convert_from_bytes(data, dpi=300)
        except Exception as e:
            return f"[PDF conversion error: {e}]"
        for page in pages:
            process_page(page.convert("RGB"))
            all_lines.append("")   # separator between PDF pages
    else:
        try:
            img = Image.open(io.BytesIO(data)).convert("RGB")
            process_page(img)
        except Exception as e:
            return f"[Image error: {e}]"

    if not all_lines or all(not ln for ln in all_lines):
        return "(no Khmer text detected)"

    # Final filter: keep only Khmer lines, preserve blank separators
    raw      = "\n".join(all_lines)
    filtered = filter_lyrics(raw)
    return filtered.strip() if filtered.strip() else "(no Khmer text detected)"


# ── Song-slide PPTX builder ──────────────────────────────────────

class SongSettings(BaseModel):
    font_family:    str   = "Khmer OS Battambang"
    font_size:      int   = Field(56, ge=8, le=120)
    bold:           bool  = True
    align:          str   = "center"
    line_spacing:   float = Field(1.5, ge=0.5, le=4.0)
    verse_spacing:  float = Field(1.0, ge=0.0, le=3.0)   # extra space between verse groups
    bg_color:       str   = "#ffffff"
    text_color:     str   = "#000000"
    ref_color:      str   = "#555555"
    ref_font_size:  int   = Field(22, ge=8, le=72)
    verse_num_color: str  = "#ffcc44"
    chorus_r_color:  str  = "#ef4444"
    padding:        float = Field(0.45, ge=0.05, le=2.0)
    lines_per_slide: int  = Field(2, ge=1, le=10)


class SongGenerateRequest(BaseModel):
    lyrics_text: str
    settings:    SongSettings = SongSettings()
    song_title:  str = ""


class SongPreviewRequest(BaseModel):
    lyrics_text: str
    settings:    SongSettings = SongSettings()
    song_title:  str = ""


def parse_song_lyrics(raw_text: str) -> list[dict]:
    """
    Parses hymn / song lyrics into structured slides:
    1. Title Slide: title, subtitle (e.g. Korean / English), category (e.g. HYMN)
    2. Verse Slides: numbered 1, 2, 3... (converted to Khmer numbers: ១, ២, ៣)
    3. Chorus Slides: 'បន្ទរ' / 'Chorus' replaced by 'R.' in red color.
    """
    raw_text = raw_text.strip()
    if not raw_text:
        return []

    lines = raw_text.splitlines()

    # Detect verse markers: e.g. '១~', '1~', '១.', '1.', '2.', '២.'
    VERSE_START_RE = re.compile(r'(?:^|\s+)([0-9០-៩]+)\s*[\~\.\:\-\)]')
    CHORUS_RE = re.compile(r'(?:បន្ទរ\s*[\:\៖]?|Chorus\s*[\:\៖]?)')

    has_verse_marker = any(VERSE_START_RE.search(l) for l in lines)
    has_chorus_marker = any(CHORUS_RE.search(l) for l in lines)

    # Hymn structure with markers
    if has_verse_marker or has_chorus_marker:
        body_start_line = len(lines)
        for i, line in enumerate(lines):
            if VERSE_START_RE.search(line.strip()):
                body_start_line = i
                break

        header_lines = [l.strip() for l in lines[:body_start_line] if l.strip()]
        body_text = ' '.join(lines[body_start_line:]).strip()

        slides = []

        # 1. Title slide
        if header_lines:
            title = header_lines[0]
            subtitle = header_lines[1] if len(header_lines) > 1 else ''
            category = ' '.join(header_lines[2:]) if len(header_lines) > 2 else ''
            slides.append({
                'is_title': True,
                'title': title,
                'subtitle': subtitle,
                'category': category,
                'prefix': '',
                'prefix_color': '',
                'lines': [title] + ([subtitle] if subtitle else []) + ([category] if category else []),
            })

        if body_text:
            MARKER_RE = re.compile(
                r'(?:(?:^|\s+)(?P<verse>[0-9០-៩]+)\s*[\~\.\:\-\)]\s*)|(?:(?:^|\s+)(?P<chorus>បន្ទរ\s*[\:\៖]?|Chorus\s*[\:\៖]?)\s*)'
            )
            matches = list(MARKER_RE.finditer(body_text))
            if matches:
                for idx, m in enumerate(matches):
                    start = m.end()
                    end = matches[idx + 1].start() if idx + 1 < len(matches) else len(body_text)
                    content = body_text[start:end].strip()
                    if not content:
                        continue
                    if m.group('verse'):
                        v_num = to_khmer_number(m.group('verse'))
                        slides.append({
                            'is_title': False,
                            'title': '',
                            'subtitle': '',
                            'category': '',
                            'prefix': v_num,
                            'prefix_color': 'verse',
                            'lines': [content],
                        })
                    elif m.group('chorus'):
                        slides.append({
                            'is_title': False,
                            'title': '',
                            'subtitle': '',
                            'category': '',
                            'prefix': 'R.',
                            'prefix_color': 'red',
                            'lines': [content],
                        })
            else:
                slides.append({
                    'is_title': False,
                    'title': '',
                    'subtitle': '',
                    'category': '',
                    'prefix': '',
                    'prefix_color': '',
                    'lines': [body_text],
                })
        return slides

    # Fallback: plain lyrics
    slides = []
    blocks = re.split(r'\n\s*\n', raw_text)
    for block in blocks:
        b_lines = [l.strip() for l in block.splitlines() if l.strip()]
        if b_lines:
            slides.append({
                'is_title': False,
                'title': '',
                'subtitle': '',
                'category': '',
                'prefix': '',
                'prefix_color': '',
                'lines': b_lines,
            })
    return slides


def _build_song_pptx(lyrics_text: str, s: SongSettings, song_title: str = "") -> bytes:
    prs = Presentation()
    prs.slide_width  = Inches(13.33)
    prs.slide_height = Inches(7.5)

    align    = ALIGN_MAP.get(s.align, PP_ALIGN.CENTER)
    text_rgb = _hex_to_rgb(s.text_color)
    ref_rgb  = _hex_to_rgb(s.ref_color)
    bg_rgb   = _hex_to_rgb(s.bg_color)
    red_rgb  = (239, 68, 68)  # Red for chorus 'R.'
    vnum_rgb = _hex_to_rgb(getattr(s, 'verse_num_color', '#ffcc44'))

    pad = Inches(s.padding)
    sw  = prs.slide_width
    sh  = prs.slide_height

    slides_data = parse_song_lyrics(lyrics_text)
    if not slides_data:
        slides_data = [{'is_title': False, 'prefix': '', 'prefix_color': '', 'lines': ['(no lyrics found)']}]

    if not song_title and slides_data and slides_data[0].get('is_title'):
        song_title = slides_data[0].get('title', '')

    for s_idx, sd in enumerate(slides_data):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        bg_fill = slide.background.fill
        bg_fill.solid()
        bg_fill.fore_color.rgb = RGBColor(*bg_rgb)

        if sd.get('is_title'):
            # ── Beautiful Centered Title Slide ──
            body_w = sw - pad * 2
            body_h = Inches(4.5)
            top_pos = (sh - body_h) / 2
            txBox = slide.shapes.add_textbox(int(pad), int(top_pos), int(body_w), int(body_h))
            tf = txBox.text_frame
            tf.word_wrap = True

            # Title
            p0 = tf.paragraphs[0]
            p0.alignment = PP_ALIGN.CENTER
            _set_line_spacing(p0, 1.2)
            run_title = p0.add_run()
            run_title.text = sd.get('title', '')
            run_title.font.size = Pt(min(s.font_size + 6, 72))
            run_title.font.bold = True
            run_title.font.color.rgb = RGBColor(*text_rgb)
            _fix_font(run_title, s.font_family)

            # Subtitle (Korean / English / etc.)
            if sd.get('subtitle'):
                p1 = tf.add_paragraph()
                p1.alignment = PP_ALIGN.CENTER
                _set_line_spacing(p1, 1.2)
                run_sub = p1.add_run()
                run_sub.text = sd.get('subtitle', '')
                run_sub.font.size = Pt(max(s.font_size - 14, 20))
                run_sub.font.color.rgb = RGBColor(*ref_rgb)
                _fix_font(run_sub, s.font_family)

            # Category / tag (e.g. HYMN)
            if sd.get('category'):
                p2 = tf.add_paragraph()
                p2.alignment = PP_ALIGN.CENTER
                _set_line_spacing(p2, 1.2)
                run_cat = p2.add_run()
                run_cat.text = sd.get('category', '')
                run_cat.font.size = Pt(max(s.font_size - 22, 16))
                run_cat.font.bold = True
                run_cat.font.color.rgb = RGBColor(91, 80, 232)
                _fix_font(run_cat, s.font_family)
            continue

        # ── Verse or Chorus Slide ──
        body_h = sh - pad * 2 - (Inches(0.65) if song_title else Inches(0.1))
        txBox  = slide.shapes.add_textbox(int(pad), int(pad), int(sw - pad * 2), int(body_h))
        tf     = txBox.text_frame
        tf.word_wrap = True

        prefix = sd.get('prefix', '')
        prefix_color = sd.get('prefix_color', '')

        for i, line in enumerate(sd.get('lines', [])):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = align
            _set_line_spacing(p, s.line_spacing)

            # Prefix run (only on first line)
            if i == 0 and prefix:
                pre_run = p.add_run()
                pre_run.text = prefix + "  "
                pre_run.font.size = Pt(s.font_size)
                pre_run.font.bold = True
                # Red for chorus 'R.', verse_num_color for verses
                if prefix == 'R.' or prefix_color == 'red':
                    pre_run.font.color.rgb = RGBColor(*red_rgb)
                else:
                    pre_run.font.color.rgb = RGBColor(*vnum_rgb)
                _fix_font(pre_run, s.font_family)

            run = p.add_run()
            run.text = line
            run.font.size = Pt(s.font_size)
            run.font.bold = s.bold
            run.font.color.rgb = RGBColor(*text_rgb)
            _fix_font(run, s.font_family)

        # Song title / reference at bottom-right
        if song_title and s_idx > 0:
            _add_textbox(
                slide, [song_title],
                pad, sh - Inches(0.65),
                sw - pad * 2, Inches(0.55),
                s.font_family, s.ref_font_size, False,
                ref_rgb, PP_ALIGN.RIGHT, 1.0,
            )

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)

    # Embed Khmer font if applicable
    pptx_bytes = buf.read()
    ttf_filename = FONT_FILE_MAP.get(s.font_family)
    if ttf_filename:
        ttf_path = FONT_DIR / ttf_filename
        if ttf_path.exists():
            pptx_bytes = _embed_font(pptx_bytes, s.font_family, ttf_filename, ttf_path.read_bytes())

    return pptx_bytes
    if ttf_filename:
        ttf_path = FONT_DIR / ttf_filename
        if ttf_path.exists():
            pptx_bytes = _embed_font(pptx_bytes, s.font_family, ttf_filename, ttf_path.read_bytes())

    return pptx_bytes


def _embed_font(pptx_bytes: bytes, font_name: str, ttf_filename: str, font_data: bytes) -> bytes:
    """Inject a TTF into the PPTX zip as an embedded font."""
    p_ns  = "http://schemas.openxmlformats.org/presentationml/2006/main"
    r_ns  = "http://schemas.openxmlformats.org/package/2006/relationships"
    r2_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
    font_rel_type  = f"{r2_ns}/font"
    font_zip_path  = f"ppt/fonts/{ttf_filename}"
    rels_path      = "ppt/_rels/presentation.xml.rels"
    prs_path_zip   = "ppt/presentation.xml"
    ct_path        = "[Content_Types].xml"

    src = io.BytesIO(pptx_bytes)
    dst = io.BytesIO()

    with zipfile.ZipFile(src, "r") as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        # Patch rels
        rels_root = etree.fromstring(zin.read(rels_path))
        existing  = {el.get("Id") for el in rels_root}
        rid = "rFnt1"
        n   = 1
        while rid in existing:
            n  += 1
            rid = f"rFnt{n}"
        new_rel = etree.SubElement(rels_root, f"{{{r_ns}}}Relationship")
        new_rel.set("Id",     rid)
        new_rel.set("Type",   font_rel_type)
        new_rel.set("Target", f"fonts/{ttf_filename}")
        patched_rels = etree.tostring(rels_root, xml_declaration=True, encoding="UTF-8", standalone=True)

        # Patch presentation.xml
        prs_root = etree.fromstring(zin.read(prs_path_zip))
        lst = prs_root.find(f"{{{p_ns}}}embeddedFontLst")
        if lst is None:
            lst = etree.SubElement(prs_root, f"{{{p_ns}}}embeddedFontLst")
        ef  = etree.SubElement(lst, f"{{{p_ns}}}embeddedFont")
        fnt = etree.SubElement(ef,  f"{{{p_ns}}}font")
        fnt.set("typeface", font_name)
        reg = etree.SubElement(ef,  f"{{{p_ns}}}regular")
        reg.set(f"{{{r2_ns}}}id", rid)
        patched_prs = etree.tostring(prs_root, xml_declaration=True, encoding="UTF-8", standalone=True)

        # Patch Content_Types
        ct_root   = etree.fromstring(zin.read(ct_path))
        part_name = f"/ppt/fonts/{ttf_filename}"
        font_ct   = "application/vnd.openxmlformats-officedocument.presentationml.font"
        if not any(el.get("PartName") == part_name for el in ct_root.findall(f"{{{ct_ns}}}Override")):
            ov = etree.SubElement(ct_root, f"{{{ct_ns}}}Override")
            ov.set("PartName",    part_name)
            ov.set("ContentType", font_ct)
        patched_ct = etree.tostring(ct_root, xml_declaration=True, encoding="UTF-8", standalone=True)

        patched = {rels_path, prs_path_zip, ct_path}
        for item in zin.infolist():
            if item.filename not in patched:
                zout.writestr(item, zin.read(item.filename))
        zout.writestr(rels_path,     patched_rels)
        zout.writestr(prs_path_zip,  patched_prs)
        zout.writestr(ct_path,       patched_ct)
        zout.writestr(font_zip_path, font_data)

    dst.seek(0)
    return dst.read()


def _build_song_preview_slides(lyrics_text: str, s: SongSettings, song_title: str) -> list[dict]:
    slides = parse_song_lyrics(lyrics_text)
    if not slides:
        return [{"is_title": False, "prefix": "", "prefix_color": "", "lines": ["(no lyrics)"], "title": song_title}]

    # In preview, attach title/reference to slides
    first_title = slides[0].get("title", "") if (slides and slides[0].get("is_title")) else song_title
    for i, s_item in enumerate(slides):
        if not s_item.get("is_title"):
            s_item["title"] = song_title or first_title
    return slides


# ── Song OCR endpoint ────────────────────────────────────────────
@app.post("/api/song/extract")
async def song_extract(file: UploadFile = File(...)):
    """Accept a PDF/JPG/PNG upload, run OCR, return extracted lyrics as JSON."""
    ext = Path(file.filename or "").suffix.lower()
    if ext not in (".pdf", ".png", ".jpg", ".jpeg"):
        return {"error": "Only PDF, PNG, JPG files are supported."}, 400
    data = await file.read()
    text = extract_lyrics_from_bytes(data, ext)
    return {"text": text}


@app.post("/api/song/preview")
def song_preview(req: SongPreviewRequest):
    """Return list of slide objects for live preview."""
    slides = _build_song_preview_slides(req.lyrics_text, req.settings, req.song_title)
    return {"slides": slides, "total": len(slides)}


@app.post("/api/song/generate")
def song_generate(req: SongGenerateRequest):
    """Build and stream a PPTX from the provided lyrics + settings."""
    pptx_bytes = _build_song_pptx(req.lyrics_text, req.settings, req.song_title)
    filename   = f"song_slides_{uuid.uuid4().hex[:8]}.pptx"
    return StreamingResponse(
        io.BytesIO(pptx_bytes),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


# ═══════════════════════════════════════════════════════════════════
#  Merge Slides — combine multiple PPTX files into one
#  Strategy: low-level ZIP surgery so every slide's XML, images,
#  fonts, themes, and background are preserved 100% faithfully.
# ═══════════════════════════════════════════════════════════════════

import copy as _copy
import base64 as _b64

# ── Faithful slide copy via ZIP-level surgery ────────────────────

def _next_rId(existing_ids: set[str]) -> str:
    """Return the next unused rId string (rId1, rId2, …)."""
    n = 1
    while f"rId{n}" in existing_ids:
        n += 1
    return f"rId{n}"


def merge_pptx_files(files_bytes: list[bytes]) -> bytes:
    """
    Merge a list of PPTX byte-strings into one PPTX preserving every
    slide's background, images, fonts, themes, and text styles exactly.

    Technique: open the first file as a mutable ZIP (the destination),
    then for each subsequent file and each slide:
      1. Copy the slide XML → ppt/slides/slideN.xml
      2. Copy its _rels      → ppt/slides/_rels/slideN.xml.rels
      3. Copy every referenced part (images, audio, video, fonts, …)
         into the dest ZIP, remapping rIds to avoid collisions
      4. Register the slide in ppt/presentation.xml (sldIdLst) and
         ppt/_rels/presentation.xml.rels
      5. Add a Content-Type Override for the new slide
    """
    if not files_bytes:
        raise ValueError("No files provided")

    NS_RELS = "http://schemas.openxmlformats.org/package/2006/relationships"
    NS_CT   = "http://schemas.openxmlformats.org/package/2006/content-types"
    NS_P    = "http://schemas.openxmlformats.org/presentationml/2006/main"
    NS_R    = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    SLIDE_CT = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
    SLIDE_REL_TYPE = f"{NS_R}/slide"

    # ── Work entirely in memory ──────────────────────────────────
    dest_buf = io.BytesIO(files_bytes[0])

    # We'll build a fresh output ZIP so we can freely rewrite entries
    out_buf = io.BytesIO()

    with zipfile.ZipFile(dest_buf, "r") as base_zip:
        # Snapshot all existing entries
        existing_names: set[str] = set(base_zip.namelist())

        # Parse the base presentation XML
        prs_xml_bytes = base_zip.read("ppt/presentation.xml")
        prs_root      = etree.fromstring(prs_xml_bytes)

        # Parse base presentation rels
        prs_rels_bytes = base_zip.read("ppt/_rels/presentation.xml.rels")
        prs_rels_root  = etree.fromstring(prs_rels_bytes)

        # Parse Content_Types
        ct_bytes = base_zip.read("[Content_Types].xml")
        ct_root  = etree.fromstring(ct_bytes)

        # Collect all existing parts as {name: bytes}
        parts: dict[str, bytes] = {n: base_zip.read(n) for n in existing_names}

        # Find the highest existing slide number in ppt/slides/slideN.xml
        slide_re = re.compile(r"^ppt/slides/slide(\d+)\.xml$")
        existing_slide_nums = [
            int(m.group(1)) for n in existing_names if (m := slide_re.match(n))
        ]
        next_slide_num = (max(existing_slide_nums) + 1) if existing_slide_nums else 1

        # Highest existing sldId value in presentation sldIdLst
        sldIdLst = prs_root.find(f"{{{NS_P}}}sldIdLst")
        if sldIdLst is None:
            sldIdLst = etree.SubElement(prs_root, f"{{{NS_P}}}sldIdLst")
        existing_sld_ids = {
            int(el.get("id", 0)) for el in sldIdLst.findall(f"{{{NS_P}}}sldId")
        }
        next_sld_id = (max(existing_sld_ids) + 1) if existing_sld_ids else 256

        # Existing rIds in presentation rels
        existing_prs_rIds = {
            el.get("Id", "") for el in prs_rels_root.findall(f"{{{NS_RELS}}}Relationship")
        }

        # ── Process each additional source file ─────────────────
        for src_bytes in files_bytes[1:]:
            with zipfile.ZipFile(io.BytesIO(src_bytes), "r") as src_zip:
                src_names = set(src_zip.namelist())

                # Detect slides in this source
                src_slide_re = re.compile(r"^ppt/slides/slide(\d+)\.xml$")
                src_slides = sorted(
                    [n for n in src_names if src_slide_re.match(n)],
                    key=lambda n: int(src_slide_re.match(n).group(1)),
                )

                for src_slide_path in src_slides:
                    slide_num_str = src_slide_re.match(src_slide_path).group(1)
                    src_slide_rels_path = f"ppt/slides/_rels/slide{slide_num_str}.xml.rels"

                    # Read slide XML
                    slide_xml_bytes = src_zip.read(src_slide_path)
                    slide_root      = etree.fromstring(slide_xml_bytes)

                    # Read slide rels (may not exist for simple slides)
                    src_rels_root = None
                    rId_remap: dict[str, str] = {}  # old rId → new rId

                    if src_slide_rels_path in src_names:
                        src_rels_bytes = src_zip.read(src_slide_rels_path)
                        src_rels_root  = etree.fromstring(src_rels_bytes)

                        # Build new rels XML, copying all referenced parts
                        new_rels_root = etree.Element(f"{{{NS_RELS}}}Relationships")

                        for rel_el in src_rels_root.findall(f"{{{NS_RELS}}}Relationship"):
                            old_rid    = rel_el.get("Id", "")
                            rel_type   = rel_el.get("Type", "")
                            target     = rel_el.get("Target", "")
                            target_mode = rel_el.get("TargetMode", "")

                            if target_mode == "External":
                                # Keep external rels as-is, just remap Id
                                new_rid = _next_rId(
                                    {e.get("Id") for e in new_rels_root} | set(existing_prs_rIds)
                                )
                                rId_remap[old_rid] = new_rid
                                new_el = etree.SubElement(new_rels_root, f"{{{NS_RELS}}}Relationship")
                                new_el.set("Id", new_rid)
                                new_el.set("Type", rel_type)
                                new_el.set("Target", target)
                                new_el.set("TargetMode", "External")
                                continue

                            # Resolve path relative to ppt/slides/
                            if target.startswith("/"):
                                abs_target = target.lstrip("/")
                            else:
                                # Join with the slide's directory then normalise ../
                                raw = "ppt/slides/" + target
                                parts_list = raw.split("/")
                                norm: list[str] = []
                                for p in parts_list:
                                    if p == "..":
                                        if norm:
                                            norm.pop()
                                    elif p and p != ".":
                                        norm.append(p)
                                abs_target = "/".join(norm)

                            # Is it a slide layout reference? point to dest layout
                            if "slideLayout" in rel_type or "slideLayout" in abs_target:
                                # Re-use the first slideLayout from the dest presentation
                                dest_layout_rels = [
                                    e for e in prs_rels_root.findall(f"{{{NS_RELS}}}Relationship")
                                    if "slideLayout" in e.get("Target", "")
                                       or "slideLayout" in e.get("Type", "")
                                ]
                                # Find layout from first dest slide's rels
                                first_dest_slide_rels_path = f"ppt/slides/_rels/slide{existing_slide_nums[0] if existing_slide_nums else 1}.xml.rels"
                                layout_target = "../slideLayouts/slideLayout1.xml"  # safe default
                                if first_dest_slide_rels_path in parts:
                                    fds_rels = etree.fromstring(parts[first_dest_slide_rels_path])
                                    for frel in fds_rels.findall(f"{{{NS_RELS}}}Relationship"):
                                        if "slideLayout" in frel.get("Type", ""):
                                            layout_target = frel.get("Target", layout_target)
                                            break

                                new_rid = _next_rId(
                                    {e.get("Id") for e in new_rels_root}
                                )
                                rId_remap[old_rid] = new_rid
                                new_el = etree.SubElement(new_rels_root, f"{{{NS_RELS}}}Relationship")
                                new_el.set("Id", new_rid)
                                new_el.set("Type", rel_type)
                                new_el.set("Target", layout_target)
                                continue

                            # Copy the actual part bytes across
                            if abs_target in src_names:
                                part_bytes = src_zip.read(abs_target)
                                # Avoid filename collisions by appending _mN suffix
                                dest_target = abs_target
                                if abs_target in parts:
                                    # If bytes are identical reuse the existing part
                                    if parts[abs_target] == part_bytes:
                                        pass  # reuse same path — dest_target stays as abs_target
                                    else:
                                        # Pick a fresh name
                                        stem, _, ext_part = abs_target.rpartition(".")
                                        n = 2
                                        candidate = abs_target
                                        while candidate in parts:
                                            candidate = f"{stem}_m{n}.{ext_part}"
                                            n += 1
                                        dest_target = candidate
                                        parts[dest_target] = part_bytes
                                else:
                                    parts[dest_target] = part_bytes

                                # Build relative target back from ppt/slides/
                                # dest_target is like "ppt/media/image1.png" → "../media/image1.png"
                                dest_parts = dest_target.split("/")
                                rel_target = "../" + "/".join(dest_parts[1:])

                                new_rid = _next_rId(
                                    {e.get("Id") for e in new_rels_root}
                                )
                                rId_remap[old_rid] = new_rid
                                new_el = etree.SubElement(new_rels_root, f"{{{NS_RELS}}}Relationship")
                                new_el.set("Id", new_rid)
                                new_el.set("Type", rel_type)
                                new_el.set("Target", rel_target)

                                # Add Content-Type override if missing
                                if not any(
                                    ov.get("PartName") == f"/{dest_target}"
                                    for ov in ct_root.findall(f"{{{NS_CT}}}Override")
                                ):
                                    # Guess content type from extension
                                    ext_map = {
                                        "png": "image/png", "jpg": "image/jpeg",
                                        "jpeg": "image/jpeg", "gif": "image/gif",
                                        "bmp": "image/bmp", "tiff": "image/tiff",
                                        "wmf": "image/x-wmf", "emf": "image/x-emf",
                                        "mp4": "video/mp4", "mp3": "audio/mpeg",
                                        "wav": "audio/wav",
                                    }
                                    file_ext = dest_target.rsplit(".", 1)[-1].lower()
                                    ct_val = ext_map.get(file_ext)
                                    if ct_val:
                                        ov = etree.SubElement(ct_root, f"{{{NS_CT}}}Override")
                                        ov.set("PartName", f"/{dest_target}")
                                        ov.set("ContentType", ct_val)

                    # ── Patch rIds inside the slide XML ─────────
                    slide_xml_str = slide_xml_bytes.decode("utf-8")
                    for old_rid, new_rid in rId_remap.items():
                        # Replace all attribute-value occurrences of the rId.
                        # XML attributes always appear as ="rId3" or ='rId3'.
                        # Use both quote styles to be safe.
                        slide_xml_str = slide_xml_str.replace(f'="{old_rid}"', f'="{new_rid}"')
                        slide_xml_str = slide_xml_str.replace(f"='{old_rid}'", f"='{new_rid}'")
                    slide_xml_bytes = slide_xml_str.encode("utf-8")

                    # ── Register new slide in dest ZIP ───────────
                    new_slide_path      = f"ppt/slides/slide{next_slide_num}.xml"
                    new_slide_rels_path = f"ppt/slides/_rels/slide{next_slide_num}.xml.rels"

                    parts[new_slide_path] = slide_xml_bytes
                    if src_rels_root is not None:
                        parts[new_slide_rels_path] = etree.tostring(
                            new_rels_root,
                            xml_declaration=True, encoding="UTF-8", standalone=True,
                        )
                    else:
                        # No rels in source — create a minimal one pointing at a slide layout
                        min_rels = etree.Element(f"{{{NS_RELS}}}Relationships")
                        layout_target = "../slideLayouts/slideLayout1.xml"
                        first_snum = existing_slide_nums[0] if existing_slide_nums else 1
                        fds_path = f"ppt/slides/_rels/slide{first_snum}.xml.rels"
                        if fds_path in parts:
                            try:
                                fds_r = etree.fromstring(parts[fds_path])
                                for frel in fds_r.findall(f"{{{NS_RELS}}}Relationship"):
                                    if "slideLayout" in frel.get("Type", ""):
                                        layout_target = frel.get("Target", layout_target)
                                        break
                            except Exception:
                                pass
                        lo_el = etree.SubElement(min_rels, f"{{{NS_RELS}}}Relationship")
                        lo_el.set("Id", "rId1")
                        lo_el.set("Type", f"{NS_R}/slideLayout")
                        lo_el.set("Target", layout_target)
                        parts[new_slide_rels_path] = etree.tostring(
                            min_rels,
                            xml_declaration=True, encoding="UTF-8", standalone=True,
                        )

                    # Content-Type for slide
                    if not any(
                        ov.get("PartName") == f"/{new_slide_path}"
                        for ov in ct_root.findall(f"{{{NS_CT}}}Override")
                    ):
                        ov = etree.SubElement(ct_root, f"{{{NS_CT}}}Override")
                        ov.set("PartName", f"/{new_slide_path}")
                        ov.set("ContentType", SLIDE_CT)

                    # Presentation rels — register slide
                    new_prs_rid = _next_rId(existing_prs_rIds)
                    existing_prs_rIds.add(new_prs_rid)
                    new_prs_rel = etree.SubElement(
                        prs_rels_root, f"{{{NS_RELS}}}Relationship"
                    )
                    new_prs_rel.set("Id", new_prs_rid)
                    new_prs_rel.set("Type", SLIDE_REL_TYPE)
                    new_prs_rel.set("Target", f"slides/slide{next_slide_num}.xml")

                    # sldIdLst entry
                    sld_id_el = etree.SubElement(sldIdLst, f"{{{NS_P}}}sldId")
                    sld_id_el.set("id", str(next_sld_id))
                    sld_id_el.set(f"{{{NS_R}}}id", new_prs_rid)

                    existing_slide_nums.append(next_slide_num)
                    next_slide_num += 1
                    next_sld_id    += 1

        # ── Write back mutated XML ───────────────────────────────
        parts["ppt/presentation.xml"] = etree.tostring(
            prs_root, xml_declaration=True, encoding="UTF-8", standalone=True,
        )
        parts["ppt/_rels/presentation.xml.rels"] = etree.tostring(
            prs_rels_root, xml_declaration=True, encoding="UTF-8", standalone=True,
        )
        parts["[Content_Types].xml"] = etree.tostring(
            ct_root, xml_declaration=True, encoding="UTF-8", standalone=True,
        )

        # ── Write final ZIP ──────────────────────────────────────
        with zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as out_zip:
            for name, data in parts.items():
                out_zip.writestr(name, data)

    out_buf.seek(0)
    return out_buf.read()


# ── Render slide thumbnails using LibreOffice / pdf2image ────────

def _render_slide_thumbnails(pptx_bytes: bytes, width_px: int = 960) -> list[str]:
    """
    Convert every slide in a PPTX to a base64-encoded PNG thumbnail.
    Uses LibreOffice headless → PDF → Pillow rasterise pipeline.
    Falls back to a coloured placeholder if LibreOffice is unavailable.
    Returns list of base64 PNG strings (one per slide).
    """
    import subprocess, shutil, tempfile

    thumbs: list[str] = []
    libreoffice = shutil.which("libreoffice") or shutil.which("soffice")

    if not libreoffice:
        # Fallback: extract background colour + text from python-pptx and
        # render a simple Pillow image so the user still sees something useful.
        return _render_slides_pillow_fallback(pptx_bytes, width_px)

    with tempfile.TemporaryDirectory() as tmp:
        pptx_path = Path(tmp) / "input.pptx"
        pptx_path.write_bytes(pptx_bytes)

        result = subprocess.run(
            [
                libreoffice, "--headless", "--convert-to", "png",
                "--outdir", tmp, str(pptx_path),
            ],
            capture_output=True, timeout=60,
        )

        if result.returncode != 0:
            return _render_slides_pillow_fallback(pptx_bytes, width_px)

        # LibreOffice names output as input.png (single) or input1.png, input2.png…
        png_files = sorted(
            Path(tmp).glob("input*.png"),
            key=lambda p: int(re.search(r"\d+", p.stem.replace("input", "") or "0").group() or 0),
        )
        if not png_files:
            return _render_slides_pillow_fallback(pptx_bytes, width_px)

        from PIL import Image as PILImage
        for png_path in png_files:
            img = PILImage.open(png_path).convert("RGB")
            # Resize to target width maintaining aspect ratio
            h = int(img.height * width_px / img.width)
            img = img.resize((width_px, h), PILImage.LANCZOS)
            buf = io.BytesIO()
            img.save(buf, format="PNG", optimize=True)
            thumbs.append(_b64.b64encode(buf.getvalue()).decode())

    return thumbs


def _render_slides_pillow_fallback(pptx_bytes: bytes, width_px: int = 960) -> list[str]:
    """
    Fallback renderer: uses python-pptx to extract each slide's content and
    draws a faithful Pillow image — correct background colour, images at the
    right position/size, and text at the correct font size and colour.
    """
    from PIL import Image as PILImage, ImageDraw, ImageFont as PILFont

    # ── Font cache: use Khmer TTF from the app's fonts/ directory ──────
    _FONT_DIR = Path(__file__).parent / "fonts"
    _font_cache: dict[int, object] = {}

    def _get_font(size_px: int) -> object:
        size_px = max(6, min(size_px, 400))
        if size_px in _font_cache:
            return _font_cache[size_px]
        for ttf_name in [
            "KhmerOSBattambang-Regular.ttf",
            "KhmerOS_.ttf",
            "KhmerOSMoulLight.ttf",
            "KhmerOS_muol.ttf",
        ]:
            ttf_path = _FONT_DIR / ttf_name
            if ttf_path.exists():
                try:
                    font = PILFont.truetype(str(ttf_path), size=size_px)
                    _font_cache[size_px] = font
                    return font
                except Exception:
                    pass
        try:
            font = PILFont.load_default(size=size_px)
        except TypeError:
            font = PILFont.load_default()
        _font_cache[size_px] = font
        return font

    def _safe_rgb(fill) -> tuple | None:
        """Return (r,g,b) from a fill object, or None if unavailable."""
        try:
            from pptx.enum.dml import MSO_FILL
            # BACKGROUND(5) and NONE have no fore_color
            if fill.type is None or fill.type.name in ("BACKGROUND", "NONE"):
                return None
            c = fill.fore_color.rgb
            return (c.red, c.green, c.blue)
        except Exception:
            return None

    prs       = Presentation(io.BytesIO(pptx_bytes))
    sw_emu    = float(prs.slide_width)
    sh_emu    = float(prs.slide_height)
    # slide height in points (1 pt = 914400/72 EMU)
    slide_h_pt = sh_emu / 914400 * 72
    aspect     = sh_emu / sw_emu if sw_emu else 9 / 16
    height_px  = int(width_px * aspect)

    # Scale EMU → pixels
    def ex(emu: float) -> int: return int(emu / sw_emu * width_px)
    def ey(emu: float) -> int: return int(emu / sh_emu * height_px)
    # Scale font points → pixels relative to canvas height
    def fpt(pt: float) -> int: return max(6, int(pt / slide_h_pt * height_px))

    thumbs: list[str] = []

    for slide in prs.slides:
        # ── Background ──────────────────────────────────────────
        bg_rgb = (255, 255, 255)  # default white (most slides are white)
        try:
            bg_color = _safe_rgb(slide.background.fill)
            if bg_color is not None:
                bg_rgb = bg_color
        except Exception:
            pass

        img  = PILImage.new("RGB", (width_px, height_px), bg_rgb)
        draw = ImageDraw.Draw(img)

        for shape in slide.shapes:
            sx = ex(shape.left)
            sy = ey(shape.top)
            sw = ex(shape.width)
            sh = ey(shape.height)

            # ── Image shapes ──────────────────────────────────
            is_pic = False
            try:
                _ = shape.image  # raises if not a picture
                is_pic = True
            except Exception:
                pass

            if is_pic:
                try:
                    from PIL import Image as _PIL2
                    pic = _PIL2.open(io.BytesIO(shape.image.blob))
                    if sw > 0 and sh > 0:
                        pic = pic.resize((sw, sh), _PIL2.LANCZOS)
                        if pic.mode == "RGBA":
                            bg_patch = _PIL2.new("RGBA", (sw, sh), bg_rgb + (255,))
                            bg_patch.alpha_composite(pic)
                            img.paste(bg_patch.convert("RGB"), (sx, sy))
                        else:
                            img.paste(pic.convert("RGB"), (sx, sy))
                except Exception:
                    pass
                continue

            # ── Filled rectangle (non-text, non-image) ────────
            if not shape.has_text_frame:
                try:
                    col = _safe_rgb(shape.fill)
                    if col is not None and sw > 0 and sh > 0:
                        draw.rectangle([sx, sy, sx + sw, sy + sh], fill=col)
                except Exception:
                    pass
                continue

            # ── Text frame ────────────────────────────────────
            tf = shape.text_frame
            # Inherit paragraph/shape-level defaults for font size & colour
            def _para_font_pt(para) -> float:
                """Best-effort paragraph-level font size in pt."""
                try:
                    s = para.runs[0].font.size
                    if s:
                        return s.pt
                except Exception:
                    pass
                try:
                    s = tf.paragraphs[0].runs[0].font.size
                    if s:
                        return s.pt
                except Exception:
                    pass
                return 24.0  # sensible default

            def _run_color(run, para_default=(0, 0, 0)) -> tuple:
                try:
                    c = run.font.color.rgb
                    return (c.red, c.green, c.blue)
                except Exception:
                    pass
                # fall back to theme colour luminance
                try:
                    from pptx.dml.color import _Colormap
                    pass
                except Exception:
                    pass
                return para_default

            line_y = sy
            for para in tf.paragraphs:
                # Collect runs (preserve spaces between runs)
                runs = [(r.text, r) for r in para.runs if r.text]
                if not runs:
                    # empty para = line spacing
                    def_pt = _para_font_pt(para)
                    line_y += fpt(def_pt * 1.2)
                    continue

                max_h = 0
                run_x = sx
                for text, run in runs:
                    try:
                        pt = run.font.size.pt if run.font.size else _para_font_pt(para)
                    except Exception:
                        pt = _para_font_pt(para)
                    px = fpt(pt)
                    col = _run_color(run)
                    font_obj = _get_font(px)
                    draw.text((run_x, line_y), text, font=font_obj, fill=col)
                    try:
                        bb = draw.textbbox((run_x, line_y), text, font=font_obj)
                        run_x = bb[2]
                        max_h = max(max_h, bb[3] - bb[1])
                    except Exception:
                        run_x += px * len(text) // 2
                        max_h = max(max_h, px)
                line_y += max_h + max(2, int(max_h * 0.15))

        buf = io.BytesIO()
        img.save(buf, format="PNG", optimize=True)
        thumbs.append(_b64.b64encode(buf.getvalue()).decode())

    return thumbs


@app.post("/api/merge/count-slides")
async def merge_count_slides(file: UploadFile = File(...)):
    """Return the number of slides in an uploaded PPTX."""
    ext = Path(file.filename or "").suffix.lower()
    if ext != ".pptx":
        from fastapi import HTTPException
        raise HTTPException(status_code=400, detail="Only .pptx files are supported.")
    data = await file.read()
    prs  = Presentation(io.BytesIO(data))
    return {"slide_count": len(prs.slides), "filename": file.filename}


@app.post("/api/merge/preview-slides")
async def merge_preview_slides(file: UploadFile = File(...)):
    """
    Render all slides of a single PPTX as base64 PNG thumbnails.
    Returns: { filename, slide_count, thumbnails: [base64_png, ...] }
    """
    from fastapi import HTTPException
    ext = Path(file.filename or "").suffix.lower()
    if ext != ".pptx":
        raise HTTPException(status_code=400, detail="Only .pptx files are supported.")
    data   = await file.read()
    thumbs = _render_slides_pillow_fallback(data, width_px=800)
    return {
        "filename":    file.filename,
        "slide_count": len(thumbs),
        "thumbnails":  thumbs,
    }


@app.post("/api/merge/generate")
async def merge_generate(
    files: list[UploadFile] = File(...),
    output_name: str = Form("merged_slides"),
):
    """
    Accept multiple PPTX files, merge all their slides in order,
    and return the combined PPTX as a download.
    """
    from fastapi import HTTPException

    if len(files) < 2:
        raise HTTPException(status_code=400, detail="Please upload at least 2 PPTX files.")

    files_bytes: list[bytes] = []
    for uf in files:
        ext = Path(uf.filename or "").suffix.lower()
        if ext != ".pptx":
            raise HTTPException(status_code=400, detail=f"'{uf.filename}' is not a .pptx file.")
        files_bytes.append(await uf.read())

    merged = merge_pptx_files(files_bytes)

    safe_name = re.sub(r'[^\w\-]', '_', output_name.strip()) or "merged_slides"
    filename  = f"{safe_name}.pptx"

    return StreamingResponse(
        io.BytesIO(merged),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


# ═══════════════════════════════════════════════════════════════════
#  Melody Slide — PDF / Image Staff + Lyric Piece Cropper & Slides
# ═══════════════════════════════════════════════════════════════════

def _row_dark_density_melody(gray: np.ndarray, threshold: int = 240) -> np.ndarray:
    return (gray < threshold).mean(axis=1).astype(float)


def _moving_avg_melody(arr: np.ndarray, w: int = 9) -> np.ndarray:
    return np.convolve(arr, np.ones(w) / w, mode="same")


def _gap_spans_melody(is_blank: np.ndarray, min_px: int) -> list:
    spans, in_g, g0 = [], False, 0
    for y, b in enumerate(is_blank):
        if b and not in_g:
            in_g, g0 = True, y
        elif not b and in_g:
            in_g = False
            if y - g0 >= min_px:
                spans.append((g0, y))
    if in_g and len(is_blank) - g0 >= min_px:
        spans.append((g0, len(is_blank)))
    return spans


def _detect_melody_raw_strips(image: Image.Image,
                              white_thr: int = 240,
                              min_gap_px: int = 10,
                              min_block_px: int = 35) -> list[tuple[int, int]]:
    gray    = np.array(image.convert("L"))
    density = _moving_avg_melody(_row_dark_density_melody(gray, white_thr), w=9)
    blank   = density < 0.005
    strips, prev = [], 0
    for g0, g1 in _gap_spans_melody(blank, min_gap_px):
        if g0 - prev >= min_block_px:
            strips.append((prev, g0))
        prev = g1
    h = gray.shape[0]
    if h - prev >= min_block_px:
        strips.append((prev, h))
    return strips


def _pair_melody_notation_lyrics(strips: list[tuple[int, int]], gray: np.ndarray,
                                 white_thr: int = 240) -> list[tuple[int, int]]:
    if len(strips) < 2:
        return strips

    def dens(t, b):
        return (gray[t:b] < white_thr).mean() if b > t else 0.0

    def inter_gap(i):
        return strips[i + 1][0] - strips[i][1] if i + 1 < len(strips) else 99999

    out, i = [], 0
    while i < len(strips):
        t0, b0 = strips[i]
        if i + 1 < len(strips):
            t1, b1  = strips[i + 1]
            g_in    = inter_gap(i)
            g_out   = inter_gap(i + 1)
            d0, d1  = dens(t0, b0), dens(t1, b1)
            tight   = g_in < g_out * 0.60 and g_in < 110
            stave_above_lyric = g_in < 90 and d0 > d1 * 1.15
            if tight or stave_above_lyric:
                out.append((t0, b1))
                i += 2
                continue
        out.append(strips[i])
        i += 1
    return out


def _make_melody_crop(page_rgb: Image.Image,
                      crop_box: tuple[int, int, int, int] | tuple[int, int],
                      pad: int = 15) -> Image.Image:
    """
    Produce a cropped PIL image with white padding.
    Supports either (top, bottom) for full-width horizontal crop,
    or (left, top, right, bottom) for 2D box crop.
    """
    w, h = page_rgb.size
    if len(crop_box) == 2:
        top, bottom = crop_box
        t = max(0, top - pad)
        b = min(h, bottom + pad)
        crop = page_rgb.crop((0, t, w, b))
    else:
        l, t, r, b = crop_box
        crop = page_rgb.crop((
            max(0, l - pad),
            max(0, t - pad),
            min(w, r + pad),
            min(h, b + pad),
        ))

    from PIL import ImageOps
    return ImageOps.expand(crop, border=20, fill=(255, 255, 255))


def _add_melody_image_slide(prs: Presentation, img_bytes: bytes, bg_hex: str = "#ffffff") -> None:
    """Add a slide and center/scale the cropped melody piece to fit presentation slide."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg_rgb = _hex_to_rgb(bg_hex)
    bg = slide.background.fill
    bg.solid()
    bg.fore_color.rgb = RGBColor(*bg_rgb)

    sw, sh = prs.slide_width, prs.slide_height

    im = Image.open(io.BytesIO(img_bytes))
    iw, ih = im.size
    aspect = iw / ih

    max_w = sw * 0.94
    max_h = sh * 0.88
    fit_w = max_w
    fit_h = fit_w / aspect
    if fit_h > max_h:
        fit_h = max_h
        fit_w = fit_h * aspect

    left = (sw - fit_w) / 2
    top  = (sh - fit_h) / 2

    # Save to temp buffer for python-pptx add_picture
    buf = io.BytesIO(img_bytes)
    slide.shapes.add_picture(buf, int(left), int(top), int(fit_w), int(fit_h))


@app.post("/api/melody/process")
async def melody_process(
    file: UploadFile = File(...),
    lines_per_slide: int = Form(1),
    dpi: int = Form(200),
):
    """
    Accept an uploaded PDF or image file of a melody/hymn sheet,
    render all pages to PNG, detect or extract pieces/strips (staff + lyrics),
    and return pages + pieces with base64 images so the user can review,
    reorder, remove, or manually select pieces.
    """
    ext = Path(file.filename or "").suffix.lower()
    if ext not in (".pdf", ".png", ".jpg", ".jpeg", ".webp"):
        raise HTTPException(status_code=400, detail="Only PDF, PNG, JPG, or WEBP files are supported.")

    file_bytes = await file.read()
    page_images: list[Image.Image] = []

    if ext == ".pdf":
        try:
            from pdf2image import convert_from_bytes
            page_images = convert_from_bytes(file_bytes, dpi=dpi)
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"PDF conversion failed: {e}")
    else:
        try:
            img = Image.open(io.BytesIO(file_bytes)).convert("RGB")
            page_images = [img]
        except Exception as e:
            raise HTTPException(status_code=400, detail=f"Image reading failed: {e}")

    result_pages = []
    result_pieces = []
    piece_counter = 0

    for page_idx, page in enumerate(page_images):
        page_rgb = page.convert("RGB")
        w, h = page_rgb.size
        gray = np.array(page_rgb.convert("L"))

        # Save web-sized page preview (max width 1200)
        preview_page = page_rgb.copy()
        preview_scale = 1.0
        if w > 1200:
            preview_scale = 1200 / w
            preview_page = preview_page.resize((1200, int(h * preview_scale)), Image.Resampling.LANCZOS)

        p_buf = io.BytesIO()
        preview_page.save(p_buf, format="JPEG", quality=85)
        page_b64 = f"data:image/jpeg;base64,{base64.b64encode(p_buf.getvalue()).decode('utf-8')}"

        result_pages.append({
            "page_index": page_idx,
            "width": w,
            "height": h,
            "preview_data_url": page_b64,
        })

        # Automatic melody piece detection (staff rows paired with lyrics)
        raw_strips = _detect_melody_raw_strips(page_rgb, white_thr=240, min_gap_px=10, min_block_px=35)
        if not raw_strips:
            # Fallback: slice into equal vertical sections
            n_sec = max(3, h // 280)
            sec_h = h // n_sec
            raw_strips = [(i * sec_h, (i + 1) * sec_h) for i in range(n_sec)]

        singable = _pair_melody_notation_lyrics(raw_strips, gray)

        # Group singable lines per slide
        groups: list[tuple[int, int]] = []
        lps = max(1, min(lines_per_slide, 6))
        for start_idx in range(0, len(singable), lps):
            chunk = singable[start_idx : start_idx + lps]
            groups.append((chunk[0][0], chunk[-1][1]))

        for grp_top, grp_bot in groups:
            crop_im = _make_melody_crop(page_rgb, (grp_top, grp_bot), pad=16)
            if crop_im.size[1] < 30:
                continue

            c_buf = io.BytesIO()
            crop_im.save(c_buf, format="PNG")
            c_bytes = c_buf.getvalue()
            c_b64 = f"data:image/png;base64,{base64.b64encode(c_bytes).decode('utf-8')}"

            piece_counter += 1
            result_pieces.append({
                "id": str(uuid.uuid4())[:8],
                "order": piece_counter,
                "page_index": page_idx,
                "box": [0, int(grp_top), int(w), int(grp_bot)],
                "image_data_url": c_b64,
                "label": f"Piece {piece_counter}",
            })

    return {
        "filename": file.filename,
        "page_count": len(page_images),
        "pages": result_pages,
        "pieces": result_pieces,
    }


class MelodyCropRequest(BaseModel):
    page_index: int
    box: list[int]   # [left, top, right, bottom] in original image coordinates
    page_data_url: Optional[str] = None


@app.post("/api/melody/crop-piece")
async def melody_crop_piece(
    file: Optional[UploadFile] = File(None),
    page_index: int = Form(0),
    box_json: str = Form("[0,0,100,100]"),
    dpi: int = Form(200),
):
    """Crop an arbitrary user-selected rectangle from a page and return the piece."""
    import json
    try:
        box = json.loads(box_json)
    except Exception:
        raise HTTPException(status_code=400, detail="Invalid box JSON")

    if not file:
        raise HTTPException(status_code=400, detail="File is required")

    file_bytes = await file.read()
    ext = Path(file.filename or "").suffix.lower()

    if ext == ".pdf":
        from pdf2image import convert_from_bytes
        pages = convert_from_bytes(file_bytes, dpi=dpi, first_page=page_index + 1, last_page=page_index + 1)
        if not pages:
            raise HTTPException(status_code=400, detail="Page not found in PDF")
        page_img = pages[0].convert("RGB")
    else:
        page_img = Image.open(io.BytesIO(file_bytes)).convert("RGB")

    l, t, r, b = box
    crop_im = _make_melody_crop(page_img, (l, t, r, b), pad=10)
    c_buf = io.BytesIO()
    crop_im.save(c_buf, format="PNG")
    c_bytes = c_buf.getvalue()
    c_b64 = f"data:image/png;base64,{base64.b64encode(c_bytes).decode('utf-8')}"

    return {
        "id": str(uuid.uuid4())[:8],
        "image_data_url": c_b64,
        "box": box,
        "page_index": page_index,
    }


class MelodyGenerateRequest(BaseModel):
    title: str = "melody_slides"
    bg_color: str = "#ffffff"
    pieces: list[dict] = []  # list of { image_data_url: str, label?: str }


@app.post("/api/melody/generate")
async def melody_generate(req: MelodyGenerateRequest):
    """
    Generate a 16:9 PowerPoint presentation (.pptx) where each piece (melody staff + lyric cut)
    is placed centered and scaled up on its own slide for optimal church projector display.
    """
    if not req.pieces:
        raise HTTPException(status_code=400, detail="No melody pieces to generate slides for.")

    prs = Presentation()
    prs.slide_width  = Inches(13.33)
    prs.slide_height = Inches(7.5)

    for piece in req.pieces:
        data_url = piece.get("image_data_url", "")
        if not data_url or "," not in data_url:
            continue
        try:
            b64_str = data_url.split(",", 1)[1]
            img_bytes = base64.b64decode(b64_str)
            _add_melody_image_slide(prs, img_bytes, bg_hex=req.bg_color)
        except Exception as e:
            continue

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)

    safe_title = re.sub(r'[^\w\-]', '_', req.title.strip()) or "melody_slides"
    filename = f"{safe_title}.pptx"

    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


# ═══════════════════════════════════════════════════════════════════
#  Bible Reader Proxy (same-origin iframe for wordproject.org)
# ═══════════════════════════════════════════════════════════════════

# Browsers block cross-origin iframe JS, so we serve wordproject.org through
# a high-performance same-origin proxy and inject a script that forwards the
# user's selection to the parent app (the Bible Text input).

import time
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

_BIBLE_PROXY_BASE = "https://www.wordproject.org"

# Reusable connection pool to wordproject.org for fast keep-alive requests
_proxy_session = requests.Session()
_proxy_adapter = HTTPAdapter(
    pool_connections=25,
    pool_maxsize=50,
    max_retries=Retry(total=2, backoff_factor=0.1, status_forcelist=[502, 503, 504]),
)
_proxy_session.mount("https://", _proxy_adapter)
_proxy_session.mount("http://", _proxy_adapter)

# In-memory asset cache: url -> (content_bytes, media_type, headers_dict, expires_at)
_PROXY_CACHE: dict[str, tuple[bytes, str, dict[str, str], float]] = {}
_MAX_PROXY_CACHE_ENTRIES = 600

_STATIC_EXTENSIONS = (
    ".css", ".js", ".ttf", ".woff", ".woff2", ".eot", ".otf",
    ".png", ".jpg", ".jpeg", ".gif", ".ico", ".svg", ".webp",
)

# Selection bridge: injected into every proxied page.
# Only captures and puts text into the slide input when the user RELEASES the mouse
# (removes hold of mouse). While the user is dragging / holding the mouse down,
# it never prematurely captures or adds partial text.
_SELECTION_SCRIPT = b"""
<script>
(function () {
  // Fix wordproject.org script bug on pages without #textBody
  try {
    if (window.$) {
      $(function () {
        if (!$('#textBody').length) {
          $('<div id="textBody" style="display:none"></div>').appendTo('body');
        }
        if (!$('#fadeout').length) {
          $('<div id="fadeout" style="display:none"></div>').appendTo('body');
        }
      });
    }
  } catch (e) {}

  var isPointerDown = false;
  var currentSelectedText = '';
  var hasActiveSelection = false;

  function onPointerDown() {
    isPointerDown = true;
  }

  function onPointerUp() {
    isPointerDown = false;
    // Wait for the browser selection to settle after mouseup
    setTimeout(finishSelection, 40);
  }

  function finishSelection() {
    // If user is still holding the mouse down, do NOT capture yet!
    if (isPointerDown) {
      return;
    }

    var sel = window.getSelection();
    var text = sel ? sel.toString().trim() : '';

    if (!text) {
      // User cleared the selection / unselected
      if (hasActiveSelection) {
        hasActiveSelection = false;
        currentSelectedText = '';
        try {
          parent.postMessage({ source: 'wordproject-proxy', type: 'deselection' }, '*');
        } catch (e) {}
      }
      return;
    }

    // If identical to active selection, do not re-send (prevents duplicate)
    if (text === currentSelectedText) {
      return;
    }

    // User finished dragging and released hold of mouse -> capture complete text!
    currentSelectedText = text;
    hasActiveSelection = true;
    try {
      parent.postMessage({ source: 'wordproject-proxy', type: 'selection', text: text }, '*');
    } catch (e) {}
  }

  // Pointer down: mark as holding mouse
  document.addEventListener('mousedown', onPointerDown, true);
  document.addEventListener('pointerdown', onPointerDown, true);
  document.addEventListener('touchstart', onPointerDown, true);

  // Pointer up: mouse hold released -> now capture!
  document.addEventListener('mouseup', onPointerUp, true);
  document.addEventListener('pointerup', onPointerUp, true);
  document.addEventListener('touchend', onPointerUp, true);

  // Keyboard selection (Shift + Arrow keys)
  document.addEventListener('keyup', function (e) {
    if (e.key === 'Shift' || e.key === 'ArrowLeft' || e.key === 'ArrowRight' || e.key === 'ArrowUp' || e.key === 'ArrowDown') {
      setTimeout(finishSelection, 50);
    }
  });

  // Selectionchange guard: NEVER fire while mouse is held down
  document.addEventListener('selectionchange', function () {
    if (isPointerDown) {
      return;
    }
    clearTimeout(finishSelection._t);
    finishSelection._t = setTimeout(finishSelection, 150);
  });
})();
</script>
"""


def _proxy_html(html: bytes) -> bytes:
    """Insert the selection bridge before </body> (or append if absent)."""
    low = html.lower()
    idx = low.rfind(b"</body>")
    if idx == -1:
        return html + _SELECTION_SCRIPT
    return html[:idx] + _SELECTION_SCRIPT + html[idx:]


@app.get("/api/bible-proxy/{path:path}")
def bible_proxy(path: str, request: Request):
    """
    High-performance proxy for wordproject.org.
    - Synchronous def so FastAPI dispatches to threadpool for true parallel downloads.
    - Keep-alive connection pooling via requests.Session.
    - In-memory caching for static assets (fonts, css, js, images) and HTML.
    - Browser Cache-Control headers for instant local loading.
    """
    url = f"{_BIBLE_PROXY_BASE}/{path}"
    query = request.url.query
    if query:
        url += f"?{query}"

    now = time.time()
    is_static = any(path.lower().endswith(ext) for ext in _STATIC_EXTENSIONS)

    # Check in-memory cache
    if url in _PROXY_CACHE:
        cached_content, cached_media_type, cached_headers, expires_at = _PROXY_CACHE[url]
        if now < expires_at:
            return Response(content=cached_content, media_type=cached_media_type, headers=cached_headers)

    try:
        resp = _proxy_session.get(
            url,
            timeout=25,
            headers={
                "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
                "Referer": "https://www.wordproject.org/",
                "Accept-Encoding": "gzip, deflate",
            },
        )
    except requests.RequestException:
        raise HTTPException(status_code=502, detail="Could not reach wordproject.org")

    if resp.status_code != 200:
        return Response(content=resp.content, status_code=resp.status_code)

    content_type = resp.headers.get("content-type", "application/octet-stream")
    body = resp.content

    if "text/html" in content_type:
        body = _proxy_html(body)
        headers = {
            "Cache-Control": "public, max-age=3600",
            "Access-Control-Allow-Origin": "*",
        }
        ttl = 3600  # 1 hour cache for HTML
    else:
        headers = {
            "Cache-Control": "public, max-age=604800, immutable" if is_static else "public, max-age=3600",
            "Access-Control-Allow-Origin": "*",
        }
        ttl = 604800 if is_static else 3600

    # Cache response in memory
    if len(_PROXY_CACHE) > _MAX_PROXY_CACHE_ENTRIES:
        _PROXY_CACHE.clear()
    _PROXY_CACHE[url] = (body, content_type, headers, now + ttl)

    return Response(content=body, media_type=content_type, headers=headers)


# ── Static frontend (Vite build) ─────────────────────────────────
# Mount the built React app so a single deployment serves everything.
# On Render: build.sh runs `npm run build` before gunicorn starts,
# so frontend/dist exists by the time this code runs.
_FRONTEND_DIST = Path(__file__).parent / "frontend" / "dist"

def _setup_static() -> None:
    """Mount static files only if the built frontend exists."""
    import sys
    print(f"[static] __file__      = {__file__}", file=sys.stderr)
    print(f"[static] _FRONTEND_DIST = {_FRONTEND_DIST}", file=sys.stderr)
    print(f"[static] dist exists   = {_FRONTEND_DIST.exists()}", file=sys.stderr)

    if not _FRONTEND_DIST.exists():
        print("[static] No frontend/dist — running in API-only mode", file=sys.stderr)
        return  # running locally without a build — API-only mode

    assets_dir = _FRONTEND_DIST / "assets"
    if assets_dir.exists():
        app.mount(
            "/assets",
            StaticFiles(directory=str(assets_dir)),
            name="assets",
        )
        print(f"[static] Mounted /assets -> {assets_dir}", file=sys.stderr)

    # SPA catch-all — must be defined at module level (not inside a conditional)
    # so gunicorn workers inherit it correctly after fork.

@app.get("/{full_path:path}", include_in_schema=False)
async def serve_spa(full_path: str):
    """Catch-all: serve index.html for any path not matched by /api routes."""
    index = _FRONTEND_DIST / "index.html"
    if not index.exists():
        return {"detail": "Frontend not built"}
    candidate = _FRONTEND_DIST / full_path
    if candidate.is_file():
        return FileResponse(str(candidate))
    return FileResponse(str(index))

_setup_static()
