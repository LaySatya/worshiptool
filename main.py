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
from functools import lru_cache
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


# ── Khmer spelling / script helpers ─────────────────────────────
# Khmer combining signs that must attach to the preceding base character.
# OCR output frequently inserts a space between a consonant and its vowel /
# coeng / diacritic (e.g. "ស្ ត្រៃ" instead of "ស្ត្រី"), which reads as broken
# spelling.  Any space that immediately follows (or precedes) one of these is
# removed so the word is rejoined.
_COMBINING_RE = re.compile(
    r'[\u17B4-\u17D3\u17DD]'          # dependent vowels, signs & numerals
    r'|\u17D2'                        # coeng (subscript marker)
)

# A leading dependent sign means the previous word was split — glue it back.
# Two dependent signs may follow each other inside one cluster (e.g. អ + ា + ន
# where OCR puts a space after the vowel), so the whole run is consumed.
_SPACE_BEFORE_COMBINING_RE = re.compile(
    r'\s+(?=[\u17B4-\u17D3]+)'
)

# Dependent vowel/diacritic followed by a space and then the next consonant of
# the same syllable — "អា ន" -> "អាន".  Applied by _glue_incomplete() only when
# the fragment before the space cannot be a complete Khmer word.
# NOTE: coeng (\u17D2) is excluded — a space after a coeng is ALWAYS wrong.
_SIGN_THEN_SPACE_RE = re.compile(
    r'(?<=[\u17B4-\u17D3])\s+(?=[\u1780-\u17A2])'
)


_KHMER_CONSONANTS = set(chr(c) for c in range(0x1780, 0x17A3))

# Dependent vowels (U+17B6..U+17C5), the ំ ះ signs (U+17C6, U+17C7) and
# diacritics (U+17C8..U+17D3) — i.e. everything that attaches to a consonant.
_KHMER_VOWELS_OR_SIGNS = set(
    chr(c) for c in list(range(0x17B6, 0x17D4))
)


@lru_cache(maxsize=4096)
def _is_incomplete_khmer_fragment(token: str) -> bool:
    """
    True when `token` is a syllable that OCR cut in half, i.e. the following
    space is NOT a real word break.

    Khmer syllables are written as:  consonant (coeng consonant)* vowel sign*.
    If OCR drops the space the syllable breaks immediately after a consonant
    that has no vowel attached yet, e.g.

        "អ ា ន"      -> "អ" is a bare consonant            -> fragment
        "ក្ ន"      -> trailing coeng without sub-consonant -> fragment
        "ស្ ត្រី"    -> trailing coeng                      -> fragment
        "ព្រ ះ"      -> trailing coeng                      -> fragment

    A real Khmer word ends with a vowel sign or a bare consonant-letter that
    IS a word ("ខ្ញុំ", "ចូល", "ទៅ", "ស្រឡាញ់") — those are left untouched.

    Only UNAMBIGUOUS split markers are flagged, so a real word is never glued
    onto its neighbour:

      • the token ends with the coeng marker (\u17D2) — a coeng can never end
        a word, so the next cluster continues the same syllable ("ស្ត្រី ស្");
      • the token is a LONE dependent sign ("ា", "ំ", "់") — it cannot stand
        on its own, so it belongs to the previous cluster.

    Nothing else is touched.  That means languages with legitimate one-letter
    words, and every normal word break, are preserved exactly as OCR saw them.
    """
    if not token:
        return False
    if token.endswith('\u17D2'):
        return True                      # dangling coeng -> never a word end
    if len(token) == 1 and token not in _KHMER_CONSONANTS:
        return True                      # lone vowel sign / diacritic
    return False

# Coeng followed by an optional space before its sub-consonant.
_COENG_SPACE_RE = re.compile(r'\u17D2\s+(?=[\u1780-\u17A2])')

# Dependent vowel signs only (U+17B6..U+17C5) — used to spot a vowel-final
# syllable that OCR cut off before its closing consonant.
_KHMER_VOWELS = set(chr(c) for c in range(0x17B6, 0x17C6))

# Very common Khmer words that OCR frequently splits into two fragments.
# Re-joining is safe because the fragments are not independent words.
# (Order matters only for readability — matches are on the joined result.)
_KHMER_SPLIT_WORDS: dict[str, tuple[str, ...]] = {
    "ដែល": ("ដ ែល", "ដែ ល", "ដែ ល"),
    "មាន": ("មា ន", "ម ាន"),
    "គឺ": ("គ ឺ",),
    "នេះ": ("នេ ះ", "ន េះ"),
    "នោះ": ("នោ ះ", "ន ោះ"),
    "ទាំង": ("ទា ំង", "ទ ំាង"),
    "យើង": ("យ ើង", "យើ ង"),
    "ព្រះ": ("ព្រ ះ", "ព្ រះ"),
    "គ្រិស្ត": ("គ្រិ ស្ត", "គ្រ ិស្ត"),
    "អម្ចាស់": ("អម្ចា ស់", "អ ម្ចាស់"),
    "ក្នុង": ("ក្នុ ង", "ក្ន ំង"),
}


# Whole-word split repairs: "មា ន" -> "មាន".  These are ordinary words whose
# two halves are not separate words, so the space is always OCR noise.
_WORD_SPLIT_PAIRS: list[tuple[re.Pattern, str]] = [
    (re.compile(re.escape(frag)), word)
    for word, frags in _KHMER_SPLIT_WORDS.items()
    for frag in frags
]


def _fix_khmer_spacing(s: str) -> str:
    """
    Re-join Khmer clusters that OCR split with stray whitespace.

    "ស្ត្រី សូរយើង"  -> fine (space between words)
    "ស្ ត្រៃសូរ យើង" -> "ស្ត្រៃសូរយើង" (spaces inside a cluster removed)
    "បែរ ចិត្ត ំ"      -> "បែរចិត្តំ"    (lone combining sign re-attached)
    "អ ា ន"           -> "អាន"
    "ក ់"            -> "ក់"
    "ក្រ ែង"         -> "ក្រែង"
    "ក្ ន"           -> "ក្ន"
    "១ ។"            -> "១។"
    "ព្រ ៖"          -> "ព្រ៖"
    "ខ្ញុំ ចូល ទៅ ក្នុង" -> unchanged (real word spacing is kept)
    """
    if not s:
        return s
    # A token ending in a dependent sign that is NOT a complete word is an
    # incomplete syllable; the following consonant belongs to it
    # ("អា ន" -> "អាន").  Real words are left untouched ("ខ្ញុំ ចូល").
    def _glue_incomplete(m: re.Match) -> str:
        start = m.start()
        j = start
        while j > 0 and not s_prev[j - 1].isspace():
            j -= 1
        token = s_prev[j:start]
        if not token:
            return m.group(0)
        if _is_incomplete_khmer_fragment(token):
            return ''
        # Short vowel-final fragment + a following fragment-looking next token
        # is a syllable split mid-way ("អា ន", "សើ យើង").  Rejoin only when the
        # next token is itself short (1-2 chars), so real words stay separate.
        k = start
        while k < len(s_prev) and s_prev[k].isspace():
            k += 1
        e = k
        while e < len(s_prev) and not s_prev[e].isspace():
            e += 1
        nxt = s_prev[k:e]
        if (
            len(token) == 2
            and token[0] in _KHMER_CONSONANTS
            and token[-1] in _KHMER_VOWELS
            and nxt
            # next token must not be a real independent word (2+ chars starting
            # with a consonant is a word, so leave it alone)
            and len(nxt) <= 2
        ):
            return ''
        # NOTE: a 2-char consonant+vowel token ("សើ", "ទៅ") is indistinguishable
        # from a real one-word break, so it is deliberately left alone rather
        # than risking gluing two genuine words together.
        return m.group(0)

    s_prev = s
    s = _SIGN_THEN_SPACE_RE.sub(_glue_incomplete, s_prev)
    # coeng + space + sub-consonant -> coeng + sub-consonant
    s = _COENG_SPACE_RE.sub('\u17D2', s)
    # space before a dependent vowel / diacritic -> remove it
    s = _SPACE_BEFORE_COMBINING_RE.sub('', s)

    # A trailing syllable-final mark (់ ៈ ៉ ៊ ់ ៌ ៍ ៎ ៏ ័ ៑ ...) may never
    # follow a space: the space was inserted INSIDE the syllable.
    #
    #     "អម្ចា ស់"  ->  "អម្ចាស់"
    #     "ណា ក់"    ->  "ណាក់"
    #
    # Only ់ (U+17CB, the syllable-final consonant mark) is treated as a split
    # marker.  It can never begin a Khmer word, so a space before
    # "consonant + ់" is always an OCR glitch.
    #
    # Deliberately NOT included:
    #   ៏ U+17CF — real words begin with it, e.g. "ដ៏" (that/which)
    #    ្ U+17D2 — "ក្នុង" legitimately starts with a subscript cluster
    #   ំ ះ — "ដំណឹង", "ព្រះ" are ordinary word starts
    s = re.sub(r'\s+(?=[\u1780-\u17A2]\u17CB)', '', s)
    # …and a space directly before a dependent sign is equally invalid.
    s = re.sub(r'\s+(?=[\u17B4-\u17D3])', '', s)
    # Khmer sentence/clause marks never stand alone after a space
    s = re.sub(r'\s+([\u17D4-\u17DA])', r'\1', s)
    # Repair common words that OCR split into two fragments ("មា ន" -> "មាន").
    for pat, word in _WORD_SPLIT_PAIRS:
        s = pat.sub(word, s)
    # Collapse the gaps left behind
    s = re.sub(r'\s{2,}', ' ', s)
    return s.strip()


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
# These are common misreads: slashes, angles, brackets, punctuation artifacts.
# NOTE: Khmer punctuation (។ ៕ ៖ ៗ …) is deliberately NOT listed here.
# NOTE: '~' is NOT noise — a hymn verse marker is written "១~" and stripping
# it produces "១  នៅពេញ..." in the final text.
_OCR_NOISE_RE = re.compile(r'[/\\<>«»›‹|(){}\[\]#@&*^`]+')

# OCR output normalisation: ASCII punctuation used where Khmer punctuation
# belongs.  Only very safe, well-known substitutions are applied.
_OCR_MARK_FIXES: list[tuple[re.Pattern, str]] = [
    # Repeated full stop ".." is OCR for the Khmer full stop
    (re.compile(r'\.{2,}'), '\u17D4'),                     # ។
    # Full-width / look-alike punctuation normalised to ASCII
    (re.compile(r'[，、]'), ','),
    (re.compile(r'[；]'), ';'),
    # A ':' printed straight after a Khmer consonant is OCR for ះ (U+17C7),
    # as in "ព្រ:គុណ" for "ព្រះគុណ".  A real colon in Khmer text is ៖ (U+17D6)
    # and never follows a consonant without a space, so this is unambiguous.
    (re.compile(r'(?<=[\u1780-\u17A2])\s*[:ៈ]'), '\u17C7'),
]

# Repetition mark ៗ (U+17D7): OCR commonly prints a quote character after the
# repeated syllable.  Handled with a callable so the group reference and the
# literal Khmer codepoint can be combined safely (a '\\1\u17D7' template is
# rejected by Python's re module as a bad escape).
_REPEAT_MARK_RE = re.compile(r'([\u1780-\u17A2])["“”]')
_REPEAT_MARK = '\u17D7'


def _normalise_ocr_marks(s: str) -> str:
    """Restore Khmer punctuation that OCR replaced with ASCII symbols."""
    for pat, repl in _OCR_MARK_FIXES:
        s = pat.sub(repl, s)
    s = _REPEAT_MARK_RE.sub(lambda m: m.group(1) + _REPEAT_MARK, s)
    return s

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
    s = _normalise_ocr_marks(s)
    s = _fix_khmer_spacing(s)
    s = re.sub(r' {2,}', ' ', s)
    return s.strip()


def _syllable_evidence(s: str) -> int:
    """
    Count how many genuine Khmer syllables `s` contains.

    A written Khmer syllable is a base consonant (optionally followed by coeng
    sub-consonants) plus optional vowel signs.  Real lyrics are made of these;
    notation/solfege rows and OCR smear produce runs with no base consonant at
    all (only vowel signs and diacritics), or endless coeng chains.

    This works even when the whole line has no spaces, which is common for
    Khmer song lyrics.

    Returns the number of syllables detected.
    """
    count = 0
    i = 0
    n = len(s)
    while i < n:
        c = s[i]
        if c in _KHMER_CONSONANTS:
            j = i + 1
            # absorb coeng + sub-consonant pairs, then vowel signs / diacritics
            while j < n:
                if s[j] == '\u17D2' and j + 1 < n and s[j + 1] in _KHMER_CONSONANTS:
                    j += 2
                elif s[j] in _KHMER_VOWELS_OR_SIGNS:
                    j += 1
                else:
                    break
            count += 1
            i = j
        else:
            i += 1
    return count


_SOLFEGE_VOWEL = '\u17B8'   # ី (ii) — every note name ends with it

# Known Khmer note names (do re mi fa sol la ti) as written in hymn books
_SOLFEGE_NAMES = {
    'ដូ', 'ឌូ', 'រី', 'មី', 'ហ្វា', 'ហ្វា', 'សូ', 'ឡា', 'ធី', 'វី',
}


def _split_khmer_syllables(s: str) -> list[str]:
    """Split a Khmer string into written syllables (base + marks + coeng)."""
    out: list[str] = []
    i, n = 0, len(s)
    while i < n:
        c = s[i]
        if c in _KHMER_CONSONANTS:
            j = i + 1
            while j < n:
                if s[j] == '\u17D2' and j + 1 < n and s[j + 1] in _KHMER_CONSONANTS:
                    j += 2
                elif s[j] in _KHMER_VOWELS_OR_SIGNS:
                    j += 1
                else:
                    break
            out.append(s[i:j])
            i = j
        else:
            i += 1
    return out


_KHMER_DIGITS_SET = set("០១២៣៤៥៦៧៨៩")


def _fix_verse_marker(s: str) -> str:
    """
    Repair a verse number that OCR merged with the '~' / '.' separator.

    In hymn books a verse starts with "១~" or "1.".  OCR frequently reads the
    '~' as another digit, giving "១០~..." / "១០នៅ..." — a wrong verse number.

    Rules applied only at the START of a line:
      • "១០" / "១៣" etc. followed by a word -> keep the first digit only
        (verse numbers never run into the lyric without a separator)
      • strip a trailing '~' or '.' left in the text
    """
    # A verse may start with digits and no separator at all ("១០នៅ..." where
    # the '~' was read as ០).  Match the digit run optionally followed by a
    # separator, and require Khmer text immediately after it.
    m = re.match(r'^([០-៩0-9]+)\s*([~\.\:\)])?\s*(?=[\u1780-\u17FF])', s)
    if not m:
        return s
    num = m.group(1)
    sep = m.group(2)
    rest = s[m.end():]
    if not rest:
        return s

    # A real separator was recognised -> keep it exactly as printed
    # ("២." must stay "២.").  Only invent a '~' when none was found.
    if sep:
        return num + sep + rest

    # No separator: the lyric runs straight into the digits.  A Khmer verse
    # number is a single digit, so keep the first.
    if len(num) >= 2:
        num = num[0]
    return num + '~' + rest


def _is_solfege_row(s: str) -> bool:
    """
    True when the line is a melody-notation row rather than lyrics.

    Detection: the line has at least 4 syllables, and >= 60 % of them are
    short solfege notes — either a known note name (រី / មី / ហ្វា / ឌូ ...)
    or a very short syllable ending in the 'ី' vowel.
    """
    syls = _split_khmer_syllables(s)
    if len(syls) < 4:
        return False
    note_like = 0
    for syl in syls:
        bare = syl.rstrip(''.join(_KHMER_VOWELS_OR_SIGNS) + '\u17D2')
        if syl in _SOLFEGE_NAMES:
            note_like += 1
        elif len(syl) <= 4 and syl.endswith(_SOLFEGE_VOWEL):
            note_like += 1
    return note_like / len(syls) >= 0.60


def _is_lyric_line(s: str) -> bool:
    """
    Return True only if this line looks like a genuine Khmer lyric line.

    IMPORTANT: Khmer lyrics are frequently written WITHOUT spaces between
    words (and OCR often drops them anyway), so a valid line can arrive as ONE
    single unbroken Khmer token.  The gates below are therefore based on the
    STRUCTURE of the script (consonant/vowel rhythm, repetition, density) and
    never on "how many space-separated words are there".

    Rejection criteria (any one -> reject):
      1. No Khmer characters at all.
      2. Fewer than 4 Khmer characters (stray glyph noise).
      3. Repetitive character distribution (OCR smear / solfege).
      4. Khmer chars < 55 % of all non-space chars (mixed notation/ascii).
      5. Notation-only line: almost no real syllables (see _syllable_evidence).
      6. Repeating 2-3 char pattern run ("ោយោយោយ", "a ba a ba").
      7. A long line whose character bigrams are dominated by one pair
         (staff-line / rhythm notation run).
    """
    if not _line_is_khmer(s):
        return False

    k_count = _khmer_char_count(s)
    if k_count < 4:
        return False

    if _is_repetitive(s):
        return False

    # Khmer density gate — notation rows are mostly ASCII / symbols
    total_chars = len(s.replace(' ', ''))
    if total_chars > 0 and k_count / total_chars < 0.55:
        return False

    # A real Khmer lyric line contains several proper consonant+vowel
    # syllables.  Pure notation / solfege rows contain almost none.
    if _syllable_evidence(s) < 2:
        return False

    # Solfege / note-name row (រី មី វី ហ្វា ឌូ ...).  These are SHORT syllables
    # built from a consonant + the 'ី' vowel (or a '្វ'-cluster), repeated many
    # times.  Real lyrics use a much wider variety of shapes.
    if _is_solfege_row(s):
        return False

    tokens = s.split()

    # Repeating short pattern run: "ោយោយោយ", "កាកាកា"
    for tok in tokens:
        if len(tok) >= 6 and _REPEAT_SUB_RE.search(tok):
            return False

    # Long token dominated by a single character bigram -> notation run.
    # A genuine Khmer word mixes many different clusters.
    for tok in tokens:
        if len(tok) >= 14:
            chars = list(tok)
            bigrams = [chars[i] + chars[i+1] for i in range(len(chars)-1)]
            if bigrams:
                top_bg_count = Counter(bigrams).most_common(1)[0][1]
                if top_bg_count / len(bigrams) >= 0.30:
                    return False

    # Rhythm rows written as identical repeated syllables ("pi pi pi pi")
    if len(tokens) >= 4:
        for i in range(len(tokens) - 1):
            if tokens[i] == tokens[i + 1] and len(tokens[i]) <= 3:
                return False
        if _has_repeated_bigram(tokens):
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

# Tesseract's LSTM engine degrades badly (and can return NOTHING) on very
# large images.  Anything above this width is downscaled first — accuracy is
# actually BETTER at ~2500-4000 px than at 6000+ px.
_MAX_OCR_WIDTH = 3600
_MAX_OCR_HEIGHT = 4800


def _limit_ocr_size(img: Image.Image) -> Image.Image:
    """
    Downscale an image so Tesseract can actually process it.

    The scale factor is computed from BOTH dimensions and applied to both, so
    the aspect ratio is preserved.  (Clamping width and height independently
    squashes wide, short lyric strips and destroys the glyph shapes.)
    """
    w, h = img.size
    if w <= 0 or h <= 0:
        return img
    scale = min(_MAX_OCR_WIDTH / w, _MAX_OCR_HEIGHT / h, 1.0)
    if scale < 1.0:
        new_w = max(1, int(round(w * scale)))
        new_h = max(1, int(round(h * scale)))
        img = img.resize((new_w, new_h), Image.LANCZOS)
    return img


def _preprocess_for_ocr(img: Image.Image, target_width: int = 2400) -> Image.Image:
    """
    Prepare a lyric-strip image for Tesseract (Khmer-optimised).

    Measured behaviour of Tesseract's Khmer model:
      • It reads Khmer perfectly when glyph height is ~30-55 px.
      • Accuracy collapses above roughly 64 px per glyph ("ស្ត" -> "សព្ត")
        and on extreme aspect ratios (a 2600x181 strip failed completely).

    So the image is scaled ONCE, proportionally, to land glyph height in the
    sweet spot — never stretched on one axis, which distorts every syllable.

      1. Greyscale
      2. Proportional resize so the strip is at most TARGET_W wide AND the
         text line is a comfortable height (both limits, same scale factor)
      3. Auto-contrast (stretch histogram to full 0-255 range)
      4. Unsharp-mask (crisper edges, doesn't over-thicken thin strokes)
      5. Otsu binarise (global threshold — robust on clean printed lyrics and
         does not erode the thin strokes the way local-mean did)
      6. White border (Tesseract hates glyphs at the very edge)
    """
    img = img.convert("L")

    w, h = img.size
    if w <= 0 or h <= 0:
        return img

    # Step 2 — ONE proportional scale.
    #   • cap the width so Tesseract's layout analyser can cope
    #   • boost small text so glyph height reaches a readable size
    scale = 1.0
    if w > target_width:
        scale = target_width / w
    elif h < 40:
        # A very short strip (thin line of text) — enlarge it proportionally
        scale = min(60.0 / h, target_width / w)

    if abs(scale - 1.0) > 1e-3:
        new_w = max(1, int(round(w * scale)))
        new_h = max(1, int(round(h * scale)))
        img = img.resize((new_w, new_h), Image.LANCZOS)

    # Guard: never exceed Tesseract's comfortable size on EITHER axis
    img = _limit_ocr_size(img)

    # Step 3 — auto-contrast
    img = ImageOps.autocontrast(img, cutoff=2)

    # Step 4 — unsharp mask (radius 2, percent 150, threshold 3)
    img = img.filter(ImageFilter.UnsharpMask(radius=2, percent=150, threshold=3))

    # Step 5 — Otsu binarise.
    #   A global threshold keeps thin Khmer strokes intact.  The previous
    #   local-mean threshold (mean - 18) eroded them and turned "ស្ត" into "សព្ត".
    arr = np.array(img, dtype=np.uint8)
    hist = np.bincount(arr.ravel(), minlength=256).astype(np.float64)
    total = arr.size
    if total:
        sum_all = float(np.dot(np.arange(256), hist))
        sum_b = 0.0
        w_b = 0.0
        best_var = -1.0
        thr = 127
        for i in range(256):
            w_b += hist[i]
            if w_b == 0:
                continue
            w_f = total - w_b
            if w_f == 0:
                break
            sum_b += i * hist[i]
            m_b = sum_b / w_b
            m_f = (sum_all - sum_b) / w_f
            var = w_b * w_f * (m_b - m_f) ** 2
            if var > best_var:
                best_var = var
                thr = i
    else:
        thr = 127
    img = Image.fromarray(((arr > thr) * 255).astype(np.uint8), mode="L")

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
    #           uniform density variance (regular pattern of staff lines).
    #           Guards: a real staff band is TALL (5 staff lines + note heads
    #           span >= 40 px at 150 dpi) and contains at least 3 rows with
    #           long horizontal runs.  A single line of large title text would
    #           otherwise be misread as notation.
    arr_d = np.array(row_densities)
    mean_d = float(arr_d.mean())
    std_d  = float(arr_d.std())
    if mean_d > 0.08 and std_d < 0.06 and h_band >= 36 and long_run_rows >= 3:
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


def _classify_strips(page_rgb: Image.Image) -> dict:
    """
    Split a music-sheet page into three kinds of horizontal region:

      title   — the song-title block at the TOP of the page (song number,
                Khmer title, and often a Korean / English subtitle).  This is
                NOT lyric text: it becomes a picture on the first slide.
      music   — staff / melody notation rows (and solfege note rows).
      lyrics  — the Khmer lyric lines that sit UNDER the melody.

    Returns {"title": crop|None, "lyrics": [crop, ...], "has_music": bool,
             "lyric_bands": [(top,bot), ...]}

    The rule is positional: everything above the FIRST melody row is the
    title block.  Everything below it that is not itself notation is lyrics.
    """
    gray = np.array(page_rgb.convert("L"))
    h, w = gray.shape

    result = {"title": None, "title_band": None, "lyrics": [],
              "has_music": False, "lyric_bands": []}

    strips = _detect_content_strips(gray)
    if len(strips) < 2:
        # Flat / photocopied page — no structure to exploit.
        return result

    # 1. Label every strip as music or not
    labelled: list[tuple[int, int, bool]] = []
    for top, bot in strips:
        is_music = _is_staff_strip(gray, top, bot, page_width=w)
        labelled.append((top, bot, is_music))

    first_music_idx = next((i for i, (_, _, m) in enumerate(labelled) if m), None)
    result["has_music"] = first_music_idx is not None

    # 2. Title block = all non-music strips ABOVE the first melody row.
    if first_music_idx is not None and first_music_idx > 0:
        head = labelled[:first_music_idx]
        t = max(0, head[0][0] - 12)
        b = min(h, head[-1][1] + 12)
        # Only treat it as a title when it is a compact block, not most of
        # the page (a lyric-only page has no title block).
        if (b - t) <= h * 0.42:
            result["title"] = page_rgb.crop((0, t, w, b))
            result["title_band"] = (t, b)

    # 3. Lyrics = non-music strips BELOW the first melody row.
    #
    #    IMPORTANT: pad vertically only into BLANK space.  A fixed pad pulls
    #    in the neighbouring staff lines / note heads, and OCR then reads them
    #    as noise ("- -- »\"- ... ដល់មនុស្ស...").  So the pad is clipped to the
    #    gap between this strip and the music rows around it.
    start_idx = first_music_idx if first_music_idx is not None else 0
    base_pad = 22          # room for Khmer vowel signs above/below the baseline
    for idx in range(start_idx, len(labelled)):
        top, bot, is_music = labelled[idx]
        if is_music:
            continue

        # How far can we expand before hitting a music strip or page edge?
        up_limit = 0
        for k in range(idx - 1, -1, -1):
            if labelled[k][2]:           # music strip directly above
                up_limit = labelled[k][1]
                break
        down_limit = h
        for k in range(idx + 1, len(labelled)):
            if labelled[k][2]:           # music strip directly below
                down_limit = labelled[k][0]
                break

        t = max(up_limit, top - base_pad)
        b = min(down_limit, bot + base_pad)
        crop = page_rgb.crop((0, t, w, b))

        # Scrub residual staff lines only.  Do NOT try to crop further to a
        # "text band": Khmer vowel signs occupy the same rows as note heads, so
        # any row-trimming erases the diacritics and shreds the word.  Leftover
        # note heads are removed by the text filter after OCR instead.
        crop = _erase_staff_lines(crop)

        result["lyrics"].append(crop)
        result["lyric_bands"].append((top, bot))

    return result


def _erase_staff_lines(img: Image.Image) -> Image.Image:
    """
    Whitened out any full-width horizontal rules (staff lines) accidentally
    included in a lyric crop, and drop isolated specks (note heads).
    """
    gray = np.array(img.convert("L"))
    h, w = gray.shape
    if h == 0 or w == 0:
        return img

    dark = gray < 170
    # A staff line: dark on > 55 % of the row, in one long run.
    row_frac = dark.mean(axis=1)
    staff_rows = np.where(row_frac > 0.55)[0]

    out = gray.copy()
    if len(staff_rows):
        for r in staff_rows:
            lo = max(0, r - 2)
            hi = min(h, r + 3)
            out[lo:hi, :] = 255

    # NOTE: do NOT try to trim "specks" by density banding — Khmer vowel signs
    # sit on the same rows as note heads and get erased with them, destroying
    # the lyric.  Staff lines are already removed above; any remaining note
    # heads are handled by the text filter after OCR.
    return Image.fromarray(out).convert(img.mode)


def _crop_to_text_band(img: Image.Image) -> Image.Image:
    """
    Tighten a lyric crop to the single row-range that actually holds text.

    Called after the crop already contains exactly one lyric line plus, at the
    edges, leftovers from the neighbouring staff (note heads or the tail of the
    line above).  Those leftovers both break OCR and inject junk like
    '” “””” ” ” ”“”'.

    Instead of guessing from ink density (which erases Khmer vowel signs), this
    uses the fact that TEXT spans a wide horizontal extent with many separate
    strokes, while note heads are small isolated round blobs:

      • count horizontal runs of dark pixels per row
      • a text row has many runs (one per letter stroke) AND a wide bounding
        extent; a note-head row has few runs / narrow extent
      • keep the longest contiguous block of text-like rows
    """
    gray = np.array(img.convert("L"))
    h, w = gray.shape
    if h < 4 or w < 4:
        return img

    dark = gray < 180
    run_counts = np.zeros(h, dtype=int)
    extents = np.zeros(h, dtype=float)

    for r in range(h):
        row = dark[r]
        if not row.any():
            continue
        # number of dark->light transitions = number of separate strokes
        d = np.diff(row.astype(np.int8))
        run_counts[r] = int(np.count_nonzero(d == 1))
        xs = np.where(row)[0]
        extents[r] = (xs[-1] - xs[0]) / w if len(xs) else 0.0

    # A real lyric row: many strokes spread across most of the crop width.
    text_like = (run_counts >= 14) & (extents >= 0.35)

    if not text_like.any():
        return img

    rows = np.where(text_like)[0]
    # Split into contiguous groups (allow 2-row gaps for vowel signs)
    breaks = np.where(np.diff(rows) > 3)[0]
    groups = np.split(rows, breaks + 1)

    # Prefer the group with the widest horizontal extent (the lyric line),
    # falling back to the longest group.
    def score(g):
        return (float(np.mean(extents[g])), len(g))
    band = max(groups, key=score)

    t = max(0, int(band[0]) - 4)
    b = min(h, int(band[-1]) + 5)
    return img.crop((0, t, w, b))


def _extract_lyric_strips(page_rgb: Image.Image) -> list[Image.Image]:
    """
    Return cropped PIL images of ONLY the lyric (Khmer text) rows that sit
    UNDER the melody notation, discarding the title block and all staff rows.

    Falls back to "every non-notation strip" when no melody row is present
    (a plain lyric sheet), so ordinary text-only pages still work.
    """
    info = _classify_strips(page_rgb)
    if info["lyrics"]:
        return info["lyrics"]

    # No melody detected at all -> treat non-music strips as lyrics.
    gray = np.array(page_rgb.convert("L"))
    h, w = gray.shape
    strips = _detect_content_strips(gray)
    if len(strips) < 2:
        return []

    title_band = info.get("title_band")
    crops: list[Image.Image] = []
    pad = 24
    for top, bot in strips:
        if _is_staff_strip(gray, top, bot, page_width=w):
            continue
        if title_band and not (bot < title_band[0] or top > title_band[1]):
            continue   # this strip is part of the title block
        t = max(0, top - pad)
        b = min(h, bot + pad)
        crops.append(page_rgb.crop((0, t, w, b)))
    return crops


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


def _trim_edge_junk(s: str) -> str:
    """
    Strip OCR junk from the START and END of a line.

    When a lyric crop picks up a neighbouring staff, OCR emits the note heads as
    symbols before/after the real text, e.g.

        '˝ ” “””” ” ” ”“” ដល់មនុស្សជាច្រើននៅគ្រប់ជាតិសាសន៍ “” ” '

    The lyric itself is intact in the middle, so everything outside the first
    and last Khmer character is removed (keeping a verse marker before it and
    Khmer punctuation after it).
    """
    first = None
    for i, c in enumerate(s):
        if _KHMER_CHAR_RE.match(c):
            first = i
            break
    if first is None:
        return s.strip()

    last = first
    for i in range(len(s) - 1, -1, -1):
        if _KHMER_CHAR_RE.match(s[i]):
            last = i
            break

    core = s[first:last + 1]
    lead = s[:first]
    tail = s[last + 1:]

    lead_keep = re.search(r'([០-៩0-9]+\s*[~\.\:\)]?\s*)$', lead)
    tail_keep = re.match(r'([\u17D4-\u17DA\.\,\!\?]+)', tail)

    out = (
        (lead_keep.group(1) if lead_keep else '')
        + core
        + (tail_keep.group(1) if tail_keep else '')
    ).strip()

    return _drop_polluted_prefix(out)


def _drop_polluted_prefix(s: str) -> str:
    """
    Remove note-head junk that OCR printed BEFORE the real lyric text.

    Example produced by a crop that caught the neighbouring staff:

      'មៗ ” 7 ”” “” ”” ”២ៗ” 7 ”” ””“ ” ”“” ” ដល់មនុស្សជាច្រើននៅគ្រប់ជាតិសាសន៍'

    The junk is a run of SHORT tokens (1-3 chars, mostly non-Khmer symbols)
    followed by the actual lyric.  We find the first token long enough to be
    real words and keep from there, provided the tail clearly dominates.
    """
    tokens = s.split()
    if len(tokens) < 5:
        return s

    # Anchor on the first LONG token: real Khmer words run 5+ characters,
    # whereas note-head junk is always 1-3 characters.
    start = None
    for i, t in enumerate(tokens):
        if len(t) >= 5 and _khmer_char_count(t) >= 4:
            start = i
            break
    if start is None or start == 0:
        return s

    head_syls = len(_split_khmer_syllables(' '.join(tokens[:start])))
    tail_syls = len(_split_khmer_syllables(' '.join(tokens[start:])))
    if tail_syls >= 6 and tail_syls > head_syls * 2:
        return ' '.join(tokens[start:])
    return s


def _fix_ocr_text(text: str) -> str:
    """
    Post-process raw Tesseract output to fix the most common Khmer OCR errors:
      1. Unicode normalisation (NFC) — combining marks properly attached
      2. Re-join Khmer clusters that OCR split with spaces (correct spelling)
      3. Remove Cyrillic/Latin characters that sneak in between Khmer characters
      4. Strip OCR noise symbols and normalise Khmer punctuation
      5. Collapse multiple spaces
    """
    import unicodedata
    text = unicodedata.normalize("NFC", text)
    # Re-attach split clusters: "ស្ ត្រៃ", "ក្រ ែង", "អ ា ន" -> "ស្ត្រៃ", "ក្រែង", "អាន"
    text = _fix_khmer_spacing(text)
    # Remove stray non-Khmer chars embedded inside Khmer words
    text = _STRIP_NON_KHMER_IN_WORD_RE.sub('', text)
    # Normalise leftover ASCII punctuation back to Khmer forms
    text = _normalise_ocr_marks(text)
    # Collapse spaces
    text = re.sub(r' {2,}', ' ', text)
    return text


def _khmer_char_count_in(text: str) -> int:
    """Count total Khmer characters in a multi-line string."""
    return len(_KHMER_CHAR_RE.findall(text))


# Characters that never legitimately appear inside Khmer lyrics.
_JUNK_RE = re.compile(r'[=<>\*\^\|_~`\\/\[\]{}]')


def _pick_best_ocr_candidate(candidates: list[str]) -> str:
    """
    Choose the most trustworthy OCR candidate.

     1. Normalise each candidate (trim / collapse spaces) and drop junk-heavy
        ones that clearly failed (e.g. staff-line noise).
     2. Group identical / near-identical readings and score each group by
        (how many configs produced it, then its quality score).
     3. Return the best representative of the winning group.

    Independent agreement is a much stronger accuracy signal than length.
    """
    if not candidates:
        return ""

    cleaned = []
    for c in candidates:
        text = " ".join(c.split())
        if not text:
            continue
        # Reject candidates that are mostly junk (staff-line / symbol noise)
        junk = len(_JUNK_RE.findall(text))
        if junk and junk / max(len(text), 1) > 0.06:
            continue
        cleaned.append(text)

    if not cleaned:
        return max(candidates, key=_ocr_quality_score)

    # Group by similarity: same length +/- 2 and >= 85 % character match
    groups: list[list[str]] = []
    for text in cleaned:
        placed = False
        for g in groups:
            ref = g[0]
            if abs(len(text) - len(ref)) <= max(2, len(ref) // 12):
                same = sum(1 for a, b in zip(text, ref) if a == b)
                if same / max(len(ref), 1) >= 0.85:
                    g.append(text)
                    placed = True
                    break
        if not placed:
            groups.append([text])

    # Rank groups: quality FIRST, then agreement, then length.
    #
    # Vote count alone is not enough: two configs can agree on "បន្ទល" while a
    # third correctly reads "បន្ទូល".  The reading with the fewest broken
    # clusters (dropped vowel signs / coeng) is objectively better.  Agreement
    # is the tie-breaker, which is what makes a random OCR corruption lose.
    def group_rank(g: list[str]) -> tuple:
        best = max(g, key=_ocr_quality_score)
        return (_ocr_quality_score(best), len(g), len(best))

    groups.sort(key=group_rank, reverse=True)
    return max(groups[0], key=_ocr_quality_score)


def _broken_cluster_count(text: str) -> int:
    """
    Count Khmer clusters that are structurally INVALID — the signature of a
    dropped vowel sign.

    "បន្ទល" is what you get when OCR loses the ូ vowel from "បន្ទូល": the
    syllable ends in a bare consonant.  "ព្រះ" ends the same way and is fine,
    so this only counts clusters that are clearly malformed:

      • a coeng (្រ) not followed by a consonant
      • a dependent vowel sign appearing with NO base consonant before it
    """
    broken = 0
    for i, c in enumerate(text):
        if c == '\u17D2':
            if i + 1 >= len(text) or text[i + 1] not in _KHMER_CONSONANTS:
                broken += 1
        elif c in _KHMER_VOWELS_OR_SIGNS:
            j = i - 1
            while j >= 0 and text[j] in _KHMER_VOWELS_OR_SIGNS:
                j -= 1
            if j < 0 or text[j] not in _KHMER_CONSONANTS:
                if j < 0 or not text[j].isspace():
                    broken += 1
    return broken


def _ocr_quality_score(text: str) -> float:
    """
    Score an OCR candidate by how much REAL, WELL-FORMED Khmer it contains.

      • every Khmer character                        -> +1.0
      • every valid Khmer syllable                   -> +0.5
      • every broken cluster (dropped vowel/coeng)   -> -2.5
      • every structurally impossible cluster        -> -4.0
      • every ASCII letter/digit mixed in            -> -1.5
      • every junk symbol (= < > * | _ ~ / ...)      -> -3.0

    The broken/invalid penalties are what make "បន្ទូល" beat "បន្ទល" and
    "ទាំង" beat "ធទាំង".
    """
    khmer = _khmer_char_count_in(text)
    syllables = _syllable_evidence(text)
    broken = _broken_cluster_count(text)
    ascii_alnum = len(re.findall(r'[A-Za-z0-9]', text))
    junk = len(_JUNK_RE.findall(text))
    return (
        khmer * 1.0
        + syllables * 0.5
        - broken * 2.5
        - ascii_alnum * 1.5
        - junk * 3.0
    )


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


@lru_cache(maxsize=1)
def _script_khmer_model() -> tuple[str, str] | None:
    """Return (script_lang, tessdata_dir) if the script/Khmer model exists."""
    tess_prefix = os.path.dirname(os.path.dirname(_get_tesseract()))  # e.g. /opt/homebrew
    script_dir  = os.path.join(tess_prefix, "share", "tessdata")
    if os.path.exists(os.path.join(script_dir, "script", "Khmer.traineddata")):
        return "script/Khmer", script_dir
    return None


def _ocr_candidates(processed: 'Image.Image') -> list[str]:
    """
    Run every OCR strategy CONCURRENTLY and return the non-empty results.

    Tesseract releases the GIL inside its C extension, so running the four
    attempts in a thread pool is ~3-4x faster than the previous sequential
    loop, and the winning candidate is exactly the same.

      A. khm,            psm 7   (single line)
      B. khm,            psm 6   (uniform block)
      C. script/Khmer,   psm 6   (script-level model)
      D. khm+eng,        psm 6   (chord-row guard)
    """
    # Always run BOTH page-segmentation modes.  They do not agree, and the one
    # that reads a dropped vowel sign correctly differs per line (psm 7 gets
    # "បន្ទូល" where psm 6 gets "បន្ទល").  Running both in parallel costs no
    # wall-clock time and lets the quality scorer pick the complete reading.
    strategies: list[tuple[str, str, str | None]] = [
        ("khm", "7", None),
        ("khm", "6", None),
    ]
    script = _script_khmer_model()
    if script:
        strategies.append((script[0], "6", script[1]))
    strategies.append(("khm+eng", "6", None))

    # De-duplicate identical lang/psm pairs
    seen: set[tuple[str, str]] = set()
    uniq: list[tuple[str, str, str | None]] = []
    for lang, psm, tdir in strategies:
        key = (lang, psm)
        if key in seen:
            continue
        seen.add(key)
        uniq.append((lang, psm, tdir))

    from concurrent.futures import ThreadPoolExecutor
    with ThreadPoolExecutor(max_workers=len(uniq)) as pool:
        results = list(pool.map(
            lambda args: _ocr_attempt(processed, args[0], args[1], args[2]), uniq
        ))
    return [r for r in results if r]


# Below this quality score a strip is considered a failed OCR read and is
# retried at other scales.
_GOOD_ENOUGH = 25.0


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

    def _run(candidate_img: 'Image.Image') -> list[str]:
        cands = _ocr_candidates(candidate_img)
        if not cands:
            return []
        chosen = _pick_best_ocr_candidate(cands)
        return [
            _fix_verse_marker(_trim_edge_junk(_fix_ocr_text(ln.strip())))
            for ln in chosen.splitlines()
            if ln.strip() and _line_is_khmer(ln)
        ]

    # ── Multi-scale OCR with agreement voting ────
    #
    # Tesseract's Khmer model is sensitive to glyph size: the same line reads
    # correctly at one scale and gains a stray letter at another ("ទាំង" vs
    # "ធទាំង").  Which scale wins differs per strip, so no single default is
    # right.
    #
    # The reliable signal is AGREEMENT: a stray character introduced by a bad
    # scale appears in that scale only, while a genuine reading is reproduced
    # independently by the other scales.  So each scale votes, and the reading
    # with the most votes wins — the same principle used for the Tesseract
    # configs, applied across render scales.
    from concurrent.futures import ThreadPoolExecutor

    # Scores are compared, not counted, so the scales only need to be
    # genuinely different sizes.  1200 was measured to be consistently worse
    # and outvoted better readings, so it is not used.
    scales = (1700, 2400)

    def _try_scale(target_w: int) -> tuple[tuple[int, int], list[str]]:
        try:
            pre = _preprocess_for_ocr(img, target_width=target_w)
            return pre.size, _run(pre)
        except Exception:
            return (0, 0), []

    with ThreadPoolExecutor(max_workers=len(scales)) as pool:
        results = list(pool.map(_try_scale, scales))

    # If two scales produced identical images their readings are the same
    # opinion and must not be counted twice.
    seen_sizes: set[tuple[int, int]] = set()
    readings: list[list[str]] = []
    for size, lines in results:
        if not lines or size in seen_sizes:
            continue
        seen_sizes.add(size)
        readings.append(lines)
    if not readings:
        return ""
    if len(readings) == 1:
        return "\n".join(readings[0])

    # Group near-identical readings (same line count, >= 90 % char match).
    def _join(r: list[str]) -> str:
        return "\n".join(r)

    def _similar(a: list[str], b: list[str]) -> bool:
        if len(a) != len(b):
            return False
        for la, lb in zip(a, b):
            if abs(len(la) - len(lb)) > max(2, len(la) // 10):
                return False
            same = sum(1 for x, y in zip(la, lb) if x == y)
            if same / max(len(la), 1) < 0.90:
                return False
        return True

    groups: list[list[list[str]]] = []
    for r in readings:
        for g in groups:
            if _similar(g[0], r):
                g.append(r)
                break
        else:
            groups.append([r])

    # Quality first (fewest broken clusters, best density), then agreement.
    #
    # Agreement alone is not enough: several scales can agree on a corrupted
    # reading while one scale is correct.  The quality score is what separates
    # "ទាំង" from "ធទាំង"; agreement only breaks near-ties.
    def _rank(g: list[list[str]]) -> tuple:
        best = max(g, key=lambda r: _ocr_quality_score(_join(r)))
        return (_ocr_quality_score(_join(best)), len(g))

    groups.sort(key=_rank, reverse=True)
    winner = max(groups[0], key=lambda r: _ocr_quality_score(_join(r)))
    return _join(winner)


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

    from concurrent.futures import ThreadPoolExecutor
    strategies: list[tuple[str, str, str | None]] = [
        ("khm", "3", None),
        ("khm+eng", "3", None),
    ]
    if script_lang:
        strategies.append((script_lang, "3", script_dir))

    with ThreadPoolExecutor(max_workers=len(strategies)) as pool:
        results = list(pool.map(
            lambda a: _ocr_attempt(processed, a[0], a[1], a[2]), strategies
        ))
    candidates = [r for r in results if r]

    if not candidates:
        return ""

    best = _pick_best_ocr_candidate(candidates)
    lines = [
        _fix_verse_marker(_fix_ocr_text(ln.strip()))
        for ln in best.splitlines()
        if ln.strip() and _line_is_khmer(ln)
    ]
    return "\n".join(lines)


# ── Main extraction entry point ─────────────────────────────────

def _encode_png(img: Image.Image) -> str:
    """Encode a PIL image as a base64 PNG data URL."""
    buf = io.BytesIO()
    img.save(buf, format="PNG", optimize=True)
    return "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode("ascii")


def _encode_thumb(img: Image.Image, max_w: int = 1100) -> str:
    """Encode a page image as a reasonably-sized base64 JPEG for display."""
    page = img.copy()
    if page.width > max_w:
        h = int(page.height * (max_w / page.width))
        page = page.resize((max_w, max(h, 1)), Image.LANCZOS)
    buf = io.BytesIO()
    page.convert("RGB").save(buf, format="JPEG", quality=82, optimize=True)
    return "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode("ascii")


def _analyze_page(page_rgb: Image.Image) -> dict:
    """
    OCR only the Khmer lyric rows that sit UNDER the melody notation.

    The song-title block at the top (title / Korean subtitle / hymn number) is
    deliberately NOT OCR'd. It is returned as an image crop so it can be placed
    on the first slide instead of polluting the lyric text box.
    """
    from concurrent.futures import ThreadPoolExecutor as _TPE

    info = _classify_strips(page_rgb)
    strips = info["lyrics"]

    if strips:
        # OCR every lyric strip of this page concurrently, keep paper order
        with _TPE(max_workers=min(8, len(strips))) as pool:
            texts = list(pool.map(_ocr_strip, strips))
        lines: list[str] = []
        for text in texts:
            lines.extend(ln for ln in text.splitlines() if ln.strip())
    else:
        # Fallback: full-page OCR filtered to Khmer only
        text = _ocr_full_page_khmer(page_rgb)
        lines = [ln for ln in text.splitlines() if ln.strip()]

    return {
        "lines": lines,
        "title": info.get("title"),
        "has_music": info.get("has_music", False),
    }


def extract_lyrics_from_bytes(data: bytes, ext: str) -> dict:
    """
    Accept raw file bytes + extension (.pdf / .png / .jpg / .jpeg).

    Returns a dict:
      text        - ONLY the Khmer lyric lines found under the melody
      pages       - base64 JPEG previews of every page (so the UI can show
                    the uploaded PDF content)
      title_image - base64 PNG of the detected song-title block (first slide)
      page_count  - number of pages
      has_music   - whether melody notation was found on any page

    Pipeline per page:
      1. Convert PDF -> RGB at 150 dpi (accuracy sweet spot for Khmer)
      2. Classify strips: title block / notation rows / lyric rows
      3. OCR ONLY the lyric rows under the melody (title block excluded)
      4. Run filter_lyrics() as a final pass to remove any remaining junk
    """
    from concurrent.futures import ThreadPoolExecutor

    empty = {"pages": [], "title_image": "", "page_count": 0, "has_music": False}

    if ext == ".pdf":
        try:
            from pdf2image import convert_from_bytes
            # 150 dpi is the accuracy sweet spot for Khmer: high enough to
            # resolve vowel signs, low enough that Tesseract's LSTM engine does
            # not choke on an oversized page (300 dpi produced garbage).
            pages = convert_from_bytes(data, dpi=150)
        except Exception as e:
            return {"text": f"[PDF conversion error: {e}]", **empty}
        page_images = [p.convert("RGB") for p in pages]
    else:
        try:
            page_images = [Image.open(io.BytesIO(data)).convert("RGB")]
        except Exception as e:
            return {"text": f"[Image error: {e}]", **empty}

    # Pages are independent -> process in parallel, preserve original order
    with ThreadPoolExecutor(max_workers=min(4, max(1, len(page_images)))) as pool:
        parsed = list(pool.map(_analyze_page, page_images))

    all_lines: list[str] = []
    for p in parsed:
        all_lines.extend(p["lines"])
        all_lines.append("")   # separator between pages

    title_crop = next((p["title"] for p in parsed if p.get("title")), None)

    filtered = filter_lyrics("\n".join(all_lines)).strip()
    if not filtered:
        filtered = "(no Khmer text detected)"

    return {
        "text": filtered,
        "pages": [_encode_thumb(img) for img in page_images],
        "title_image": _encode_png(title_crop) if title_crop is not None else "",
        "page_count": len(page_images),
        "has_music": any(p.get("has_music") for p in parsed),
    }


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
    verse_num_color: str  = "#ef4444"
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
    # Verse numbers default to red, matching the chorus 'R.' marker
    vnum_rgb = _hex_to_rgb(getattr(s, 'verse_num_color', '#ef4444'))

    pad = Inches(s.padding)
    sw  = prs.slide_width
    sh  = prs.slide_height

    slides_data = parse_song_lyrics(lyrics_text)
    if not slides_data:
        slides_data = [{'is_title': False, 'prefix': '', 'prefix_color': '', 'lines': ['(no lyrics found)']}]
    # NOTE: one parsed verse/chorus == one slide.  Text is NOT split across
    # continuation slides — a verse always stays together on a single page.

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
    # One verse/chorus == one slide (no continuation splitting).

    # In preview, attach title/reference to slides
    first_title = slides[0].get("title", "") if (slides and slides[0].get("is_title")) else song_title
    for i, s_item in enumerate(slides):
        if not s_item.get("is_title"):
            s_item["title"] = song_title or first_title
    return slides


# ── Song OCR endpoint ────────────────────────────────────────────
@app.post("/api/song/extract")
async def song_extract(file: UploadFile = File(...)):
    """
    Accept a PDF / image / text upload and return extracted lyrics as JSON.

    Response: { text, pages[], title_image, page_count, has_music }
      • text         — Khmer lyrics found UNDER the melody (title block excluded)
      • pages        — base64 page previews so the UI can display the PDF
      • title_image  — base64 crop of the song-title block for the first slide
    """
    ext = Path(file.filename or "").suffix.lower()
    if ext not in (".pdf", ".png", ".jpg", ".jpeg", ".webp", ".txt"):
        return {"error": "Only PDF, PNG, JPG, WEBP and TXT files are supported."}, 400

    data = await file.read()

    # Plain text: no OCR, no page image
    if ext == ".txt":
        try:
            return {
                "text": data.decode("utf-8", errors="replace"),
                "pages": [], "title_image": "", "page_count": 0, "has_music": False,
            }
        except Exception as e:
            return {"error": f"Could not read text file: {e}"}, 400

    return extract_lyrics_from_bytes(data, ext)


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
#  Slide Editor — bulk-restyle EVERY slide of an uploaded PPTX
#  (background, font, size, colour, outline, alignment, spacing …)
# ═══════════════════════════════════════════════════════════════════

class SlideRestyleSettings(BaseModel):
    """Formatting applied to every slide in the deck.  All fields optional:
    a field left as ``None`` is NOT changed."""
    # Background
    bg_color:      Optional[str]   = None
    bg_image:      Optional[str]   = None   # data URL, applied to every slide
    # Text
    font_family:   Optional[str]   = None
    font_size:     Optional[int]   = Field(None, ge=6, le=200)
    text_color:    Optional[str]   = None
    bold:          Optional[bool]  = None
    italic:        Optional[bool]  = None
    underline:     Optional[bool]  = None
    align:         Optional[str]   = None   # left | center | right | justify
    line_spacing:  Optional[float] = Field(None, ge=0.5, le=4.0)
    # Text outline (PowerPoint "text outline")
    outline_color: Optional[str]   = None
    outline_width: Optional[float] = Field(None, ge=0.0, le=20.0)   # points
    # Text shadow
    shadow:        Optional[bool]  = None
    # Scaling of the whole text box content
    scale_font_by_pct: Optional[float] = Field(None, ge=25, le=400)
    # Vertical anchoring of text inside its box: top | middle | bottom
    vertical_anchor: Optional[str]   = None


class SlideRestyleRequest(BaseModel):
    settings: SlideRestyleSettings = SlideRestyleSettings()
    # Restrict changes to selected slide indices (1-based). Empty = all.
    slides: list[int] = []


# ── Low-level run / paragraph formatting helpers ──────────────────

def _iter_slide_runs(prs):
    """Yield (slide, shape, paragraph, run) for every text run in the deck."""
    for slide in prs.slides:
        for shape in slide.shapes:
            if not getattr(shape, "has_text_frame", False):
                continue
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    yield slide, shape, para, run


def _set_run_outline(run, color_hex: str, width_pt: float) -> None:
    """
    Apply a PowerPoint text OUTLINE (not a shape outline) to a run.

    python-pptx has no API for this, so the <a:ln> element is written
    directly on the run properties, followed by a solid fill.
    """
    rPr = run._r.get_or_add_rPr()
    for el in rPr.findall(qn("a:ln")):
        rPr.remove(el)

    ln = etree.SubElement(rPr, qn("a:ln"))
    ln.set("w", str(int(max(width_pt, 0.25) * 12700)))   # points -> EMU
    fill = etree.SubElement(ln, qn("a:solidFill"))
    clr = etree.SubElement(fill, qn("a:srgbClr"))
    clr.set("val", color_hex.lstrip("#").upper()[:6])
    # Round joins look far better on Khmer glyphs than mitered corners
    etree.SubElement(ln, qn("a:round"))

    # <a:ln> must sit before the fill elements in rPr — python-pptx appends in
    # schema order for solidFill, so re-insert ln at the correct position.
    rPr.remove(ln)
    insert_at = 0
    for i, child in enumerate(list(rPr)):
        tag = etree.QName(child).localname
        if tag in ("ln", "noFill", "solidFill", "gradFill", "blipFill",
                   "pattFill", "grpFill", "effectLst", "effectDag",
                   "highlight", "uLnTx", "uLn", "uFillTx", "uFill"):
            break
        insert_at = i + 1
    rPr.insert(insert_at, ln)


def _set_run_shadow(run) -> None:
    """Add a subtle drop shadow to a run's <a:effectLst>."""
    rPr = run._r.get_or_add_rPr()
    for el in rPr.findall(qn("a:effectLst")):
        rPr.remove(el)
    eff = etree.SubElement(rPr, qn("a:effectLst"))
    shdw = etree.SubElement(eff, qn("a:outerShdw"))
    shdw.set("blurRad", "38100")
    shdw.set("dist", "25400")
    shdw.set("dir", "2700000")
    shdw.set("algn", "tl")
    shdw.set("rotWithShape", "0")
    clr = etree.SubElement(shdw, qn("a:srgbClr"))
    clr.set("val", "000000")
    alpha = etree.SubElement(clr, qn("a:alpha"))
    alpha.set("val", "45000")


def _apply_slide_background(slide, color_hex: str | None) -> None:
    """Solid-fill the slide background."""
    if not color_hex:
        return
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(*_hex_to_rgb(color_hex))


def _add_background_image(prs, slide, image_bytes: bytes) -> None:
    """Place an image covering the entire slide, behind existing content."""
    from pptx.util import Emu
    pic = slide.shapes.add_picture(
        io.BytesIO(image_bytes), 0, 0,
        width=Emu(prs.slide_width), height=Emu(prs.slide_height),
    )
    # Move the picture to the very back of the z-order: right after the
    # mandatory nvGrpSpPr / grpSpPr elements, whatever their position.
    sp_tree = slide.shapes._spTree
    sp_tree.remove(pic._element)
    insert_at = 0
    for i, child in enumerate(list(sp_tree)):
        tag = etree.QName(child).localname
        if tag in ("nvGrpSpPr", "grpSpPr"):
            insert_at = i + 1
    sp_tree.insert(insert_at, pic._element)


def restyle_pptx(data: bytes, s: SlideRestyleSettings, only_slides: list[int] | None = None) -> bytes:
    """Apply bulk formatting to an existing PPTX and return the new bytes."""
    prs = Presentation(io.BytesIO(data))

    # Decode the optional background image once
    bg_img_bytes = None
    if s.bg_image:
        try:
            payload = s.bg_image.split(",", 1)[1] if "," in s.bg_image else s.bg_image
            bg_img_bytes = base64.b64decode(payload)
        except Exception:
            bg_img_bytes = None

    targets = set(only_slides or [])

    for idx, slide in enumerate(prs.slides, start=1):
        if targets and idx not in targets:
            continue

        # ── Background ──────────────────────────────────────────────
        if bg_img_bytes:
            _add_background_image(prs, slide, bg_img_bytes)
        elif s.bg_color:
            _apply_slide_background(slide, s.bg_color)

        # ── Text formatting ─────────────────────────────────────────
        for shape in slide.shapes:
            if not getattr(shape, "has_text_frame", False):
                continue
            if s.vertical_anchor:
                from pptx.enum.text import MSO_VERTICAL_ANCHOR
                anchor_map = {
                    "top": MSO_VERTICAL_ANCHOR.TOP,
                    "middle": MSO_VERTICAL_ANCHOR.MIDDLE,
                    "bottom": MSO_VERTICAL_ANCHOR.BOTTOM,
                }
                a = anchor_map.get(s.vertical_anchor.lower())
                if a is not None:
                    shape.text_frame.vertical_anchor = a
                    # PowerPoint honours anchorCtr for true horizontal centring
                    shape.text_frame.word_wrap = True

            for para in shape.text_frame.paragraphs:
                if s.align:
                    para.alignment = ALIGN_MAP.get(s.align)
                if s.line_spacing is not None:
                    _set_line_spacing(para, s.line_spacing)

                for run in para.runs:
                    if not run.text.strip():
                        continue
                    if s.font_family:
                        run.font.name = s.font_family
                        _fix_font(run, s.font_family)
                    if s.font_size is not None:
                        run.font.size = Pt(s.font_size)
                    elif s.scale_font_by_pct:
                        base = run.font.size
                        if base is not None:
                            run.font.size = Pt(
                                max(6, int(base.pt * s.scale_font_by_pct / 100))
                            )
                    if s.font_size is not None and s.scale_font_by_pct:
                        run.font.size = Pt(max(6, int(s.font_size * s.scale_font_by_pct / 100)))
                    if s.text_color:
                        run.font.color.rgb = RGBColor(*_hex_to_rgb(s.text_color))
                    if s.bold is not None:
                        run.font.bold = s.bold
                    if s.italic is not None:
                        run.font.italic = s.italic
                    if s.underline is not None:
                        run.font.underline = s.underline
                    if s.outline_color and (s.outline_width or 0) > 0:
                        _set_run_outline(run, s.outline_color, s.outline_width)
                    if s.shadow:
                        _set_run_shadow(run)

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out.read()


def inspect_pptx(data: bytes) -> dict:
    """Summarise the slides so the editor UI can show thumbnails + counts."""
    prs = Presentation(io.BytesIO(data))
    slides_info = []
    for i, slide in enumerate(prs.slides, start=1):
        text_shapes = 0
        pictures = 0
        for shape in slide.shapes:
            if getattr(shape, "has_text_frame", False) and shape.text_frame.text.strip():
                text_shapes += 1
            if shape.shape_type == 13:   # PICTURE
                pictures += 1
        slides_info.append({
            "index": i,
            "text_shapes": text_shapes,
            "pictures": pictures,
        })
    return {
        "slide_count": len(slides_info),
        "slides": slides_info,
        "width": prs.slide_width,
        "height": prs.slide_height,
        "aspect": round(prs.slide_width / prs.slide_height, 4) if prs.slide_height else 1.7778,
    }


@app.post("/api/slides/inspect")
async def slides_inspect(file: UploadFile = File(...)):
    """Return slide count + per-slide summary + rendered thumbnails."""
    if Path(file.filename or "").suffix.lower() != ".pptx":
        raise HTTPException(status_code=400, detail="Only .pptx files are supported.")
    data = await file.read()
    try:
        info = inspect_pptx(data)
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Could not read PPTX: {e}")
    info["thumbnails"] = _render_slides_pillow_fallback(data, width_px=900)
    info["filename"] = file.filename
    return info


@app.post("/api/slides/apply")
async def slides_apply(
    file: UploadFile = File(...),
    settings: str = Form("{}"),
    slides: str = Form("[]"),
    preview: str = Form("false"),
):
    """
    Apply bulk formatting to every (or selected) slide and return the PPTX.

    When ``preview=true`` the result is NOT downloaded; it is returned as
    base64 thumbnails so the UI can show the change before saving.
    """
    import json as _json

    if Path(file.filename or "").suffix.lower() != ".pptx":
        raise HTTPException(status_code=400, detail="Only .pptx files are supported.")

    data = await file.read()
    try:
        s = SlideRestyleSettings(**_json.loads(settings or "{}"))
        only = _json.loads(slides or "[]")
        only = [int(x) for x in only] if only else []
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Bad settings: {e}")

    try:
        new_bytes = restyle_pptx(data, s, only)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Restyle failed: {e}")

    if preview.lower() in ("true", "1", "yes"):
        return {
            "slide_count": len(Presentation(io.BytesIO(new_bytes)).slides),
            "thumbnails": _render_slides_pillow_fallback(new_bytes, width_px=900),
        }

    return StreamingResponse(
        io.BytesIO(new_bytes),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": "attachment; filename=\"edited_slides.pptx\""},
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
            # BACKGROUND(5) and NONE have no fore_color
            if fill.type is None or fill.type.name in ("BACKGROUND", "NONE"):
                return None
            c = fill.fore_color.rgb
            return (c.red, c.green, c.blue)
        except Exception:
            return None

    def _slide_bg_rgb(slide) -> tuple | None:
        """
        Read the slide's explicit background colour from the XML.

        slide.background.fill is unreliable: for an explicitly-set <p:bg> the
        fill type can still report BACKGROUND, so the renderer would fall back
        to white and the thumbnail would not match the real deck.  Reading
        srgbClr directly from <p:bg><p:bgPr><a:solidFill> is exact.
        """
        try:
            csld = slide._element.find(qn("p:cSld"))
            if csld is None:
                return None
            bg = csld.find(qn("p:bg"))
            if bg is None:
                return None
            bg_pr = bg.find(qn("p:bgPr"))
            if bg_pr is None:
                return None
            solid = bg_pr.find(qn("a:solidFill"))
            if solid is None:
                return None
            clr = solid.find(qn("a:srgbClr"))
            if clr is None:
                return None
            val = (clr.get("val") or "").strip()
            if len(val) != 6:
                return None
            return tuple(int(val[i:i + 2], 16) for i in (0, 2, 4))
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
            bg_color = _slide_bg_rgb(slide) or _safe_rgb(slide.background.fill)
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
                # 1) normal API path (an explicit srgbClr)
                try:
                    c = run.font.color.rgb
                    if c is not None:
                        return (c.red, c.green, c.blue)
                except Exception:
                    pass
                # 2) read srgbClr straight from the run's XML.  Needed because
                # font.color.rgb raises for some explicit solid fills, which
                # would silently render the text black.
                try:
                    rPr = run._r.find(qn("a:rPr"))
                    if rPr is not None:
                        solid = rPr.find(qn("a:solidFill"))
                        if solid is not None:
                            clr = solid.find(qn("a:srgbClr"))
                            if clr is not None:
                                val = (clr.get("val") or "").strip()
                                if len(val) == 6:
                                    return tuple(int(val[i:i + 2], 16) for i in (0, 2, 4))
                except Exception:
                    pass
                return para_default

            # Vertical anchoring: PowerPoint centres text inside the box when
            # anchor is MIDDLE.  Without this the thumbnail always draws from
            # the top of the shape, which looks nothing like the real slide.
            anchor = None
            try:
                a = tf.vertical_anchor
                anchor = a.name if a is not None else None
            except Exception:
                anchor = None

            total_h = 0
            if anchor == "MIDDLE":
                for para in tf.paragraphs:
                    runs_p = [r for r in para.runs if r.text]
                    if not runs_p:
                        total_h += fpt(_para_font_pt(para) * 1.2)
                        continue
                    try:
                        pt = runs_p[0].font.size.pt if runs_p[0].font.size else _para_font_pt(para)
                    except Exception:
                        pt = _para_font_pt(para)
                    total_h += int(fpt(pt) * 1.35)

            line_y = sy
            if anchor == "MIDDLE" and total_h < sh:
                line_y = sy + (sh - total_h) // 2
            elif anchor == "BOTTOM":
                line_y = sy + max(0, sh - total_h)

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
