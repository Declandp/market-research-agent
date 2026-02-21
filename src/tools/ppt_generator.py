"""
PowerPoint Report Generator
============================
Converts the markdown competitive intelligence report into a
professional PowerPoint presentation.
BCG-inspired consulting theme: clean, minimal, authoritative.
"""

import os
import re
from datetime import datetime

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE

# ── BCG-Inspired Theme ───────────────────────────────────
# Clean whites, deep greens, minimal accents
BG_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BG_OFF_WHITE = RGBColor(0xFA, 0xFA, 0xFA)
TITLE_BG = RGBColor(0x00, 0x3B, 0x2D)         # BCG deep green
TITLE_BG_ALT = RGBColor(0x00, 0x2E, 0x23)     # Darker green

BLACK = RGBColor(0x1A, 0x1A, 0x1A)
DARK_TEXT = RGBColor(0x2D, 0x2D, 0x2D)
MED_TEXT = RGBColor(0x5A, 0x5A, 0x5A)
LIGHT_TEXT = RGBColor(0x8C, 0x8C, 0x8C)

BCG_GREEN = RGBColor(0x00, 0x6B, 0x4F)        # Primary accent
BCG_GREEN_LIGHT = RGBColor(0xE6, 0xF2, 0xEF)  # Light green tint
BCG_GREEN_MED = RGBColor(0x00, 0x8A, 0x68)    # Medium green
ACCENT_RED = RGBColor(0xC4, 0x2B, 0x2B)       # For threats/warnings
ACCENT_AMBER = RGBColor(0xB8, 0x6E, 0x00)     # For caution
ACCENT_TEAL = RGBColor(0x00, 0x7A, 0x87)      # For opportunities

RULE_COLOR = RGBColor(0xD4, 0xD4, 0xD4)       # Thin divider lines
CARD_BG = RGBColor(0xF5, 0xF7, 0xF6)          # Subtle card background
TABLE_HEADER = RGBColor(0x00, 0x3B, 0x2D)
TABLE_ROW = RGBColor(0xFF, 0xFF, 0xFF)
TABLE_ALT = RGBColor(0xF5, 0xF7, 0xF6)

FONT_TITLE = "Georgia"
FONT_BODY = "Calibri"
SLIDE_WIDTH = Inches(13.333)
SLIDE_HEIGHT = Inches(7.5)


def _get_blank_layout(prs):
    try:
        return prs.slide_layouts[6]
    except IndexError:
        return prs.slide_layouts[-1]


def _clean(text: str) -> str:
    """Strip markdown formatting."""
    text = re.sub(r'\*\*(.+?)\*\*', r'\1', text)
    text = re.sub(r'\*(.+?)\*', r'\1', text)
    text = re.sub(r'`(.+?)`', r'\1', text)
    text = re.sub(r'\[(.+?)\]\(.+?\)', r'\1', text)
    return text.strip()


def _bg(slide, color=BG_WHITE):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def _rect(slide, left, top, w, h, color):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, w, h)
    s.fill.solid()
    s.fill.fore_color.rgb = color
    s.line.fill.background()
    return s


def _line(slide, left, top, width):
    """Add a thin horizontal rule."""
    return _rect(slide, left, top, width, Inches(0.015), RULE_COLOR)


def _text(slide, left, top, w, h, text, size=16, color=DARK_TEXT,
          bold=False, align=PP_ALIGN.LEFT, font=FONT_BODY):
    txBox = slide.shapes.add_textbox(left, top, w, h)
    tf = txBox.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    p = tf.paragraphs[0]
    p.text = _clean(text)
    p.font.size = Pt(size)
    p.font.color.rgb = color
    p.font.bold = bold
    p.font.name = font
    p.alignment = align
    return txBox


def _bullets(tf, items, size=13, color=DARK_TEXT, label_color=BCG_GREEN):
    """Add bullet list. Full text, no truncation."""
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        clean = item.strip().lstrip("*-•")
        clean = re.sub(r'^\d+[\.\)]\s*', '', clean)
        clean = _clean(clean)

        p.font.size = Pt(size)
        p.font.color.rgb = color
        p.font.name = FONT_BODY
        p.space_before = Pt(2)
        p.space_after = Pt(6)

        if ":" in clean:
            parts = clean.split(":", 1)
            r1 = p.add_run()
            r1.text = parts[0] + ":"
            r1.font.size = Pt(size)
            r1.font.color.rgb = label_color
            r1.font.bold = True
            r1.font.name = FONT_BODY
            if len(parts) > 1:
                r2 = p.add_run()
                r2.text = " " + parts[1].strip()
                r2.font.size = Pt(size)
                r2.font.color.rgb = color
                r2.font.name = FONT_BODY
        else:
            p.text = clean


def _section_header(slide, text):
    """Consulting-style section header: thin green bar + title + rule."""
    _rect(slide, Inches(0.8), Inches(0.45), Inches(0.05), Inches(0.4), BCG_GREEN)
    _text(slide, Inches(1.0), Inches(0.35), Inches(10), Inches(0.6),
          text, size=24, color=BLACK, bold=True, font=FONT_TITLE)
    _line(slide, Inches(0.8), Inches(1.0), Inches(11.7))


def _parse_sections(report_text: str) -> dict:
    sections = {}
    current_key = "intro"
    current_content = []
    for line in report_text.split("\n"):
        if line.startswith("## "):
            if current_content:
                sections[current_key] = "\n".join(current_content)
            current_key = re.sub(r'^#+\s*', '', line).strip().rstrip(".")
            current_content = []
        elif line.startswith("# ") and current_key == "intro":
            sections["title"] = re.sub(r'^#+\s*', '', line).strip()
        else:
            current_content.append(line)
    if current_content:
        sections[current_key] = "\n".join(current_content)
    return sections


def _get_bullets(text: str) -> list[str]:
    bullets = []
    for line in text.split("\n"):
        s = line.strip()
        if s and (s.startswith(("-", "*", "•")) or re.match(r'^\d+[\.\)]', s)):
            c = re.sub(r'^[-*•]\s*', '', s)
            c = re.sub(r'^\d+[\.\)]\s*', '', c)
            c = _clean(c)
            if c and len(c) > 3:
                bullets.append(c)
    return bullets


def _get_table(text: str):
    headers, rows = [], []
    for line in text.split("\n"):
        s = line.strip()
        if "|" in s and not re.match(r'^\|[\s\-:|]+\|$', s):
            cells = [c.strip() for c in s.split("|") if c.strip()]
            if not headers:
                headers = cells
            else:
                rows.append(cells)
    return headers, rows


# ── Slide Builders ────────────────────────────────────────

def _slide_title(prs, company, competitors, date_str):
    slide = prs.slides.add_slide(_get_blank_layout(prs))
    _bg(slide, TITLE_BG)

    # Subtle top line
    _rect(slide, Inches(0), Inches(0), SLIDE_WIDTH, Inches(0.04),
          RGBColor(0x00, 0x8A, 0x68))

    _text(slide, Inches(1.5), Inches(1.6), Inches(10), Inches(0.6),
          "Competitive Intelligence Report", size=16,
          color=RGBColor(0x7F, 0xBF, 0xAD), font=FONT_BODY)

    _text(slide, Inches(1.5), Inches(2.2), Inches(10), Inches(1.5),
          company, size=44, color=BG_WHITE, bold=True, font=FONT_TITLE,
          align=PP_ALIGN.LEFT)

    # Thin white rule
    _rect(slide, Inches(1.5), Inches(3.9), Inches(3), Inches(0.02), BG_WHITE)

    _text(slide, Inches(1.5), Inches(4.2), Inches(10), Inches(0.5),
          f"Competitors: {', '.join(competitors)}", size=14,
          color=RGBColor(0x7F, 0xBF, 0xAD), font=FONT_BODY)

    _text(slide, Inches(1.5), Inches(4.8), Inches(10), Inches(0.4),
          date_str, size=12, color=RGBColor(0x5A, 0x9A, 0x8A), font=FONT_BODY)

    # Confidential footer
    _text(slide, Inches(1.5), Inches(6.5), Inches(10), Inches(0.3),
          "CONFIDENTIAL", size=9, color=RGBColor(0x5A, 0x9A, 0x8A),
          font=FONT_BODY, bold=True)

    _rect(slide, Inches(0), Inches(7.46), SLIDE_WIDTH, Inches(0.04),
          RGBColor(0x00, 0x8A, 0x68))


def _slide_exec_summary(prs, section_text):
    slide = prs.slides.add_slide(_get_blank_layout(prs))
    _bg(slide)
    _section_header(slide, "Executive Summary")

    bullets = _get_bullets(section_text)
    if not bullets:
        bullets = [_clean(l.strip()) for l in section_text.split("\n")
                   if l.strip() and not l.strip().startswith("#")]

    if len(bullets) > 5:
        mid = len(bullets) // 2
        # Left column
        txBox = slide.shapes.add_textbox(
            Inches(0.8), Inches(1.3), Inches(5.7), Inches(5.7))
        tf = txBox.text_frame
        tf.word_wrap = True
        _bullets(tf, bullets[:mid], size=13)

        # Vertical divider
        _rect(slide, Inches(6.65), Inches(1.3), Inches(0.015), Inches(5.5), RULE_COLOR)

        # Right column
        txBox = slide.shapes.add_textbox(
            Inches(6.9), Inches(1.3), Inches(5.7), Inches(5.7))
        tf = txBox.text_frame
        tf.word_wrap = True
        _bullets(tf, bullets[mid:mid+6], size=13)
    else:
        txBox = slide.shapes.add_textbox(
            Inches(0.8), Inches(1.3), Inches(11.7), Inches(5.7))
        tf = txBox.text_frame
        tf.word_wrap = True
        _bullets(tf, bullets[:8], size=14)


def _slide_market(prs, section_text):
    slide = prs.slides.add_slide(_get_blank_layout(prs))
    _bg(slide)
    _section_header(slide, "Market Landscape")

    paragraphs = [p.strip() for p in section_text.split("\n")
                  if p.strip() and not p.strip().startswith("#")]
    bullets = _get_bullets(section_text)

    body = []
    for p in paragraphs:
        if not p.startswith(("-", "*", "•")) and not re.match(r'^\d+[\.\)]', p):
            c = _clean(p)
            if c and len(c) > 10:
                body.append(c)

    if body:
        _text(slide, Inches(0.8), Inches(1.3), Inches(5.5), Inches(5.5),
              "\n\n".join(body[:4]), size=12, color=MED_TEXT)

    if bullets:
        # Vertical divider
        _rect(slide, Inches(6.65), Inches(1.3), Inches(0.015), Inches(5.5), RULE_COLOR)

        _text(slide, Inches(7.0), Inches(1.3), Inches(5), Inches(0.4),
              "Key Trends", size=16, color=BCG_GREEN, bold=True, font=FONT_TITLE)
        _line(slide, Inches(7.0), Inches(1.75), Inches(5.5))

        txBox = slide.shapes.add_textbox(
            Inches(7.0), Inches(1.95), Inches(5.5), Inches(5.0))
        tf = txBox.text_frame
        tf.word_wrap = True
        _bullets(tf, bullets[:6], size=12)


def _slide_competitor(prs, name, section_text):
    slide = prs.slides.add_slide(_get_blank_layout(prs))
    _bg(slide)
    _section_header(slide, f"Competitor: {name}")

    # Threat badge
    threat_match = re.search(r'[Tt]hreat\s*[Ll]evel.*?(\d+)/10', section_text)
    threat_level = int(threat_match.group(1)) if threat_match else 5
    badge_color = BCG_GREEN if threat_level <= 3 else (ACCENT_AMBER if threat_level <= 6 else ACCENT_RED)
    _rect(slide, Inches(10.8), Inches(0.35), Inches(1.8), Inches(0.55), badge_color)
    _text(slide, Inches(10.8), Inches(0.38), Inches(1.8), Inches(0.5),
          f"Threat: {threat_level}/10", size=13, color=BG_WHITE, bold=True,
          align=PP_ALIGN.CENTER)

    bullets = _get_bullets(section_text)
    if not bullets:
        bullets = [_clean(l.strip()) for l in section_text.split("\n")
                   if l.strip() and not l.strip().startswith("#")]

    mid = max(1, len(bullets) // 2)

    # Left: Overview
    _text(slide, Inches(0.8), Inches(1.3), Inches(5), Inches(0.4),
          "Overview", size=16, color=BCG_GREEN, bold=True, font=FONT_TITLE)
    _line(slide, Inches(0.8), Inches(1.75), Inches(5.5))
    if bullets[:mid]:
        txBox = slide.shapes.add_textbox(
            Inches(0.8), Inches(1.95), Inches(5.5), Inches(5.0))
        tf = txBox.text_frame
        tf.word_wrap = True
        _bullets(tf, bullets[:mid][:7], size=12)

    # Vertical divider
    _rect(slide, Inches(6.65), Inches(1.3), Inches(0.015), Inches(5.5), RULE_COLOR)

    # Right: Strengths & Weaknesses
    _text(slide, Inches(7.0), Inches(1.3), Inches(5), Inches(0.4),
          "Strengths & Weaknesses", size=16, color=BCG_GREEN, bold=True, font=FONT_TITLE)
    _line(slide, Inches(7.0), Inches(1.75), Inches(5.5))
    if bullets[mid:]:
        txBox = slide.shapes.add_textbox(
            Inches(7.0), Inches(1.95), Inches(5.5), Inches(5.0))
        tf = txBox.text_frame
        tf.word_wrap = True
        _bullets(tf, bullets[mid:][:7], size=12)


def _slide_matrix(prs, section_text):
    slide = prs.slides.add_slide(_get_blank_layout(prs))
    _bg(slide)
    _section_header(slide, "Competitive Comparison")

    headers, rows = _get_table(section_text)

    if not headers or not rows:
        bullets = _get_bullets(section_text)
        if bullets:
            txBox = slide.shapes.add_textbox(
                Inches(0.8), Inches(1.3), Inches(11.7), Inches(5.7))
            tf = txBox.text_frame
            tf.word_wrap = True
            _bullets(tf, bullets[:10], size=13)
        return

    nc = len(headers)
    nr = min(len(rows) + 1, 8)
    rows = rows[:nr - 1]

    tw = Inches(min(11.5, nc * 2.2))
    th = Inches(0.5 * nr)
    ts = slide.shapes.add_table(nr, nc, Inches(0.8), Inches(1.3), tw, th)
    table = ts.table

    cw = int(tw / nc)
    for col in table.columns:
        col.width = cw

    for i, h in enumerate(headers):
        cell = table.cell(0, i)
        cell.text = _clean(h)
        cell.fill.solid()
        cell.fill.fore_color.rgb = TABLE_HEADER
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        for p in cell.text_frame.paragraphs:
            p.font.size = Pt(11)
            p.font.color.rgb = BG_WHITE
            p.font.bold = True
            p.font.name = FONT_BODY

    for r, rd in enumerate(rows):
        bg = TABLE_ROW if r % 2 == 0 else TABLE_ALT
        for c in range(nc):
            cell = table.cell(r + 1, c)
            cell.text = _clean(rd[c]) if c < len(rd) else ""
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            for p in cell.text_frame.paragraphs:
                p.font.size = Pt(10)
                p.font.color.rgb = DARK_TEXT
                p.font.name = FONT_BODY


def _slide_swot(prs, section_text):
    slide = prs.slides.add_slide(_get_blank_layout(prs))
    _bg(slide)
    _section_header(slide, "SWOT Analysis")

    swot = {"Strengths": [], "Weaknesses": [], "Opportunities": [], "Threats": []}
    current = None

    for line in section_text.split("\n"):
        stripped = line.strip()
        lower = stripped.lower()

        matched = None
        for key in swot:
            if key.lower() in lower:
                matched = key
                break

        if matched:
            cm = re.search(r'(?:strengths|weaknesses|opportunities|threats)\s*:?\s*:(.+)',
                           stripped, re.IGNORECASE)
            if cm:
                items = [i.strip().rstrip(".") for i in re.split(r'[,;]', cm.group(1)) if i.strip()]
                swot[matched].extend(items)
                current = matched
            elif ":" in stripped or stripped.startswith("#") or stripped.startswith("*"):
                current = matched
        elif current and stripped.startswith(("-", "*", "•")):
            clean = re.sub(r'^[-*•]\s*', '', stripped)
            clean = _clean(clean).rstrip(".")
            sub = None
            for key in swot:
                if clean.lower().startswith(key.lower()):
                    sub = key
                    break
            if sub:
                cp = clean.find(":")
                if cp > 0:
                    items = [i.strip().rstrip(".") for i in re.split(r'[,;]', clean[cp+1:]) if i.strip()]
                    swot[sub].extend(items)
            elif clean and len(clean) > 3:
                swot[current].append(clean)

    grid = [
        (Inches(0.8),  Inches(1.3), "Strengths",     BCG_GREEN,    swot["Strengths"]),
        (Inches(6.8),  Inches(1.3), "Weaknesses",    ACCENT_RED,   swot["Weaknesses"]),
        (Inches(0.8),  Inches(4.2), "Opportunities", ACCENT_TEAL,  swot["Opportunities"]),
        (Inches(6.8),  Inches(4.2), "Threats",        ACCENT_AMBER, swot["Threats"]),
    ]

    for left, top, title, color, items in grid:
        # Card background
        _rect(slide, left, top, Inches(5.7), Inches(2.7), CARD_BG)
        # Top accent bar
        _rect(slide, left, top, Inches(5.7), Inches(0.04), color)
        # Title
        _text(slide, left + Inches(0.2), top + Inches(0.12),
              Inches(5), Inches(0.35), title,
              size=15, color=color, bold=True, font=FONT_TITLE)
        # Items
        if items:
            txBox = slide.shapes.add_textbox(
                left + Inches(0.2), top + Inches(0.55), Inches(5.3), Inches(2.0))
            tf = txBox.text_frame
            tf.word_wrap = True
            _bullets(tf, items[:5], size=11, color=MED_TEXT, label_color=color)
        else:
            _text(slide, left + Inches(0.2), top + Inches(0.7),
                  Inches(5), Inches(0.3), "No data available",
                  size=10, color=LIGHT_TEXT)


def _slide_opps_threats(prs, section_text):
    slide = prs.slides.add_slide(_get_blank_layout(prs))
    _bg(slide)
    _section_header(slide, "Opportunities & Threats")

    opps, threats = [], []
    cur = None
    for line in section_text.split("\n"):
        s = line.strip().lower()
        if "opportunit" in s and (":" in s or s.startswith("#")):
            cur = "o"
        elif "threat" in s and (":" in s or s.startswith("#")):
            cur = "t"
        elif "gap" in s and (":" in s or s.startswith("#")):
            cur = None

        raw = line.strip()
        is_b = raw.startswith(("-", "*", "•")) or re.match(r'^\d+[\.\)]', raw)
        if cur and is_b:
            c = re.sub(r'^[-*•]\s*', '', raw)
            c = re.sub(r'^\d+[\.\)]\s*', '', c)
            c = _clean(c)
            if c and len(c) > 3:
                (opps if cur == "o" else threats).append(c)

    # Left: Opportunities
    _text(slide, Inches(0.8), Inches(1.3), Inches(5.5), Inches(0.4),
          "Opportunities", size=16, color=ACCENT_TEAL, bold=True, font=FONT_TITLE)
    _line(slide, Inches(0.8), Inches(1.75), Inches(5.5))
    if opps:
        txBox = slide.shapes.add_textbox(
            Inches(0.8), Inches(1.95), Inches(5.5), Inches(5.0))
        tf = txBox.text_frame
        tf.word_wrap = True
        _bullets(tf, opps[:6], size=13, label_color=ACCENT_TEAL)

    # Divider
    _rect(slide, Inches(6.65), Inches(1.3), Inches(0.015), Inches(5.5), RULE_COLOR)

    # Right: Threats
    _text(slide, Inches(7.0), Inches(1.3), Inches(5.5), Inches(0.4),
          "Threats", size=16, color=ACCENT_RED, bold=True, font=FONT_TITLE)
    _line(slide, Inches(7.0), Inches(1.75), Inches(5.5))
    if threats:
        txBox = slide.shapes.add_textbox(
            Inches(7.0), Inches(1.95), Inches(5.5), Inches(5.0))
        tf = txBox.text_frame
        tf.word_wrap = True
        _bullets(tf, threats[:6], size=13, label_color=ACCENT_RED)


def _slide_recommendations(prs, section_text):
    slide = prs.slides.add_slide(_get_blank_layout(prs))
    _bg(slide)
    _section_header(slide, "Strategic Recommendations")

    timeframes = [
        ("Immediate (30 Days)", BCG_GREEN, []),
        ("Short-Term (3 Months)", ACCENT_TEAL, []),
        ("Long-Term (12 Months)", RGBColor(0x4A, 0x5A, 0x6A), []),
    ]

    idx = None
    for line in section_text.split("\n"):
        s = line.strip().lower()
        if "immediate" in s or "30 day" in s or "next 30" in s:
            idx = 0
        elif "short" in s or "3 month" in s or "next 3" in s:
            idx = 1
        elif "long" in s or "12 month" in s or "next 12" in s:
            idx = 2

        raw = line.strip()
        is_b = raw.startswith(("-", "*", "•")) or re.match(r'^\d+[\.\)]', raw)
        if idx is not None and is_b:
            c = re.sub(r'^[-*•]\s*', '', raw)
            c = re.sub(r'^\d+[\.\)]\s*', '', c)
            c = _clean(c)
            c = re.sub(r'\[.\]', '', c).strip()
            if c and len(c) > 3:
                timeframes[idx][2].append(c)

    col_w = Inches(3.8)
    gap = Inches(0.15)
    for i, (label, color, items) in enumerate(timeframes):
        left = Inches(0.8) + (col_w + gap) * i
        # Column header
        _rect(slide, left, Inches(1.3), col_w, Inches(0.04), color)
        _text(slide, left + Inches(0.1), Inches(1.45), col_w - Inches(0.2),
              Inches(0.4), label, size=14, color=color, bold=True, font=FONT_TITLE)
        _line(slide, left, Inches(1.9), col_w)

        if items:
            txBox = slide.shapes.add_textbox(
                left + Inches(0.1), Inches(2.1), col_w - Inches(0.2), Inches(4.8))
            tf = txBox.text_frame
            tf.word_wrap = True
            _bullets(tf, items[:5], size=12, color=MED_TEXT, label_color=color)

    # Vertical dividers
    _rect(slide, Inches(0.8) + col_w + gap / 2, Inches(1.3),
          Inches(0.015), Inches(5.5), RULE_COLOR)
    _rect(slide, Inches(0.8) + (col_w + gap) * 2 - gap / 2, Inches(1.3),
          Inches(0.015), Inches(5.5), RULE_COLOR)


def _slide_closing(prs):
    slide = prs.slides.add_slide(_get_blank_layout(prs))
    _bg(slide, TITLE_BG)

    _rect(slide, Inches(0), Inches(0), SLIDE_WIDTH, Inches(0.04),
          RGBColor(0x00, 0x8A, 0x68))

    _text(slide, Inches(1.5), Inches(2.5), Inches(10), Inches(0.8),
          "Thank You", size=44, color=BG_WHITE, bold=True, font=FONT_TITLE,
          align=PP_ALIGN.LEFT)

    _rect(slide, Inches(1.5), Inches(3.5), Inches(2.5), Inches(0.02), BG_WHITE)

    _text(slide, Inches(1.5), Inches(3.8), Inches(10), Inches(0.5),
          "AI Market Research Agent", size=18,
          color=RGBColor(0x7F, 0xBF, 0xAD), bold=True)

    _text(slide, Inches(1.5), Inches(4.4), Inches(10), Inches(0.4),
          "Multi-Agent Intelligence System", size=14,
          color=RGBColor(0x5A, 0x9A, 0x8A))

    _text(slide, Inches(1.5), Inches(6.5), Inches(10), Inches(0.3),
          "CONFIDENTIAL", size=9, color=RGBColor(0x5A, 0x9A, 0x8A),
          bold=True)

    _rect(slide, Inches(0), Inches(7.46), SLIDE_WIDTH, Inches(0.04),
          RGBColor(0x00, 0x8A, 0x68))


# ── Main Generator ────────────────────────────────────────

def generate_pptx(report_text: str, company: str, competitors: list[str],
                  output_path: str) -> str:
    output_dir = os.path.dirname(output_path)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)

    prs = Presentation()
    prs.slide_width = SLIDE_WIDTH
    prs.slide_height = SLIDE_HEIGHT

    date_str = datetime.now().strftime("%B %d, %Y")
    sections = _parse_sections(report_text)

    _slide_title(prs, company, competitors, date_str)

    for key in sections:
        if "executive" in key.lower() or "summary" in key.lower():
            _slide_exec_summary(prs, sections[key])
            break

    for key in sections:
        if "market" in key.lower() and "landscape" in key.lower():
            _slide_market(prs, sections[key])
            break

    for key in sections:
        if "competitor" in key.lower() and "profile" in key.lower():
            st = sections[key]
            for comp in competitors:
                if comp.lower() in st.lower():
                    _slide_competitor(prs, comp, st)
                    break
            else:
                _slide_competitor(prs, ", ".join(competitors), st)
            break

    for key in sections:
        if "competitive" in key.lower() and "analy" in key.lower():
            _slide_matrix(prs, sections[key])
            if "swot" in sections[key].lower():
                _slide_swot(prs, sections[key])
            break

    for key in sections:
        if "swot" in key.lower():
            _slide_swot(prs, sections[key])
            break

    for key in sections:
        if "opportunit" in key.lower() or "threat" in key.lower():
            _slide_opps_threats(prs, sections[key])
            break

    for key in sections:
        if "recommend" in key.lower() or "strateg" in key.lower():
            _slide_recommendations(prs, sections[key])
            break

    _slide_closing(prs)

    prs.save(output_path)
    return output_path
