"""
PowerPoint Report Generator
============================
Converts the markdown competitive intelligence report into a
professional PowerPoint presentation with a bright, modern theme.
"""

import os
import re
from datetime import datetime

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE

# ── Bright Theme Colors ──────────────────────────────────
BG_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BG_LIGHT = RGBColor(0xF8, 0xF9, 0xFA)       # Very light gray
CARD_BG = RGBColor(0xF0, 0xF4, 0xF8)         # Light blue-gray card
TITLE_BG = RGBColor(0x1A, 0x56, 0xDB)        # Bold blue for title slide
DARK_TEXT = RGBColor(0x1E, 0x29, 0x3B)        # Near-black text
MED_TEXT = RGBColor(0x4B, 0x55, 0x63)         # Medium gray text
LIGHT_TEXT = RGBColor(0x6B, 0x72, 0x80)       # Light gray text

ACCENT_BLUE = RGBColor(0x1A, 0x56, 0xDB)     # Primary blue
ACCENT_GREEN = RGBColor(0x05, 0x96, 0x69)     # Success green
ACCENT_RED = RGBColor(0xDC, 0x26, 0x26)       # Danger red
ACCENT_YELLOW = RGBColor(0xD9, 0x77, 0x06)    # Warning amber
ACCENT_PURPLE = RGBColor(0x7C, 0x3A, 0xED)    # Purple

TABLE_HEADER_BG = RGBColor(0x1A, 0x56, 0xDB)
TABLE_ROW_BG = RGBColor(0xFF, 0xFF, 0xFF)
TABLE_ALT_BG = RGBColor(0xF0, 0xF4, 0xF8)
TABLE_BORDER = RGBColor(0xE2, 0xE8, 0xF0)

FONT_NAME = "Calibri"
SLIDE_WIDTH = Inches(13.333)
SLIDE_HEIGHT = Inches(7.5)


def _get_blank_layout(prs):
    """Get a blank slide layout safely."""
    try:
        return prs.slide_layouts[6]
    except IndexError:
        return prs.slide_layouts[-1]


def _clean_markdown(text: str) -> str:
    """Remove common markdown formatting from text."""
    text = re.sub(r'\*\*(.+?)\*\*', r'\1', text)
    text = re.sub(r'\*(.+?)\*', r'\1', text)
    text = re.sub(r'`(.+?)`', r'\1', text)
    text = re.sub(r'\[(.+?)\]\(.+?\)', r'\1', text)
    return text.strip()


def _set_slide_bg(slide, color=BG_WHITE):
    """Set solid background color for a slide."""
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def _add_shape(slide, left, top, width, height, color):
    """Add a colored rectangle shape."""
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    return shape


def _add_rounded_card(slide, left, top, width, height, color=CARD_BG):
    """Add a rounded rectangle card background."""
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.color.rgb = RGBColor(0xE2, 0xE8, 0xF0)
    shape.line.width = Pt(1)
    return shape


def _add_text(slide, left, top, width, height, text, font_size=18,
              color=DARK_TEXT, bold=False, alignment=PP_ALIGN.LEFT):
    """Add a text box with styled text and auto-shrink."""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    p = tf.paragraphs[0]
    p.text = _clean_markdown(text)
    p.font.size = Pt(font_size)
    p.font.color.rgb = color
    p.font.bold = bold
    p.font.name = FONT_NAME
    p.alignment = alignment
    return txBox


def _add_bullets(text_frame, items, font_size=14, color=DARK_TEXT,
                 bold_label=False, accent_color=ACCENT_BLUE):
    """Add bullet points to a text frame. No truncation — text wraps naturally."""
    for i, item in enumerate(items):
        p = text_frame.paragraphs[0] if i == 0 else text_frame.add_paragraph()

        # Clean markdown
        clean = item.strip().lstrip("*-•")
        clean = re.sub(r'^\d+[\.\)]\s*', '', clean)
        clean = _clean_markdown(clean)

        p.font.size = Pt(font_size)
        p.font.color.rgb = color
        p.font.name = FONT_NAME
        p.space_before = Pt(2)
        p.space_after = Pt(4)
        p.level = 0

        if bold_label and ":" in clean:
            parts = clean.split(":", 1)
            run1 = p.add_run()
            run1.text = parts[0] + ":"
            run1.font.size = Pt(font_size)
            run1.font.color.rgb = accent_color
            run1.font.bold = True
            run1.font.name = FONT_NAME
            if len(parts) > 1:
                run2 = p.add_run()
                run2.text = " " + parts[1].strip()
                run2.font.size = Pt(font_size)
                run2.font.color.rgb = color
                run2.font.name = FONT_NAME
        else:
            p.text = clean


def _section_title(slide, left, top, width, text, font_size=28):
    """Add a section title with colored accent bar."""
    _add_shape(slide, left, top + Inches(0.05), Inches(0.06), Inches(0.45), ACCENT_BLUE)
    _add_text(slide, left + Inches(0.2), top, width, Inches(0.6),
              text, font_size=font_size, color=DARK_TEXT, bold=True)


def _parse_sections(report_text: str) -> dict:
    """Parse markdown report into sections by ## headers."""
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


def _extract_bullets(text: str) -> list[str]:
    """Extract bullet points from markdown text."""
    bullets = []
    for line in text.split("\n"):
        stripped = line.strip()
        if stripped and (stripped.startswith(("-", "*", "•")) or
                         re.match(r'^\d+[\.\)]', stripped)):
            clean = re.sub(r'^[-*•]\s*', '', stripped)
            clean = re.sub(r'^\d+[\.\)]\s*', '', clean)
            clean = _clean_markdown(clean)
            if clean and len(clean) > 3:
                bullets.append(clean)
    return bullets


def _extract_table_data(text: str) -> tuple[list[str], list[list[str]]]:
    """Extract table headers and rows from markdown table."""
    headers = []
    rows = []
    for line in text.split("\n"):
        stripped = line.strip()
        if "|" in stripped and not re.match(r'^\|[\s\-:|]+\|$', stripped):
            cells = [c.strip() for c in stripped.split("|") if c.strip()]
            if not headers:
                headers = cells
            else:
                rows.append(cells)
    return headers, rows


# ── Slide Builders ────────────────────────────────────────

def _build_title_slide(prs, company, competitors, date_str):
    """Slide 1: Bold blue title slide."""
    slide = prs.slides.add_slide(_get_blank_layout(prs))
    _set_slide_bg(slide, TITLE_BG)

    # Top accent stripe
    _add_shape(slide, Inches(0), Inches(0), SLIDE_WIDTH, Inches(0.08), BG_WHITE)

    _add_text(slide, Inches(1.5), Inches(1.5), Inches(10), Inches(0.8),
              "COMPETITIVE INTELLIGENCE REPORT", font_size=18,
              color=RGBColor(0xBF, 0xDB, 0xFE), bold=True, alignment=PP_ALIGN.CENTER)

    _add_text(slide, Inches(1.5), Inches(2.3), Inches(10), Inches(1.5),
              company, font_size=48, color=BG_WHITE, bold=True,
              alignment=PP_ALIGN.CENTER)

    # Divider line
    _add_shape(slide, Inches(5.5), Inches(4.0), Inches(2.3), Inches(0.04), BG_WHITE)

    comp_text = f"Competitors Analyzed: {', '.join(competitors)}"
    _add_text(slide, Inches(1.5), Inches(4.3), Inches(10), Inches(0.6),
              comp_text, font_size=16, color=RGBColor(0xBF, 0xDB, 0xFE),
              alignment=PP_ALIGN.CENTER)

    _add_text(slide, Inches(1.5), Inches(5.0), Inches(10), Inches(0.5),
              date_str, font_size=14, color=RGBColor(0x93, 0xB8, 0xEF),
              alignment=PP_ALIGN.CENTER)

    # Bottom accent
    _add_shape(slide, Inches(0), Inches(7.42), SLIDE_WIDTH, Inches(0.08), BG_WHITE)


def _build_executive_summary(prs, section_text):
    """Slide 2: Executive Summary."""
    slide = prs.slides.add_slide(_get_blank_layout(prs))
    _set_slide_bg(slide, BG_WHITE)

    _section_title(slide, Inches(0.8), Inches(0.4), Inches(11), "Executive Summary")

    bullets = _extract_bullets(section_text)
    if not bullets:
        bullets = [_clean_markdown(l.strip()) for l in section_text.split("\n")
                   if l.strip() and not l.strip().startswith("#")]

    if len(bullets) > 5:
        mid = len(bullets) // 2
        # Left card
        _add_rounded_card(slide, Inches(0.8), Inches(1.2), Inches(5.7), Inches(5.8))
        txBox = slide.shapes.add_textbox(Inches(1.1), Inches(1.5), Inches(5.2), Inches(5.3))
        tf = txBox.text_frame
        tf.word_wrap = True
        _add_bullets(tf, bullets[:mid], font_size=13, bold_label=True)

        # Right card
        _add_rounded_card(slide, Inches(6.8), Inches(1.2), Inches(5.7), Inches(5.8))
        txBox = slide.shapes.add_textbox(Inches(7.1), Inches(1.5), Inches(5.2), Inches(5.3))
        tf = txBox.text_frame
        tf.word_wrap = True
        _add_bullets(tf, bullets[mid:mid+6], font_size=13, bold_label=True)
    else:
        _add_rounded_card(slide, Inches(0.8), Inches(1.2), Inches(11.7), Inches(5.8))
        txBox = slide.shapes.add_textbox(Inches(1.1), Inches(1.5), Inches(11.2), Inches(5.3))
        tf = txBox.text_frame
        tf.word_wrap = True
        _add_bullets(tf, bullets[:8], font_size=14, bold_label=True)


def _build_market_landscape(prs, section_text):
    """Slide 3: Market Landscape Overview."""
    slide = prs.slides.add_slide(_get_blank_layout(prs))
    _set_slide_bg(slide, BG_WHITE)

    _section_title(slide, Inches(0.8), Inches(0.4), Inches(11), "Market Landscape")

    paragraphs = [p.strip() for p in section_text.split("\n")
                  if p.strip() and not p.strip().startswith("#")]
    bullets = _extract_bullets(section_text)

    # Body text (left)
    text_content = []
    for p in paragraphs:
        if not p.startswith(("-", "*", "•")) and not re.match(r'^\d+[\.\)]', p):
            clean = _clean_markdown(p)
            if clean and len(clean) > 10:
                text_content.append(clean)

    if text_content:
        _add_rounded_card(slide, Inches(0.8), Inches(1.2), Inches(5.7), Inches(5.8))
        _add_text(slide, Inches(1.1), Inches(1.4), Inches(5.2), Inches(5.4),
                  "\n\n".join(text_content[:4]), font_size=12, color=MED_TEXT)

    # Key trends (right)
    if bullets:
        _add_rounded_card(slide, Inches(6.8), Inches(1.2), Inches(5.7), Inches(5.8))
        _add_text(slide, Inches(7.1), Inches(1.4), Inches(5), Inches(0.5),
                  "Key Trends", font_size=18, color=ACCENT_BLUE, bold=True)
        txBox = slide.shapes.add_textbox(Inches(7.1), Inches(2.0), Inches(5.2), Inches(4.8))
        tf = txBox.text_frame
        tf.word_wrap = True
        _add_bullets(tf, bullets[:6], font_size=12, bold_label=True)


def _build_competitor_profile(prs, name, section_text):
    """Slide 4+: Individual competitor profile."""
    slide = prs.slides.add_slide(_get_blank_layout(prs))
    _set_slide_bg(slide, BG_WHITE)

    _section_title(slide, Inches(0.8), Inches(0.4), Inches(10),
                   f"Competitor Profile: {name}")

    # Threat badge
    threat_match = re.search(r'[Tt]hreat\s*[Ll]evel.*?(\d+)/10', section_text)
    threat_level = int(threat_match.group(1)) if threat_match else 5
    threat_color = ACCENT_GREEN if threat_level <= 3 else (ACCENT_YELLOW if threat_level <= 6 else ACCENT_RED)
    _add_rounded_card(slide, Inches(10.5), Inches(0.35), Inches(2.3), Inches(0.7), threat_color)
    _add_text(slide, Inches(10.5), Inches(0.4), Inches(2.3), Inches(0.6),
              f"Threat: {threat_level}/10", font_size=16, color=BG_WHITE, bold=True,
              alignment=PP_ALIGN.CENTER)

    bullets = _extract_bullets(section_text)
    if not bullets:
        bullets = [_clean_markdown(l.strip()) for l in section_text.split("\n")
                   if l.strip() and not l.strip().startswith("#")]

    mid = max(1, len(bullets) // 2)

    # Left: Overview
    _add_rounded_card(slide, Inches(0.8), Inches(1.3), Inches(5.6), Inches(5.7))
    _add_text(slide, Inches(1.0), Inches(1.45), Inches(5), Inches(0.5),
              "Overview", font_size=18, color=ACCENT_BLUE, bold=True)
    if bullets[:mid]:
        txBox = slide.shapes.add_textbox(Inches(1.0), Inches(2.1), Inches(5.2), Inches(4.7))
        tf = txBox.text_frame
        tf.word_wrap = True
        _add_bullets(tf, bullets[:mid][:6], font_size=12, bold_label=True)

    # Right: Strengths & Weaknesses
    _add_rounded_card(slide, Inches(6.8), Inches(1.3), Inches(5.6), Inches(5.7))
    _add_text(slide, Inches(7.0), Inches(1.45), Inches(5), Inches(0.5),
              "Strengths & Weaknesses", font_size=18, color=ACCENT_BLUE, bold=True)
    if bullets[mid:]:
        txBox = slide.shapes.add_textbox(Inches(7.0), Inches(2.1), Inches(5.2), Inches(4.7))
        tf = txBox.text_frame
        tf.word_wrap = True
        _add_bullets(tf, bullets[mid:][:6], font_size=12, bold_label=True)


def _build_competitive_matrix(prs, section_text):
    """Slide: Competitive comparison matrix as a table."""
    slide = prs.slides.add_slide(_get_blank_layout(prs))
    _set_slide_bg(slide, BG_WHITE)

    _section_title(slide, Inches(0.8), Inches(0.4), Inches(11), "Competitive Analysis")

    headers, rows = _extract_table_data(section_text)

    if not headers or not rows:
        bullets = _extract_bullets(section_text)
        if bullets:
            _add_rounded_card(slide, Inches(0.8), Inches(1.2), Inches(11.7), Inches(5.8))
            txBox = slide.shapes.add_textbox(Inches(1.1), Inches(1.5), Inches(11.2), Inches(5.3))
            tf = txBox.text_frame
            tf.word_wrap = True
            _add_bullets(tf, bullets[:10], font_size=13, bold_label=True)
        return

    num_cols = len(headers)
    num_rows = min(len(rows) + 1, 8)
    rows = rows[:num_rows - 1]

    table_width = Inches(min(11.5, num_cols * 2.2))
    table_height = Inches(0.5 * num_rows)

    table_shape = slide.shapes.add_table(
        num_rows, num_cols,
        Inches(0.8), Inches(1.4),
        table_width, table_height
    )
    table = table_shape.table

    col_w = int(table_width / num_cols)
    for col in table.columns:
        col.width = col_w

    # Header row
    for i, header in enumerate(headers):
        cell = table.cell(0, i)
        cell.text = _clean_markdown(header)
        cell.fill.solid()
        cell.fill.fore_color.rgb = TABLE_HEADER_BG
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        for p in cell.text_frame.paragraphs:
            p.font.size = Pt(12)
            p.font.color.rgb = BG_WHITE
            p.font.bold = True
            p.font.name = FONT_NAME

    # Data rows
    for r, row_data in enumerate(rows):
        bg = TABLE_ROW_BG if r % 2 == 0 else TABLE_ALT_BG
        for c in range(num_cols):
            cell = table.cell(r + 1, c)
            cell.text = _clean_markdown(row_data[c]) if c < len(row_data) else ""
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            for p in cell.text_frame.paragraphs:
                p.font.size = Pt(11)
                p.font.color.rgb = DARK_TEXT
                p.font.name = FONT_NAME


def _build_swot(prs, section_text):
    """Slide: SWOT Analysis in 2x2 grid."""
    slide = prs.slides.add_slide(_get_blank_layout(prs))
    _set_slide_bg(slide, BG_WHITE)

    _section_title(slide, Inches(0.8), Inches(0.4), Inches(11), "SWOT Analysis")

    swot = {"Strengths": [], "Weaknesses": [], "Opportunities": [], "Threats": []}
    current = None

    for line in section_text.split("\n"):
        stripped = line.strip()
        lower = stripped.lower()

        matched_key = None
        for key in swot:
            if key.lower() in lower:
                matched_key = key
                break

        if matched_key:
            colon_match = re.search(
                r'(?:strengths|weaknesses|opportunities|threats)\s*:?\s*:(.+)',
                stripped, re.IGNORECASE)
            if colon_match:
                items_text = colon_match.group(1).strip()
                items = [i.strip().rstrip(".") for i in re.split(r'[,;]', items_text) if i.strip()]
                swot[matched_key].extend(items)
                current = matched_key
            elif ":" in stripped or stripped.startswith("#") or stripped.startswith("*"):
                current = matched_key
        elif current and stripped.startswith(("-", "*", "•")):
            clean = re.sub(r'^[-*•]\s*', '', stripped)
            clean = _clean_markdown(clean).rstrip(".")

            sub_match = None
            for key in swot:
                if clean.lower().startswith(key.lower()):
                    sub_match = key
                    break
            if sub_match:
                colon_pos = clean.find(":")
                if colon_pos > 0:
                    items_text = clean[colon_pos + 1:].strip()
                    items = [i.strip().rstrip(".") for i in re.split(r'[,;]', items_text) if i.strip()]
                    swot[sub_match].extend(items)
            elif clean and len(clean) > 3:
                swot[current].append(clean)

    # 2x2 grid with colored top bars
    grid = [
        (Inches(0.8), Inches(1.2), "Strengths", ACCENT_GREEN, swot["Strengths"]),
        (Inches(6.8), Inches(1.2), "Weaknesses", ACCENT_RED, swot["Weaknesses"]),
        (Inches(0.8), Inches(4.1), "Opportunities", ACCENT_BLUE, swot["Opportunities"]),
        (Inches(6.8), Inches(4.1), "Threats", ACCENT_YELLOW, swot["Threats"]),
    ]

    for left, top, title, color, items in grid:
        _add_rounded_card(slide, left, top, Inches(5.7), Inches(2.7))
        # Colored top bar
        _add_shape(slide, left, top, Inches(5.7), Inches(0.06), color)
        # Title
        _add_text(slide, left + Inches(0.2), top + Inches(0.15),
                  Inches(5), Inches(0.4), title,
                  font_size=16, color=color, bold=True)
        # Items
        if items:
            txBox = slide.shapes.add_textbox(
                left + Inches(0.2), top + Inches(0.6), Inches(5.3), Inches(1.9))
            tf = txBox.text_frame
            tf.word_wrap = True
            _add_bullets(tf, items[:5], font_size=11, color=MED_TEXT, accent_color=color)
        else:
            _add_text(slide, left + Inches(0.2), top + Inches(0.8),
                      Inches(5), Inches(0.4), "No data available",
                      font_size=11, color=LIGHT_TEXT)


def _build_opportunities_threats(prs, section_text):
    """Slide: Opportunities & Threats ranked lists."""
    slide = prs.slides.add_slide(_get_blank_layout(prs))
    _set_slide_bg(slide, BG_WHITE)

    _section_title(slide, Inches(0.8), Inches(0.4), Inches(11), "Opportunities & Threats")

    opportunities = []
    threats = []
    current_list = None
    for line in section_text.split("\n"):
        stripped = line.strip().lower()
        if "opportunit" in stripped and (":" in stripped or stripped.startswith("#")):
            current_list = "opp"
        elif "threat" in stripped and (":" in stripped or stripped.startswith("#")):
            current_list = "threat"
        elif "gap" in stripped and (":" in stripped or stripped.startswith("#")):
            current_list = None

        raw = line.strip()
        is_bullet = raw.startswith(("-", "*", "•")) or re.match(r'^\d+[\.\)]', raw)
        if current_list and is_bullet:
            clean = re.sub(r'^[-*•]\s*', '', raw)
            clean = re.sub(r'^\d+[\.\)]\s*', '', clean)
            clean = _clean_markdown(clean)
            if clean and len(clean) > 3:
                if current_list == "opp":
                    opportunities.append(clean)
                elif current_list == "threat":
                    threats.append(clean)

    # Left: Opportunities
    _add_rounded_card(slide, Inches(0.8), Inches(1.2), Inches(5.7), Inches(5.8))
    _add_shape(slide, Inches(0.8), Inches(1.2), Inches(5.7), Inches(0.06), ACCENT_GREEN)
    _add_text(slide, Inches(1.1), Inches(1.4), Inches(5), Inches(0.5),
              "Opportunities", font_size=20, color=ACCENT_GREEN, bold=True)
    if opportunities:
        txBox = slide.shapes.add_textbox(Inches(1.1), Inches(2.1), Inches(5.2), Inches(4.7))
        tf = txBox.text_frame
        tf.word_wrap = True
        _add_bullets(tf, opportunities[:5], font_size=13, bold_label=True,
                     accent_color=ACCENT_GREEN)

    # Right: Threats
    _add_rounded_card(slide, Inches(6.8), Inches(1.2), Inches(5.7), Inches(5.8))
    _add_shape(slide, Inches(6.8), Inches(1.2), Inches(5.7), Inches(0.06), ACCENT_RED)
    _add_text(slide, Inches(7.1), Inches(1.4), Inches(5), Inches(0.5),
              "Threats", font_size=20, color=ACCENT_RED, bold=True)
    if threats:
        txBox = slide.shapes.add_textbox(Inches(7.1), Inches(2.1), Inches(5.2), Inches(4.7))
        tf = txBox.text_frame
        tf.word_wrap = True
        _add_bullets(tf, threats[:5], font_size=13, bold_label=True,
                     accent_color=ACCENT_RED)


def _build_recommendations(prs, section_text):
    """Slide: Strategic Recommendations with timeline columns."""
    slide = prs.slides.add_slide(_get_blank_layout(prs))
    _set_slide_bg(slide, BG_WHITE)

    _section_title(slide, Inches(0.8), Inches(0.4), Inches(11), "Strategic Recommendations")

    timeframes = [
        ("Next 30 Days", ACCENT_GREEN, []),
        ("Next 3 Months", ACCENT_BLUE, []),
        ("Next 12 Months", ACCENT_PURPLE, []),
    ]

    current_idx = None
    for line in section_text.split("\n"):
        stripped = line.strip().lower()
        if "immediate" in stripped or "30 day" in stripped or "next 30" in stripped:
            current_idx = 0
        elif "short" in stripped or "3 month" in stripped or "next 3" in stripped:
            current_idx = 1
        elif "long" in stripped or "12 month" in stripped or "next 12" in stripped:
            current_idx = 2

        raw = line.strip()
        is_bullet = raw.startswith(("-", "*", "•")) or re.match(r'^\d+[\.\)]', raw)
        if current_idx is not None and is_bullet:
            clean = re.sub(r'^[-*•]\s*', '', raw)
            clean = re.sub(r'^\d+[\.\)]\s*', '', clean)
            clean = _clean_markdown(clean)
            clean = re.sub(r'\[.\]', '', clean).strip()
            if clean and len(clean) > 3:
                timeframes[current_idx][2].append(clean)

    col_width = Inches(3.8)
    gap = Inches(0.15)
    for i, (label, color, items) in enumerate(timeframes):
        left = Inches(0.8) + (col_width + gap) * i
        _add_rounded_card(slide, left, Inches(1.2), col_width, Inches(5.8))
        _add_shape(slide, left, Inches(1.2), col_width, Inches(0.06), color)
        _add_text(slide, left + Inches(0.2), Inches(1.4), col_width - Inches(0.4),
                  Inches(0.5), label, font_size=18, color=color, bold=True)

        if items:
            txBox = slide.shapes.add_textbox(
                left + Inches(0.2), Inches(2.1), col_width - Inches(0.4), Inches(4.7))
            tf = txBox.text_frame
            tf.word_wrap = True
            _add_bullets(tf, items[:5], font_size=12, color=MED_TEXT,
                         bold_label=True, accent_color=color)


def _build_closing_slide(prs):
    """Final slide with branding."""
    slide = prs.slides.add_slide(_get_blank_layout(prs))
    _set_slide_bg(slide, TITLE_BG)

    _add_shape(slide, Inches(0), Inches(0), SLIDE_WIDTH, Inches(0.08), BG_WHITE)

    _add_text(slide, Inches(1), Inches(2.3), Inches(11), Inches(1),
              "Thank You", font_size=48, color=BG_WHITE, bold=True,
              alignment=PP_ALIGN.CENTER)

    _add_shape(slide, Inches(5.5), Inches(3.5), Inches(2.3), Inches(0.04), BG_WHITE)

    _add_text(slide, Inches(1), Inches(3.8), Inches(11), Inches(0.6),
              "AI Market Research Agent", font_size=20,
              color=RGBColor(0xBF, 0xDB, 0xFE), bold=True, alignment=PP_ALIGN.CENTER)

    _add_text(slide, Inches(1), Inches(4.5), Inches(11), Inches(0.6),
              "Multi-Agent Intelligence System  |  Powered by CrewAI",
              font_size=16, color=RGBColor(0x93, 0xB8, 0xEF), alignment=PP_ALIGN.CENTER)

    _add_shape(slide, Inches(0), Inches(7.42), SLIDE_WIDTH, Inches(0.08), BG_WHITE)


# ── Main Generator ────────────────────────────────────────

def generate_pptx(report_text: str, company: str, competitors: list[str],
                  output_path: str) -> str:
    """
    Convert a markdown report into a professional PowerPoint presentation.

    Args:
        report_text: The raw markdown report content
        company: Company name
        competitors: List of competitor names
        output_path: Path to save the .pptx file

    Returns:
        The output file path
    """
    output_dir = os.path.dirname(output_path)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)

    prs = Presentation()
    prs.slide_width = SLIDE_WIDTH
    prs.slide_height = SLIDE_HEIGHT

    date_str = datetime.now().strftime("%B %d, %Y")
    sections = _parse_sections(report_text)

    # Slide 1: Title
    _build_title_slide(prs, company, competitors, date_str)

    # Slide 2: Executive Summary
    for key in sections:
        if "executive" in key.lower() or "summary" in key.lower():
            _build_executive_summary(prs, sections[key])
            break

    # Slide 3: Market Landscape
    for key in sections:
        if "market" in key.lower() and "landscape" in key.lower():
            _build_market_landscape(prs, sections[key])
            break

    # Slide 4+: Competitor Profiles
    for key in sections:
        if "competitor" in key.lower() and "profile" in key.lower():
            section_text = sections[key]
            for comp in competitors:
                if comp.lower() in section_text.lower():
                    _build_competitor_profile(prs, comp, section_text)
                    break
            else:
                _build_competitor_profile(prs, ", ".join(competitors), section_text)
            break

    # Slide: Competitive Analysis / Matrix
    for key in sections:
        if "competitive" in key.lower() and "analy" in key.lower():
            _build_competitive_matrix(prs, sections[key])
            if "swot" in sections[key].lower():
                _build_swot(prs, sections[key])
            break

    # Slide: SWOT (standalone)
    for key in sections:
        if "swot" in key.lower():
            _build_swot(prs, sections[key])
            break

    # Slide: Opportunities & Threats
    for key in sections:
        if "opportunit" in key.lower() or "threat" in key.lower():
            _build_opportunities_threats(prs, sections[key])
            break

    # Slide: Recommendations
    for key in sections:
        if "recommend" in key.lower() or "strateg" in key.lower():
            _build_recommendations(prs, sections[key])
            break

    # Closing slide
    _build_closing_slide(prs)

    prs.save(output_path)
    return output_path
