"""
PowerPoint Report Generator
============================
Converts the markdown competitive intelligence report into a
professional PowerPoint presentation with dark theme styling.
"""

import os
import re
from datetime import datetime

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# ── Theme Colors ──────────────────────────────────────────
BG_DARK = RGBColor(0x1B, 0x2A, 0x4A)       # Dark navy
BG_DARKER = RGBColor(0x0F, 0x1A, 0x30)     # Darker navy
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_GRAY = RGBColor(0xBD, 0xC3, 0xC7)
ACCENT_BLUE = RGBColor(0x34, 0x98, 0xDB)
ACCENT_GREEN = RGBColor(0x2E, 0xCC, 0x71)
ACCENT_RED = RGBColor(0xE7, 0x4C, 0x3C)
ACCENT_YELLOW = RGBColor(0xF3, 0x9C, 0x12)
ACCENT_PURPLE = RGBColor(0x9B, 0x59, 0xB6)
TABLE_HEADER_BG = RGBColor(0x2C, 0x3E, 0x50)
TABLE_ROW_BG = RGBColor(0x1E, 0x30, 0x50)
TABLE_ALT_BG = RGBColor(0x15, 0x24, 0x3E)
CARD_BG = RGBColor(0x15, 0x24, 0x3E)

FONT_NAME = "Calibri"
SLIDE_WIDTH = Inches(13.333)
SLIDE_HEIGHT = Inches(7.5)

# Max characters per bullet to prevent overflow
MAX_BULLET_CHARS = 120


def _get_blank_layout(prs):
    """Get a blank slide layout safely."""
    # Try layout 6 (blank) first, then fall back to last layout
    try:
        layout = prs.slide_layouts[6]
    except IndexError:
        layout = prs.slide_layouts[-1]
    return layout


def _truncate(text: str, max_chars: int = MAX_BULLET_CHARS) -> str:
    """Truncate text to max_chars, adding ellipsis if needed."""
    if len(text) <= max_chars:
        return text
    return text[:max_chars - 3].rstrip() + "..."


def _clean_markdown(text: str) -> str:
    """Remove common markdown formatting from text."""
    text = re.sub(r'\*\*(.+?)\*\*', r'\1', text)  # **bold**
    text = re.sub(r'\*(.+?)\*', r'\1', text)       # *italic*
    text = re.sub(r'`(.+?)`', r'\1', text)         # `code`
    text = re.sub(r'\[(.+?)\]\(.+?\)', r'\1', text)  # [link](url)
    return text.strip()


def _set_slide_bg(slide, color=BG_DARK):
    """Set solid background color for a slide."""
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color


def _add_shape_bg(slide, left, top, width, height, color):
    """Add a colored rectangle shape as a background element."""
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    return shape


def _add_text_box(slide, left, top, width, height, text, font_size=18,
                  color=WHITE, bold=False, alignment=PP_ALIGN.LEFT, font_name=FONT_NAME):
    """Add a text box with styled text."""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = _clean_markdown(text)
    p.font.size = Pt(font_size)
    p.font.color.rgb = color
    p.font.bold = bold
    p.font.name = font_name
    p.alignment = alignment
    return txBox


def _add_bullet_list(text_frame, items, font_size=16, color=WHITE,
                     bold_first=False, max_chars=MAX_BULLET_CHARS):
    """Add bullet points to a text frame."""
    for i, item in enumerate(items):
        if i == 0:
            p = text_frame.paragraphs[0]
        else:
            p = text_frame.add_paragraph()
        # Clean markdown formatting and truncate
        clean = item.strip().lstrip("*-•")
        clean = re.sub(r'^\d+[\.\)]\s*', '', clean)
        clean = _clean_markdown(clean)
        clean = _truncate(clean, max_chars)

        p.font.size = Pt(font_size)
        p.font.color.rgb = color
        p.font.name = FONT_NAME
        p.space_after = Pt(4)
        p.level = 0

        if bold_first and ":" in clean:
            # Make text before colon bold via runs
            parts = clean.split(":", 1)
            run1 = p.add_run()
            run1.text = parts[0] + ":"
            run1.font.size = Pt(font_size)
            run1.font.color.rgb = ACCENT_BLUE
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


def _parse_sections(report_text: str) -> dict:
    """Parse markdown report into sections by ## headers."""
    sections = {}
    current_key = "intro"
    current_content = []

    for line in report_text.split("\n"):
        if line.startswith("## "):
            if current_content:
                sections[current_key] = "\n".join(current_content)
            # Clean the section key: strip # and surrounding whitespace/dots
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
            if clean and len(clean) > 3:  # Skip tiny fragments
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


def _slide_title(slide, left, top, width, text, font_size=32, color=WHITE):
    """Add a section title with accent bar."""
    # Accent bar
    _add_shape_bg(slide, left, top, Inches(0.08), Inches(0.6), ACCENT_BLUE)
    # Title text
    _add_text_box(slide, left + Inches(0.25), top, width, Inches(0.7),
                  text, font_size=font_size, color=color, bold=True)


# ── Slide Builders ────────────────────────────────────────

def _build_title_slide(prs, company, competitors, date_str):
    """Slide 1: Title slide with company name and report info."""
    slide = prs.slides.add_slide(_get_blank_layout(prs))  # blank
    _set_slide_bg(slide, BG_DARKER)

    # Top accent line
    _add_shape_bg(slide, Inches(0), Inches(0), SLIDE_WIDTH, Inches(0.06), ACCENT_BLUE)

    # Company name
    _add_text_box(slide, Inches(1), Inches(1.8), Inches(11), Inches(1.2),
                  "COMPETITIVE INTELLIGENCE REPORT", font_size=20,
                  color=ACCENT_BLUE, bold=True, alignment=PP_ALIGN.CENTER)

    _add_text_box(slide, Inches(1), Inches(2.6), Inches(11), Inches(1.5),
                  company, font_size=44, color=WHITE, bold=True,
                  alignment=PP_ALIGN.CENTER)

    # Divider
    _add_shape_bg(slide, Inches(5.5), Inches(4.2), Inches(2.3), Inches(0.04), ACCENT_BLUE)

    # Competitors
    comp_text = f"Competitors: {', '.join(competitors)}"
    _add_text_box(slide, Inches(1), Inches(4.5), Inches(11), Inches(0.6),
                  comp_text, font_size=18, color=LIGHT_GRAY,
                  alignment=PP_ALIGN.CENTER)

    # Date
    _add_text_box(slide, Inches(1), Inches(5.2), Inches(11), Inches(0.5),
                  date_str, font_size=16, color=LIGHT_GRAY,
                  alignment=PP_ALIGN.CENTER)

    # Bottom accent
    _add_shape_bg(slide, Inches(0), Inches(7.44), SLIDE_WIDTH, Inches(0.06), ACCENT_BLUE)


def _build_executive_summary(prs, section_text):
    """Slide 2: Executive Summary with key findings."""
    slide = prs.slides.add_slide(_get_blank_layout(prs))
    _set_slide_bg(slide)

    _slide_title(slide, Inches(0.8), Inches(0.5), Inches(11), "Executive Summary")

    bullets = _extract_bullets(section_text)
    if not bullets:
        bullets = [_clean_markdown(line.strip()) for line in section_text.split("\n")
                   if line.strip() and not line.strip().startswith("#")]

    # Split into two columns if many bullets
    if len(bullets) > 5:
        left_items = bullets[:len(bullets)//2]
        right_items = bullets[len(bullets)//2:]

        # Left column
        _add_shape_bg(slide, Inches(0.8), Inches(1.3), Inches(5.7), Inches(5.7), CARD_BG)
        txBox = slide.shapes.add_textbox(Inches(1.1), Inches(1.5), Inches(5.2), Inches(5.3))
        tf = txBox.text_frame
        tf.word_wrap = True
        _add_bullet_list(tf, left_items[:6], font_size=14, bold_first=True, max_chars=90)

        # Right column
        _add_shape_bg(slide, Inches(6.8), Inches(1.3), Inches(5.7), Inches(5.7), CARD_BG)
        txBox = slide.shapes.add_textbox(Inches(7.1), Inches(1.5), Inches(5.2), Inches(5.3))
        tf = txBox.text_frame
        tf.word_wrap = True
        _add_bullet_list(tf, right_items[:6], font_size=14, bold_first=True, max_chars=90)
    else:
        _add_shape_bg(slide, Inches(0.8), Inches(1.3), Inches(11.7), Inches(5.7), CARD_BG)
        txBox = slide.shapes.add_textbox(Inches(1.1), Inches(1.5), Inches(11.2), Inches(5.3))
        tf = txBox.text_frame
        tf.word_wrap = True
        _add_bullet_list(tf, bullets[:8], font_size=15, bold_first=True)


def _build_market_landscape(prs, section_text):
    """Slide 3: Market Landscape Overview."""
    slide = prs.slides.add_slide(_get_blank_layout(prs))
    _set_slide_bg(slide)

    _slide_title(slide, Inches(0.8), Inches(0.5), Inches(11), "Market Landscape")

    # Split into subsections
    paragraphs = [p.strip() for p in section_text.split("\n") if p.strip()
                  and not p.strip().startswith("#")]
    bullets = _extract_bullets(section_text)

    # Main text (left side)
    text_content = []
    for p in paragraphs:
        if not p.startswith(("-", "*", "•")) and not re.match(r'^\d+[\.\)]', p):
            clean = _clean_markdown(p)
            if clean and len(clean) > 10:
                text_content.append(_truncate(clean, 200))

    if text_content:
        _add_shape_bg(slide, Inches(0.8), Inches(1.3), Inches(5.7), Inches(5.7), CARD_BG)
        _add_text_box(slide, Inches(1.1), Inches(1.5), Inches(5.2), Inches(5.3),
                      "\n\n".join(text_content[:3]), font_size=13, color=LIGHT_GRAY)

    # Key trends on right
    if bullets:
        _add_shape_bg(slide, Inches(6.8), Inches(1.3), Inches(5.7), Inches(5.7), CARD_BG)
        _add_text_box(slide, Inches(7.1), Inches(1.5), Inches(5), Inches(0.5),
                      "Key Trends", font_size=20, color=ACCENT_BLUE, bold=True)

        txBox = slide.shapes.add_textbox(Inches(7.1), Inches(2.2), Inches(5.2), Inches(4.6))
        tf = txBox.text_frame
        tf.word_wrap = True
        _add_bullet_list(tf, bullets[:6], font_size=13, color=WHITE, bold_first=True, max_chars=80)


def _build_competitor_profile(prs, name, section_text):
    """Slide 4+: Individual competitor profile."""
    slide = prs.slides.add_slide(_get_blank_layout(prs))
    _set_slide_bg(slide)

    _slide_title(slide, Inches(0.8), Inches(0.5), Inches(11),
                 f"Competitor Profile: {name}")

    # Extract threat level
    threat_match = re.search(r'[Tt]hreat\s*[Ll]evel.*?(\d+)/10', section_text)
    threat_level = int(threat_match.group(1)) if threat_match else 5

    # Threat indicator (top right)
    threat_color = ACCENT_GREEN if threat_level <= 3 else (ACCENT_YELLOW if threat_level <= 6 else ACCENT_RED)
    _add_shape_bg(slide, Inches(10.5), Inches(0.5), Inches(2.3), Inches(0.7), threat_color)
    _add_text_box(slide, Inches(10.5), Inches(0.5), Inches(2.3), Inches(0.7),
                  f"Threat: {threat_level}/10", font_size=18, color=WHITE, bold=True,
                  alignment=PP_ALIGN.CENTER)

    bullets = _extract_bullets(section_text)
    if not bullets:
        bullets = [_clean_markdown(line.strip()) for line in section_text.split("\n")
                   if line.strip() and not line.strip().startswith("#")]

    # Left column: Overview
    mid = max(1, len(bullets) // 2)
    overview_items = bullets[:mid]
    analysis_items = bullets[mid:]

    # Left card
    _add_shape_bg(slide, Inches(0.8), Inches(1.4), Inches(5.6), Inches(5.5), CARD_BG)
    _add_text_box(slide, Inches(1.0), Inches(1.5), Inches(5), Inches(0.5),
                  "Overview", font_size=20, color=ACCENT_BLUE, bold=True)
    if overview_items:
        txBox = slide.shapes.add_textbox(Inches(1.0), Inches(2.2), Inches(5.2), Inches(4.5))
        tf = txBox.text_frame
        tf.word_wrap = True
        _add_bullet_list(tf, overview_items[:6], font_size=13, bold_first=True, max_chars=85)

    # Right card
    _add_shape_bg(slide, Inches(6.8), Inches(1.4), Inches(5.6), Inches(5.5), CARD_BG)
    _add_text_box(slide, Inches(7.0), Inches(1.5), Inches(5), Inches(0.5),
                  "Strengths & Weaknesses", font_size=20, color=ACCENT_BLUE, bold=True)
    if analysis_items:
        txBox = slide.shapes.add_textbox(Inches(7.0), Inches(2.2), Inches(5.2), Inches(4.5))
        tf = txBox.text_frame
        tf.word_wrap = True
        _add_bullet_list(tf, analysis_items[:6], font_size=13, bold_first=True, max_chars=85)


def _build_competitive_matrix(prs, section_text):
    """Slide: Competitive comparison matrix as a table."""
    slide = prs.slides.add_slide(_get_blank_layout(prs))
    _set_slide_bg(slide)

    _slide_title(slide, Inches(0.8), Inches(0.5), Inches(11), "Competitive Analysis")

    headers, rows = _extract_table_data(section_text)

    if not headers or not rows:
        # Fallback to bullets
        bullets = _extract_bullets(section_text)
        if bullets:
            _add_shape_bg(slide, Inches(0.8), Inches(1.3), Inches(11.7), Inches(5.7), CARD_BG)
            txBox = slide.shapes.add_textbox(Inches(1.1), Inches(1.5), Inches(11.2), Inches(5.3))
            tf = txBox.text_frame
            tf.word_wrap = True
            _add_bullet_list(tf, bullets[:10], font_size=14, bold_first=True)
        return

    num_cols = len(headers)
    num_rows = min(len(rows) + 1, 8)  # +1 for header, cap at 8 rows
    rows = rows[:num_rows - 1]  # Trim rows to match

    # Calculate table dimensions (use explicit Inches to avoid EMU math issues)
    table_width = Inches(min(11.5, num_cols * 2.2))
    row_height_val = 0.45
    table_height = Inches(row_height_val * num_rows)

    table_shape = slide.shapes.add_table(
        num_rows, num_cols,
        Inches(0.8), Inches(1.5),
        table_width, table_height
    )
    table = table_shape.table

    # Set column widths evenly
    single_col_width = int(table_width / num_cols)
    for col in table.columns:
        col.width = single_col_width

    # Style header row
    for i, header in enumerate(headers):
        cell = table.cell(0, i)
        cell.text = _clean_markdown(header)
        cell.fill.solid()
        cell.fill.fore_color.rgb = TABLE_HEADER_BG
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        for p in cell.text_frame.paragraphs:
            p.font.size = Pt(12)
            p.font.color.rgb = ACCENT_BLUE
            p.font.bold = True
            p.font.name = FONT_NAME

    # Style data rows
    for r, row_data in enumerate(rows):
        bg = TABLE_ROW_BG if r % 2 == 0 else TABLE_ALT_BG
        for c in range(num_cols):
            cell = table.cell(r + 1, c)
            cell_text = _clean_markdown(row_data[c]) if c < len(row_data) else ""
            cell.text = _truncate(cell_text, 40)
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            for p in cell.text_frame.paragraphs:
                p.font.size = Pt(11)
                p.font.color.rgb = WHITE
                p.font.name = FONT_NAME


def _build_swot(prs, section_text):
    """Slide: SWOT Analysis in 2x2 grid."""
    slide = prs.slides.add_slide(_get_blank_layout(prs))
    _set_slide_bg(slide)

    _slide_title(slide, Inches(0.8), Inches(0.5), Inches(11), "SWOT Analysis")

    # Parse SWOT items - handles multiple formats:
    # Format A: Separate bullets under each heading
    # Format B: "- **Strengths:** item1, item2, item3" (comma-separated in one line)
    swot = {"Strengths": [], "Weaknesses": [], "Opportunities": [], "Threats": []}
    current = None

    for line in section_text.split("\n"):
        stripped = line.strip()
        lower = stripped.lower()

        # Check if this line is a SWOT category header or inline item
        matched_key = None
        for key in swot:
            if key.lower() in lower:
                matched_key = key
                break

        if matched_key:
            # Check for inline format: "- **Strengths:** item1, item2, item3"
            # or "**Strengths:** item1, item2"
            colon_match = re.search(r'(?:strengths|weaknesses|opportunities|threats)\s*:?\s*:(.+)',
                                     stripped, re.IGNORECASE)
            if colon_match:
                items_text = colon_match.group(1).strip()
                # Split on commas and periods to get individual items
                items = [i.strip().rstrip(".") for i in re.split(r'[,;]', items_text) if i.strip()]
                swot[matched_key].extend(items)
                current = matched_key
            elif ":" in stripped or stripped.startswith("#") or stripped.startswith("*"):
                current = matched_key
        elif current and stripped.startswith(("-", "*", "•")):
            clean = re.sub(r'^[-*•]\s*', '', stripped)
            clean = _clean_markdown(clean).rstrip(".")

            # Check if this bullet has a SWOT keyword with items after colon
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

    # 2x2 grid
    grid = [
        (Inches(0.8), Inches(1.4), "Strengths", ACCENT_GREEN, swot["Strengths"]),
        (Inches(6.8), Inches(1.4), "Weaknesses", ACCENT_RED, swot["Weaknesses"]),
        (Inches(0.8), Inches(4.2), "Opportunities", ACCENT_BLUE, swot["Opportunities"]),
        (Inches(6.8), Inches(4.2), "Threats", ACCENT_YELLOW, swot["Threats"]),
    ]

    for left, top, title, color, items in grid:
        # Background box
        _add_shape_bg(slide, left, top, Inches(5.7), Inches(2.6), CARD_BG)
        # Color accent bar at top
        _add_shape_bg(slide, left, top, Inches(5.7), Inches(0.06), color)
        # Title
        _add_text_box(slide, left + Inches(0.2), top + Inches(0.15),
                      Inches(5), Inches(0.4), title,
                      font_size=18, color=color, bold=True)
        # Items
        if items:
            txBox = slide.shapes.add_textbox(
                left + Inches(0.2), top + Inches(0.6), Inches(5.3), Inches(1.8))
            tf = txBox.text_frame
            tf.word_wrap = True
            _add_bullet_list(tf, items[:4], font_size=12, color=LIGHT_GRAY, max_chars=70)
        else:
            # Show placeholder if no items parsed
            _add_text_box(slide, left + Inches(0.2), top + Inches(0.8),
                          Inches(5), Inches(0.4), "No data available",
                          font_size=12, color=LIGHT_GRAY)


def _build_opportunities_threats(prs, section_text):
    """Slide: Opportunities & Threats ranked lists."""
    slide = prs.slides.add_slide(_get_blank_layout(prs))
    _set_slide_bg(slide)

    _slide_title(slide, Inches(0.8), Inches(0.5), Inches(11), "Opportunities & Threats")

    # Parse into opportunities and threats
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

    # Left: Opportunities (green)
    _add_shape_bg(slide, Inches(0.8), Inches(1.4), Inches(5.7), Inches(5.5), CARD_BG)
    _add_shape_bg(slide, Inches(0.8), Inches(1.4), Inches(5.7), Inches(0.06), ACCENT_GREEN)
    _add_text_box(slide, Inches(1.0), Inches(1.6), Inches(5), Inches(0.5),
                  "Opportunities", font_size=22, color=ACCENT_GREEN, bold=True)
    if opportunities:
        txBox = slide.shapes.add_textbox(Inches(1.0), Inches(2.3), Inches(5.3), Inches(4.4))
        tf = txBox.text_frame
        tf.word_wrap = True
        _add_bullet_list(tf, opportunities[:5], font_size=13, color=WHITE,
                         bold_first=True, max_chars=80)

    # Right: Threats (red)
    _add_shape_bg(slide, Inches(6.8), Inches(1.4), Inches(5.7), Inches(5.5), CARD_BG)
    _add_shape_bg(slide, Inches(6.8), Inches(1.4), Inches(5.7), Inches(0.06), ACCENT_RED)
    _add_text_box(slide, Inches(7.0), Inches(1.6), Inches(5), Inches(0.5),
                  "Threats", font_size=22, color=ACCENT_RED, bold=True)
    if threats:
        txBox = slide.shapes.add_textbox(Inches(7.0), Inches(2.3), Inches(5.3), Inches(4.4))
        tf = txBox.text_frame
        tf.word_wrap = True
        _add_bullet_list(tf, threats[:5], font_size=13, color=WHITE,
                         bold_first=True, max_chars=80)


def _build_recommendations(prs, section_text):
    """Slide: Strategic Recommendations with timeline."""
    slide = prs.slides.add_slide(_get_blank_layout(prs))
    _set_slide_bg(slide)

    _slide_title(slide, Inches(0.8), Inches(0.5), Inches(11), "Strategic Recommendations")

    # Parse into 30-day, 3-month, 12-month
    timeframes = [
        ("30 Days", ACCENT_GREEN, []),
        ("3 Months", ACCENT_BLUE, []),
        ("12 Months", ACCENT_PURPLE, []),
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

    # Three columns
    col_width = Inches(3.8)
    for i, (label, color, items) in enumerate(timeframes):
        left = Inches(0.8) + (col_width + Inches(0.15)) * i
        _add_shape_bg(slide, left, Inches(1.4), col_width, Inches(5.5), CARD_BG)
        _add_shape_bg(slide, left, Inches(1.4), col_width, Inches(0.06), color)
        _add_text_box(slide, left + Inches(0.2), Inches(1.6), col_width - Inches(0.4),
                      Inches(0.5), label, font_size=20, color=color, bold=True)

        if items:
            txBox = slide.shapes.add_textbox(
                left + Inches(0.2), Inches(2.3), col_width - Inches(0.4), Inches(4.4))
            tf = txBox.text_frame
            tf.word_wrap = True
            _add_bullet_list(tf, items[:4], font_size=12, color=LIGHT_GRAY,
                             bold_first=True, max_chars=65)


def _build_closing_slide(prs):
    """Final slide with branding."""
    slide = prs.slides.add_slide(_get_blank_layout(prs))
    _set_slide_bg(slide, BG_DARKER)

    _add_shape_bg(slide, Inches(0), Inches(0), SLIDE_WIDTH, Inches(0.06), ACCENT_BLUE)

    _add_text_box(slide, Inches(1), Inches(2.5), Inches(11), Inches(1),
                  "Thank You", font_size=44, color=WHITE, bold=True,
                  alignment=PP_ALIGN.CENTER)

    _add_shape_bg(slide, Inches(5.5), Inches(3.7), Inches(2.3), Inches(0.04), ACCENT_BLUE)

    _add_text_box(slide, Inches(1), Inches(4.0), Inches(11), Inches(0.6),
                  "AI Market Research Agent", font_size=20,
                  color=ACCENT_BLUE, bold=True, alignment=PP_ALIGN.CENTER)

    _add_text_box(slide, Inches(1), Inches(4.7), Inches(11), Inches(0.6),
                  "Multi-Agent Intelligence System  |  Powered by CrewAI",
                  font_size=16, color=LIGHT_GRAY, alignment=PP_ALIGN.CENTER)

    _add_shape_bg(slide, Inches(0), Inches(7.44), SLIDE_WIDTH, Inches(0.06), ACCENT_BLUE)


# ── Main Generator ────────────────────────────────────────

def generate_pptx(report_text: str, company: str, competitors: list[str],
                  output_path: str) -> str:
    """
    Convert a markdown report into a professional PowerPoint presentation.

    Args:
        report_text: The raw markdown report content
        company: Company name
        competitors: List of competitor names
        output_path: Path to save the .pptx file (e.g. output/report.pptx)

    Returns:
        The output file path
    """
    # Ensure output directory exists
    output_dir = os.path.dirname(output_path)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)

    prs = Presentation()
    prs.slide_width = SLIDE_WIDTH
    prs.slide_height = SLIDE_HEIGHT

    date_str = datetime.now().strftime("%B %d, %Y")

    # Parse report sections
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
            # Try to split by competitor name subsections
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
            # Also try to build SWOT from same section
            if "swot" in sections[key].lower():
                _build_swot(prs, sections[key])
            break

    # Slide: SWOT (standalone if exists)
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

    # Save
    prs.save(output_path)
    return output_path
