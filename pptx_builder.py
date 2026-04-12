import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from chart_generator import generate_chart, make_progress_rings

# ── Brand colors ─────────────────────────────────────────────────────────
RED        = RGBColor(0xEF, 0x44, 0x44)
DARK       = RGBColor(0x2C, 0x2C, 0x2C)
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_GREY = RGBColor(0xF5, 0xF5, 0xF5)
MID_GREY   = RGBColor(0xA0, 0xA0, 0xA0)
ACCENT     = RGBColor(0x1A, 0x1A, 0x2E)
CARD_BG    = RGBColor(0xF9, 0xF9, 0xF9)

SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.5)

THEMES = {
    "red": {
        "primary": RGBColor(0xEF, 0x44, 0x44),
        "dark":    RGBColor(0x2C, 0x2C, 0x2C),
        "light":   RGBColor(0xF5, 0xF5, 0xF5),
        "mid":     RGBColor(0xA0, 0xA0, 0xA0),
    },
    "green": {
        "primary": RGBColor(0x4A, 0x7C, 0x59),
        "dark":    RGBColor(0x1E, 0x3A, 0x2F),
        "light":   RGBColor(0xF0, 0xF4, 0xF1),
        "mid":     RGBColor(0x7A, 0x9A, 0x84),
    },
    "blue": {
        "primary": RGBColor(0x1A, 0x6B, 0xB5),
        "dark":    RGBColor(0x0D, 0x2B, 0x4E),
        "light":   RGBColor(0xF0, 0xF4, 0xFA),
        "mid":     RGBColor(0x7A, 0x9A, 0xC0),
    },
}

def set_theme(theme_name="green"):
    global RED, DARK, LIGHT_GREY, MID_GREY
    t = THEMES.get(theme_name, THEMES["green"])
    RED        = t["primary"]
    DARK       = t["dark"]
    LIGHT_GREY = t["light"]
    MID_GREY   = t["mid"]


# ── Helpers ──────────────────────────────────────────────────────────────
def _add_text(slide, text, left, top, width, height,
              font_name="Calibri", font_size=14, bold=False,
              color=None, align=PP_ALIGN.LEFT, wrap=True):
    txb = slide.shapes.add_textbox(left, top, width, height)
    tf  = txb.text_frame
    tf.word_wrap = wrap
    p   = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = str(text)
    run.font.name  = font_name
    run.font.size  = Pt(font_size)
    run.font.bold  = bold
    run.font.color.rgb = color if color else DARK
    return txb


def _add_rect(slide, left, top, width, height, fill_color,
              line_color=None, line_width=Pt(0)):
    shape = slide.shapes.add_shape(1, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if line_color:
        shape.line.color.rgb = line_color
        shape.line.width = line_width
    else:
        shape.line.fill.background()
    return shape


def _add_circle(slide, left, top, size, fill_color, text, font_size=11):
    circle = slide.shapes.add_shape(9, left, top, size, size)
    circle.fill.solid()
    circle.fill.fore_color.rgb = fill_color
    circle.line.fill.background()
    tf = circle.text_frame
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    run = tf.paragraphs[0].add_run()
    run.text = str(text)
    run.font.name  = "Calibri"
    run.font.size  = Pt(font_size)
    run.font.bold  = True
    run.font.color.rgb = WHITE


def _slide_header(slide, title, bg_color=DARK):
    """Enhanced header with decorative elements."""
    # Left red strip
    _add_rect(slide, 0, 0, Inches(0.12), SLIDE_H, RED)
    # Main header bar
    _add_rect(slide, Inches(0.12), 0, SLIDE_W - Inches(0.12), Inches(1.2), bg_color)
    # Subtle secondary rect — slightly lighter, right portion only
    _add_rect(slide, SLIDE_W - Inches(3.5), 0, Inches(3.5), Inches(1.2),
              RGBColor(0x3C, 0x3C, 0x3C) if bg_color == DARK else
              RGBColor(0x2A, 0x4A, 0x3A))
    # Title
    _add_text(slide, title,
              left=Inches(0.35), top=Inches(0.17),
              width=Inches(9.5), height=Inches(0.85),
              font_size=24, bold=True, color=WHITE)
    # Red underline
    _add_rect(slide, Inches(0.35), Inches(1.18), Inches(3.8), Inches(0.06), RED)
    # Decorative top-right circle
    _add_circle(slide, SLIDE_W - Inches(0.85), Inches(0.28),
                Inches(0.6), RED, "", font_size=6)


# ══════════════════════════════════════════════════════════════════════════
# LAYOUT 1 — TITLE SLIDE
# ══════════════════════════════════════════════════════════════════════════
def build_title_slide(slide, title, subtitle):
    """Visually rich title slide with layered background."""
    # Base white
    _add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, WHITE)

    # Dark background block — right side
    _add_rect(slide, SLIDE_W - Inches(5.5), 0, Inches(5.5), SLIDE_H, DARK)

    # Large decorative diamond — overlapping the two halves
    diamond = slide.shapes.add_shape(4, Inches(6.8), Inches(0.3),
                                     Inches(5.2), Inches(7.0))
    diamond.fill.solid()
    diamond.fill.fore_color.rgb = RGBColor(0x3A, 0x3A, 0x3A)
    diamond.line.fill.background()

    # Second smaller diamond — offset for depth
    diamond2 = slide.shapes.add_shape(4, Inches(8.5), Inches(1.5),
                                      Inches(3.2), Inches(4.5))
    diamond2.fill.solid()
    diamond2.fill.fore_color.rgb = RED
    diamond2.line.fill.background()

    # Left red vertical strip
    _add_rect(slide, 0, 0, Inches(0.12), SLIDE_H, RED)

    # Top-left red corner accent
    _add_rect(slide, Inches(0.12), 0, Inches(3.5), Inches(0.08), RED)

    # Bottom dark band
    _add_rect(slide, 0, SLIDE_H - Inches(0.6), SLIDE_W, Inches(0.6), DARK)

    # Red horizontal line accent
    _add_rect(slide, Inches(0.4), Inches(4.15), Inches(5.2), Inches(0.07), RED)

    # Small decorative circles
    _add_circle(slide, Inches(0.4), Inches(1.2), Inches(0.3), RED, "", font_size=6)
    _add_circle(slide, Inches(0.4), Inches(1.7), Inches(0.18),
                RGBColor(0xCC, 0xCC, 0xCC), "", font_size=6)

    # Title — large and bold
    _add_text(slide, title,
              left=Inches(0.5), top=Inches(1.8),
              width=Inches(7.2), height=Inches(2.1),
              font_size=40, bold=True, color=RED)

    # Subtitle
    _add_text(slide, subtitle,
              left=Inches(0.5), top=Inches(4.3),
              width=Inches(7.0), height=Inches(1.0),
              font_size=15, color=DARK)

    # Bottom left — "Powered by AI" style tag
    _add_text(slide, "AI-Generated Presentation  •  Confidential",
              left=Inches(0.5), top=SLIDE_H - Inches(0.5),
              width=Inches(6.0), height=Inches(0.4),
              font_size=9, color=MID_GREY)

# ══════════════════════════════════════════════════════════════════════════
# LAYOUT 2 — EXECUTIVE SUMMARY
# ══════════════════════════════════════════════════════════════════════════
def build_executive_summary_slide(slide, title, key_points):
    _add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, LIGHT_GREY)
    _add_rect(slide, 0, 0, SLIDE_W, Inches(1.5), RED)
    _add_rect(slide, 0, 0, Inches(0.12), SLIDE_H, RGBColor(0xCC, 0x20, 0x20))
    _add_rect(slide, 0, SLIDE_H - Inches(0.45), SLIDE_W, Inches(0.45), DARK)

    _add_text(slide, title.upper(),
              left=Inches(0.5), top=Inches(0.32),
              width=Inches(12), height=Inches(0.82),
              font_size=28, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    n       = min(len(key_points), 5)
    card_w  = Inches(2.32)
    card_h  = Inches(5.3)
    gap     = Inches(0.17)
    total_w = n * card_w + (n - 1) * gap
    start_x = (SLIDE_W - total_w) / 2

    for i, point in enumerate(key_points[:n]):
        x = start_x + i * (card_w + gap)
        y = Inches(1.6)

        _add_rect(slide, x + Inches(0.04), y + Inches(0.04),
                  card_w, card_h, RGBColor(0xCC, 0xCC, 0xCC))
        _add_rect(slide, x, y, card_w, card_h, WHITE,
                  line_color=RGBColor(0xE0, 0xE0, 0xE0), line_width=Pt(1))

        bar_color = RED if i % 2 == 0 else DARK
        _add_rect(slide, x, y, card_w, Inches(0.1), bar_color)

        _add_circle(slide, x + Inches(0.88), y + Inches(0.2),
                    Inches(0.56), bar_color, str(i + 1), font_size=11)

        parts  = point.split(":", 1) if ":" in point else [point, ""]
        header = parts[0].strip()
        body   = parts[1].strip() if len(parts) > 1 else ""

        _add_text(slide, header,
                  left=x + Inches(0.1), top=y + Inches(0.9),
                  width=card_w - Inches(0.2), height=Inches(0.75),
                  font_size=13, bold=True, color=DARK, align=PP_ALIGN.CENTER)

        _add_rect(slide, x + Inches(0.4), y + Inches(1.72),
                  card_w - Inches(0.8), Inches(0.04), RGBColor(0xEE, 0xEE, 0xEE))

        if body:
            _add_text(slide, body,
                      left=x + Inches(0.12), top=y + Inches(1.85),
                      width=card_w - Inches(0.24), height=Inches(3.2),
                      font_size=11, color=MID_GREY, align=PP_ALIGN.CENTER)

        _add_circle(slide, x + card_w / 2 - Inches(0.12),
                    y + card_h - Inches(0.35),
                    Inches(0.24), bar_color, "", font_size=6)


# ══════════════════════════════════════════════════════════════════════════
# LAYOUT 3 — SPLIT PANEL
# ══════════════════════════════════════════════════════════════════════════
def build_split_panel_slide(slide, title, key_points, context_text=""):
    _add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, WHITE)
    _add_rect(slide, 0, 0, Inches(0.12), SLIDE_H, RED)

    panel_w = Inches(4.3)
    _add_rect(slide, Inches(0.12), 0, panel_w, SLIDE_H, DARK)

    tri = slide.shapes.add_shape(5, Inches(0.12), SLIDE_H - Inches(2.2),
                                 Inches(2.0), Inches(2.2))
    tri.fill.solid()
    tri.fill.fore_color.rgb = RGBColor(0x3A, 0x10, 0x10)
    tri.line.fill.background()

    _add_rect(slide, Inches(0.35), Inches(0.4),
              Inches(0.35), Inches(0.35), RED)

    _add_text(slide, title,
              left=Inches(0.35), top=Inches(1.1),
              width=Inches(3.8), height=Inches(2.8),
              font_size=28, bold=True, color=WHITE)

    _add_rect(slide, Inches(0.35), Inches(4.05),
              Inches(2.2), Inches(0.07), RED)

    context = context_text or ""
    if context:
        _add_text(slide, context,
                  left=Inches(0.35), top=Inches(4.25),
                  width=Inches(3.8), height=Inches(2.8),
                  font_size=12, color=MID_GREY)

    right_x = Inches(4.7)
    right_w = SLIDE_W - right_x - Inches(0.25)

    n = min(len(key_points), 5)
    if n == 0:
        return

    available_h = SLIDE_H - Inches(0.3)
    gap = Inches(0.14)
    card_h = (available_h - gap * (n - 1)) / n
    start_y = Inches(0.15)

    for i, point in enumerate(key_points[:n]):
        y = start_y + i * (card_h + gap)

        if ":" in point:
            parts = point.split(":", 1)
            label = parts[0].strip()
            desc  = parts[1].strip()
        else:
            label = ""
            desc  = point

        _add_rect(slide, right_x, y, right_w, card_h,
                  LIGHT_GREY if i % 2 == 0 else WHITE,
                  line_color=RGBColor(0xE8, 0xE8, 0xE8), line_width=Pt(0.5))

        _add_rect(slide, right_x, y, Inches(0.07), card_h, RED)

        _add_circle(slide, right_x + Inches(0.1), y + card_h / 2 - Inches(0.26),
                    Inches(0.52), RED, str(i + 1), font_size=9)

        if label:
            _add_text(slide, label,
                      left=right_x + Inches(0.75), top=y + Inches(0.1),
                      width=right_w - Inches(0.85), height=Inches(0.38),
                      font_size=13, bold=True, color=RED)
            _add_text(slide, desc,
                      left=right_x + Inches(0.75), top=y + Inches(0.42),
                      width=right_w - Inches(0.85), height=card_h - Inches(0.5),
                      font_size=12, color=DARK)
        else:
            _add_text(slide, desc,
                      left=right_x + Inches(0.75), top=y + Inches(0.12),
                      width=right_w - Inches(0.85), height=card_h - Inches(0.2),
                      font_size=13, color=DARK)


# ══════════════════════════════════════════════════════════════════════════
# LAYOUT 4 — TIMELINE
# ══════════════════════════════════════════════════════════════════════════
def build_timeline_slide(slide, title, key_points):
    _add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, WHITE)
    _slide_header(slide, title)

    n = min(len(key_points), 6)
    if n == 0:
        return

    spine_y = Inches(3.8)
    spine_x = Inches(0.8)
    spine_w = SLIDE_W - Inches(1.2)
    spine_h = Inches(0.06)
    _add_rect(slide, spine_x, spine_y, spine_w, spine_h, RED)

    dot_size = Inches(0.28)
    item_w   = spine_w / n

    for i, point in enumerate(key_points[:n]):
        parts = point.split(":", 1) if ":" in point else [f"Step {i+1}", point]
        year  = parts[0].strip()
        event = parts[1].strip() if len(parts) > 1 else ""

        cx = spine_x + i * item_w + item_w / 2
        dot_x = cx - dot_size / 2
        dot_y = spine_y - dot_size / 2 + spine_h / 2
        circle = slide.shapes.add_shape(9, dot_x, dot_y, dot_size, dot_size)
        circle.fill.solid()
        circle.fill.fore_color.rgb = RED
        circle.line.color.rgb = WHITE
        circle.line.width = Pt(1.5)

        text_w = item_w - Inches(0.1)
        text_x = cx - text_w / 2

        if i % 2 == 0:
            _add_text(slide, year,
                      left=text_x, top=Inches(2.6),
                      width=text_w, height=Inches(0.45),
                      font_size=15, bold=True, color=RED,
                      align=PP_ALIGN.CENTER)
            _add_rect(slide, cx - Inches(0.02), Inches(3.05),
                      Inches(0.04), spine_y - Inches(3.05), MID_GREY)
            _add_text(slide, event,
                      left=text_x, top=spine_y + Inches(0.35),
                      width=text_w, height=Inches(1.8),
                      font_size=11, color=DARK, align=PP_ALIGN.CENTER)
        else:
            _add_text(slide, year,
                      left=text_x, top=spine_y + Inches(0.35),
                      width=text_w, height=Inches(0.45),
                      font_size=15, bold=True, color=RED,
                      align=PP_ALIGN.CENTER)
            _add_rect(slide, cx - Inches(0.02), spine_y + Inches(0.8),
                      Inches(0.04), Inches(0.5), MID_GREY)
            _add_text(slide, event,
                      left=text_x, top=spine_y + Inches(0.85),
                      width=text_w, height=Inches(2.5),
                      font_size=11, color=DARK, align=PP_ALIGN.CENTER)

# ══════════════════════════════════════════════════════════════════════════
# LAYOUT 5 — KPI STATS
# ══════════════════════════════════════════════════════════════════════════
def build_kpi_stats_slide(slide, title, key_points):
    _add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, WHITE)
    _slide_header(slide, title)

    _add_rect(slide, 0, Inches(4.5), SLIDE_W, Inches(3.0),
              RGBColor(0xF8, 0xF8, 0xF8))

    n = min(len(key_points), 4)
    if n == 0:
        return

    card_w  = (SLIDE_W - Inches(1.0)) / n
    card_h  = Inches(5.3)
    start_x = Inches(0.5)
    start_y = Inches(1.35)

    for i, point in enumerate(key_points[:n]):
        parts   = [p.strip() for p in point.split(":", 2)]
        label   = parts[0] if len(parts) > 0 else ""
        value   = parts[1] if len(parts) > 1 else ""
        context = parts[2] if len(parts) > 2 else ""

        x = start_x + i * card_w

        _add_rect(slide, x + Inches(0.05), start_y + Inches(0.05),
                  card_w - Inches(0.15), card_h, RGBColor(0xDD, 0xDD, 0xDD))

        bg = LIGHT_GREY if i % 2 == 0 else WHITE
        _add_rect(slide, x, start_y, card_w - Inches(0.15), card_h, bg,
                  line_color=RGBColor(0xE0, 0xE0, 0xE0), line_width=Pt(1))

        _add_rect(slide, x, start_y, card_w - Inches(0.15), Inches(0.14), RED)

        _add_text(slide, label.upper(),
                  left=x + Inches(0.12), top=start_y + Inches(0.25),
                  width=card_w - Inches(0.3), height=Inches(0.55),
                  font_size=12, bold=True, color=MID_GREY,
                  align=PP_ALIGN.CENTER)

        _add_rect(slide, x + Inches(0.3), start_y + Inches(0.85),
                  card_w - Inches(0.75), Inches(0.04),
                  RGBColor(0xDD, 0xDD, 0xDD))

        _add_text(slide, value,
                  left=x + Inches(0.08), top=start_y + Inches(0.95),
                  width=card_w - Inches(0.2), height=Inches(1.7),
                  font_size=44, bold=True, color=RED,
                  align=PP_ALIGN.CENTER)

        bar_w_total = card_w - Inches(0.6)
        _add_rect(slide, x + Inches(0.25), start_y + Inches(2.75),
                  bar_w_total, Inches(0.1), RGBColor(0xEE, 0xEE, 0xEE))
        fill_pct = max(0.1, 0.75 - (i * 0.1))
        _add_rect(slide, x + Inches(0.25), start_y + Inches(2.75),
                  bar_w_total * fill_pct, Inches(0.1), RED)

        if context:
            _add_text(slide, context,
                      left=x + Inches(0.12), top=start_y + Inches(3.0),
                      width=card_w - Inches(0.3), height=Inches(2.1),
                      font_size=12, color=DARK, align=PP_ALIGN.CENTER)

        _add_circle(slide, x + card_w / 2 - Inches(0.42),
                    start_y + card_h - Inches(0.55),
                    Inches(0.38), DARK, str(i + 1), font_size=8)


# ══════════════════════════════════════════════════════════════════════════
# LAYOUT 6 — TWO COLUMN COMPARE
# ══════════════════════════════════════════════════════════════════════════
def build_two_col_compare_slide(slide, title, key_points, metadata=None):
    if metadata is None:
        metadata = {}

    left_label  = metadata.get("left_label", "Column A")
    right_label = metadata.get("right_label", "Column B")

    _add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, WHITE)
    _slide_header(slide, title)

    mid     = SLIDE_W / 2
    col_w   = mid - Inches(0.7)
    start_y = Inches(1.4)

    half = max(1, len(key_points) // 2)
    left_points  = key_points[:half]
    right_points = key_points[half:half * 2]

    for col_idx, (label, points, x_start) in enumerate([
        (left_label,  left_points,  Inches(0.5)),
        (right_label, right_points, mid + Inches(0.2))
    ]):
        col_color = RED if col_idx == 0 else DARK

        _add_rect(slide, x_start, start_y, col_w, Inches(0.55), col_color)
        _add_text(slide, label.upper(),
                  left=x_start + Inches(0.1), top=start_y + Inches(0.08),
                  width=col_w - Inches(0.2), height=Inches(0.4),
                  font_size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

        item_h = Inches(0.88)
        gap    = Inches(0.12)
        for j, point in enumerate(points[:4]):
            y = start_y + Inches(0.7) + j * (item_h + gap)
            _add_rect(slide, x_start, y, col_w, item_h, LIGHT_GREY,
                      line_color=RGBColor(0xE0, 0xE0, 0xE0), line_width=Pt(0.5))
            _add_rect(slide, x_start, y, Inches(0.07), item_h, col_color)
            _add_text(slide, point,
                      left=x_start + Inches(0.18), top=y + Inches(0.1),
                      width=col_w - Inches(0.25), height=item_h - Inches(0.15),
                      font_size=12, color=DARK)

    _add_rect(slide, mid - Inches(0.05), start_y, Inches(0.06),
              SLIDE_H - start_y - Inches(0.3), RGBColor(0xE0, 0xE0, 0xE0))


# ══════════════════════════════════════════════════════════════════════════
# LAYOUT 7 — GRID 4 COLUMN
# ══════════════════════════════════════════════════════════════════════════
def build_grid_4col_slide(slide, title, key_points):
    _add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, LIGHT_GREY)
    _slide_header(slide, title, bg_color=DARK)

    while len(key_points) < 4:
        key_points.append("—")

    col_w   = (SLIDE_W - Inches(0.8)) / 4
    col_h   = Inches(5.3)
    start_x = Inches(0.4)
    start_y = Inches(1.5)
    gap     = Inches(0.07)

    col_colors = [RED, DARK, RGBColor(0x55, 0x55, 0x55), RGBColor(0x77, 0x77, 0x77)]

    for i, point in enumerate(key_points[:4]):
        x = start_x + i * (col_w + gap)

        parts  = point.split(":", 1) if ":" in point else [f"Category {i+1}", point]
        header = parts[0].strip()
        body   = parts[1].strip() if len(parts) > 1 else ""

        _add_rect(slide, x, start_y, col_w, col_h, WHITE,
                  line_color=RGBColor(0xDD, 0xDD, 0xDD), line_width=Pt(1))

        _add_rect(slide, x, start_y, col_w, Inches(0.65), col_colors[i])
        _add_text(slide, header,
                  left=x + Inches(0.08), top=start_y + Inches(0.1),
                  width=col_w - Inches(0.16), height=Inches(0.5),
                  font_size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

        items = [b.strip() for b in re.split(r'[|\n]', body) if b.strip()] if body else []
        if not items and body:
            items = [body]

        item_h = Inches(0.7)
        for j, item in enumerate(items[:5]):
            iy = start_y + Inches(0.75) + j * (item_h + Inches(0.05))
            _add_rect(slide, x + Inches(0.08), iy,
                      col_w - Inches(0.16), item_h, LIGHT_GREY,
                      line_color=RGBColor(0xEE, 0xEE, 0xEE), line_width=Pt(0.5))
            _add_text(slide, item,
                      left=x + Inches(0.14), top=iy + Inches(0.08),
                      width=col_w - Inches(0.28), height=item_h - Inches(0.1),
                      font_size=11, color=DARK, align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════════════════
# LAYOUT 8 — DATA TABLE
# ══════════════════════════════════════════════════════════════════════════
def build_data_table_slide(slide, title, table_data):
    _add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, WHITE)
    _slide_header(slide, title)

    if not table_data or not table_data.get("headers"):
        _add_text(slide, "No table data available.",
                  left=Inches(0.5), top=Inches(2.0),
                  width=Inches(12), height=Inches(1.0),
                  font_size=14, color=MID_GREY)
        return

    headers = table_data["headers"]
    rows    = table_data["rows"]
    n_cols  = len(headers)
    col_w   = (SLIDE_W - Inches(0.8)) / n_cols
    row_h   = Inches(0.62)
    start_x = Inches(0.4)
    start_y = Inches(1.4)

    for j, h in enumerate(headers):
        x = start_x + j * col_w
        _add_rect(slide, x, start_y, col_w, row_h, RED)
        _add_text(slide, str(h),
                  left=x + Inches(0.08), top=start_y + Inches(0.08),
                  width=col_w - Inches(0.12), height=row_h - Inches(0.1),
                  font_size=12, bold=True, color=WHITE)

    for i, row in enumerate(rows[:8]):
        y      = start_y + (i + 1) * row_h
        bg_col = LIGHT_GREY if i % 2 == 0 else WHITE
        for j, cell in enumerate(row[:n_cols]):
            x = start_x + j * col_w
            _add_rect(slide, x, y, col_w, row_h, bg_col,
                      line_color=RGBColor(0xE0, 0xE0, 0xE0), line_width=Pt(0.5))
            _add_text(slide, str(cell),
                      left=x + Inches(0.08), top=y + Inches(0.08),
                      width=col_w - Inches(0.12), height=row_h - Inches(0.1),
                      font_size=11, color=DARK)


# ══════════════════════════════════════════════════════════════════════════
# LAYOUT 9 — CHART
# ══════════════════════════════════════════════════════════════════════════
def build_chart_slide(slide, title, chart_path):
    """Chart slide with left info panel + chart on right."""
    _add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, WHITE)
    _add_rect(slide, 0, 0, Inches(0.12), SLIDE_H, RED)

    # Left dark panel with title
    panel_w = Inches(3.2)
    _add_rect(slide, Inches(0.12), 0, panel_w, SLIDE_H, DARK)

    # Decorative element in left panel
    _add_rect(slide, Inches(0.35), Inches(0.4),
              Inches(0.3), Inches(0.3), RED)

    # Title in left panel
    _add_text(slide, title,
              left=Inches(0.35), top=Inches(1.0),
              width=Inches(2.8), height=Inches(2.5),
              font_size=22, bold=True, color=WHITE)

    # Red accent line
    _add_rect(slide, Inches(0.35), Inches(3.7),
              Inches(1.8), Inches(0.06), RED)

    # Subtitle in left panel
    _add_text(slide, "Data visualization from\ndocument analysis",
              left=Inches(0.35), top=Inches(3.9),
              width=Inches(2.8), height=Inches(1.0),
              font_size=10, color=MID_GREY)

    # Small stat circle at bottom of panel
    _add_circle(slide, Inches(1.0), Inches(5.8),
                Inches(0.8), RED, "📊", font_size=14)

    # Chart area — right side
    chart_x = Inches(3.5)
    chart_w = SLIDE_W - chart_x - Inches(0.2)

    # Light background for chart area
    _add_rect(slide, chart_x, Inches(0.15), chart_w, SLIDE_H - Inches(0.3),
              LIGHT_GREY)

    if chart_path and os.path.exists(chart_path):
        slide.shapes.add_picture(
            chart_path,
            left=chart_x + Inches(0.1),
            top=Inches(0.25),
            width=chart_w - Inches(0.2),
            height=SLIDE_H - Inches(0.5)
        )
    else:
        _add_text(slide, "Chart could not be generated.",
                  left=chart_x, top=Inches(3.0),
                  width=chart_w, height=Inches(1.0),
                  font_size=14, color=MID_GREY, align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════════════════
# LAYOUT 10 — CONCLUSION
# ══════════════════════════════════════════════════════════════════════════
def build_kpi_visual_slide(slide, title, parsed_doc):
    """Dark dashboard slide with circular progress rings."""
    _add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, DARK)
    _add_rect(slide, 0, 0, Inches(0.12), SLIDE_H, RED)
    _add_rect(slide, Inches(0.12), 0, SLIDE_W - Inches(0.12), Inches(1.2),
              RGBColor(0x1A, 0x1A, 0x1A))

    _add_text(slide, title,
              left=Inches(0.35), top=Inches(0.17),
              width=Inches(12.5), height=Inches(0.85),
              font_size=24, bold=True, color=WHITE)

    _add_rect(slide, Inches(0.35), Inches(1.18), Inches(3.8), Inches(0.06), RED)

    # Extract stats from parsed_doc
    stats = parsed_doc.get("key_stats", [])
    rings_data = []
    for idx, s in enumerate(stats[:4]):
        raw = str(s["value"])
        unit = "%" if "%" in raw else \
               "B" if "B" in raw or "billion" in raw.lower() else \
               "M" if "M" in raw or "million" in raw.lower() else ""
        nums = re.findall(r'[\d]+\.?\d*', raw)
        if nums:
            try:
                val = float(nums[0])
                # Use different fill percentages per ring for visual variety
                fill_pcts = [0.85, 0.65, 0.75, 0.45]
                max_val = val / fill_pcts[idx % 4]
                label = s["label"][:12]
                rings_data.append((label, val, max_val, unit))
            except:
                continue

    # Fallback — extract from tables if no key_stats
    if not rings_data and parsed_doc.get("tables"):
        for table in parsed_doc["tables"][:3]:
            for row in table.get("rows", [])[:4]:
                for cell in row[1:]:
                    nums = re.findall(r'[\d]+\.?\d*', str(cell))
                    if nums:
                        try:
                            val = float(nums[0])
                            if val > 0:
                                rings_data.append((str(row[0])[:12], val, val * 1.5, ""))
                                break
                        except:
                            continue
                if len(rings_data) >= 4:
                    break
            if len(rings_data) >= 4:
                break

    if rings_data:
        chart_path = make_progress_rings(rings_data, "rings.png")
        if chart_path and os.path.exists(chart_path):
            slide.shapes.add_picture(
                chart_path,
                left=Inches(0.5),
                top=Inches(1.5),
                width=Inches(12.3),
                height=Inches(5.5)
            )
    else:
        _add_text(slide, "Key metrics visualization unavailable for this document.",
                  left=Inches(0.5), top=Inches(3.5),
                  width=Inches(12), height=Inches(1.0),
                  font_size=14, color=MID_GREY, align=PP_ALIGN.CENTER)


def build_conclusion_slide(slide, title, key_points):
    _add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, DARK)
    _add_rect(slide, 0, 0, Inches(0.12), SLIDE_H, RED)

    _add_text(slide, title,
              left=Inches(0.5), top=Inches(0.35),
              width=Inches(12), height=Inches(0.9),
              font_size=30, bold=True, color=WHITE)

    _add_rect(slide, Inches(0.5), Inches(1.25), Inches(3.5), Inches(0.06), RED)

    for i, point in enumerate(key_points[:5]):
        y = Inches(1.6) + i * Inches(1.0)
        _add_rect(slide, Inches(0.5), y, SLIDE_W - Inches(1.0), Inches(0.85),
                  RGBColor(0x3C, 0x3C, 0x3C))
        _add_rect(slide, Inches(0.5), y, Inches(0.08), Inches(0.85), RED)
        _add_text(slide, "→",
                  left=Inches(0.72), top=y + Inches(0.14),
                  width=Inches(0.5), height=Inches(0.5),
                  font_size=16, bold=True, color=RED)
        _add_text(slide, point,
                  left=Inches(1.35), top=y + Inches(0.14),
                  width=Inches(11.5), height=Inches(0.6),
                  font_size=14, color=WHITE)


# ── Helpers ───────────────────────────────────────────────────────────────
def add_slide_number(slide, number):
    _add_text(slide, str(number),
              left=SLIDE_W - Inches(0.65), top=SLIDE_H - Inches(0.42),
              width=Inches(0.5), height=Inches(0.35),
              font_size=9, color=MID_GREY, align=PP_ALIGN.RIGHT)


def find_table(parsed_doc, title_hint):
    if not parsed_doc.get("tables"):
        return None
    if not title_hint:
        return parsed_doc["tables"][0]
    
    title_hint_lower = title_hint.lower()
    
    # Exact match first
    for t in parsed_doc["tables"]:
        if t["title"].lower() == title_hint_lower:
            return t
    
    # Word overlap match
    hint_words = set(title_hint_lower.split())
    best_match = None
    best_score = 0
    for t in parsed_doc["tables"]:
        table_words = set(t["title"].lower().split())
        score = len(hint_words & table_words)
        if score > best_score:
            best_score = score
            best_match = t
    
    if best_score >= 2:
        return best_match
    
    return parsed_doc["tables"][0]

import re


# ══════════════════════════════════════════════════════════════════════════
# MAIN BUILDER
# ══════════════════════════════════════════════════════════════════════════
def build_presentation(slide_plan, parsed_doc, template_path, output_path):
    prs = Presentation()
    prs.slide_width  = SLIDE_W
    prs.slide_height = SLIDE_H
    blank_layout = prs.slide_layouts[5]

    print(f"\n🏗  Building {len(slide_plan)} slides...")

    chart_path = None

    for slide_info in slide_plan:
        layout  = slide_info.get("layout_type", "split_panel")
        title   = slide_info.get("title", "")
        points  = slide_info.get("key_points", [])
        meta    = slide_info.get("metadata", {})
        use_tbl = slide_info.get("use_table")
        num     = slide_info.get("slide_number", 1)

        print(f"   Slide {num}: [{layout}] {title[:50]}")

        slide = prs.slides.add_slide(blank_layout)

        if layout == "title":
            subtitle = points[0] if points else parsed_doc.get("subtitle", "")
            build_title_slide(slide, title, subtitle)
        elif layout == "executive_summary":
            build_executive_summary_slide(slide, title, points)
        elif layout == "split_panel":
            context = meta.get("context", "")
            build_split_panel_slide(slide, title, points, context)
        elif layout == "timeline":
            build_timeline_slide(slide, title, points)
        elif layout == "kpi_stats":
            build_kpi_stats_slide(slide, title, points)
        elif layout == "two_col_compare":
            build_two_col_compare_slide(slide, title, points, meta)
        elif layout == "grid_4col":
            build_grid_4col_slide(slide, title, points)
        elif layout == "data_table":
            table_data = find_table(parsed_doc, use_tbl or title)
            build_data_table_slide(slide, title, table_data)
        elif layout == "chart":
            chart_filename = f"chart_{num}.png"
            table_data = find_table(parsed_doc, use_tbl or title)
            print(f"DEBUG use_tbl: {use_tbl}")
            print(f"DEBUG title: {title}")
            print(f"DEBUG matched table: {table_data['title'] if table_data else None}")
            chart_path = generate_chart(parsed_doc, chart_filename, table_data)
            build_chart_slide(slide, title, chart_path)
        elif layout == "kpi_visual":
             build_kpi_visual_slide(slide, title, parsed_doc)
        elif layout == "conclusion":
            build_conclusion_slide(slide, title, points)
        else:
            build_split_panel_slide(slide, title, points)

        if layout != "title":
            add_slide_number(slide, num)

    prs.save(output_path)
    print(f"\n✅ Saved: {output_path}")




# ── Entry point ───────────────────────────────────────────────────────────
if __name__ == "__main__":
    import sys
    from parser import parse_markdown
    from slide_planner import plan_slides

    API_KEY = "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"

    MD_FILE = sys.argv[1] if len(sys.argv) > 1 else "Accenture Tech Acquisition Analysis.md"
    OUTPUT  = sys.argv[2] if len(sys.argv) > 2 else MD_FILE.replace(".md", ".pptx")

    print(f"Input:  {MD_FILE}")
    print(f"Output: {OUTPUT}")

    print("Parsing markdown...")
    doc = parse_markdown(MD_FILE)

    print("Planning slides...")
    plan = plan_slides(doc, API_KEY)

    print("Building presentation...")
    build_presentation(plan, doc, None, OUTPUT)