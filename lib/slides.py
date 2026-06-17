"""Generador de presentaciones .pptx con diseño cálido de panadería.

Port parametrizado de make_julio.py. Función principal:

    build_deck(items, deck_title, subtitle=None) -> io.BytesIO

`items` es una lista de dicts con las claves:
    date, name, icon, formato, red, pilar, tipo, objetivo, estado,
    borrador (bool), guion (list[str] | None), copy (str | None)
"""

import io

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# ─── Colores ────────────────────────────────────────────────────────────────
BROWN    = RGBColor(0x4A, 0x2C, 0x0A)
TERRA    = RGBColor(0xC4, 0x62, 0x2D)
GOLD     = RGBColor(0xD4, 0x92, 0x28)
CREAM    = RGBColor(0xFD, 0xF6, 0xEC)
WARM_MED = RGBColor(0xF5, 0xE3, 0xCC)
MED_BR   = RGBColor(0x6A, 0x42, 0x20)
WHITE    = RGBColor(0xFF, 0xFF, 0xFF)
FOOTER_C = RGBColor(0x3A, 0x1E, 0x06)
DRAFT_H  = RGBColor(0xB5, 0xA9, 0xA0)
DRAFT_B  = RGBColor(0xF2, 0xEF, 0xEC)
GRAY_T   = RGBColor(0x88, 0x80, 0x78)
COPY_C   = RGBColor(0x7A, 0x5A, 0x38)
GOLD_FT  = RGBColor(0xBB, 0x99, 0x55)
DRAFT_SUB = RGBColor(0xCC, 0xC0, 0xB8)


# ─── Helpers ─────────────────────────────────────────────────────────────────
def _add_rect(slide, x, y, w, h, color):
    s = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid()
    s.fill.fore_color.rgb = color
    s.line.fill.background()
    return s


def _add_text(slide, text, x, y, w, h, size, bold=False, italic=False,
              color=WHITE, align=PP_ALIGN.LEFT, font="Calibri"):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.italic = italic
    r.font.color.rgb = color
    r.font.name = font
    return tb


def _add_multiline(slide, lines, x, y, w, h, size, color, font="Calibri",
                   space_before=1, space_after=1):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_before = Pt(space_before)
        p.space_after = Pt(space_after)
        r = p.add_run()
        r.text = line
        r.font.size = Pt(size)
        r.font.color.rgb = color
        r.font.name = font
    return tb


def _add_footer(slide, footer_text):
    _add_rect(slide, 0, 5.225, 10, 0.4, FOOTER_C)
    _add_text(slide, footer_text, 0.3, 5.235, 9.4, 0.37, 9.5,
              color=GOLD_FT, font="Calibri")


# ─── Construcción del deck ─────────────────────────────────────────────────────
def build_deck(items, deck_title, subtitle="La Portena  ·  Panaderia & Pasteleria"):
    """Construye el .pptx y devuelve un BytesIO listo para descargar."""
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(5.625)
    blank = prs.slide_layouts[6]

    footer_text = f"{subtitle.split('·')[0].strip()}  |  {deck_title}"

    # ── Portada ────────────────────────────────────────────────────────────────
    slide = prs.slides.add_slide(blank)
    _add_rect(slide, 0, 0, 10, 5.625, BROWN)
    _add_rect(slide, 0, 0, 0.35, 5.625, TERRA)
    _add_rect(slide, 0.55, 2.33, 8.9, 0.04, GOLD)

    _add_text(slide, "Planificacion de Contenidos",
              0.55, 0.75, 9.2, 1.3, 44, bold=True,
              color=WHITE, align=PP_ALIGN.CENTER, font="Georgia")
    _add_text(slide, deck_title,
              0.55, 2.45, 9.2, 0.85, 36,
              color=GOLD, align=PP_ALIGN.CENTER, font="Georgia")
    _add_text(slide, subtitle,
              0.55, 3.42, 9.2, 0.65, 18, italic=True,
              color=CREAM, align=PP_ALIGN.CENTER, font="Georgia")
    _add_text(slide, f"{len(items)} piezas de contenido programadas",
              0.55, 4.25, 9.2, 0.5, 13,
              color=RGBColor(0xAA, 0x88, 0x66), align=PP_ALIGN.CENTER, font="Calibri")

    # ── Slides de contenido ──────────────────────────────────────────────────
    for item in items:
        slide = prs.slides.add_slide(blank)
        is_draft = item.get("borrador", False)
        hdr_color = DRAFT_H if is_draft else TERRA
        bg_color = DRAFT_B if is_draft else CREAM

        _add_rect(slide, 0, 0, 10, 5.625, bg_color)
        _add_rect(slide, 0, 0, 10, 1.05, hdr_color)
        if not is_draft:
            _add_rect(slide, 0, 1.05, 10, 0.04, GOLD)

        _add_text(slide, item.get("date", ""), 0.25, 0.1, 1.7, 0.85, 28,
                  bold=True, color=WHITE, font="Georgia")
        _add_text(slide, f"{item.get('icon', '')}  {item.get('name', '')}".strip(),
                  2.05, 0.12, 6.4, 0.8, 22,
                  bold=True, color=WHITE, font="Georgia")

        badge_bg = RGBColor(0xFF, 0xF0, 0xE0) if not is_draft else RGBColor(0xDD, 0xD8, 0xD5)
        _add_rect(slide, 8.52, 0.27, 1.26, 0.5, badge_bg)
        _add_text(slide, item.get("estado", ""), 8.52, 0.27, 1.26, 0.5, 9,
                  bold=True, color=BROWN, align=PP_ALIGN.CENTER, font="Calibri")

        _add_footer(slide, footer_text)

        # ── Slides borrador ──────────────────────────────────────────────────
        if is_draft:
            _add_text(slide, "CONTENIDO PENDIENTE",
                      1.0, 1.9, 8.0, 1.1, 30, bold=True,
                      color=DRAFT_H, align=PP_ALIGN.CENTER, font="Georgia")
            sub = "  |  ".join([s for s in [item.get("formato"), item.get("red")] if s])
            _add_text(slide, sub, 1.0, 3.2, 8.0, 0.6, 15,
                      color=GRAY_T, align=PP_ALIGN.CENTER, font="Calibri")
            _add_text(slide, "Guion pendiente de elaboracion",
                      1.0, 3.9, 8.0, 0.5, 12, italic=True,
                      color=DRAFT_SUB, align=PP_ALIGN.CENTER, font="Calibri")
            continue

        # ── Columna de metadata (izquierda) ──────────────────────────────────
        _add_rect(slide, 0, 1.09, 3.25, 4.135, WARM_MED)
        _add_rect(slide, 3.25, 1.09, 0.045, 4.135, GOLD)

        meta_rows = [
            ("FORMATO",    item.get("formato") or "—"),
            ("RED SOCIAL", item.get("red") or "—"),
            ("PILAR",      item.get("pilar") or "—"),
            ("TIPO",       item.get("tipo") or "—"),
            ("OBJETIVO",   item.get("objetivo") or "—"),
        ]
        y = 1.22
        for label, value in meta_rows:
            _add_text(slide, label, 0.25, y, 2.9, 0.22, 7.5,
                      bold=True, color=TERRA, font="Calibri")
            _add_text(slide, value, 0.25, y + 0.2, 2.9, 0.4, 10.5,
                      color=MED_BR, font="Calibri")
            y += 0.72

        # ── Columna de contenido (derecha) ───────────────────────────────────
        rx, rw = 3.35, 6.45
        guion = item.get("guion") or []
        copy = item.get("copy")
        has_guion = bool(guion)
        has_copy = bool(copy)

        if has_guion and has_copy:
            gy, gh = 1.13, 2.35
            cy, ch = 3.6, 1.5
        elif has_guion:
            gy, gh = 1.13, 3.95
            cy, ch = 0, 0
        else:
            gy, gh = 0, 0
            cy, ch = 1.13, 3.95

        if has_guion:
            _add_text(slide, "GUION", rx, gy, rw, 0.24, 7.5,
                      bold=True, color=TERRA, font="Calibri")
            n = len(guion)
            fsize = 11 if n <= 3 else (10 if n <= 5 else 9)
            bullets = ["·  " + l for l in guion]
            _add_multiline(slide, bullets, rx, gy + 0.24, rw, gh - 0.24,
                           fsize, MED_BR, space_before=2, space_after=2)
            if has_copy:
                _add_rect(slide, rx, 3.53, rw, 0.03, GOLD)

        if has_copy:
            _add_text(slide, "COPY", rx, cy, rw, 0.24, 7.5,
                      bold=True, color=TERRA, font="Calibri")
            _add_multiline(slide, [copy], rx, cy + 0.24, rw, ch - 0.24,
                           9, COPY_C, space_before=2, space_after=2)

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf
