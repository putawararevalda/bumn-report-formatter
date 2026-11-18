"""
BUMN & DAMRI Report Deck Builder (bumn-report-formatter).

Builds 16:9 .pptx executive report decks on top of assets/template.pptx
following the official BUMN Taskforce & Perum DAMRI design system.
Every slide draws native, editable PowerPoint shapes, tables, and charts.

Usage:
    import sys; sys.path.insert(0, "<skill-dir>/scripts")
    from bumn_deck import BUMNDeck

    d = BUMNDeck(lang="id", footer="Taskforce BP BUMN · Perum DAMRI")
    d.cover(
        title="Taskforce BP BUMN - Perum Damri",
        subtitle="AI for Public Services: DAMRI's Transformation Roadmap for Public Transportation",
        meta="Desember 2026 · AI & Analytics Transformation"
    )
    ...
    d.save("output.pptx")
"""
from __future__ import annotations

import math
import os
import re
from typing import Any, Dict, Iterable, List, Optional, Sequence, Union

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

HERE = os.path.dirname(os.path.abspath(__file__))
SKILL_DIR = os.path.dirname(HERE)
TEMPLATE_PATH = os.path.join(SKILL_DIR, "assets", "template.pptx")
LOGOS_DIR = os.path.join(SKILL_DIR, "assets", "logos")
BACKGROUNDS_DIR = os.path.join(SKILL_DIR, "assets", "backgrounds")

# --------------------------------------------------------------------------
# Brand Tokens & Color Palette
# --------------------------------------------------------------------------
NAVY_PRIMARY = "002060"     # BUMN Navy / Brand primary
NAVY_HEADLINE = "003C79"    # Slide headline blue
BLUE_ACCENT = "174EA6"      # Section label blue
BLUE_SLATE = "123C79"       # Module sub-header & timeline headers
DARK_SLATE = "0E2841"       # Card headlines & dark elements
BODY_DARK = "1F2A37"        # Body copy dark gray
MUTED_TEXT = "5B6675"       # Subtitles, data labels, secondary info
MUTED_GRAY = "595959"       # Footnotes, subtle text

CARD_ACTIVE = "E6EDF6"      # Active card tint (Quick wins / highlighted)
CARD_NEUTRAL = "F3F5F8"     # Standard card background & panel fill
PHASE_BAND_1 = "DCE6F2"     # Timeline phase band tint 1
PHASE_BAND_2 = "EEF2F7"     # Timeline phase band tint 2
WHITE = "FFFFFF"            # Pure white

BORDER_BLUE = "B9C7DE"      # Border rule
BORDER_LIGHT = "E7E6E6"     # Light separator
BORDER_MID = "BFBFBF"       # Mid separator

ALERT_RED = "C00000"        # "We are here" marker, WIP badges, urgent callout
SUCCESS_GREEN = "1E7E34"    # Positive trend
WARNING_AMBER = "D97706"    # Neutral/alert trend

# Spektrum series colors for charts
SERIES_COLORS = [NAVY_PRIMARY, BLUE_ACCENT, "0070C0", "2B7FFF", ALERT_RED, SUCCESS_GREEN, WARNING_AMBER, "4B5563"]

# Fonts: Universal Arial family
FONT_FAMILY = "Arial"

# Slide dimensions (16:9 widescreen)
SLIDE_W = 13.333
SLIDE_H = 7.500

# Content canvas boundaries
CONTENT_X = 0.40
CONTENT_Y = 1.15
CONTENT_W = 12.53
CONTENT_BOTTOM = 6.90

LABELS = {
    "id": {
        "toc": "Daftar Isi",
        "takeaway": "Insight Utama",
        "key_insights": "Key insights",
        "effort": "EFFORT",
        "impact": "IMPACT",
        "high": "High",
        "low": "Low",
        "quick_wins": "Quick wins",
        "major_projects": "Major projects",
        "low_hanging": "Low-hanging fruit",
        "hygiene": "Hygiene factors",
        "data_needed": "DATA NEEDED",
        "outcome": "OUTCOME",
        "release": "RELEASE",
        "we_are_here": "We are here",
        "phase": "Phase",
        "activity": "Activity",
        "thanks": "Terima kasih",
        "step": "Tahap",
    },
    "en": {
        "toc": "Table of Contents",
        "takeaway": "Key Takeaway",
        "key_insights": "Key insights",
        "effort": "EFFORT",
        "impact": "IMPACT",
        "high": "High",
        "low": "Low",
        "quick_wins": "Quick wins",
        "major_projects": "Major projects",
        "low_hanging": "Low-hanging fruit",
        "hygiene": "Hygiene factors",
        "data_needed": "DATA NEEDED",
        "outcome": "OUTCOME",
        "release": "RELEASE",
        "we_are_here": "We are here",
        "phase": "Phase",
        "activity": "Activity",
        "thanks": "Thank You",
        "step": "Step",
    },
}


def rgb(hex_code: str) -> RGBColor:
    return RGBColor.from_string(hex_code.lstrip("#").upper())


class BUMNDeck:
    """Builder for BUMN Taskforce & DAMRI executive presentation decks."""

    def __init__(
        self,
        lang: str = "id",
        footer: str = "Taskforce BP BUMN · Perum DAMRI",
        template_path: Optional[str] = None,
    ):
        self.lang = lang if lang in LABELS else "id"
        self.footer_text = footer
        self.t = LABELS[self.lang]
        tpl = template_path or TEMPLATE_PATH
        if not os.path.exists(tpl):
            raise FileNotFoundError(f"Template not found at: {tpl}")
        self.prs = Presentation(tpl)
        # Clear existing slides if any
        slide_ids = self.prs.slides._sldIdLst
        for s in list(slide_ids):
            self.prs.part.drop_rel(s.rId)
            slide_ids.remove(s)
        self.slide_index = 0

    # --------------------------------------------------------------------------
    # Internal helpers
    # --------------------------------------------------------------------------
    def _add_slide(self, layout_idx: int = 5):
        """Adds a slide using Title Only layout (layout 5) or Blank (layout 6)."""
        idx = layout_idx if layout_idx < len(self.prs.slide_layouts) else 5
        layout = self.prs.slide_layouts[idx]
        slide = self.prs.slides.add_slide(layout)
        # Remove any unwanted default placeholders
        for ph in list(slide.placeholders):
            sp = ph._element
            sp.getparent().remove(sp)
        self.slide_index += 1
        return slide

    def _header(
        self,
        slide,
        title: str,
        section: Optional[str] = None,
        badge: Optional[str] = None,
        badge_color: str = ALERT_RED,
    ):
        """Draws standard section label and headline at official coordinates."""
        # 1. Section label (top)
        sec_text = section or "Transformation Plan - Dynamic Pricing"
        tb_sec = slide.shapes.add_textbox(Inches(0.40), Inches(0.25), Inches(8.33), Inches(0.24))
        tf_sec = tb_sec.text_frame
        tf_sec.word_wrap = True
        tf_sec.margin_left = tf_sec.margin_top = tf_sec.margin_right = tf_sec.margin_bottom = 0
        p_sec = tf_sec.paragraphs[0]
        p_sec.text = sec_text
        p_sec.font.name = FONT_FAMILY
        p_sec.font.size = Pt(13)
        p_sec.font.bold = True
        p_sec.font.color.rgb = rgb(BLUE_ACCENT)

        # 2. Headline (main title)
        tb_title = slide.shapes.add_textbox(Inches(0.40), Inches(0.46), Inches(10.50), Inches(0.48))
        tf_title = tb_title.text_frame
        tf_title.word_wrap = True
        tf_title.margin_left = tf_title.margin_top = tf_title.margin_right = tf_title.margin_bottom = 0
        p_title = tf_title.paragraphs[0]
        p_title.text = title
        p_title.font.name = FONT_FAMILY
        p_title.font.size = Pt(18)
        p_title.font.bold = True
        p_title.font.color.rgb = rgb(NAVY_HEADLINE)

        # 3. Optional status / WIP / release badge (top right)
        if badge:
            bx = Inches(11.20)
            by = Inches(0.28)
            bw = Inches(1.70)
            bh = Inches(0.34)
            sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, bx, by, bw, bh)
            sh.fill.solid()
            sh.fill.fore_color.rgb = rgb(badge_color)
            sh.line.color.rgb = rgb(badge_color)
            tf = sh.text_frame
            tf.word_wrap = False
            tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
            p = tf.paragraphs[0]
            p.text = badge
            p.alignment = PP_ALIGN.CENTER
            p.font.name = FONT_FAMILY
            p.font.size = Pt(10)
            p.font.bold = True
            p.font.color.rgb = rgb(WHITE)

    def _footer(self, slide):
        """Draws page number and footer note."""
        tb_num = slide.shapes.add_textbox(Inches(12.50), Inches(7.18), Inches(0.70), Inches(0.22))
        tf_num = tb_num.text_frame
        tf_num.margin_left = tf_num.margin_top = tf_num.margin_right = tf_num.margin_bottom = 0
        p_num = tf_num.paragraphs[0]
        p_num.text = str(self.slide_index)
        p_num.alignment = PP_ALIGN.RIGHT
        p_num.font.name = FONT_FAMILY
        p_num.font.size = Pt(9.5)
        p_num.font.color.rgb = rgb(MUTED_TEXT)

    def _takeaway(self, slide, text: Optional[str]):
        """Draws bottom key takeaway banner."""
        if not text:
            return
        y = Inches(6.50)
        h = Inches(0.40)
        x = Inches(0.40)
        w = Inches(12.53)

        bg = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
        bg.fill.solid()
        bg.fill.fore_color.rgb = rgb(CARD_ACTIVE)
        bg.line.color.rgb = rgb(BORDER_BLUE)
        bg.line.width = Pt(0.75)

        # Accent bar on left edge
        bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, Inches(0.08), h)
        bar.fill.solid()
        bar.fill.fore_color.rgb = rgb(NAVY_PRIMARY)
        bar.line.fill.background()

        tb = slide.shapes.add_textbox(x + Inches(0.18), y + Inches(0.05), w - Inches(0.30), h - Inches(0.10))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT

        prefix = f"{self.t['takeaway']}: "
        r1 = p.add_run()
        r1.text = prefix
        r1.font.name = FONT_FAMILY
        r1.font.size = Pt(10)
        r1.font.bold = True
        r1.font.color.rgb = rgb(NAVY_PRIMARY)

        self._format_markdown_runs(p, text, font_size=Pt(10), font_color=rgb(BODY_DARK))

    def _format_markdown_runs(self, paragraph, text: str, font_size=Pt(10), font_color=None, base_bold=False):
        """Parses **bold text** syntax and appends formatted runs."""
        col = font_color or rgb(BODY_DARK)
        tokens = re.split(r"(\*\*.*?\*\*)", text)
        for token in tokens:
            if not token:
                continue
            r = paragraph.add_run()
            r.font.name = FONT_FAMILY
            r.font.size = font_size
            if token.startswith("**") and token.endswith("**") and len(token) >= 4:
                r.text = token[2:-2]
                r.font.bold = True
                r.font.color.rgb = col
            else:
                r.text = token
                r.font.bold = base_bold
                r.font.color.rgb = col

    # --------------------------------------------------------------------------
    # Public Slide Methods
    # --------------------------------------------------------------------------
    def cover(
        self,
        title: str = "Taskforce BP BUMN - Perum Damri",
        subtitle: str = "AI for Public Services: DAMRI's Transformation Roadmap for Public Transportation",
        meta: Optional[str] = None,
        partner_logo: Optional[str] = None,
    ):
        """Slide 1: Official BUMN / DAMRI Cover Slide."""
        slide = self._add_slide(0)

        # 1. Background image
        bg_img = os.path.join(BACKGROUNDS_DIR, "cover-damri.jpg")
        if os.path.exists(bg_img):
            slide.shapes.add_picture(bg_img, Inches(0), Inches(0), Inches(SLIDE_W), Inches(SLIDE_H))
        else:
            bg_rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(SLIDE_W), Inches(SLIDE_H))
            bg_rect.fill.solid()
            bg_rect.fill.fore_color.rgb = rgb(NAVY_PRIMARY)

        # 2. Logos (top-right)
        bumn_logo = os.path.join(LOGOS_DIR, "bumn-white.png")
        damri_logo = os.path.join(LOGOS_DIR, "damri-white.png")

        if os.path.exists(bumn_logo):
            slide.shapes.add_picture(bumn_logo, Inches(9.80), Inches(0.26), Inches(1.20), Inches(0.58))
        if os.path.exists(damri_logo):
            slide.shapes.add_picture(damri_logo, Inches(11.20), Inches(0.38), Inches(1.80), Inches(0.32))

        if partner_logo:
            p_logo_path = os.path.join(LOGOS_DIR, f"{partner_logo}.png")
            if os.path.exists(p_logo_path):
                slide.shapes.add_picture(p_logo_path, Inches(8.30), Inches(0.30), width=Inches(1.20))

        # 3. Bottom Navy Banner (Rectangle 5)
        banner_y = Inches(5.80)
        banner_h = Inches(1.35)
        banner = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), banner_y, Inches(SLIDE_W), banner_h)
        banner.fill.solid()
        banner.fill.fore_color.rgb = rgb(NAVY_PRIMARY)
        banner.line.fill.background()

        # 4. Title in banner
        tb_title = slide.shapes.add_textbox(Inches(0.40), banner_y + Inches(0.15), Inches(12.50), Inches(0.45))
        tf_title = tb_title.text_frame
        tf_title.word_wrap = True
        tf_title.margin_left = tf_title.margin_top = tf_title.margin_right = tf_title.margin_bottom = 0
        p1 = tf_title.paragraphs[0]
        p1.text = title
        p1.font.name = FONT_FAMILY
        p1.font.size = Pt(24)
        p1.font.bold = True
        p1.font.color.rgb = rgb(WHITE)

        # 5. Subtitle in banner
        sub_full = subtitle
        if meta:
            sub_full = f"{subtitle}  |  {meta}"
        tb_sub = slide.shapes.add_textbox(Inches(0.40), banner_y + Inches(0.68), Inches(12.50), Inches(0.40))
        tf_sub = tb_sub.text_frame
        tf_sub.word_wrap = True
        tf_sub.margin_left = tf_sub.margin_top = tf_sub.margin_right = tf_sub.margin_bottom = 0
        p2 = tf_sub.paragraphs[0]
        p2.text = sub_full
        p2.font.name = FONT_FAMILY
        p2.font.size = Pt(13)
        p2.font.color.rgb = rgb(CARD_ACTIVE)

        return slide

    def matrix_2x2(
        self,
        quadrants: Dict[str, Dict[str, Any]],
        insights: Optional[Sequence[Dict[str, str]]] = None,
        title: str = "Prioritization Framework (Impact vs Effort)",
        section: Optional[str] = "Transformation Plan - Dynamic Pricing",
        x_label: str = "EFFORT",
        y_label: str = "IMPACT",
        takeaway: Optional[str] = None,
        badge: Optional[str] = None,
    ):
        """Slide 2 Archetype: 2x2 Matrix with 4 Quadrants & Key Insights Sidebar."""
        slide = self._add_slide(5)
        self._header(slide, title, section, badge)

        # Sub-header above matrix
        tb_mh = slide.shapes.add_textbox(Inches(0.40), Inches(1.15), Inches(8.40), Inches(0.24))
        tf_mh = tb_mh.text_frame
        tf_mh.margin_left = tf_mh.margin_top = tf_mh.margin_right = tf_mh.margin_bottom = 0
        pmh = tf_mh.paragraphs[0]
        pmh.text = "Use-case prioritization by impact and effort"
        pmh.font.name = FONT_FAMILY
        pmh.font.size = Pt(12)
        pmh.font.bold = True
        pmh.font.color.rgb = rgb(NAVY_PRIMARY)

        # Sub-header line
        sh_line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.40), Inches(1.42), Inches(8.40), Inches(0.015))
        sh_line.fill.solid()
        sh_line.fill.fore_color.rgb = rgb(BORDER_BLUE)
        sh_line.line.fill.background()

        # Matrix geometry
        mx = Inches(0.90)
        my = Inches(1.65)
        qw = Inches(3.85)
        qh_top = Inches(2.25)
        qh_bot = Inches(2.20)
        gap = Inches(0.18)

        # Axes titles
        tb_y = slide.shapes.add_textbox(Inches(0.25), my + Inches(1.80), Inches(0.60), Inches(0.30))
        tb_y.rotation = 270
        tb_y.text_frame.paragraphs[0].text = y_label
        tb_y.text_frame.paragraphs[0].font.size = Pt(9.5)
        tb_y.text_frame.paragraphs[0].font.bold = True
        tb_y.text_frame.paragraphs[0].font.color.rgb = rgb(NAVY_PRIMARY)

        tb_x = slide.shapes.add_textbox(mx + Inches(3.20), my + qh_top + qh_bot + gap + Inches(0.05), Inches(1.50), Inches(0.25))
        tb_x.text_frame.paragraphs[0].text = x_label
        tb_x.text_frame.paragraphs[0].font.size = Pt(9.5)
        tb_x.text_frame.paragraphs[0].font.bold = True
        tb_x.text_frame.paragraphs[0].font.color.rgb = rgb(NAVY_PRIMARY)

        # Quadrant config: A (top-left), B (top-right), C (bot-left), D (bot-right)
        coords = {
            "A": (mx, my, qw, qh_top, True),
            "B": (mx + qw + gap, my, qw, qh_top, False),
            "C": (mx, my + qh_top + gap, qw, qh_bot, False),
            "D": (mx + qw + gap, my + qh_top + gap, qw, qh_bot, False),
        }

        for q_key, (qx, qy, w, h, is_active) in coords.items():
            data = quadrants.get(q_key, {})
            title_q = data.get("title", f"Quadrant {q_key}")
            sub_q = data.get("subtitle", "")
            count_q = str(data.get("count", len(data.get("items", []))))
            items_q = data.get("items", [])

            # Card box
            box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, qx, qy, w, h)
            box.fill.solid()
            box.fill.fore_color.rgb = rgb(CARD_ACTIVE if is_active else CARD_NEUTRAL)
            box.line.color.rgb = rgb(BORDER_BLUE if is_active else BORDER_LIGHT)
            box.line.width = Pt(1 if is_active else 0.5)

            if is_active:
                # Accent bar on left edge
                act_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, qx, qy, Inches(0.06), h)
                act_bar.fill.solid()
                act_bar.fill.fore_color.rgb = rgb(NAVY_PRIMARY)
                act_bar.line.fill.background()

            # Badge pill
            badge_sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, qx + Inches(0.16), qy + Inches(0.14), Inches(0.26), Inches(0.26))
            badge_sh.fill.solid()
            badge_sh.fill.fore_color.rgb = rgb(NAVY_PRIMARY if is_active else MUTED_TEXT)
            badge_sh.line.fill.background()
            p_b = badge_sh.text_frame.paragraphs[0]
            p_b.text = q_key
            p_b.alignment = PP_ALIGN.CENTER
            p_b.font.size = Pt(10)
            p_b.font.bold = True
            p_b.font.color.rgb = rgb(WHITE)

            # Title & Sub
            tb_t = slide.shapes.add_textbox(qx + Inches(0.48), qy + Inches(0.10), Inches(2.50), Inches(0.40))
            tf_t = tb_t.text_frame
            tf_t.margin_left = tf_t.margin_top = tf_t.margin_right = tf_t.margin_bottom = 0
            pt = tf_t.paragraphs[0]
            pt.text = title_q
            pt.font.size = Pt(11.5)
            pt.font.bold = True
            pt.font.color.rgb = rgb(NAVY_PRIMARY)

            if sub_q:
                ps = tf_t.add_paragraph()
                ps.text = sub_q
                ps.font.size = Pt(8.5)
                ps.font.color.rgb = rgb(MUTED_TEXT)

            # Count
            tb_c = slide.shapes.add_textbox(qx + w - Inches(0.60), qy + Inches(0.10), Inches(0.50), Inches(0.35))
            tb_c.text_frame.margin_left = tb_c.text_frame.margin_top = tb_c.text_frame.margin_right = tb_c.text_frame.margin_bottom = 0
            pc = tb_c.text_frame.paragraphs[0]
            pc.text = count_q
            pc.alignment = PP_ALIGN.RIGHT
            pc.font.size = Pt(16)
            pc.font.bold = True
            pc.font.color.rgb = rgb(NAVY_PRIMARY if is_active else MUTED_TEXT)

            # Divider line inside card
            q_div = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, qx + Inches(0.16), qy + Inches(0.55), w - Inches(0.32), Inches(0.01))
            q_div.fill.solid()
            q_div.fill.fore_color.rgb = rgb(BORDER_BLUE if is_active else BORDER_LIGHT)
            q_div.line.fill.background()

            # Items list
            tb_l = slide.shapes.add_textbox(qx + Inches(0.16), qy + Inches(0.62), w - Inches(0.32), h - Inches(0.68))
            tf_l = tb_l.text_frame
            tf_l.word_wrap = True
            tf_l.margin_left = tf_l.margin_top = tf_l.margin_right = tf_l.margin_bottom = 0
            for i, itm in enumerate(items_q[:6]):
                p_i = tf_l.paragraphs[0] if i == 0 else tf_l.add_paragraph()
                num_str = f"{i+1:02d}  "
                r_n = p_i.add_run()
                r_n.text = num_str
                r_n.font.size = Pt(9.5)
                r_n.font.bold = True
                r_n.font.color.rgb = rgb(NAVY_PRIMARY if is_active else MUTED_TEXT)
                self._format_markdown_runs(p_i, str(itm), font_size=Pt(9.5), font_color=rgb(BODY_DARK))

        # Right-side Insights Sidebar
        sx = Inches(9.10)
        sy = Inches(1.15)
        sw = Inches(3.83)
        sh_h = Inches(5.15)

        sb = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, sx, sy, sw, sh_h)
        sb.fill.solid()
        sb.fill.fore_color.rgb = rgb(CARD_NEUTRAL)
        sb.line.color.rgb = rgb(BORDER_LIGHT)

        # Header in sidebar
        tb_sh = slide.shapes.add_textbox(sx + Inches(0.20), sy + Inches(0.15), sw - Inches(0.40), Inches(0.25))
        tb_sh.text_frame.margin_left = tb_sh.text_frame.margin_top = tb_sh.text_frame.margin_right = tb_sh.text_frame.margin_bottom = 0
        psh = tb_sh.text_frame.paragraphs[0]
        psh.text = self.t["key_insights"]
        psh.font.size = Pt(12)
        psh.font.bold = True
        psh.font.color.rgb = rgb(NAVY_PRIMARY)

        # Line under sidebar header
        s_line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, sx + Inches(0.20), sy + Inches(0.42), sw - Inches(0.40), Inches(0.015))
        s_line.fill.solid()
        s_line.fill.fore_color.rgb = rgb(BORDER_BLUE)
        s_line.line.fill.background()

        # Insight items
        ins_items = insights or [
            {"badge": "A", "title": "Quick wins", "text": "Improve occupancy and competitor visibility with existing data; no system integration required."},
            {"badge": "B", "title": "Major projects", "text": "The core revenue engines requiring automated pricing integrations and dynamic pricing models."},
            {"badge": "C", "title": "Low-hanging fruit", "text": "Customer segmentation & governance supporting the pricing rollouts."},
            {"badge": "D", "title": "Hygiene factors", "text": "Scenario simulator & ancillary optimization for long-term expansion."},
        ]

        curr_y = sy + Inches(0.55)
        for ins in ins_items:
            b_val = ins.get("badge", "•")
            t_val = ins.get("title", "")
            txt_val = ins.get("text", "")

            # Badge pill
            i_badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, sx + Inches(0.20), curr_y, Inches(0.24), Inches(0.24))
            i_badge.fill.solid()
            i_badge.fill.fore_color.rgb = rgb(NAVY_PRIMARY if b_val == "A" else MUTED_TEXT)
            i_badge.line.fill.background()
            p_ib = i_badge.text_frame.paragraphs[0]
            p_ib.text = b_val
            p_ib.alignment = PP_ALIGN.CENTER
            p_ib.font.size = Pt(8.5)
            p_ib.font.bold = True
            p_ib.font.color.rgb = rgb(WHITE)

            tb_it = slide.shapes.add_textbox(sx + Inches(0.50), curr_y - Inches(0.02), sw - Inches(0.70), Inches(0.85))
            tf_it = tb_it.text_frame
            tf_it.word_wrap = True
            tf_it.margin_left = tf_it.margin_top = tf_it.margin_right = tf_it.margin_bottom = 0
            pit = tf_it.paragraphs[0]
            pit.text = t_val
            pit.font.size = Pt(10)
            pit.font.bold = True
            pit.font.color.rgb = rgb(NAVY_PRIMARY)

            pit_desc = tf_it.add_paragraph()
            self._format_markdown_runs(pit_desc, txt_val, font_size=Pt(9), font_color=rgb(BODY_DARK))

            curr_y += Inches(1.10)

        self._takeaway(slide, takeaway)
        self._footer(slide)
        return slide

    def module_cards(
        self,
        cards: Sequence[Dict[str, Any]],
        preview_image: Optional[str] = None,
        preview_title: str = "05   Monitoring & alerting dashboard",
        preview_caption: str = "This dashboard covers analytical modules (use cases 01–04)",
        release_info: Optional[Dict[str, str]] = None,
        title: str = "Quick Wins & Output Release 1",
        section: Optional[str] = "Transformation Plan - Dynamic Pricing",
        takeaway: Optional[str] = None,
        badge: Optional[str] = None,
    ):
        """Slide 3 Archetype: 2x2 Analytical Module Cards with Release Bar & Dashboard Preview."""
        slide = self._add_slide(5)
        self._header(slide, title, section, badge)

        # Left header
        tb_lh = slide.shapes.add_textbox(Inches(0.40), Inches(1.05), Inches(7.10), Inches(0.24))
        tb_lh.text_frame.margin_left = tb_lh.text_frame.margin_top = tb_lh.text_frame.margin_right = tb_lh.text_frame.margin_bottom = 0
        plh = tb_lh.text_frame.paragraphs[0]
        plh.text = "01–04   Analytical modules"
        plh.font.size = Pt(12)
        plh.font.bold = True
        plh.font.color.rgb = rgb(NAVY_PRIMARY)

        # Right header
        tb_rh = slide.shapes.add_textbox(Inches(7.80), Inches(1.05), Inches(5.15), Inches(0.24))
        tb_rh.text_frame.margin_left = tb_rh.text_frame.margin_top = tb_rh.text_frame.margin_right = tb_rh.text_frame.margin_bottom = 0
        prh = tb_rh.text_frame.paragraphs[0]
        prh.text = preview_title
        prh.font.size = Pt(12)
        prh.font.bold = True
        prh.font.color.rgb = rgb(NAVY_PRIMARY)

        # 4 cards geometry (2 cols x 2 rows)
        c_w = Inches(3.50)
        c_h = Inches(2.20)
        x_col0 = Inches(0.40)
        x_col1 = Inches(4.05)
        y_row0 = Inches(1.40)
        y_row1 = Inches(3.75)

        card_coords = [
            (x_col0, y_row0),
            (x_col1, y_row0),
            (x_col0, y_row1),
            (x_col1, y_row1),
        ]

        for i, card in enumerate(cards[:4]):
            cx, cy = card_coords[i]
            c_num = card.get("num", f"{i+1:02d}")
            c_title = card.get("title", f"Module {c_num}")
            c_obj = card.get("objective", "")
            c_data = card.get("data", "")
            c_outcome = card.get("outcome", "")

            # Card background
            cb = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, cy, c_w, c_h)
            cb.fill.solid()
            cb.fill.fore_color.rgb = rgb(CARD_NEUTRAL)
            cb.line.color.rgb = rgb(BORDER_LIGHT)

            # Left accent pill
            bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, cx, cy, Inches(0.04), c_h)
            bar.fill.solid()
            bar.fill.fore_color.rgb = rgb(BLUE_ACCENT)
            bar.line.fill.background()

            # Number badge
            nb = slide.shapes.add_textbox(cx + Inches(0.12), cy + Inches(0.08), Inches(0.35), Inches(0.25))
            nb.text_frame.margin_left = nb.text_frame.margin_top = nb.text_frame.margin_right = nb.text_frame.margin_bottom = 0
            p_nb = nb.text_frame.paragraphs[0]
            p_nb.text = str(c_num)
            p_nb.font.size = Pt(11)
            p_nb.font.bold = True
            p_nb.font.color.rgb = rgb(NAVY_PRIMARY)

            # Title
            tb_ct = slide.shapes.add_textbox(cx + Inches(0.45), cy + Inches(0.08), c_w - Inches(0.55), Inches(0.25))
            tb_ct.text_frame.margin_left = tb_ct.text_frame.margin_top = tb_ct.text_frame.margin_right = tb_ct.text_frame.margin_bottom = 0
            pct = tb_ct.text_frame.paragraphs[0]
            pct.text = c_title
            pct.font.size = Pt(11)
            pct.font.bold = True
            pct.font.color.rgb = rgb(NAVY_PRIMARY)

            # Objective
            tb_co = slide.shapes.add_textbox(cx + Inches(0.12), cy + Inches(0.36), c_w - Inches(0.24), Inches(0.36))
            tb_co.text_frame.word_wrap = True
            tb_co.text_frame.margin_left = tb_co.text_frame.margin_top = tb_co.text_frame.margin_right = tb_co.text_frame.margin_bottom = 0
            pco = tb_co.text_frame.paragraphs[0]
            self._format_markdown_runs(pco, c_obj, font_size=Pt(9), font_color=rgb(BODY_DARK))

            # DATA NEEDED label & text
            tb_dl = slide.shapes.add_textbox(cx + Inches(0.12), cy + Inches(0.78), c_w - Inches(0.24), Inches(0.60))
            tf_dl = tb_dl.text_frame
            tf_dl.word_wrap = True
            tf_dl.margin_left = tf_dl.margin_top = tf_dl.margin_right = tf_dl.margin_bottom = 0
            pdl = tf_dl.paragraphs[0]
            pdl.text = self.t["data_needed"]
            pdl.font.size = Pt(8)
            pdl.font.bold = True
            pdl.font.color.rgb = rgb(MUTED_TEXT)
            pdt = tf_dl.add_paragraph()
            self._format_markdown_runs(pdt, c_data, font_size=Pt(8.5), font_color=rgb(BODY_DARK))

            # OUTCOME label & text
            tb_ol = slide.shapes.add_textbox(cx + Inches(0.12), cy + Inches(1.45), c_w - Inches(0.24), Inches(0.65))
            tf_ol = tb_ol.text_frame
            tf_ol.word_wrap = True
            tf_ol.margin_left = tf_ol.margin_top = tf_ol.margin_right = tf_ol.margin_bottom = 0
            pol = tf_ol.paragraphs[0]
            pol.text = self.t["outcome"]
            pol.font.size = Pt(8)
            pol.font.bold = True
            pol.font.color.rgb = rgb(MUTED_TEXT)
            pot = tf_ol.add_paragraph()
            self._format_markdown_runs(pot, c_outcome, font_size=Pt(8.5), font_color=rgb(BODY_DARK))

        # Bottom Release Bar
        rel_info = release_info or {
            "release": "RELEASE 1",
            "beta": "Beta   W2 Nov 26",
            "prototype": "Prototype   W2 Dec 26",
        }
        rx = Inches(0.40)
        ry = Inches(6.12)
        rw = Inches(7.15)
        rh = Inches(0.40)

        rb_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, rx, ry, rw, rh)
        rb_box.fill.solid()
        rb_box.fill.fore_color.rgb = rgb(CARD_ACTIVE)
        rb_box.line.color.rgb = rgb(BORDER_BLUE)

        # Release tag pill
        rt_pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, rx + Inches(0.10), ry + Inches(0.06), Inches(1.40), Inches(0.28))
        rt_pill.fill.solid()
        rt_pill.fill.fore_color.rgb = rgb(NAVY_PRIMARY)
        rt_pill.line.fill.background()
        p_rp = rt_pill.text_frame.paragraphs[0]
        p_rp.text = rel_info.get("release", "RELEASE 1")
        p_rp.alignment = PP_ALIGN.CENTER
        p_rp.font.size = Pt(9)
        p_rp.font.bold = True
        p_rp.font.color.rgb = rgb(WHITE)

        # Milestones text
        tb_rm = slide.shapes.add_textbox(rx + Inches(1.65), ry + Inches(0.08), rw - Inches(1.75), Inches(0.26))
        tf_rm = tb_rm.text_frame
        tf_rm.margin_left = tf_rm.margin_top = tf_rm.margin_right = tf_rm.margin_bottom = 0
        p_rm = tf_rm.paragraphs[0]
        p_rm.text = f"{rel_info.get('beta', '')}     •     {rel_info.get('prototype', '')}"
        p_rm.font.size = Pt(9.5)
        p_rm.font.bold = True
        p_rm.font.color.rgb = rgb(NAVY_PRIMARY)

        # Right Preview Box / Screenshot Placeholder
        px = Inches(7.80)
        py = Inches(1.40)
        pw = Inches(5.13)
        ph = Inches(5.12)

        p_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, px, py, pw, ph)
        p_box.fill.solid()
        p_box.fill.fore_color.rgb = rgb(CARD_NEUTRAL)
        p_box.line.color.rgb = rgb(BORDER_LIGHT)

        if preview_image and os.path.exists(preview_image):
            slide.shapes.add_picture(preview_image, px + Inches(0.15), py + Inches(0.15), width=pw - Inches(0.30))
        else:
            # Subtle dashboard placeholder artwork
            tb_ph = slide.shapes.add_textbox(px + Inches(0.50), py + Inches(1.80), pw - Inches(1.00), Inches(1.20))
            tf_ph = tb_ph.text_frame
            tf_ph.word_wrap = True
            pph = tf_ph.paragraphs[0]
            pph.alignment = PP_ALIGN.CENTER
            pph.text = f"[ {preview_caption} ]"
            pph.font.size = Pt(11)
            pph.font.color.rgb = rgb(MUTED_TEXT)

        self._takeaway(slide, takeaway)
        self._footer(slide)
        return slide

    def architecture_flow(
        self,
        stages: Sequence[Dict[str, Any]],
        title: str = "Implementation Architecture & Data Flow",
        section: Optional[str] = "Transformation Plan - Dynamic Pricing",
        takeaway: Optional[str] = None,
        badge: Optional[str] = None,
    ):
        """Slide 4 Archetype: End-to-End Pipeline (Data Sources -> AI/ML Engine -> Deployment)."""
        slide = self._add_slide(5)
        self._header(slide, title, section, badge)

        n = len(stages)
        if n == 0:
            return slide

        gap = Inches(0.25)
        avail_w = CONTENT_W - (gap * (n - 1))
        col_w = avail_w / n
        col_h = Inches(4.80)
        y = CONTENT_Y + Inches(0.20)

        for i, stg in enumerate(stages):
            x = CONTENT_X + i * (col_w + gap)
            stg_num = str(stg.get("num", f"{i+1}"))
            stg_name = stg.get("title", f"Stage {i+1}")
            stg_cat = stg.get("category", "")
            items = stg.get("items", [])

            # Stage card
            sc = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, col_w, col_h)
            sc.fill.solid()
            sc.fill.fore_color.rgb = rgb(CARD_ACTIVE if i == 1 else CARD_NEUTRAL)
            sc.line.color.rgb = rgb(BORDER_BLUE if i == 1 else BORDER_LIGHT)

            # Stage number circle
            circ = slide.shapes.add_shape(MSO_SHAPE.OVAL, x + Inches(0.15), y + Inches(0.15), Inches(0.36), Inches(0.36))
            circ.fill.solid()
            circ.fill.fore_color.rgb = rgb(NAVY_PRIMARY)
            circ.line.fill.background()
            p_circ = circ.text_frame.paragraphs[0]
            p_circ.text = stg_num
            p_circ.alignment = PP_ALIGN.CENTER
            p_circ.font.size = Pt(11)
            p_circ.font.bold = True
            p_circ.font.color.rgb = rgb(WHITE)

            # Stage title
            tb_st = slide.shapes.add_textbox(x + Inches(0.60), y + Inches(0.12), col_w - Inches(0.70), Inches(0.40))
            tf_st = tb_st.text_frame
            tf_st.word_wrap = True
            tf_st.margin_left = tf_st.margin_top = tf_st.margin_right = tf_st.margin_bottom = 0
            pst = tf_st.paragraphs[0]
            pst.text = stg_name
            pst.font.size = Pt(11.5)
            pst.font.bold = True
            pst.font.color.rgb = rgb(NAVY_PRIMARY)

            if stg_cat:
                pst_sub = tf_st.add_paragraph()
                pst_sub.text = stg_cat
                pst_sub.font.size = Pt(9)
                pst_sub.font.color.rgb = rgb(MUTED_TEXT)

            # Line under stage header
            s_div = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x + Inches(0.15), y + Inches(0.62), col_w - Inches(0.30), Inches(0.015))
            s_div.fill.solid()
            s_div.fill.fore_color.rgb = rgb(BORDER_BLUE if i == 1 else BORDER_LIGHT)
            s_div.line.fill.background()

            # Items inside stage
            tb_items = slide.shapes.add_textbox(x + Inches(0.15), y + Inches(0.75), col_w - Inches(0.30), col_h - Inches(0.90))
            tf_items = tb_items.text_frame
            tf_items.word_wrap = True
            tf_items.margin_left = tf_items.margin_top = tf_items.margin_right = tf_items.margin_bottom = 0

            for j, itm in enumerate(items):
                p = tf_items.paragraphs[0] if j == 0 else tf_items.add_paragraph()
                p.space_after = Pt(6)
                if isinstance(itm, dict):
                    h_text = itm.get("heading", "")
                    b_text = itm.get("desc", "")
                    rh = p.add_run()
                    rh.text = f"•  {h_text}\n"
                    rh.font.bold = True
                    rh.font.size = Pt(9.5)
                    rh.font.color.rgb = rgb(NAVY_PRIMARY)
                    self._format_markdown_runs(p, f"   {b_text}", font_size=Pt(9), font_color=rgb(BODY_DARK))
                else:
                    r = p.add_run()
                    r.text = "•  "
                    r.font.bold = True
                    r.font.size = Pt(9.5)
                    r.font.color.rgb = rgb(BLUE_ACCENT)
                    self._format_markdown_runs(p, str(itm), font_size=Pt(9.5), font_color=rgb(BODY_DARK))

            # Connector arrow to next stage
            if i < n - 1:
                arr_x = x + col_w + Inches(0.05)
                arr_y = y + col_h / 2 - Inches(0.15)
                arr = slide.shapes.add_shape(MSO_SHAPE.OVAL, arr_x, arr_y, Inches(0.22), Inches(0.22))
                arr.fill.solid()
                arr.fill.fore_color.rgb = rgb(BLUE_ACCENT)
                arr.line.fill.background()
                p_arr = arr.text_frame.paragraphs[0]
                p_arr.text = ">"
                p_arr.alignment = PP_ALIGN.CENTER
                p_arr.font.size = Pt(9)
                p_arr.font.bold = True
                p_arr.font.color.rgb = rgb(WHITE)

        self._takeaway(slide, takeaway)
        self._footer(slide)
        return slide

    def timeline_gantt(
        self,
        phases: Sequence[Dict[str, Any]],
        tasks: Sequence[Dict[str, Any]],
        milestones: Optional[Sequence[Dict[str, Any]]] = None,
        current_point: Optional[Dict[str, Any]] = None,
        title: str = "Timeline delivery: Prototype launch in Dec W2",
        section: Optional[str] = "Transformation Plan - Dynamic Pricing",
        note: Optional[str] = "Beta is followed by a review before the Dec W2 prototype launch",
        takeaway: Optional[str] = None,
        badge: Optional[str] = None,
    ):
        """Slide 6 Archetype: Delivery Timeline / Gantt Roadmap with Phase Bands & Milestone Pin."""
        slide = self._add_slide(5)
        self._header(slide, title, section, badge)

        # Header for columns
        x_phase = Inches(0.56)
        x_task = Inches(1.65)
        w_phase = Inches(1.00)
        w_task = Inches(3.20)
        x_grid = Inches(5.00)
        w_grid = Inches(7.70)
        y_top = Inches(1.50)

        tb_hp = slide.shapes.add_textbox(x_phase, y_top - Inches(0.30), w_phase, Inches(0.25))
        tb_hp.text_frame.margin_left = tb_hp.text_frame.margin_top = tb_hp.text_frame.margin_right = tb_hp.text_frame.margin_bottom = 0
        php = tb_hp.text_frame.paragraphs[0]
        php.text = self.t["phase"]
        php.font.size = Pt(11)
        php.font.bold = True
        php.font.color.rgb = rgb(BLUE_SLATE)

        tb_ht = slide.shapes.add_textbox(x_task, y_top - Inches(0.30), w_task, Inches(0.25))
        tb_ht.text_frame.margin_left = tb_ht.text_frame.margin_top = tb_ht.text_frame.margin_right = tb_ht.text_frame.margin_bottom = 0
        pht = tb_ht.text_frame.paragraphs[0]
        pht.text = self.t["activity"]
        pht.font.size = Pt(11)
        pht.font.bold = True
        pht.font.color.rgb = rgb(BLUE_SLATE)

        # Draw tasks & rows
        num_tasks = len(tasks)
        row_h = Inches(4.50) / max(num_tasks, 1)

        for i, tsk in enumerate(tasks):
            ry = y_top + i * row_h
            t_name = tsk.get("name", f"Task {i+1}")
            start_frac = float(tsk.get("start", 0.0))  # 0.0 to 1.0
            dur_frac = float(tsk.get("dur", 0.25))

            # Task text
            tb_t = slide.shapes.add_textbox(x_task, ry + Inches(0.05), w_task, row_h - Inches(0.10))
            tf_t = tb_t.text_frame
            tf_t.word_wrap = True
            tf_t.margin_left = tf_t.margin_top = tf_t.margin_right = tf_t.margin_bottom = 0
            pt = tf_t.paragraphs[0]
            pt.text = t_name
            pt.font.size = Pt(10)
            pt.font.color.rgb = rgb(DARK_SLATE)

            # Row divider line
            r_sep = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x_task, ry + row_h, w_task + w_grid + Inches(0.20), Inches(0.01))
            r_sep.fill.solid()
            r_sep.fill.fore_color.rgb = rgb(BORDER_LIGHT if i % 2 == 0 else BORDER_MID)
            r_sep.line.fill.background()

            # Schedule bar in grid
            bar_x = x_grid + start_frac * w_grid
            bar_w = dur_frac * w_grid
            bar_h = row_h * 0.45
            bar_y = ry + (row_h - bar_h) / 2

            s_bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, bar_x, bar_y, bar_w, bar_h)
            s_bar.fill.solid()
            s_bar.fill.fore_color.rgb = rgb(BLUE_ACCENT if i % 2 == 0 else NAVY_PRIMARY)
            s_bar.line.fill.background()

        # Draw Phase Bands on left
        curr_py = y_top
        for p_idx, ph in enumerate(phases):
            p_name = ph.get("name", f"Phase {p_idx+1}")
            p_span = int(ph.get("span", 2))  # number of rows spanned
            band_h = p_span * row_h - Inches(0.04)

            band = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x_phase, curr_py + Inches(0.02), w_phase, band_h)
            band.fill.solid()
            band.fill.fore_color.rgb = rgb(PHASE_BAND_1 if p_idx % 2 == 0 else PHASE_BAND_2)
            band.line.fill.background()

            tf_b = band.text_frame
            tf_b.word_wrap = True
            tf_b.margin_left = tf_b.margin_right = Inches(0.08)
            pb = tf_b.paragraphs[0]
            pb.text = p_name
            pb.alignment = PP_ALIGN.CENTER
            pb.font.size = Pt(9.5)
            pb.font.bold = True
            pb.font.color.rgb = rgb(BLUE_SLATE)

            curr_py += p_span * row_h

        # "We are here" red milestone pin
        cp = current_point or {"frac": 0.70, "label": "We are here"}
        c_frac = float(cp.get("frac", 0.70))
        c_label = cp.get("label", self.t["we_are_here"])
        pin_x = x_grid + c_frac * w_grid

        # Vertical dotted/solid pin line
        v_line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, pin_x, y_top - Inches(0.10), Inches(0.02), Inches(4.70))
        v_line.fill.solid()
        v_line.fill.fore_color.rgb = rgb(ALERT_RED)
        v_line.line.fill.background()

        # Pin label at bottom
        pin_pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, pin_x - Inches(0.60), y_top + Inches(4.65), Inches(1.25), Inches(0.26))
        pin_pill.fill.solid()
        pin_pill.fill.fore_color.rgb = rgb(ALERT_RED)
        pin_pill.line.fill.background()
        p_pin = pin_pill.text_frame.paragraphs[0]
        p_pin.text = c_label
        p_pin.alignment = PP_ALIGN.CENTER
        p_pin.font.size = Pt(9)
        p_pin.font.bold = True
        p_pin.font.color.rgb = rgb(WHITE)

        # Note at bottom left
        if note:
            tb_note = slide.shapes.add_textbox(x_phase, y_top + Inches(4.65), Inches(6.00), Inches(0.25))
            tb_note.text_frame.margin_left = tb_note.text_frame.margin_top = tb_note.text_frame.margin_right = tb_note.text_frame.margin_bottom = 0
            p_note = tb_note.text_frame.paragraphs[0]
            p_note.text = note
            p_note.font.size = Pt(9.5)
            p_note.font.color.rgb = rgb(MUTED_GRAY)

        self._takeaway(slide, takeaway)
        self._footer(slide)
        return slide

    def kpis(
        self,
        items: Sequence[Dict[str, Any]],
        title: str = "Key Performance Indicators",
        section: Optional[str] = "Performance Review",
        takeaway: Optional[str] = None,
        badge: Optional[str] = None,
    ):
        """2–4 High-Impact KPI Tiles."""
        slide = self._add_slide(5)
        self._header(slide, title, section, badge)

        n = min(len(items), 4)
        if n == 0:
            return slide

        gap = Inches(0.25)
        card_w = (CONTENT_W - (gap * (n - 1))) / n
        card_h = Inches(4.50)
        y = CONTENT_Y + Inches(0.35)

        for i, itm in enumerate(items[:n]):
            x = CONTENT_X + i * (card_w + gap)
            val = str(itm.get("value", "0"))
            unit = itm.get("unit", "")
            lbl = itm.get("label", f"Metric {i+1}")
            trend = itm.get("trend", "")
            trend_type = itm.get("trend_type", "up")  # up, down, neutral
            desc = itm.get("desc", "")

            # Card fill
            cb = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, card_w, card_h)
            cb.fill.solid()
            cb.fill.fore_color.rgb = rgb(CARD_ACTIVE if i == 0 else CARD_NEUTRAL)
            cb.line.color.rgb = rgb(BORDER_BLUE if i == 0 else BORDER_LIGHT)

            # Top accent bar
            bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, card_w, Inches(0.06))
            bar.fill.solid()
            bar.fill.fore_color.rgb = rgb(NAVY_PRIMARY if i == 0 else BLUE_ACCENT)
            bar.line.fill.background()

            # Value & Unit
            tb_val = slide.shapes.add_textbox(x + Inches(0.20), y + Inches(0.40), card_w - Inches(0.40), Inches(0.80))
            tf_val = tb_val.text_frame
            tf_val.word_wrap = True
            tf_val.margin_left = tf_val.margin_top = tf_val.margin_right = tf_val.margin_bottom = 0
            p_val = tf_val.paragraphs[0]
            rv = p_val.add_run()
            rv.text = val
            rv.font.size = Pt(32)
            rv.font.bold = True
            rv.font.color.rgb = rgb(NAVY_PRIMARY)

            if unit:
                ru = p_val.add_run()
                ru.text = f" {unit}"
                ru.font.size = Pt(16)
                ru.font.bold = True
                ru.font.color.rgb = rgb(BLUE_ACCENT)

            # Trend pill
            if trend:
                tr_col = SUCCESS_GREEN if trend_type == "up" else (ALERT_RED if trend_type == "down" else WARNING_AMBER)
                tr_pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x + Inches(0.20), y + Inches(1.30), Inches(1.40), Inches(0.30))
                tr_pill.fill.solid()
                tr_pill.fill.fore_color.rgb = rgb(tr_col)
                tr_pill.line.fill.background()
                ptr = tr_pill.text_frame.paragraphs[0]
                ptr.text = trend
                ptr.alignment = PP_ALIGN.CENTER
                ptr.font.size = Pt(9.5)
                ptr.font.bold = True
                ptr.font.color.rgb = rgb(WHITE)

            # Metric Label
            tb_lbl = slide.shapes.add_textbox(x + Inches(0.20), y + Inches(1.85), card_w - Inches(0.40), Inches(0.50))
            tf_lbl = tb_lbl.text_frame
            tf_lbl.word_wrap = True
            tf_lbl.margin_left = tf_lbl.margin_top = tf_lbl.margin_right = tf_lbl.margin_bottom = 0
            pl = tf_lbl.paragraphs[0]
            pl.text = lbl
            pl.font.size = Pt(13)
            pl.font.bold = True
            pl.font.color.rgb = rgb(DARK_SLATE)

            # Divider
            div = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x + Inches(0.20), y + Inches(2.45), card_w - Inches(0.40), Inches(0.015))
            div.fill.solid()
            div.fill.fore_color.rgb = rgb(BORDER_BLUE if i == 0 else BORDER_LIGHT)
            div.line.fill.background()

            # Description
            if desc:
                tb_desc = slide.shapes.add_textbox(x + Inches(0.20), y + Inches(2.60), card_w - Inches(0.40), Inches(1.60))
                tf_desc = tb_desc.text_frame
                tf_desc.word_wrap = True
                tf_desc.margin_left = tf_desc.margin_top = tf_desc.margin_right = tf_desc.margin_bottom = 0
                pd = tf_desc.paragraphs[0]
                self._format_markdown_runs(pd, desc, font_size=Pt(10), font_color=rgb(BODY_DARK))

        self._takeaway(slide, takeaway)
        self._footer(slide)
        return slide

    def table(
        self,
        headers: Sequence[str],
        rows: Sequence[Sequence[Any]],
        title: str = "Data Summary",
        section: Optional[str] = "Summary",
        col_widths: Optional[Sequence[float]] = None,
        takeaway: Optional[str] = None,
        badge: Optional[str] = None,
    ):
        """Clean BUMN Corporate Data Table."""
        slide = self._add_slide(5)
        self._header(slide, title, section, badge)

        num_cols = len(headers)
        num_rows = len(rows) + 1
        x = CONTENT_X
        y = CONTENT_Y + Inches(0.25)
        w = CONTENT_W

        # Add native table shape
        tbl_shape = slide.shapes.add_table(num_rows, num_cols, x, y, w, Inches(4.50))
        table = tbl_shape.table

        # Column widths
        if col_widths and len(col_widths) == num_cols:
            tot = sum(col_widths)
            for c_i, cw in enumerate(col_widths):
                table.columns[c_i].width = int(Inches(w * (cw / tot)))
        else:
            eq_w = int(Inches(w / num_cols))
            for c_i in range(num_cols):
                table.columns[c_i].width = eq_w

        # Header Row
        for c_i, h in enumerate(headers):
            cell = table.cell(0, c_i)
            cell.fill.solid()
            cell.fill.fore_color.rgb = rgb(NAVY_PRIMARY)
            p = cell.text_frame.paragraphs[0]
            p.text = str(h)
            p.font.name = FONT_FAMILY
            p.font.size = Pt(10.5)
            p.font.bold = True
            p.font.color.rgb = rgb(WHITE)
            p.alignment = PP_ALIGN.LEFT

        # Data Rows
        for r_i, r_data in enumerate(rows):
            is_alt = (r_i % 2 == 1)
            for c_i, val in enumerate(r_data):
                if c_i >= num_cols:
                    break
                cell = table.cell(r_i + 1, c_i)
                cell.fill.solid()
                cell.fill.fore_color.rgb = rgb(CARD_NEUTRAL if is_alt else WHITE)
                p = cell.text_frame.paragraphs[0]
                self._format_markdown_runs(p, str(val), font_size=Pt(9.5), font_color=rgb(BODY_DARK), base_bold=(c_i == 0))

        self._takeaway(slide, takeaway)
        self._footer(slide)
        return slide

    def chart(
        self,
        kind: str,
        categories: Sequence[str],
        series: Sequence[Dict[str, Any]],
        title: str = "Analytics & Performance",
        section: Optional[str] = "Analytics",
        takeaway: Optional[str] = None,
        badge: Optional[str] = None,
    ):
        """Native PowerPoint Chart (bar, horizontal bar, stacked, line, pie, doughnut)."""
        slide = self._add_slide(5)
        self._header(slide, title, section, badge)

        chart_type_map = {
            "bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
            "barh": XL_CHART_TYPE.BAR_CLUSTERED,
            "stacked": XL_CHART_TYPE.COLUMN_STACKED,
            "line": XL_CHART_TYPE.LINE_MARKERS,
            "pie": XL_CHART_TYPE.PIE,
            "doughnut": XL_CHART_TYPE.DOUGHNUT,
        }
        xl_type = chart_type_map.get(kind.lower(), XL_CHART_TYPE.COLUMN_CLUSTERED)

        c_data = CategoryChartData()
        c_data.categories = list(categories)
        for s in series:
            c_data.add_series(s.get("name", "Series"), tuple(s.get("values", [])))

        x = CONTENT_X
        y = CONTENT_Y + Inches(0.20)
        w = CONTENT_W
        h = Inches(4.80)

        chart_shape = slide.shapes.add_chart(xl_type, x, y, w, h, c_data)
        chart_obj = chart_shape.chart
        chart_obj.has_legend = len(series) > 1
        if chart_obj.has_legend:
            chart_obj.legend.position = XL_LEGEND_POSITION.TOP
            chart_obj.legend.include_in_layout = False

        self._takeaway(slide, takeaway)
        self._footer(slide)
        return slide

    def two_column(
        self,
        title: str,
        left: Dict[str, Any],
        right: Dict[str, Any],
        section: Optional[str] = "Comparative Analysis",
        takeaway: Optional[str] = None,
        badge: Optional[str] = None,
    ):
        """Two-Column Comparison Layout."""
        slide = self._add_slide(5)
        self._header(slide, title, section, badge)

        gap = Inches(0.35)
        col_w = (CONTENT_W - gap) / 2
        col_h = Inches(4.80)
        y = CONTENT_Y + Inches(0.20)

        for i, col_data in enumerate([left, right]):
            x = CONTENT_X + i * (col_w + gap)
            c_head = col_data.get("heading", f"Column {i+1}")
            c_sub = col_data.get("sub", "")
            items = col_data.get("items", [])

            # Card box
            box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, col_w, col_h)
            box.fill.solid()
            box.fill.fore_color.rgb = rgb(CARD_ACTIVE if i == 0 else CARD_NEUTRAL)
            box.line.color.rgb = rgb(BORDER_BLUE if i == 0 else BORDER_LIGHT)

            # Left accent bar
            bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, Inches(0.06), col_h)
            bar.fill.solid()
            bar.fill.fore_color.rgb = rgb(NAVY_PRIMARY if i == 0 else BLUE_ACCENT)
            bar.line.fill.background()

            # Heading
            tb_h = slide.shapes.add_textbox(x + Inches(0.25), y + Inches(0.20), col_w - Inches(0.50), Inches(0.35))
            tb_h.text_frame.margin_left = tb_h.text_frame.margin_top = tb_h.text_frame.margin_right = tb_h.text_frame.margin_bottom = 0
            ph = tb_h.text_frame.paragraphs[0]
            ph.text = c_head
            ph.font.size = Pt(13)
            ph.font.bold = True
            ph.font.color.rgb = rgb(NAVY_PRIMARY)

            if c_sub:
                ps = tb_h.text_frame.add_paragraph()
                ps.text = c_sub
                ps.font.size = Pt(9.5)
                ps.font.color.rgb = rgb(MUTED_TEXT)

            # Divider
            div = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x + Inches(0.25), y + Inches(0.70), col_w - Inches(0.50), Inches(0.015))
            div.fill.solid()
            div.fill.fore_color.rgb = rgb(BORDER_BLUE if i == 0 else BORDER_LIGHT)
            div.line.fill.background()

            # Items
            tb_b = slide.shapes.add_textbox(x + Inches(0.25), y + Inches(0.85), col_w - Inches(0.50), col_h - Inches(1.00))
            tf_b = tb_b.text_frame
            tf_b.word_wrap = True
            tf_b.margin_left = tf_b.margin_top = tf_b.margin_right = tf_b.margin_bottom = 0

            for j, itm in enumerate(items):
                p = tf_b.paragraphs[0] if j == 0 else tf_b.add_paragraph()
                p.space_after = Pt(8)
                r_bullet = p.add_run()
                r_bullet.text = "•  "
                r_bullet.font.bold = True
                r_bullet.font.color.rgb = rgb(BLUE_ACCENT)
                self._format_markdown_runs(p, str(itm), font_size=Pt(10), font_color=rgb(BODY_DARK))

        self._takeaway(slide, takeaway)
        self._footer(slide)
        return slide

    def key_finding(
        self,
        headline: str,
        statement: str,
        body: str,
        number: Optional[str] = "01",
        section: Optional[str] = "Key Findings",
        takeaway: Optional[str] = None,
        badge: Optional[str] = None,
    ):
        """Single Strategic Key Finding Slide."""
        slide = self._add_slide(5)
        self._header(slide, headline, section, badge)

        # Main callout card
        cx = CONTENT_X
        cy = CONTENT_Y + Inches(0.25)
        cw = CONTENT_W
        ch = Inches(4.80)

        box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, cy, cw, ch)
        box.fill.solid()
        box.fill.fore_color.rgb = rgb(CARD_ACTIVE)
        box.line.color.rgb = rgb(BORDER_BLUE)

        # Left thick bar
        bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, cx, cy, Inches(0.12), ch)
        bar.fill.solid()
        bar.fill.fore_color.rgb = rgb(NAVY_PRIMARY)
        bar.line.fill.background()

        # Large pill with finding number
        if number:
            nb = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx + Inches(0.40), cy + Inches(0.40), Inches(0.80), Inches(0.36))
            nb.fill.solid()
            nb.fill.fore_color.rgb = rgb(NAVY_PRIMARY)
            nb.line.fill.background()
            pnb = nb.text_frame.paragraphs[0]
            pnb.text = number
            pnb.alignment = PP_ALIGN.CENTER
            pnb.font.size = Pt(13)
            pnb.font.bold = True
            pnb.font.color.rgb = rgb(WHITE)

        # Big statement
        tb_s = slide.shapes.add_textbox(cx + Inches(0.40), cy + Inches(0.95), cw - Inches(0.80), Inches(1.10))
        tf_s = tb_s.text_frame
        tf_s.word_wrap = True
        tf_s.margin_left = tf_s.margin_top = tf_s.margin_right = tf_s.margin_bottom = 0
        ps = tf_s.paragraphs[0]
        ps.text = statement
        ps.font.size = Pt(18)
        ps.font.bold = True
        ps.font.color.rgb = rgb(NAVY_PRIMARY)

        # Divider
        div = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, cx + Inches(0.40), cy + Inches(2.20), cw - Inches(0.80), Inches(0.015))
        div.fill.solid()
        div.fill.fore_color.rgb = rgb(BORDER_BLUE)
        div.line.fill.background()

        # Body narrative
        tb_b = slide.shapes.add_textbox(cx + Inches(0.40), cy + Inches(2.40), cw - Inches(0.80), Inches(2.00))
        tf_b = tb_b.text_frame
        tf_b.word_wrap = True
        tf_b.margin_left = tf_b.margin_top = tf_b.margin_right = tf_b.margin_bottom = 0
        pb = tf_b.paragraphs[0]
        self._format_markdown_runs(pb, body, font_size=Pt(11.5), font_color=rgb(BODY_DARK))

        self._takeaway(slide, takeaway)
        self._footer(slide)
        return slide

    def toc(
        self,
        items: Sequence[Dict[str, str]],
        title: str = "Daftar Isi",
        section: Optional[str] = "Agenda Transformasi",
    ):
        """Table of Contents / Agenda Slide."""
        slide = self._add_slide(5)
        self._header(slide, title, section)

        n = len(items)
        cols = 2 if n > 4 else 1
        items_per_col = math.ceil(n / cols)

        col_w = Inches(5.90)
        gap = Inches(0.60)
        y_start = CONTENT_Y + Inches(0.30)
        item_h = Inches(0.85)

        for i, itm in enumerate(items):
            c_idx = i // items_per_col
            r_idx = i % items_per_col
            x = CONTENT_X + c_idx * (col_w + gap)
            y = y_start + r_idx * item_h

            num_str = itm.get("num", f"{i+1:02d}")
            t_str = itm.get("title", f"Section {i+1}")
            d_str = itm.get("desc", "")

            # Item card
            card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, col_w, item_h - Inches(0.12))
            card.fill.solid()
            card.fill.fore_color.rgb = rgb(CARD_NEUTRAL)
            card.line.color.rgb = rgb(BORDER_LIGHT)

            # Number badge
            badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x + Inches(0.15), y + Inches(0.12), Inches(0.48), Inches(0.48))
            badge.fill.solid()
            badge.fill.fore_color.rgb = rgb(NAVY_PRIMARY)
            badge.line.fill.background()
            pb = badge.text_frame.paragraphs[0]
            pb.text = num_str
            pb.alignment = PP_ALIGN.CENTER
            pb.font.size = Pt(13)
            pb.font.bold = True
            pb.font.color.rgb = rgb(WHITE)

            # Text block
            tb = slide.shapes.add_textbox(x + Inches(0.75), y + Inches(0.12), col_w - Inches(0.90), item_h - Inches(0.24))
            tf = tb.text_frame
            tf.word_wrap = True
            tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
            pt = tf.paragraphs[0]
            pt.text = t_str
            pt.font.size = Pt(11.5)
            pt.font.bold = True
            pt.font.color.rgb = rgb(NAVY_PRIMARY)

            if d_str:
                pd = tf.add_paragraph()
                pd.text = d_str
                pd.font.size = Pt(9)
                pd.font.color.rgb = rgb(MUTED_TEXT)

        self._footer(slide)
        return slide

    def section(
        self,
        number: str = "01",
        title: str = "Strategi & Rencana Kerja",
        subtitle: Optional[str] = "Pilar Prioritas dan Roadmap Implementasi AI",
        section: Optional[str] = "Overview",
    ):
        """Executive Section Divider Slide."""
        slide = self._add_slide(2)

        # Big Navy card
        cx = Inches(0.40)
        cy = Inches(1.20)
        cw = Inches(12.53)
        ch = Inches(5.10)

        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, cy, cw, ch)
        card.fill.solid()
        card.fill.fore_color.rgb = rgb(NAVY_PRIMARY)
        card.line.fill.background()

        # Section number
        tb_num = slide.shapes.add_textbox(cx + Inches(0.80), cy + Inches(0.80), Inches(2.00), Inches(0.60))
        tf_num = tb_num.text_frame
        p_num = tf_num.paragraphs[0]
        p_num.text = str(number)
        p_num.font.size = Pt(36)
        p_num.font.bold = True
        p_num.font.color.rgb = rgb(BLUE_ACCENT)

        # Section divider line
        div = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, cx + Inches(0.80), cy + Inches(1.50), Inches(4.00), Inches(0.04))
        div.fill.solid()
        div.fill.fore_color.rgb = rgb(BLUE_ACCENT)
        div.line.fill.background()

        # Title
        tb_t = slide.shapes.add_textbox(cx + Inches(0.80), cy + Inches(1.80), cw - Inches(1.60), Inches(1.20))
        tf_t = tb_t.text_frame
        tf_t.word_wrap = True
        pt = tf_t.paragraphs[0]
        pt.text = title
        pt.font.size = Pt(26)
        pt.font.bold = True
        pt.font.color.rgb = rgb(WHITE)

        # Subtitle
        if subtitle:
            ps = tf_t.add_paragraph()
            ps.text = subtitle
            ps.font.size = Pt(14)
            ps.font.color.rgb = rgb(CARD_ACTIVE)

        self._footer(slide)
        return slide

    def closing(
        self,
        title: str = "Terima kasih",
        subtitle: Optional[str] = "Satgas Transformasi Digital BP BUMN & Perum DAMRI",
        contact: Optional[str] = "info@damri.co.id  ·  www.damri.co.id",
        partner_logo: Optional[str] = None,
    ):
        """Slide 7: Official Closing Slide."""
        slide = self._add_slide(0)

        # 1. Background image
        bg_img = os.path.join(BACKGROUNDS_DIR, "closing-damri.jpg")
        if os.path.exists(bg_img):
            slide.shapes.add_picture(bg_img, Inches(0), Inches(0), Inches(SLIDE_W), Inches(SLIDE_H))
        else:
            bg_rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(SLIDE_W), Inches(SLIDE_H))
            bg_rect.fill.solid()
            bg_rect.fill.fore_color.rgb = rgb(NAVY_PRIMARY)

        # 2. Logos (top-right)
        bumn_logo = os.path.join(LOGOS_DIR, "bumn-white.png")
        damri_logo = os.path.join(LOGOS_DIR, "damri-white.png")

        if os.path.exists(bumn_logo):
            slide.shapes.add_picture(bumn_logo, Inches(9.80), Inches(0.26), Inches(1.20), Inches(0.58))
        if os.path.exists(damri_logo):
            slide.shapes.add_picture(damri_logo, Inches(11.20), Inches(0.38), Inches(1.80), Inches(0.32))

        if partner_logo:
            p_logo_path = os.path.join(LOGOS_DIR, f"{partner_logo}.png")
            if os.path.exists(p_logo_path):
                slide.shapes.add_picture(p_logo_path, Inches(8.30), Inches(0.30), width=Inches(1.20))

        # 3. Bottom Banner
        banner_y = Inches(5.80)
        banner_h = Inches(1.35)
        banner = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), banner_y, Inches(SLIDE_W), banner_h)
        banner.fill.solid()
        banner.fill.fore_color.rgb = rgb(NAVY_PRIMARY)
        banner.line.fill.background()

        # 4. Title
        tb_t = slide.shapes.add_textbox(Inches(0.40), banner_y + Inches(0.18), Inches(12.50), Inches(0.50))
        tf_t = tb_t.text_frame
        tf_t.word_wrap = True
        tf_t.margin_left = tf_t.margin_top = tf_t.margin_right = tf_t.margin_bottom = 0
        pt = tf_t.paragraphs[0]
        pt.text = title
        pt.font.name = FONT_FAMILY
        pt.font.size = Pt(24)
        pt.font.bold = True
        pt.font.color.rgb = rgb(WHITE)

        # 5. Subtitle & contact
        sub_text = subtitle or ""
        if contact:
            sub_text = f"{sub_text}  |  {contact}" if sub_text else contact
        if sub_text:
            tb_s = slide.shapes.add_textbox(Inches(0.40), banner_y + Inches(0.70), Inches(12.50), Inches(0.35))
            tf_s = tb_s.text_frame
            tf_s.word_wrap = True
            tf_s.margin_left = tf_s.margin_top = tf_s.margin_right = tf_s.margin_bottom = 0
            ps = tf_s.paragraphs[0]
            ps.text = sub_text
            ps.font.name = FONT_FAMILY
            ps.font.size = Pt(12)
            ps.font.color.rgb = rgb(CARD_ACTIVE)

        return slide

    def save(self, output_path: str):
        """Saves presentation to the specified file path."""
        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        self.prs.save(output_path)
        print(f"Presentation saved successfully to: {output_path} ({self.slide_index} slides)")
