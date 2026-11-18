"""
BUMN & DAMRI Report Deck Builder.

Faithfully implements the 7 slide templates from the official BUMN Task Force PPT:
  Slide 1: Cover Slide (Bus fleet background, BUMN & DAMRI logos top-right, title banner bottom)
  Slide 2: Prioritization Matrix (2x2 Impact vs Effort + Key Insights sidebar + Confidential footer)
  Slide 3: Quick Wins & Analytical Modules (2x2 module cards + Release milestone bar + Dashboard preview)
  Slide 4: Implementation Pipeline (Data Sourcing -> Modelling & AI -> Recommendation)
  Slide 5: Flow Journey (Step-by-step operational / customer journey)
  Slide 6: Timeline Delivery Gantt (Phase bands, task rows, schedule bars, "We are here" milestone)
  Slide 7: Closing Slide (Bus fleet background, BUMN & DAMRI logos top-right, "Terima kasih" banner)

Every content slide includes the official footer:
"CONFIDENTIAL AND PROPRIETARY | Any use of this material without specific permission of State-Owned Enterprises Regulatory Agency is strictly prohibited"
"""
from __future__ import annotations

import os
import re
from typing import Any, Dict, List, Optional, Sequence, Union

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

HERE = os.path.dirname(os.path.abspath(__file__))
SKILL_DIR = os.path.dirname(HERE)
TEMPLATE_PATH = os.path.join(SKILL_DIR, "assets", "template.pptx")
LOGOS_DIR = os.path.join(SKILL_DIR, "assets", "logos")
BACKGROUNDS_DIR = os.path.join(SKILL_DIR, "assets", "backgrounds")

# --------------------------------------------------------------------------
# Official BUMN / DAMRI Color Palette
# --------------------------------------------------------------------------
NAVY_PRIMARY = "002060"     # BUMN Navy / Deep Blue
NAVY_HEADLINE = "003C79"    # Slide headline blue
BLUE_SECTION = "174EA6"     # Section label blue
BLUE_SLATE = "123C79"       # Subheaders & timeline headers
DARK_SLATE = "0E2841"       # Primary text & task names
BODY_DARK = "1F2A37"        # Body text dark
MUTED_TEXT = "5B6675"       # Subtitles, data labels, secondary info
MUTED_GRAY = "595959"       # Confidential footer text

CARD_ACTIVE = "E6EDF6"      # Active card tint (Quick wins)
CARD_NEUTRAL = "F3F5F8"     # Standard card background & panel fill
PHASE_BAND_1 = "DCE6F2"     # Timeline phase band tint 1
PHASE_BAND_2 = "EEF2F7"     # Timeline phase band tint 2
WHITE = "FFFFFF"

BORDER_BLUE = "B9C7DE"      # Border line
BORDER_LIGHT = "E7E6E6"     # Light separator
BORDER_MID = "BFBFBF"       # Mid separator
RULE_BLUE = "002060"        # Subheader rule

ALERT_RED = "C00000"        # "We are here" milestone, WIP badges
BADGE_GRAY = "8C99AA"       # Quadrant B, C, D badge fill

FONT_FAMILY = "Arial"
CONFIDENTIAL_TEXT = (
    "CONFIDENTIAL AND PROPRIETARY | Any use of this material without specific permission "
    "of State-Owned Enterprises Regulatory Agency is strictly prohibited"
)


def rgb(hex_code: str) -> RGBColor:
    return RGBColor.from_string(hex_code.lstrip("#").upper())


class BUMNDeck:
    """Deck Builder matching the 7 official BUMN Task Force slide templates."""

    def __init__(self, lang: str = "id", template_path: Optional[str] = None):
        self.lang = lang
        tpl = template_path or TEMPLATE_PATH
        if not os.path.exists(tpl):
            raise FileNotFoundError(f"Template not found at: {tpl}")
        self.prs = Presentation(tpl)
        # Clear existing slide instances
        slide_ids = self.prs.slides._sldIdLst
        for s in list(slide_ids):
            self.prs.part.drop_rel(s.rId)
            slide_ids.remove(s)
        self.slide_index = 0

    # --------------------------------------------------------------------------
    # Internal Slide & Header/Footer Helpers
    # --------------------------------------------------------------------------
    def _add_content_slide(self):
        """Creates a content slide using the official Master 1 Title Only layout."""
        # Use Master 1 Layout 5 if available, else fallback to layouts[5]
        if len(self.prs.slide_masters) > 1 and len(self.prs.slide_masters[1].slide_layouts) > 5:
            layout = self.prs.slide_masters[1].slide_layouts[5]
        else:
            layout = self.prs.slide_layouts[5]

        slide = self.prs.slides.add_slide(layout)
        # Remove any leftover master placeholder text boxes
        for ph in list(slide.placeholders):
            sp = ph._element
            sp.getparent().remove(sp)

        self.slide_index += 1
        return slide

    def _header(
        self,
        slide,
        title: str,
        section_label: Optional[str] = None,
        badge: Optional[str] = None,
        badge_color: str = ALERT_RED,
    ):
        """Standard BUMN top header (section-label at y=0.25", headline at y=0.44")."""
        # Section Label (e.g. "Transformation Plan - Dynamic Pricing")
        sec_text = section_label or "Transformation Plan - Dynamic Pricing"
        tb_sec = slide.shapes.add_textbox(Inches(0.40), Inches(0.25), Inches(8.33), Inches(0.24))
        tf_sec = tb_sec.text_frame
        tf_sec.word_wrap = True
        tf_sec.margin_left = tf_sec.margin_top = tf_sec.margin_right = tf_sec.margin_bottom = 0
        p_sec = tf_sec.paragraphs[0]
        p_sec.text = sec_text
        p_sec.font.name = FONT_FAMILY
        p_sec.font.size = Pt(14)
        p_sec.font.bold = True
        p_sec.font.color.rgb = rgb(BLUE_SECTION)

        # Headline
        tb_head = slide.shapes.add_textbox(Inches(0.40), Inches(0.44), Inches(10.50), Inches(0.47))
        tf_head = tb_head.text_frame
        tf_head.word_wrap = True
        tf_head.margin_left = tf_head.margin_top = tf_head.margin_right = tf_head.margin_bottom = 0
        p_head = tf_head.paragraphs[0]
        p_head.text = title
        p_head.font.name = FONT_FAMILY
        p_head.font.size = Pt(18)
        p_head.font.bold = True
        p_head.font.color.rgb = rgb(NAVY_HEADLINE)

        # Optional WIP / Status badge
        if badge:
            sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(10.45), Inches(0.18), Inches(1.32), Inches(0.50))
            sh.fill.solid()
            sh.fill.fore_color.rgb = rgb(badge_color)
            sh.line.color.rgb = rgb(badge_color)
            tf_b = sh.text_frame
            tf_b.word_wrap = True
            tf_b.margin_left = tf_b.margin_top = tf_b.margin_right = tf_b.margin_bottom = 0
            pb = tf_b.paragraphs[0]
            pb.text = badge
            pb.alignment = PP_ALIGN.CENTER
            pb.font.name = FONT_FAMILY
            pb.font.size = Pt(12)
            pb.font.bold = True
            pb.font.color.rgb = rgb(WHITE)

    def _footer(self, slide):
        """Draws the official BUMN confidential footer and page number."""
        # Top-right BUMN logo on content slide
        bumn_logo = os.path.join(LOGOS_DIR, "bumn-primary.png")
        if os.path.exists(bumn_logo):
            slide.shapes.add_picture(bumn_logo, Inches(12.05), Inches(0.22), Inches(1.05), Inches(0.55))

        # Bottom horizontal rule above footer
        line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.10), Inches(7.09), Inches(12.75), Inches(0.01))
        line.fill.solid()
        line.fill.fore_color.rgb = rgb(BORDER_BLUE)
        line.line.fill.background()

        # Confidential text
        tb_conf = slide.shapes.add_textbox(Inches(0.10), Inches(7.18), Inches(10.50), Inches(0.25))
        tf_conf = tb_conf.text_frame
        tf_conf.word_wrap = True
        tf_conf.margin_left = tf_conf.margin_top = tf_conf.margin_right = tf_conf.margin_bottom = 0
        p_conf = tf_conf.paragraphs[0]
        p_conf.text = CONFIDENTIAL_TEXT
        p_conf.font.name = FONT_FAMILY
        p_conf.font.size = Pt(8)
        p_conf.font.color.rgb = rgb(MUTED_GRAY)

        # Slide Number
        tb_num = slide.shapes.add_textbox(Inches(12.40), Inches(7.18), Inches(0.55), Inches(0.25))
        tf_num = tb_num.text_frame
        tf_num.margin_left = tf_num.margin_top = tf_num.margin_right = tf_num.margin_bottom = 0
        p_num = tf_num.paragraphs[0]
        p_num.text = str(self.slide_index)
        p_num.alignment = PP_ALIGN.RIGHT
        p_num.font.name = FONT_FAMILY
        p_num.font.size = Pt(9)
        p_num.font.color.rgb = rgb(MUTED_GRAY)

    def _format_runs(self, paragraph, text: str, font_size=Pt(10), font_color=None, base_bold=False):
        """Applies markdown **bold** formatting to paragraph runs."""
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
    # SLIDE 1 TEMPLATE: Cover Slide
    # --------------------------------------------------------------------------
    def cover(
        self,
        title: str = "Taskforce BP BUMN - Perum Damri",
        subtitle: str = "AI for Public Services: DAMRI's Transformation Roadmap for Public Transportation",
        bg_image: Optional[str] = None,
    ):
        """
        Slide 1: Big background picture + BUMN & DAMRI logos top-right + bottom title banner.
        Exactly matches Slide 1 in Framework Task Force BUMN.pptx.
        """
        # Layout 0: Title Slide
        layout = self.prs.slide_layouts[0]
        slide = self.prs.slides.add_slide(layout)
        for ph in list(slide.placeholders):
            sp = ph._element
            sp.getparent().remove(sp)
        self.slide_index += 1

        # 1. Big Background Image (Full Bleed 13.33 x 7.50)
        img_path = bg_image or os.path.join(BACKGROUNDS_DIR, "cover-damri.jpg")
        if os.path.exists(img_path):
            slide.shapes.add_picture(img_path, Inches(0), Inches(0), Inches(13.333), Inches(7.500))
        else:
            bg_rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.500))
            bg_rect.fill.solid()
            bg_rect.fill.fore_color.rgb = rgb(NAVY_PRIMARY)

        # 2. Logos on Top Right (Picture 6 and Picture 7)
        bumn_logo = os.path.join(LOGOS_DIR, "bumn-white.png")
        damri_logo = os.path.join(LOGOS_DIR, "damri-white.png")

        if os.path.exists(bumn_logo):
            slide.shapes.add_picture(bumn_logo, Inches(9.99), Inches(0.26), Inches(1.10), Inches(0.57))
        if os.path.exists(damri_logo):
            slide.shapes.add_picture(damri_logo, Inches(11.30), Inches(0.39), Inches(1.83), Inches(0.30))

        # 3. Bottom Banner (Rectangle 5 at y=5.92", h=1.20")
        banner = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(5.92), Inches(13.333), Inches(1.20))
        banner.fill.solid()
        banner.fill.fore_color.rgb = rgb(NAVY_PRIMARY)
        banner.line.fill.background()

        # 4. Title Text (TextBox 8)
        tb_title = slide.shapes.add_textbox(Inches(0.39), Inches(6.12), Inches(12.55), Inches(0.50))
        tf_title = tb_title.text_frame
        tf_title.word_wrap = True
        tf_title.margin_left = tf_title.margin_top = tf_title.margin_right = tf_title.margin_bottom = 0
        p_t = tf_title.paragraphs[0]
        p_t.text = title
        p_t.font.name = FONT_FAMILY
        p_t.font.size = Pt(24)
        p_t.font.bold = True
        p_t.font.color.rgb = rgb(WHITE)

        # 5. Subtitle Text (TextBox 9)
        tb_sub = slide.shapes.add_textbox(Inches(0.39), Inches(6.59), Inches(9.60), Inches(0.34))
        tf_sub = tb_sub.text_frame
        tf_sub.word_wrap = True
        tf_sub.margin_left = tf_sub.margin_top = tf_sub.margin_right = tf_sub.margin_bottom = 0
        p_s = tf_sub.paragraphs[0]
        p_s.text = subtitle
        p_s.font.name = FONT_FAMILY
        p_s.font.size = Pt(14)
        p_s.font.color.rgb = rgb(WHITE)

        return slide

    slide_1_cover = cover

    # --------------------------------------------------------------------------
    # SLIDE 2 TEMPLATE: Prioritization Framework (Impact vs Effort 2x2 Matrix)
    # --------------------------------------------------------------------------
    def prioritization_matrix(
        self,
        quadrants: Dict[str, Dict[str, Any]],
        insights: Optional[Sequence[Dict[str, str]]] = None,
        title: str = "Prioritization Framework (Impact vs Effort)",
        section_label: str = "Transformation Plan - Dynamic Pricing",
        badge: Optional[str] = None,
        x_label: str = "EFFORT",
        y_label: str = "IMPACT",
    ):
        """
        Slide 2: 2x2 Impact vs Effort Matrix + Key Insights sidebar.
        Exactly matches Slide 2 in Framework Task Force BUMN.pptx.
        """
        slide = self._add_content_slide()
        self._header(slide, title, section_label, badge)

        # Sub-header: "Use-case prioritization by impact and effort"
        tb_mh = slide.shapes.add_textbox(Inches(0.27), Inches(1.26), Inches(8.54), Inches(0.25))
        tf_mh = tb_mh.text_frame
        tf_mh.margin_left = tf_mh.margin_top = tf_mh.margin_right = tf_mh.margin_bottom = 0
        pmh = tf_mh.paragraphs[0]
        pmh.text = "Use-case prioritization by impact and effort"
        pmh.font.name = FONT_FAMILY
        pmh.font.size = Pt(14)
        pmh.font.bold = True
        pmh.font.color.rgb = rgb(NAVY_PRIMARY)

        # Rule line under matrix header
        rule_m = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.27), Inches(1.57), Inches(8.54), Inches(0.015))
        rule_m.fill.solid()
        rule_m.fill.fore_color.rgb = rgb(RULE_BLUE)
        rule_m.line.fill.background()

        # Matrix Geometry
        qw = Inches(3.97)
        qh_top = Inches(2.31)
        qh_bot = Inches(2.14)
        mx_left = Inches(0.79)
        mx_right = Inches(4.84)
        my_top = Inches(1.83)
        my_bot = Inches(4.22)

        # Y Axis (IMPACT)
        y_title = slide.shapes.add_textbox(Inches(0.18), Inches(3.00), Inches(0.40), Inches(1.00))
        y_title.rotation = 270
        y_title.text_frame.paragraphs[0].text = y_label
        y_title.text_frame.paragraphs[0].font.size = Pt(10)
        y_title.text_frame.paragraphs[0].font.bold = True
        y_title.text_frame.paragraphs[0].font.color.rgb = rgb(NAVY_PRIMARY)

        # X Axis (EFFORT)
        x_title = slide.shapes.add_textbox(Inches(4.30), Inches(6.45), Inches(1.50), Inches(0.25))
        x_title.text_frame.paragraphs[0].text = x_label
        x_title.text_frame.paragraphs[0].font.size = Pt(10)
        x_title.text_frame.paragraphs[0].font.bold = True
        x_title.text_frame.paragraphs[0].font.color.rgb = rgb(NAVY_PRIMARY)

        # Quadrant coordinates (A: top-left, B: top-right, C: bot-left, D: bot-right)
        q_coords = {
            "A": (mx_left, my_top, qw, qh_top, True),
            "B": (mx_right, my_top, qw, qh_top, False),
            "C": (mx_left, my_bot, qw, qh_bot, False),
            "D": (mx_right, my_bot, qw, qh_bot, False),
        }

        for key, (qx, qy, w, h, is_active) in q_coords.items():
            q_data = quadrants.get(key, {})
            q_title = q_data.get("title", f"Quadrant {key}")
            q_sub = q_data.get("subtitle", "")
            q_count = str(q_data.get("count", len(q_data.get("items", []))))
            q_items = q_data.get("items", [])

            # Quadrant Box
            q_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, qx, qy, w, h)
            q_box.fill.solid()
            q_box.fill.fore_color.rgb = rgb(CARD_ACTIVE if is_active else CARD_NEUTRAL)
            q_box.line.color.rgb = rgb(BORDER_LIGHT)

            if is_active:
                # Left accent rule (0.06" wide)
                accent = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, qx, qy, Inches(0.06), h)
                accent.fill.solid()
                accent.fill.fore_color.rgb = rgb(NAVY_PRIMARY)
                accent.line.fill.background()

            # Badge pill
            badge_sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, qx + Inches(0.19), qy + Inches(0.16), Inches(0.28), Inches(0.28))
            badge_sh.fill.solid()
            badge_sh.fill.fore_color.rgb = rgb(NAVY_PRIMARY if is_active else BADGE_GRAY)
            badge_sh.line.fill.background()
            pb = badge_sh.text_frame.paragraphs[0]
            pb.text = key
            pb.alignment = PP_ALIGN.CENTER
            pb.font.size = Pt(10)
            pb.font.bold = True
            pb.font.color.rgb = rgb(WHITE)

            # Title & Sub
            tb_t = slide.shapes.add_textbox(qx + Inches(0.58), qy + Inches(0.12), Inches(2.50), Inches(0.40))
            tf_t = tb_t.text_frame
            tf_t.margin_left = tf_t.margin_top = tf_t.margin_right = tf_t.margin_bottom = 0
            pt = tf_t.paragraphs[0]
            pt.text = q_title
            pt.font.size = Pt(13)
            pt.font.bold = True
            pt.font.color.rgb = rgb(NAVY_PRIMARY)

            if q_sub:
                ps = tf_t.add_paragraph()
                ps.text = q_sub
                ps.font.size = Pt(9.5)
                ps.font.color.rgb = rgb(MUTED_TEXT)

            # Count
            tb_c = slide.shapes.add_textbox(qx + w - Inches(0.70), qy + Inches(0.11), Inches(0.60), Inches(0.36))
            tb_c.text_frame.margin_left = tb_c.text_frame.margin_top = tb_c.text_frame.margin_right = tb_c.text_frame.margin_bottom = 0
            pc = tb_c.text_frame.paragraphs[0]
            pc.text = q_count
            pc.alignment = PP_ALIGN.RIGHT
            pc.font.size = Pt(20)
            pc.font.bold = True
            pc.font.color.rgb = rgb(NAVY_PRIMARY if is_active else MUTED_TEXT)

            # Rule inside card
            q_rule = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, qx + Inches(0.19), qy + Inches(0.61), w - Inches(0.38), Inches(0.01))
            q_rule.fill.solid()
            q_rule.fill.fore_color.rgb = rgb(BORDER_BLUE if is_active else BORDER_LIGHT)
            q_rule.line.fill.background()

            # Item list
            tb_l = slide.shapes.add_textbox(qx + Inches(0.19), qy + Inches(0.72), w - Inches(0.38), h - Inches(0.80))
            tf_l = tb_l.text_frame
            tf_l.word_wrap = True
            tf_l.margin_left = tf_l.margin_top = tf_l.margin_right = tf_l.margin_bottom = 0

            for i, itm in enumerate(q_items[:5]):
                p_i = tf_l.paragraphs[0] if i == 0 else tf_l.add_paragraph()
                p_i.space_after = Pt(3)
                r_num = p_i.add_run()
                r_num.text = f"{i+1:02d}\t"
                r_num.font.size = Pt(10.5)
                r_num.font.bold = True
                r_num.font.color.rgb = rgb(NAVY_PRIMARY if is_active else MUTED_TEXT)
                self._format_runs(p_i, str(itm), font_size=Pt(11), font_color=rgb(BODY_DARK))

        # Right-side Insights Sidebar (Rectangle 198)
        rx = Inches(9.14)
        ry = Inches(1.15)
        rw = Inches(3.92)
        rh = Inches(5.82)

        side_rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, rx, ry, rw, rh)
        side_rect.fill.solid()
        side_rect.fill.fore_color.rgb = rgb(CARD_NEUTRAL)
        side_rect.line.fill.background()

        # Sidebar Header: "Key insights"
        tb_sh = slide.shapes.add_textbox(Inches(9.28), Inches(1.26), Inches(3.65), Inches(0.25))
        tf_sh = tb_sh.text_frame
        tf_sh.margin_left = tf_sh.margin_top = tf_sh.margin_right = tf_sh.margin_bottom = 0
        psh = tf_sh.paragraphs[0]
        psh.text = "Key insights"
        psh.font.size = Pt(14)
        psh.font.bold = True
        psh.font.color.rgb = rgb(NAVY_PRIMARY)

        # Rule line under insights header
        rule_i = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(9.28), Inches(1.57), Inches(3.65), Inches(0.015))
        rule_i.fill.solid()
        rule_i.fill.fore_color.rgb = rgb(RULE_BLUE)
        rule_i.line.fill.background()

        # Lead sentence
        tb_lead = slide.shapes.add_textbox(Inches(9.28), Inches(1.70), Inches(3.65), Inches(0.48))
        tf_lead = tb_lead.text_frame
        tf_lead.word_wrap = True
        tf_lead.margin_left = tf_lead.margin_top = tf_lead.margin_right = tf_lead.margin_bottom = 0
        plead = tf_lead.paragraphs[0]
        plead.text = "15 use cases fall into four value pools, sequenced by revenue impact and delivery complexity"
        plead.font.size = Pt(12)
        plead.font.bold = True
        plead.font.color.rgb = rgb(NAVY_PRIMARY)

        # Insight items
        ins_data = insights or [
            {"badge": "A", "title": "Quick wins", "text": "Improve occupancy and competitor visibility with existing data; no system integration required."},
            {"badge": "B", "title": "Major projects", "text": "The core dynamic fare engines; high value but require automated pricing approval workflows."},
            {"badge": "C", "title": "Low-hanging fruit", "text": "Customer segmentation & governance supporting the pricing rollout."},
            {"badge": "D", "title": "Hygiene factors", "text": "Scenario simulator & ancillary optimization for long-term expansion."},
        ]

        curr_iy = Inches(2.35)
        for ins in ins_data[:4]:
            b_val = ins.get("badge", "•")
            t_val = ins.get("title", "")
            txt_val = ins.get("text", "")

            # Insight divider
            i_rule = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(9.28), curr_iy - Inches(0.10), Inches(3.65), Inches(0.01))
            i_rule.fill.solid()
            i_rule.fill.fore_color.rgb = rgb(BORDER_LIGHT)
            i_rule.line.fill.background()

            # Badge pill
            i_badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.28), curr_iy, Inches(0.24), Inches(0.24))
            i_badge.fill.solid()
            i_badge.fill.fore_color.rgb = rgb(NAVY_PRIMARY if b_val == "A" else BADGE_GRAY)
            i_badge.line.fill.background()
            p_ib = i_badge.text_frame.paragraphs[0]
            p_ib.text = b_val
            p_ib.alignment = PP_ALIGN.CENTER
            p_ib.font.size = Pt(8.5)
            p_ib.font.bold = True
            p_ib.font.color.rgb = rgb(WHITE)

            # Insight text
            tb_it = slide.shapes.add_textbox(Inches(9.60), curr_iy - Inches(0.02), Inches(3.30), Inches(0.85))
            tf_it = tb_it.text_frame
            tf_it.word_wrap = True
            tf_it.margin_left = tf_it.margin_top = tf_it.margin_right = tf_it.margin_bottom = 0
            pit = tf_it.paragraphs[0]
            pit.text = t_val
            pit.font.size = Pt(11)
            pit.font.bold = True
            pit.font.color.rgb = rgb(NAVY_PRIMARY)

            pit_desc = tf_it.add_paragraph()
            self._format_runs(pit_desc, txt_val, font_size=Pt(10.5), font_color=rgb(BODY_DARK))

            curr_iy += Inches(1.05)

        self._footer(slide)
        return slide

    slide_2_matrix = prioritization_matrix

    # --------------------------------------------------------------------------
    # SLIDE 3 TEMPLATE: Quick Wins & Output Release 1 (Analytical Modules)
    # --------------------------------------------------------------------------
    def quick_wins_modules(
        self,
        cards: Sequence[Dict[str, Any]],
        preview_title: str = "05   Monitoring & alerting dashboard",
        preview_caption: str = "This dashboard covered analytical modules (use case 1-4)",
        preview_image: Optional[str] = None,
        release_info: Optional[Dict[str, str]] = None,
        title: str = "Quick Wins & Output Release 1",
        section_label: str = "Transformation Plan - Dynamic Pricing",
        badge: Optional[str] = None,
    ):
        """
        Slide 3: 2x2 Analytical Module Cards + Release bar + Right Preview Dashboard.
        Exactly matches Slide 3 in Framework Task Force BUMN.pptx.
        """
        slide = self._add_content_slide()
        self._header(slide, title, section_label, badge)

        # Left subheader: "01–04   Analytical modules"
        tb_lh = slide.shapes.add_textbox(Inches(0.40), Inches(1.03), Inches(7.10), Inches(0.25))
        tf_lh = tb_lh.text_frame
        tf_lh.margin_left = tf_lh.margin_top = tf_lh.margin_right = tf_lh.margin_bottom = 0
        plh = tf_lh.paragraphs[0]
        plh.text = "01–04   Analytical modules"
        plh.font.size = Pt(14)
        plh.font.bold = True
        plh.font.color.rgb = rgb(NAVY_PRIMARY)

        # Right subheader: "05   Monitoring & alerting dashboard"
        tb_rh = slide.shapes.add_textbox(Inches(7.78), Inches(1.03), Inches(5.15), Inches(0.25))
        tf_rh = tb_rh.text_frame
        tf_rh.margin_left = tf_rh.margin_top = tf_rh.margin_right = tf_rh.margin_bottom = 0
        prh = tf_rh.paragraphs[0]
        prh.text = preview_title
        prh.font.size = Pt(14)
        prh.font.bold = True
        prh.font.color.rgb = rgb(NAVY_PRIMARY)

        # Right caption
        tb_cap = slide.shapes.add_textbox(Inches(7.78), Inches(1.37), Inches(5.15), Inches(0.19))
        tf_cap = tb_cap.text_frame
        tf_cap.margin_left = tf_cap.margin_top = tf_cap.margin_right = tf_cap.margin_bottom = 0
        pcap = tf_cap.paragraphs[0]
        pcap.text = preview_caption
        pcap.font.size = Pt(10)
        pcap.font.color.rgb = rgb(MUTED_TEXT)

        # 4 Cards Geometry (2 columns x 2 rows)
        c_w = Inches(3.44)
        card_coords = [
            (Inches(0.40), Inches(1.54), Inches(0.79)),  # Card 01
            (Inches(4.06), Inches(1.54), Inches(4.44)),  # Card 02
            (Inches(0.40), Inches(3.88), Inches(0.79)),  # Card 03
            (Inches(4.06), Inches(3.88), Inches(4.44)),  # Card 04
        ]

        for i, card in enumerate(cards[:4]):
            col_x, top_y, title_x = card_coords[i]
            c_num = card.get("num", f"{i+1:02d}")
            c_title = card.get("title", f"Module {c_num}")
            c_obj = card.get("objective", "")
            c_data = card.get("data", "")
            c_outcome = card.get("outcome", "")

            # Number Text
            tb_n = slide.shapes.add_textbox(col_x, top_y, Inches(0.36), Inches(0.25))
            tb_n.text_frame.margin_left = tb_n.text_frame.margin_top = tb_n.text_frame.margin_right = tb_n.text_frame.margin_bottom = 0
            pn = tb_n.text_frame.paragraphs[0]
            pn.text = c_num
            pn.font.size = Pt(13)
            pn.font.bold = True
            pn.font.color.rgb = rgb(NAVY_PRIMARY)

            # Title Text
            tb_t = slide.shapes.add_textbox(title_x, top_y, Inches(3.06), Inches(0.25))
            tb_t.text_frame.margin_left = tb_t.text_frame.margin_top = tb_t.text_frame.margin_right = tb_t.text_frame.margin_bottom = 0
            pt = tb_t.text_frame.paragraphs[0]
            pt.text = c_title
            pt.font.size = Pt(13)
            pt.font.bold = True
            pt.font.color.rgb = rgb(NAVY_PRIMARY)

            # Objective Text
            tb_o = slide.shapes.add_textbox(col_x, top_y + Inches(0.31), c_w, Inches(0.36))
            tf_o = tb_o.text_frame
            tf_o.word_wrap = True
            tf_o.margin_left = tf_o.margin_top = tf_o.margin_right = tf_o.margin_bottom = 0
            po = tf_o.paragraphs[0]
            self._format_runs(po, c_obj, font_size=Pt(10), font_color=rgb(BODY_DARK))

            # DATA NEEDED label
            tb_dl = slide.shapes.add_textbox(col_x, top_y + Inches(0.74), c_w, Inches(0.15))
            tb_dl.text_frame.margin_left = tb_dl.text_frame.margin_top = tb_dl.text_frame.margin_right = tb_dl.text_frame.margin_bottom = 0
            pdl = tb_dl.text_frame.paragraphs[0]
            pdl.text = "DATA NEEDED"
            pdl.font.size = Pt(8.5)
            pdl.font.bold = True
            pdl.font.color.rgb = rgb(NAVY_PRIMARY)

            # DATA NEEDED content
            tb_dt = slide.shapes.add_textbox(col_x, top_y + Inches(0.90), c_w, Inches(0.50))
            tf_dt = tb_dt.text_frame
            tf_dt.word_wrap = True
            tf_dt.margin_left = tf_dt.margin_top = tf_dt.margin_right = tf_dt.margin_bottom = 0
            pdt = tf_dt.paragraphs[0]
            self._format_runs(pdt, c_data, font_size=Pt(9), font_color=rgb(MUTED_TEXT))

            # OUTCOME label
            tb_ol = slide.shapes.add_textbox(col_x, top_y + Inches(1.46), c_w, Inches(0.15))
            tb_ol.text_frame.margin_left = tb_ol.text_frame.margin_top = tb_ol.text_frame.margin_right = tb_ol.text_frame.margin_bottom = 0
            pol = tb_ol.text_frame.paragraphs[0]
            pol.text = "OUTCOME"
            pol.font.size = Pt(8.5)
            pol.font.bold = True
            pol.font.color.rgb = rgb(NAVY_PRIMARY)

            # OUTCOME content
            tb_ot = slide.shapes.add_textbox(col_x, top_y + Inches(1.63), c_w, Inches(0.50))
            tf_ot = tb_ot.text_frame
            tf_ot.word_wrap = True
            tf_ot.margin_left = tf_ot.margin_top = tf_ot.margin_right = tf_ot.margin_bottom = 0
            pot = tf_ot.paragraphs[0]
            self._format_runs(pot, c_outcome, font_size=Pt(9), font_color=rgb(MUTED_TEXT))

        # Bottom Release Bar
        rel = release_info or {
            "tag": "RELEASE 1",
            "beta": "Beta   W2 Nov 26",
            "prototype": "Prototype   W2 Dec 26",
        }

        # Release Tag Box (x=0.40, y=6.18, w=1.53, h=0.47)
        rel_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.40), Inches(6.18), Inches(1.53), Inches(0.47))
        rel_box.fill.solid()
        rel_box.fill.fore_color.rgb = rgb(NAVY_PRIMARY)
        rel_box.line.fill.background()
        p_rb = rel_box.text_frame.paragraphs[0]
        p_rb.text = rel.get("tag", "RELEASE 1")
        p_rb.alignment = PP_ALIGN.CENTER
        p_rb.font.size = Pt(11)
        p_rb.font.bold = True
        p_rb.font.color.rgb = rgb(WHITE)

        # Release Beta Text
        tb_beta = slide.shapes.add_textbox(Inches(2.15), Inches(6.22), Inches(2.36), Inches(0.39))
        tb_beta.text_frame.margin_left = tb_beta.text_frame.margin_top = tb_beta.text_frame.margin_right = tb_beta.text_frame.margin_bottom = 0
        p_b = tb_beta.text_frame.paragraphs[0]
        p_b.text = rel.get("beta", "Beta   W2 Nov 26")
        p_b.font.size = Pt(11)
        p_b.font.bold = True
        p_b.font.color.rgb = rgb(NAVY_PRIMARY)

        # Release Prototype Text
        tb_proto = slide.shapes.add_textbox(Inches(4.64), Inches(6.22), Inches(2.78), Inches(0.39))
        tb_proto.text_frame.margin_left = tb_proto.text_frame.margin_top = tb_proto.text_frame.margin_right = tb_proto.text_frame.margin_bottom = 0
        p_p = tb_proto.text_frame.paragraphs[0]
        p_p.text = rel.get("prototype", "Prototype   W2 Dec 26")
        p_p.font.size = Pt(11)
        p_p.font.bold = True
        p_p.font.color.rgb = rgb(NAVY_PRIMARY)

        # Right Dashboard Preview Frame / Image (pos: x=7.78, y=1.58, w=5.35, h=5.35)
        px = Inches(7.78)
        py = Inches(1.58)
        pw = Inches(5.35)
        ph = Inches(5.15)

        if preview_image and os.path.exists(preview_image):
            slide.shapes.add_picture(preview_image, px, py, pw, ph)
        else:
            # Elegant placeholder card frame matching slide 3
            p_frame = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, px, py, pw, ph)
            p_frame.fill.solid()
            p_frame.fill.fore_color.rgb = rgb(CARD_NEUTRAL)
            p_frame.line.color.rgb = rgb(BORDER_LIGHT)

            tb_pf = slide.shapes.add_textbox(px + Inches(0.50), py + Inches(2.00), pw - Inches(1.00), Inches(1.00))
            tf_pf = tb_pf.text_frame
            tf_pf.word_wrap = True
            ppf = tf_pf.paragraphs[0]
            ppf.text = f"[ {preview_caption} ]"
            ppf.alignment = PP_ALIGN.CENTER
            ppf.font.size = Pt(11)
            ppf.font.color.rgb = rgb(MUTED_TEXT)

        self._footer(slide)
        return slide

    slide_3_modules = quick_wins_modules

    # --------------------------------------------------------------------------
    # SLIDE 4 TEMPLATE: Implementation Pipeline & Architecture
    # --------------------------------------------------------------------------
    def implementation_pipeline(
        self,
        stages: Optional[Sequence[Dict[str, Any]]] = None,
        title: str = "Implementation",
        section_label: str = "Transformation Plan - Dynamic Pricing",
        badge: Optional[str] = "WIP",
    ):
        """
        Slide 4: 3-Stage Pipeline (Data Sourcing -> Modelling & AI -> Recommendation).
        Exactly matches Slide 4 in Framework Task Force BUMN.pptx.
        """
        slide = self._add_content_slide()
        self._header(slide, title, section_label, badge)

        # Outer canvas container
        outer = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.34), Inches(1.13), Inches(12.72), Inches(5.75))
        outer.fill.solid()
        outer.fill.fore_color.rgb = rgb(WHITE)
        outer.line.color.rgb = rgb(BORDER_LIGHT)

        default_stages = [
            {
                "num": "1",
                "title": "Data Sourcing & Pipeline",
                "cards": [
                    {"name": "Load Factor Prediction", "desc": "Last 2 years route & capacity\nNumber of bus, operated & fulfillment\nDeparture schedule + load factor\nMaster route"},
                    {"name": "Booking Curve Analysis", "desc": "Historical order\nBooking date vs departure date\nCancellation rate"},
                    {"name": "Competitor Price Intelligence", "desc": "OTA Data Scraping vs competitor\nRoute fare matching\nCadence monitoring"},
                    {"name": "Promotion Optimization", "desc": "Pricing & tiering\nOTA Flash sale\nRevenue per branches"},
                ],
            },
            {
                "num": "2",
                "title": "Modelling & Optimization",
                "cards": [
                    {"name": "Machine Learning", "desc": "AI / ML Model\nTime-series demand forecasting\nSeat elasticity optimization"},
                    {"name": "Gen AI", "desc": "Natural language query engine\nAutomated commercial narrative\nExecutive anomaly alerts"},
                ],
            },
            {
                "num": "3",
                "title": "Recommendation",
                "cards": [
                    {"name": "Deploy Model", "desc": "Outcome per use case\nCommercial quota adjustment\nDynamic ticket pricing output"},
                ],
            },
        ]
        stgs = stages or default_stages

        # Column 1: Data Sourcing (w=4.73)
        col1_x = Inches(0.43)
        col1_w = Inches(4.73)
        col_y = Inches(1.26)
        col_h = Inches(5.45)

        # Box 1
        b1 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, col1_x, col_y, col1_w, col_h)
        b1.fill.solid()
        b1.fill.fore_color.rgb = rgb(CARD_NEUTRAL)
        b1.line.color.rgb = rgb(BORDER_LIGHT)

        # Badge Circle 1
        circ1 = slide.shapes.add_shape(MSO_SHAPE.OVAL, col1_x + Inches(0.12), col_y + Inches(0.12), Inches(0.30), Inches(0.30))
        circ1.fill.solid()
        circ1.fill.fore_color.rgb = rgb(NAVY_PRIMARY)
        circ1.line.fill.background()
        p_c1 = circ1.text_frame.paragraphs[0]
        p_c1.text = "1"
        p_c1.alignment = PP_ALIGN.CENTER
        p_c1.font.size = Pt(11)
        p_c1.font.bold = True
        p_c1.font.color.rgb = rgb(WHITE)

        # Col 1 Title
        tb_t1 = slide.shapes.add_textbox(col1_x + Inches(0.50), col_y + Inches(0.12), Inches(4.00), Inches(0.30))
        tb_t1.text_frame.margin_left = tb_t1.text_frame.margin_top = tb_t1.text_frame.margin_right = tb_t1.text_frame.margin_bottom = 0
        pt1 = tb_t1.text_frame.paragraphs[0]
        pt1.text = stgs[0].get("title", "Data Sourcing & Pipeline")
        pt1.font.size = Pt(13)
        pt1.font.bold = True
        pt1.font.color.rgb = rgb(NAVY_PRIMARY)

        # Col 1 Cards (2x2 grid inside col 1)
        c1_cards = stgs[0].get("cards", [])
        c1_grid = [
            (col1_x + Inches(0.14), col_y + Inches(0.66), Inches(2.17), Inches(2.00)),  # Top Left
            (col1_x + Inches(2.43), col_y + Inches(0.66), Inches(2.17), Inches(2.00)),  # Top Right
            (col1_x + Inches(0.14), col_y + Inches(2.80), Inches(2.17), Inches(2.35)),  # Bot Left
            (col1_x + Inches(2.43), col_y + Inches(2.80), Inches(2.17), Inches(2.35)),  # Bot Right
        ]
        for idx, (cx, cy, cw, ch) in enumerate(c1_grid[:len(c1_cards)]):
            c_info = c1_cards[idx]
            c_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, cx, cy, cw, ch)
            c_box.fill.solid()
            c_box.fill.fore_color.rgb = rgb(WHITE)
            c_box.line.color.rgb = rgb(BORDER_LIGHT)

            tb_cn = slide.shapes.add_textbox(cx + Inches(0.10), cy + Inches(0.08), cw - Inches(0.20), Inches(0.30))
            tb_cn.text_frame.margin_left = tb_cn.text_frame.margin_top = tb_cn.text_frame.margin_right = tb_cn.text_frame.margin_bottom = 0
            pcn = tb_cn.text_frame.paragraphs[0]
            pcn.text = c_info.get("name", "")
            pcn.font.size = Pt(10)
            pcn.font.bold = True
            pcn.font.color.rgb = rgb(NAVY_PRIMARY)

            tb_cd = slide.shapes.add_textbox(cx + Inches(0.10), cy + Inches(0.40), cw - Inches(0.20), ch - Inches(0.45))
            tf_cd = tb_cd.text_frame
            tf_cd.word_wrap = True
            tf_cd.margin_left = tf_cd.margin_top = tf_cd.margin_right = tf_cd.margin_bottom = 0
            pcd = tf_cd.paragraphs[0]
            pcd.text = c_info.get("desc", "")
            pcd.font.size = Pt(9)
            pcd.font.color.rgb = rgb(BODY_DARK)

        # Connector 1 (>)
        arr1 = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(5.24), Inches(3.83), Inches(0.33), Inches(0.33))
        arr1.fill.solid()
        arr1.fill.fore_color.rgb = rgb(BLUE_SECTION)
        arr1.line.fill.background()
        p_a1 = arr1.text_frame.paragraphs[0]
        p_a1.text = ">"
        p_a1.alignment = PP_ALIGN.CENTER
        p_a1.font.size = Pt(12)
        p_a1.font.bold = True
        p_a1.font.color.rgb = rgb(WHITE)

        # Column 2: Modelling & Optimization (w=3.65)
        col2_x = Inches(5.65)
        col2_w = Inches(3.65)
        b2 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, col2_x, col_y, col2_w, col_h)
        b2.fill.solid()
        b2.fill.fore_color.rgb = rgb(CARD_ACTIVE)
        b2.line.color.rgb = rgb(BORDER_BLUE)

        circ2 = slide.shapes.add_shape(MSO_SHAPE.OVAL, col2_x + Inches(0.12), col_y + Inches(0.12), Inches(0.30), Inches(0.30))
        circ2.fill.solid()
        circ2.fill.fore_color.rgb = rgb(NAVY_PRIMARY)
        circ2.line.fill.background()
        p_c2 = circ2.text_frame.paragraphs[0]
        p_c2.text = "2"
        p_c2.alignment = PP_ALIGN.CENTER
        p_c2.font.size = Pt(11)
        p_c2.font.bold = True
        p_c2.font.color.rgb = rgb(WHITE)

        tb_t2 = slide.shapes.add_textbox(col2_x + Inches(0.50), col_y + Inches(0.12), Inches(3.00), Inches(0.30))
        tb_t2.text_frame.margin_left = tb_t2.text_frame.margin_top = tb_t2.text_frame.margin_right = tb_t2.text_frame.margin_bottom = 0
        pt2 = tb_t2.text_frame.paragraphs[0]
        pt2.text = stgs[1].get("title", "Modelling & Optimization")
        pt2.font.size = Pt(13)
        pt2.font.bold = True
        pt2.font.color.rgb = rgb(NAVY_PRIMARY)

        # Col 2 Cards (Machine Learning & Gen AI)
        c2_cards = stgs[1].get("cards", [])
        if len(c2_cards) >= 1:
            box_ml = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, col2_x + Inches(0.18), col_y + Inches(0.66), Inches(3.26), Inches(2.20))
            box_ml.fill.solid()
            box_ml.fill.fore_color.rgb = rgb(WHITE)
            box_ml.line.color.rgb = rgb(BORDER_LIGHT)

            tb_ml = slide.shapes.add_textbox(col2_x + Inches(0.30), col_y + Inches(0.78), Inches(3.00), Inches(1.90))
            tf_ml = tb_ml.text_frame
            tf_ml.word_wrap = True
            p_ml1 = tf_ml.paragraphs[0]
            p_ml1.text = c2_cards[0].get("name", "Machine Learning")
            p_ml1.font.size = Pt(11)
            p_ml1.font.bold = True
            p_ml1.font.color.rgb = rgb(BLUE_SLATE)
            p_ml2 = tf_ml.add_paragraph()
            p_ml2.text = c2_cards[0].get("desc", "")
            p_ml2.font.size = Pt(9.5)
            p_ml2.font.color.rgb = rgb(BODY_DARK)

        if len(c2_cards) >= 2:
            box_gen = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, col2_x + Inches(0.18), col_y + Inches(3.00), Inches(3.26), Inches(2.15))
            box_gen.fill.solid()
            box_gen.fill.fore_color.rgb = rgb(WHITE)
            box_gen.line.color.rgb = rgb(BORDER_LIGHT)

            tb_gen = slide.shapes.add_textbox(col2_x + Inches(0.30), col_y + Inches(3.12), Inches(3.00), Inches(1.85))
            tf_gen = tb_gen.text_frame
            tf_gen.word_wrap = True
            p_g1 = tf_gen.paragraphs[0]
            p_g1.text = c2_cards[1].get("name", "Gen AI")
            p_g1.font.size = Pt(11)
            p_g1.font.bold = True
            p_g1.font.color.rgb = rgb(BLUE_SLATE)
            p_g2 = tf_gen.add_paragraph()
            p_g2.text = c2_cards[1].get("desc", "")
            p_g2.font.size = Pt(9.5)
            p_g2.font.color.rgb = rgb(BODY_DARK)

        # Connector 2 (>)
        arr2 = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(9.41), Inches(3.83), Inches(0.33), Inches(0.33))
        arr2.fill.solid()
        arr2.fill.fore_color.rgb = rgb(BLUE_SECTION)
        arr2.line.fill.background()
        p_a2 = arr2.text_frame.paragraphs[0]
        p_a2.text = ">"
        p_a2.alignment = PP_ALIGN.CENTER
        p_a2.font.size = Pt(12)
        p_a2.font.bold = True
        p_a2.font.color.rgb = rgb(WHITE)

        # Column 3: Recommendation (w=3.13)
        col3_x = Inches(9.84)
        col3_w = Inches(3.13)
        b3 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, col3_x, col_y, col3_w, col_h)
        b3.fill.solid()
        b3.fill.fore_color.rgb = rgb(CARD_NEUTRAL)
        b3.line.color.rgb = rgb(BORDER_LIGHT)

        circ3 = slide.shapes.add_shape(MSO_SHAPE.OVAL, col3_x + Inches(0.12), col_y + Inches(0.12), Inches(0.30), Inches(0.30))
        circ3.fill.solid()
        circ3.fill.fore_color.rgb = rgb(NAVY_PRIMARY)
        circ3.line.fill.background()
        p_c3 = circ3.text_frame.paragraphs[0]
        p_c3.text = "3"
        p_c3.alignment = PP_ALIGN.CENTER
        p_c3.font.size = Pt(11)
        p_c3.font.bold = True
        p_c3.font.color.rgb = rgb(WHITE)

        tb_t3 = slide.shapes.add_textbox(col3_x + Inches(0.50), col_y + Inches(0.12), Inches(2.50), Inches(0.30))
        tb_t3.text_frame.margin_left = tb_t3.text_frame.margin_top = tb_t3.text_frame.margin_right = tb_t3.text_frame.margin_bottom = 0
        pt3 = tb_t3.text_frame.paragraphs[0]
        pt3.text = stgs[2].get("title", "Recommendation") if len(stgs) > 2 else "Recommendation"
        pt3.font.size = Pt(13)
        pt3.font.bold = True
        pt3.font.color.rgb = rgb(NAVY_PRIMARY)

        c3_cards = stgs[2].get("cards", []) if len(stgs) > 2 else []
        box_rec = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, col3_x + Inches(0.14), col_y + Inches(0.66), Inches(2.85), Inches(4.50))
        box_rec.fill.solid()
        box_rec.fill.fore_color.rgb = rgb(WHITE)
        box_rec.line.color.rgb = rgb(BORDER_LIGHT)

        tb_rec = slide.shapes.add_textbox(col3_x + Inches(0.25), col_y + Inches(0.80), Inches(2.60), Inches(4.20))
        tf_rec = tb_rec.text_frame
        tf_rec.word_wrap = True
        for idx, itm in enumerate(c3_cards):
            p = tf_rec.paragraphs[0] if idx == 0 else tf_rec.add_paragraph()
            p.space_after = Pt(10)
            r1 = p.add_run()
            r1.text = f"{itm.get('name', '')}\n"
            r1.font.bold = True
            r1.font.size = Pt(11)
            r1.font.color.rgb = rgb(BLUE_SLATE)
            r2 = p.add_run()
            r2.text = itm.get("desc", "")
            r2.font.size = Pt(9.5)
            r2.font.color.rgb = rgb(BODY_DARK)

        self._footer(slide)
        return slide

    slide_4_architecture = implementation_pipeline

    # --------------------------------------------------------------------------
    # SLIDE 5 TEMPLATE: Flow Journey
    # --------------------------------------------------------------------------
    def flow_journey(
        self,
        steps: Optional[Sequence[Dict[str, Any]]] = None,
        title: str = "Flow Journey",
        section_label: str = "Transformation Plan - Dynamic Pricing",
        badge: Optional[str] = "WIP",
    ):
        """
        Slide 5: End-to-End Operational / Customer Flow Journey.
        Matches Slide 5 in Framework Task Force BUMN.pptx.
        """
        slide = self._add_content_slide()
        self._header(slide, title, section_label, badge)

        default_steps = [
            {"num": "01", "stage": "Pencarian & Permintaan", "title": "Customer Intent", "desc": "Calon penumpang mencari jadwal via DAMRI Apps atau loket; sinyal permintaan awal terekam."},
            {"num": "02", "stage": "Optimasi Tarif", "title": "Dynamic Pricing Engine", "desc": "Sistem menghitung load factor prediksi dan membandingkan tarif kompetitor OTA secara real-time."},
            {"num": "03", "stage": "Alokasi Kursi", "title": "Inventory Control", "desc": "Pemberian tier harga promo atau reguler sesuai booking curve keberangkatan armada."},
            {"num": "04", "stage": "Operasional & Rit", "title": "Dispatch & LMB Online", "desc": "Kru bus dan timer memvalidasi jumlah muatan aktual via SIMA Operasi."},
            {"num": "05", "stage": "Evaluasi & Feedback", "title": "Revenue Realization", "desc": "Rekapitulasi pendapatan per cabang dan feedback model untuk akurasi prediksi jadwal berikutnya."},
        ]
        stg_list = steps or default_steps
        n = len(stg_list)

        gap = Inches(0.20)
        card_w = (Inches(12.53) - (gap * (n - 1))) / n
        card_h = Inches(5.10)
        y = Inches(1.35)

        for i, s in enumerate(stg_list):
            x = Inches(0.40) + i * (card_w + gap)
            s_num = s.get("num", f"{i+1:02d}")
            s_stage = s.get("stage", f"Stage {i+1}")
            s_title = s.get("title", "")
            s_desc = s.get("desc", "")

            # Step card
            sc = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, card_w, card_h)
            sc.fill.solid()
            sc.fill.fore_color.rgb = rgb(CARD_ACTIVE if i == 1 else CARD_NEUTRAL)
            sc.line.color.rgb = rgb(BORDER_BLUE if i == 1 else BORDER_LIGHT)

            # Top accent bar
            top_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, card_w, Inches(0.06))
            top_bar.fill.solid()
            top_bar.fill.fore_color.rgb = rgb(NAVY_PRIMARY if i == 1 else BLUE_SECTION)
            top_bar.line.fill.background()

            # Number badge pill
            nb = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x + Inches(0.15), y + Inches(0.20), Inches(0.45), Inches(0.45))
            nb.fill.solid()
            nb.fill.fore_color.rgb = rgb(NAVY_PRIMARY if i == 1 else BLUE_SLATE)
            nb.line.fill.background()
            p_nb = nb.text_frame.paragraphs[0]
            p_nb.text = s_num
            p_nb.alignment = PP_ALIGN.CENTER
            p_nb.font.size = Pt(12)
            p_nb.font.bold = True
            p_nb.font.color.rgb = rgb(WHITE)

            # Stage name
            tb_st = slide.shapes.add_textbox(x + Inches(0.15), y + Inches(0.75), card_w - Inches(0.30), Inches(0.30))
            tb_st.text_frame.margin_left = tb_st.text_frame.margin_top = tb_st.text_frame.margin_right = tb_st.text_frame.margin_bottom = 0
            pst = tb_st.text_frame.paragraphs[0]
            pst.text = s_stage
            pst.font.size = Pt(9.5)
            pst.font.bold = True
            pst.font.color.rgb = rgb(MUTED_TEXT)

            # Step Title
            tb_t = slide.shapes.add_textbox(x + Inches(0.15), y + Inches(1.10), card_w - Inches(0.30), Inches(0.50))
            tf_t = tb_t.text_frame
            tf_t.word_wrap = True
            tf_t.margin_left = tf_t.margin_top = tf_t.margin_right = tf_t.margin_bottom = 0
            pt = tf_t.paragraphs[0]
            pt.text = s_title
            pt.font.size = Pt(12)
            pt.font.bold = True
            pt.font.color.rgb = rgb(NAVY_PRIMARY)

            # Rule inside card
            rule = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x + Inches(0.15), y + Inches(1.70), card_w - Inches(0.30), Inches(0.01))
            rule.fill.solid()
            rule.fill.fore_color.rgb = rgb(BORDER_BLUE if i == 1 else BORDER_LIGHT)
            rule.line.fill.background()

            # Description
            tb_d = slide.shapes.add_textbox(x + Inches(0.15), y + Inches(1.85), card_w - Inches(0.30), card_h - Inches(2.00))
            tf_d = tb_d.text_frame
            tf_d.word_wrap = True
            tf_d.margin_left = tf_d.margin_top = tf_d.margin_right = tf_d.margin_bottom = 0
            pd = tf_d.paragraphs[0]
            self._format_runs(pd, s_desc, font_size=Pt(10), font_color=rgb(BODY_DARK))

            # Arrow indicator between cards
            if i < n - 1:
                arr = slide.shapes.add_shape(MSO_SHAPE.OVAL, x + card_w + Inches(0.03), y + Inches(2.20), Inches(0.16), Inches(0.16))
                arr.fill.solid()
                arr.fill.fore_color.rgb = rgb(BLUE_SECTION)
                arr.line.fill.background()
                pa = arr.text_frame.paragraphs[0]
                pa.text = ">"
                pa.alignment = PP_ALIGN.CENTER
                pa.font.size = Pt(7)
                pa.font.bold = True
                pa.font.color.rgb = rgb(WHITE)

        self._footer(slide)
        return slide

    slide_5_journey = flow_journey

    # --------------------------------------------------------------------------
    # SLIDE 6 TEMPLATE: Timeline Delivery (Gantt with Phase Bands)
    # --------------------------------------------------------------------------
    def timeline_delivery(
        self,
        phases: Sequence[Dict[str, Any]],
        tasks: Sequence[Dict[str, Any]],
        current_point: Optional[Dict[str, Any]] = None,
        title: str = "Timeline delivery: Prototype launch in Dec W2",
        section_label: str = "Transformation Plan - Dynamic Pricing",
        note: str = "Beta is followed by a review before the Dec W2 prototype launch",
        badge: Optional[str] = "WIP",
    ):
        """
        Slide 6: Delivery Timeline Gantt with Phase bands & "We are here" red milestone pin.
        Exactly matches Slide 6 in Framework Task Force BUMN.pptx.
        """
        slide = self._add_content_slide()
        self._header(slide, title, section_label, badge)

        # Headers: Phase & Activity (y=1.84)
        x_phase = Inches(0.56)
        w_phase = Inches(0.97)
        x_task = Inches(1.56)
        w_task = Inches(2.59)
        x_grid = Inches(4.30)
        w_grid = Inches(8.40)
        y_top = Inches(2.11)

        # Header labels
        tb_hp = slide.shapes.add_textbox(x_phase, Inches(1.84), w_phase, Inches(0.20))
        tb_hp.text_frame.margin_left = tb_hp.text_frame.margin_top = tb_hp.text_frame.margin_right = tb_hp.text_frame.margin_bottom = 0
        php = tb_hp.text_frame.paragraphs[0]
        php.text = "Phase"
        php.font.size = Pt(11)
        php.font.bold = True
        php.font.color.rgb = rgb(BLUE_SLATE)

        tb_ha = slide.shapes.add_textbox(x_task, Inches(1.84), w_task, Inches(0.20))
        tb_ha.text_frame.margin_left = tb_ha.text_frame.margin_top = tb_ha.text_frame.margin_right = tb_ha.text_frame.margin_bottom = 0
        pha = tb_ha.text_frame.paragraphs[0]
        pha.text = "Activity"
        pha.font.size = Pt(11)
        pha.font.bold = True
        pha.font.color.rgb = rgb(BLUE_SLATE)

        # Tasks & Rows (8 tasks)
        num_tasks = len(tasks)
        row_h = Inches(0.555)  # exactly matches 4.44" / 8 = 0.555"

        for i, tsk in enumerate(tasks):
            ry = y_top + i * row_h
            t_name = tsk.get("name", f"Task {i}")
            start_frac = float(tsk.get("start", 0.0))
            dur_frac = float(tsk.get("dur", 0.20))

            # Task Text
            tb_t = slide.shapes.add_textbox(x_task, ry + Inches(0.08), w_task, row_h - Inches(0.10))
            tf_t = tb_t.text_frame
            tf_t.word_wrap = True
            tf_t.margin_left = tf_t.margin_top = tf_t.margin_right = tf_t.margin_bottom = 0
            pt = tf_t.paragraphs[0]
            pt.text = t_name
            pt.font.size = Pt(11)
            pt.font.color.rgb = rgb(DARK_SLATE)

            # Separator line across slide width (Sep 0 to Sep 7, w=11.22)
            sep = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x_task, ry + row_h, Inches(11.22), Inches(0.01))
            sep.fill.solid()
            sep.fill.fore_color.rgb = rgb(BORDER_LIGHT if i % 2 == 0 else BORDER_MID)
            sep.line.fill.background()

            # Gantt bar
            bx = x_grid + start_frac * w_grid
            bw = dur_frac * w_grid
            bh = Inches(0.24)
            by = ry + (row_h - bh) / 2

            bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, bx, by, bw, bh)
            bar.fill.solid()
            bar.fill.fore_color.rgb = rgb(BLUE_SECTION if i % 2 == 0 else NAVY_PRIMARY)
            bar.line.fill.background()

        # Phase Bands (Discover, Analyze & design, Build & test, Release)
        curr_py = y_top
        for p_idx, ph in enumerate(phases):
            p_name = ph.get("name", f"Phase {p_idx}")
            span = int(ph.get("span", 2))
            band_h = span * row_h - Inches(0.04)

            band = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x_phase, curr_py + Inches(0.02), w_phase, band_h)
            band.fill.solid()
            band.fill.fore_color.rgb = rgb(PHASE_BAND_1 if p_idx % 2 == 0 else PHASE_BAND_2)
            band.line.fill.background()

            tf_b = band.text_frame
            tf_b.word_wrap = True
            tf_b.margin_left = tf_b.margin_right = Inches(0.05)
            pb = tf_b.paragraphs[0]
            pb.text = p_name
            pb.alignment = PP_ALIGN.CENTER
            pb.font.size = Pt(10.5)
            pb.font.bold = True
            pb.font.color.rgb = rgb(BLUE_SLATE)

            curr_py += span * row_h

        # "We are here" red milestone pin
        cp = current_point or {"frac": 0.70, "label": "We are here"}
        c_frac = float(cp.get("frac", 0.70))
        c_label = cp.get("label", "We are here")
        pin_x = x_grid + c_frac * w_grid

        # Vertical line
        v_line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, pin_x, y_top - Inches(0.15), Inches(0.02), Inches(4.65))
        v_line.fill.solid()
        v_line.fill.fore_color.rgb = rgb(ALERT_RED)
        v_line.line.fill.background()

        # Red label pill at bottom
        pill = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, pin_x - Inches(0.70), Inches(6.69), Inches(1.39), Inches(0.22))
        pill.fill.solid()
        pill.fill.fore_color.rgb = rgb(ALERT_RED)
        pill.line.fill.background()
        p_pill = pill.text_frame.paragraphs[0]
        p_pill.text = c_label
        p_pill.alignment = PP_ALIGN.CENTER
        p_pill.font.size = Pt(10)
        p_pill.font.bold = True
        p_pill.font.color.rgb = rgb(WHITE)

        # Note at bottom
        tb_note = slide.shapes.add_textbox(x_phase, Inches(6.69), Inches(5.28), Inches(0.22))
        tb_note.text_frame.margin_left = tb_note.text_frame.margin_top = tb_note.text_frame.margin_right = tb_note.text_frame.margin_bottom = 0
        pn = tb_note.text_frame.paragraphs[0]
        pn.text = note
        pn.font.size = Pt(10)
        pn.font.color.rgb = rgb(MUTED_GRAY)

        self._footer(slide)
        return slide

    slide_6_timeline = timeline_delivery

    # --------------------------------------------------------------------------
    # SLIDE 7 TEMPLATE: Closing Slide
    # --------------------------------------------------------------------------
    def closing(
        self,
        title: str = "Terima kasih",
        subtitle: Optional[str] = "Taskforce BP BUMN - Perum Damri",
        contact: Optional[str] = "info@damri.co.id",
        bg_image: Optional[str] = None,
    ):
        """
        Slide 7: Bus fleet background picture + BUMN & DAMRI logos top-right + bottom "Terima kasih" banner.
        Exactly matches Slide 7 in Framework Task Force BUMN.pptx.
        """
        layout = self.prs.slide_layouts[0]
        slide = self.prs.slides.add_slide(layout)
        for ph in list(slide.placeholders):
            sp = ph._element
            sp.getparent().remove(sp)
        self.slide_index += 1

        # 1. Background image
        img_path = bg_image or os.path.join(BACKGROUNDS_DIR, "closing-damri.jpg")
        if os.path.exists(img_path):
            slide.shapes.add_picture(img_path, Inches(0), Inches(0), Inches(13.333), Inches(7.500))
        else:
            bg_rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.500))
            bg_rect.fill.solid()
            bg_rect.fill.fore_color.rgb = rgb(NAVY_PRIMARY)

        # 2. Logos on Top Right
        bumn_logo = os.path.join(LOGOS_DIR, "bumn-white.png")
        damri_logo = os.path.join(LOGOS_DIR, "damri-white.png")

        if os.path.exists(bumn_logo):
            slide.shapes.add_picture(bumn_logo, Inches(9.99), Inches(0.26), Inches(1.10), Inches(0.57))
        if os.path.exists(damri_logo):
            slide.shapes.add_picture(damri_logo, Inches(11.30), Inches(0.39), Inches(1.83), Inches(0.30))

        # 3. Bottom Banner (Rectangle 5 at y=5.92, h=1.20)
        banner = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(5.92), Inches(13.333), Inches(1.20))
        banner.fill.solid()
        banner.fill.fore_color.rgb = rgb(NAVY_PRIMARY)
        banner.line.fill.background()

        # 4. Title: "Terima kasih" (TextBox 8 at y=6.27, h=0.50)
        tb_t = slide.shapes.add_textbox(Inches(0.39), Inches(6.15), Inches(12.55), Inches(0.50))
        tf_t = tb_t.text_frame
        tf_t.word_wrap = True
        tf_t.margin_left = tf_t.margin_top = tf_t.margin_right = tf_t.margin_bottom = 0
        pt = tf_t.paragraphs[0]
        pt.text = title
        pt.font.name = FONT_FAMILY
        pt.font.size = Pt(24)
        pt.font.bold = True
        pt.font.color.rgb = rgb(WHITE)

        # 5. Optional Subtitle / Contact
        extra = []
        if subtitle:
            extra.append(subtitle)
        if contact:
            extra.append(contact)
        if extra:
            tb_s = slide.shapes.add_textbox(Inches(0.39), Inches(6.65), Inches(12.55), Inches(0.35))
            tf_s = tb_s.text_frame
            tf_s.word_wrap = True
            tf_s.margin_left = tf_s.margin_top = tf_s.margin_right = tf_s.margin_bottom = 0
            ps = tf_s.paragraphs[0]
            ps.text = "   |   ".join(extra)
            ps.font.name = FONT_FAMILY
            ps.font.size = Pt(12)
            ps.font.color.rgb = rgb(CARD_ACTIVE)

        return slide

    slide_7_closing = closing

    # --------------------------------------------------------------------------
    # Save Presentation
    # --------------------------------------------------------------------------
    def save(self, output_path: str):
        """Saves presentation file to disk."""
        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        self.prs.save(output_path)
        print(f"Presentation saved successfully to: {output_path} ({self.slide_index} slides)")
