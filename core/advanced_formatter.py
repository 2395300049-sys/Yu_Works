"""Yu_Works 高精度排版中台 — 缩进单位感知 + 行距 OOXML 同步"""
import logging
from docx.enum.text import WD_ALIGN_PARAGRAPH
from core.indent import apply_style_config_indents
from core.line_spacing import apply_line_spacing, sync_spacing_ooxml

logger = logging.getLogger("Yu_Works.Formatter")


class Yu_WorksStyleAdapter:
    def __init__(self, yu_works_style):
        self.alignment = getattr(yu_works_style, "alignment", "justify")
        self.space_before_pt = getattr(yu_works_style, "space_before_pt", 0.0)
        self.space_after_pt = getattr(yu_works_style, "space_after_pt", 0.0)
        self.line_spacing_mode = getattr(yu_works_style, "line_spacing_mode", "exact")
        self.line_spacing_pt = getattr(yu_works_style, "line_spacing_pt", 20.0)
        self.size_pt = getattr(yu_works_style, "size_pt", 12.0)

        left_cm = getattr(yu_works_style, "left_indent_cm", 0.0)
        right_cm = getattr(yu_works_style, "right_indent_cm", 0.0)
        first_line_cm = getattr(yu_works_style, "first_line_indent_cm", 0.0)
        hanging_cm = getattr(yu_works_style, "hanging_indent_cm", 0.0)

        self.left_indent_chars = left_cm
        self.left_indent_unit = "cm"
        self.right_indent_chars = right_cm
        self.right_indent_unit = "cm"

        if hanging_cm > 0:
            self.special_indent_mode = "hanging"
            self.special_indent_value = hanging_cm
            self.special_indent_unit = "cm"
            self.hanging_indent_chars = hanging_cm
            self.hanging_indent_unit = "cm"
            self.first_line_indent_chars = 0.0
            self.first_line_indent_unit = "cm"
        elif first_line_cm > 0:
            self.special_indent_mode = "first_line"
            self.special_indent_value = first_line_cm
            self.special_indent_unit = "cm"
            self.first_line_indent_chars = first_line_cm
            self.first_line_indent_unit = "cm"
            self.hanging_indent_chars = 0.0
            self.hanging_indent_unit = "cm"
        else:
            self.special_indent_mode = "none"
            self.special_indent_value = 0.0
            self.special_indent_unit = "cm"
            self.first_line_indent_chars = 0.0
            self.first_line_indent_unit = "cm"
            self.hanging_indent_chars = 0.0
            self.hanging_indent_unit = "cm"


class TypographyEngine:
    @staticmethod
    def apply_paragraph_style(paragraph, yu_works_style_config):
        if not paragraph or not yu_works_style_config:
            return False
        try:
            adapted = Yu_WorksStyleAdapter(yu_works_style_config)
            p_format = paragraph.paragraph_format
            p_element = paragraph._element

            align_map = {
                "left": WD_ALIGN_PARAGRAPH.LEFT,
                "right": WD_ALIGN_PARAGRAPH.RIGHT,
                "center": WD_ALIGN_PARAGRAPH.CENTER,
                "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
            }
            p_format.alignment = align_map.get(
                str(adapted.alignment).lower(), WD_ALIGN_PARAGRAPH.JUSTIFY)

            apply_style_config_indents(p_format, p_element, adapted)
            apply_line_spacing(p_format, adapted.line_spacing_mode,
                              adapted.line_spacing_pt)
            sync_spacing_ooxml(
                p_element,
                space_before_pt=adapted.space_before_pt,
                space_after_pt=adapted.space_after_pt,
                line_spacing_type=adapted.line_spacing_mode,
                line_spacing_value=adapted.line_spacing_pt,
            )
            p_format.widow_control = bool(getattr(yu_works_style_config, "widow_control", True))
            p_format.keep_with_next = bool(getattr(yu_works_style_config, "keep_with_next", False))
            p_format.keep_together = bool(getattr(yu_works_style_config, "keep_together", False))
            p_format.page_break_before = bool(getattr(yu_works_style_config, "page_break_before", False))

            ppr = p_element.get_or_add_pPr()
            for attr, tag in (
                    ("widow_control", "widowControl"),
                    ("keep_with_next", "keepNext"),
                    ("keep_together", "keepLines"),
                    ("page_break_before", "pageBreakBefore")):
                val = bool(getattr(yu_works_style_config, attr, attr == "widow_control"))
                el = ppr.find(f"{{http://schemas.openxmlformats.org/wordprocessingml/2006/main}}{tag}")
                if el is None:
                    from docx.oxml import OxmlElement
                    el = OxmlElement(f"w:{tag}")
                    ppr.append(el)
                if val:
                    el.attrib.pop(f"{{http://schemas.openxmlformats.org/wordprocessingml/2006/main}}val", None)
                else:
                    from docx.oxml.ns import qn
                    el.set(qn("w:val"), "0")
            return True
        except Exception as e:
            logger.error(f"TypographyEngine 注入失败: {e}")
            return False
