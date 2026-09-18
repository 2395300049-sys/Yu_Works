"""排版配置数据树——所有格式参数通过语义标签按需注入渲染管线"""
from dataclasses import dataclass, field


@dataclass
class HeadingPatternConfig:
    """标题检测规则：正则模式 + 对应标题层级"""
    pattern: str = ""
    level: int = 1


@dataclass
class MarginConfig:
    top_cm: float = 3.8
    bottom_cm: float = 3.8
    left_cm: float = 3.2
    right_cm: float = 3.2


@dataclass
class ColumnConfig:
    """分栏配置——count=1 为单栏（默认），count=2 为双栏"""
    count: int = 1
    space_cm: float = 0.0
    equal_width: bool = True
    separator: bool = False


@dataclass
class PageSetupConfig:
    paper_size: str = "A4"
    margin: MarginConfig = field(default_factory=MarginConfig)
    columns: ColumnConfig = field(default_factory=ColumnConfig)


@dataclass
class StyleConfig:
    font_cn: str = "宋体"
    font_en: str = "Times New Roman"
    size_pt: float = 12.0
    bold: bool = False
    italic: bool = False
    alignment: str = "justify"
    first_line_indent_cm: float = 0.85
    left_indent_cm: float = 0.0
    hanging_indent_cm: float = 0.0
    line_spacing_mode: str = "exact"
    line_spacing_pt: float = 20.0
    space_before_pt: float = 0.0
    space_after_pt: float = 0.0
    widow_control: bool = True
    keep_with_next: bool = False
    keep_together: bool = False
    page_break_before: bool = False


@dataclass
class SceneConfig:
    name: str = "默认格式"
    version: str = "1.0"
    page_setup: PageSetupConfig = field(default_factory=PageSetupConfig)
    styles: dict[str, StyleConfig] = field(default_factory=lambda: {
        "normal": StyleConfig(),
        "heading1": StyleConfig(
            font_cn="黑体", font_en="Times New Roman", size_pt=15.0,
            bold=False, first_line_indent_cm=0.0,
            space_before_pt=40.0, space_after_pt=20.0,
            alignment="center", keep_with_next=True,
            page_break_before=True),
        "heading2": StyleConfig(
            font_cn="黑体", font_en="Times New Roman", size_pt=14.0,
            bold=False, first_line_indent_cm=0.0,
            space_before_pt=24.0, space_after_pt=6.0,
            keep_with_next=True),
        "heading3": StyleConfig(
            font_cn="黑体", font_en="Times New Roman", size_pt=12.0,
            bold=False, first_line_indent_cm=0.0,
            space_before_pt=12.0, space_after_pt=6.0,
            keep_with_next=True),
        "code_block": StyleConfig(
            font_en="Consolas", size_pt=11.0, first_line_indent_cm=0.0,
            line_spacing_pt=20.0, alignment="left"),
        "references_body": StyleConfig(
            size_pt=10.5, first_line_indent_cm=-0.74, left_indent_cm=0.74,
            hanging_indent_cm=0.74, alignment="left"),
        "caption": StyleConfig(
            size_pt=10.5, first_line_indent_cm=0.0, space_before_pt=6.0,
            space_after_pt=6.0, alignment="center"),
    })
    # 标题检测配置——preset 可选注入，不填则使用 TitleDictionary 内置默认值
    front_matter_titles: list = field(default_factory=list)
    numbered_heading_patterns: list = field(default_factory=list)
    auto_heading_numbering: bool = True
    # 表格排版配置
    normal_table_border_mode: str = "three_line"
    table_border_width_pt: float = 0.5
    three_line_header_width_pt: float = 1.5
    three_line_bottom_width_pt: float = 0.75
    normal_table_layout_mode: str = "smart"
    normal_table_smart_levels: int = 4
    normal_table_line_spacing_mode: str = "single"
    normal_table_repeat_header: bool = True
    normal_table_cant_split_rows: bool = True
    table_text_font_cn: str = "宋体"
    table_text_font_en: str = "Times New Roman"
    table_text_size_pt: float = 10.5
    table_text_alignment: str = "left"
    table_number_alignment: str = "center"
