"""
DOCX 输入分析器
段落角色预判：代码识别、图片检测、标题模式匹配、表格格式化
"""
import re
from docx.shared import Cm
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


# 常见等宽字体（用于识别代码块）
_MONOSPACE_FONTS = {
    'consolas', 'courier', 'courier new', 'source code pro',
    'fira code', 'jetbrains mono', 'monaco', 'menlo', 'dejavu sans mono',
    'lucida console', 'inconsolata', 'cascadia code', 'ubuntu mono',
}

# 中文序号
_CN_NUM_PLAIN = r'(?:[一二三四五六七八九十]{1,3})'



def _looks_like_code(text):
    """启发式判断文本是否像代码行。"""
    stripped = text.strip()
    if not stripped:
        return False
    if re.search(r'[一-鿿]', stripped):
        return False
    code_starters = (
        'def ', 'class ', 'if ', 'for ', 'while ', 'import ', 'from ',
        'var ', 'let ', 'const ', 'function ', 'return ', 'print(',
        'public ', 'private ', 'protected ', 'static ', 'void ',
        'int ', 'string ', 'bool ', 'float ', 'double ',
        'console.', 'System.', 'using ', 'package ', '#include',
    )
    if any(stripped.startswith(kw) for kw in code_starters):
        return True
    if re.search(r'[{}();\[\]]', stripped):
        return True
    return False


def _para_has_image(para):
    """递归检查段落中是否包含图片、图形、图表等。"""
    for child in para._element.iter():
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag in ('drawing', 'pict', 'object', 'inline', 'anchor'):
            return True
    return False


def _format_table(tbl_element, config=None):
    """对深拷贝后的表格 XML 元素统一应用学术三线表规范。"""
    tblPr = tbl_element.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl_element.insert(0, tblPr)

    tblW = tblPr.find(qn('w:tblW'))
    if tblW is None:
        tblW = OxmlElement('w:tblW')
        tblPr.insert(0, tblW)
    tblpPr = tblPr.find(qn('w:tblpPr'))
    if tblpPr is not None:
        tblPr.remove(tblpPr)

    # 清除来源表格的主题样式、交替底纹和自动伸缩，避免 WPS/Word 二次排版。
    for tag in ('w:tblStyle', 'w:tblLook', 'w:tblCellSpacing', 'w:shd'):
        old = tblPr.find(qn(tag))
        if old is not None:
            tblPr.remove(old)

    tblLayout = tblPr.find(qn('w:tblLayout'))
    if tblLayout is None:
        tblLayout = OxmlElement('w:tblLayout')
        tblPr.append(tblLayout)
    tblLayout.set(qn('w:type'), 'fixed')

    jc = tblPr.find(qn('w:jc'))
    if jc is None:
        jc = OxmlElement('w:jc')
        tblPr.append(jc)
    jc.set(qn('w:val'), 'center')

    rows = tbl_element.findall(qn('w:tr'))
    repeat_header = bool(getattr(config, "normal_table_repeat_header", False)) if config else False
    cant_split = bool(getattr(config, "normal_table_cant_split_rows", False)) if config else False
    text_font_cn = getattr(config, "table_text_font_cn", "宋体") if config else "宋体"
    text_font_en = getattr(config, "table_text_font_en", "Times New Roman") if config else "Times New Roman"
    text_size_pt = float(getattr(config, "table_text_size_pt", 10.5) if config else 10.5)
    text_alignment = getattr(config, "table_text_alignment", "left") if config else "left"
    number_alignment = getattr(config, "table_number_alignment", "center") if config else "center"

    def _span(tc):
        tcPr = tc.find(qn('w:tcPr'))
        grid_span = tcPr.find(qn('w:gridSpan')) if tcPr is not None else None
        try:
            return max(1, int(grid_span.get(qn('w:val')))) if grid_span is not None else 1
        except (TypeError, ValueError):
            return 1

    def _logical_col_count(tr):
        return sum(_span(tc) for tc in tr.findall(qn('w:tc')))

    # A4/Letter 页面可用宽度，单位 DXA。三线表使用确定宽度，禁止百分比或 auto。
    paper_width_cm = {"A4": 21.0, "Letter": 21.59, "US Letter": 21.59}.get(
        getattr(getattr(config, "page_setup", None), "paper_size", "A4"), 21.0)
    margin_cfg = getattr(getattr(config, "page_setup", None), "margin", None)
    left_cm = float(getattr(margin_cfg, "left_cm", 2.0))
    right_cm = float(getattr(margin_cfg, "right_cm", 2.0))
    target_width = max(3600, int(round((paper_width_cm - left_cm - right_cm) * 567)))

    col_count = max((_logical_col_count(tr) for tr in rows), default=1)
    old_grid = tbl_element.find(qn('w:tblGrid'))
    source_widths = []
    if old_grid is not None:
        for grid_col in old_grid.findall(qn('w:gridCol')):
            try:
                source_widths.append(max(1, int(grid_col.get(qn('w:w')))))
            except (TypeError, ValueError):
                source_widths.append(1)

    if len(source_widths) != col_count or not any(source_widths):
        # 无可靠网格时按内容估算；叙述列更宽，短数字列保持紧凑。
        source_widths = [4] * col_count
        for tr in rows:
            cursor = 0
            for tc in tr.findall(qn('w:tc')):
                span = _span(tc)
                if span == 1 and cursor < col_count:
                    value = ''.join(t.text or '' for t in tc.iter(qn('w:t'))).strip()
                    visual_len = sum(2 if '\u4e00' <= ch <= '\u9fff' else 1 for ch in value)
                    source_widths[cursor] = max(source_widths[cursor], min(28, visual_len + 2))
                cursor += span

    total_source = max(1, sum(source_widths))
    min_col_width = min(720, max(360, target_width // max(1, col_count * 2)))
    grid_widths = [max(min_col_width, int(round(target_width * w / total_source)))
                   for w in source_widths]
    delta = target_width - sum(grid_widths)
    grid_widths[max(range(len(grid_widths)), key=lambda i: grid_widths[i])] += delta

    tblW.set(qn('w:w'), str(target_width))
    tblW.set(qn('w:type'), 'dxa')

    tblInd = tblPr.find(qn('w:tblInd'))
    if tblInd is None:
        tblInd = OxmlElement('w:tblInd')
        tblPr.append(tblInd)
    tblInd.set(qn('w:w'), '0')
    tblInd.set(qn('w:type'), 'dxa')

    new_grid = OxmlElement('w:tblGrid')
    for width in grid_widths:
        grid_col = OxmlElement('w:gridCol')
        grid_col.set(qn('w:w'), str(width))
        new_grid.append(grid_col)
    if old_grid is not None:
        tbl_element.replace(old_grid, new_grid)
    else:
        tbl_element.insert(1, new_grid)

    def _ensure(parent, tag):
        el = parent.find(qn(tag))
        if el is None:
            el = OxmlElement(tag)
            parent.append(el)
        return el

    def _set_on_off(parent, tag, enabled=True):
        el = parent.find(qn(tag))
        if el is None:
            el = OxmlElement(tag)
            parent.append(el)
        if enabled:
            el.attrib.pop(qn('w:val'), None)
        else:
            el.set(qn('w:val'), '0')

    def _cell_text(tc):
        return ''.join(t.text or '' for t in tc.iter(qn('w:t'))).strip()

    def _looks_numeric(value):
        value = value.strip().replace(',', '')
        return bool(re.fullmatch(r'[-+]?(\d+(\.\d+)?|\.\d+)(%|‰)?', value))

    def _jc_val(value):
        return {"left": "left", "center": "center", "right": "right",
                "justify": "both"}.get(str(value or "").lower(), "left")

    for row_idx, tr in enumerate(rows):
        trPr = tr.find(qn('w:trPr'))
        if trPr is None:
            trPr = OxmlElement('w:trPr')
            tr.insert(0, trPr)
        row_height = trPr.find(qn('w:trHeight'))
        if row_height is not None:
            trPr.remove(row_height)
        if cant_split:
            _set_on_off(trPr, 'w:cantSplit', True)
        if repeat_header and row_idx == 0:
            _set_on_off(trPr, 'w:tblHeader', True)

    for row_idx, tr in enumerate(rows):
        col_cursor = 0
        for tc in tr.findall(qn('w:tc')):
            span = _span(tc)
            cell_width = sum(grid_widths[col_cursor:min(col_count, col_cursor + span)])
            col_cursor += span
            is_header = row_idx == 0
            is_numeric_cell = _looks_numeric(_cell_text(tc))
            tcPr = tc.find(qn('w:tcPr'))
            if tcPr is None:
                tcPr = OxmlElement('w:tcPr')
                tc.insert(0, tcPr)

            tcW = tcPr.find(qn('w:tcW'))
            if tcW is None:
                tcW = OxmlElement('w:tcW')
                tcPr.insert(0, tcW)
            tcW.set(qn('w:w'), str(cell_width))
            tcW.set(qn('w:type'), 'dxa')

            for tag in ('w:shd', 'w:noWrap'):
                old = tcPr.find(qn(tag))
                if old is not None:
                    tcPr.remove(old)

            vAlign = tcPr.find(qn('w:vAlign'))
            if vAlign is None:
                vAlign = OxmlElement('w:vAlign')
                tcPr.append(vAlign)
            vAlign.set(qn('w:val'), 'center')

            tcMar = tcPr.find(qn('w:tcMar'))
            if tcMar is not None:
                tcPr.remove(tcMar)
            tcMar = OxmlElement('w:tcMar')
            for side, width in (('top', 90), ('left', 108), ('bottom', 90), ('right', 108)):
                mar = OxmlElement(f'w:{side}')
                mar.set(qn('w:w'), str(width))
                mar.set(qn('w:type'), 'dxa')
                tcMar.append(mar)
            tcPr.append(tcMar)

            for p in tc.iter(qn('w:p')):
                pPr = p.find(qn('w:pPr'))
                if pPr is None:
                    pPr = OxmlElement('w:pPr')
                    p.insert(0, pPr)
                jc = pPr.find(qn('w:jc'))
                if jc is None:
                    jc = OxmlElement('w:jc')
                    pPr.append(jc)
                jc.set(qn('w:val'), _jc_val('center' if is_header else
                                           (number_alignment if is_numeric_cell else text_alignment)))

                ind = pPr.find(qn('w:ind'))
                if ind is not None:
                    pPr.remove(ind)
                spacing = pPr.find(qn('w:spacing'))
                if spacing is None:
                    spacing = OxmlElement('w:spacing')
                    pPr.append(spacing)
                spacing.set(qn('w:before'), '0')
                spacing.set(qn('w:after'), '0')
                spacing.set(qn('w:line'), '240')
                spacing.set(qn('w:lineRule'), 'auto')

                for r in p.iter(qn('w:r')):
                    rPr = r.find(qn('w:rPr'))
                    if rPr is None:
                        rPr = OxmlElement('w:rPr')
                        r.insert(0, rPr)
                    rFonts = rPr.find(qn('w:rFonts'))
                    if rFonts is None:
                        rFonts = OxmlElement('w:rFonts')
                        rPr.insert(0, rFonts)
                    rFonts.set(qn('w:ascii'), text_font_en)
                    rFonts.set(qn('w:hAnsi'), text_font_en)
                    rFonts.set(qn('w:eastAsia'), text_font_cn)
                    color = rPr.find(qn('w:color'))
                    if color is None:
                        color = OxmlElement('w:color')
                        rPr.append(color)
                    color.set(qn('w:val'), '000000')
                    for tag in ('w:highlight', 'w:shd'):
                        old = rPr.find(qn(tag))
                        if old is not None:
                            rPr.remove(old)
                    bold = rPr.find(qn('w:b'))
                    if is_header:
                        if bold is None:
                            bold = OxmlElement('w:b')
                            rPr.append(bold)
                        bold.attrib.pop(qn('w:val'), None)
                    for tag_name in ('w:sz', 'w:szCs'):
                        sz = rPr.find(qn(tag_name))
                        if sz is None:
                            sz = OxmlElement(tag_name)
                            rPr.append(sz)
                        sz.set(qn('w:val'), str(int(round(text_size_pt * 2))))


# ── Word 自动编号解析 ────────────────────────────────────────────

def _build_numbering_maps(doc):
    num_maps = {"abstract": {}, "instance": {}}
    try:
        numbering_part = doc.part.numbering_part
        if numbering_part is None:
            return num_maps
        root = numbering_part.element
    except (AttributeError, KeyError, NotImplementedError):
        return num_maps

    for abs_num in root.findall(qn('w:abstractNum')):
        abs_str = abs_num.get(qn('w:abstractNumId'))
        if abs_str is None:
            continue
        abs_id = int(abs_str)
        levels = {}
        for lvl in abs_num.findall(qn('w:lvl')):
            ilvl_str = lvl.get(qn('w:ilvl'))
            if ilvl_str is None:
                continue
            ilvl = int(ilvl_str)
            numFmt_el = lvl.find(qn('w:numFmt'))
            lvlText_el = lvl.find(qn('w:lvlText'))
            levels[ilvl] = {
                "numFmt": numFmt_el.get(qn('w:val')) if numFmt_el is not None else "decimal",
                "lvlText": lvlText_el.get(qn('w:val')) if lvlText_el is not None else "",
            }
        num_maps["abstract"][abs_id] = levels

    for num in root.findall(qn('w:num')):
        num_id_str = num.get(qn('w:numId'))
        if num_id_str is None:
            continue
        num_id = int(num_id_str)
        abs_ref = num.find(qn('w:abstractNumId'))
        if abs_ref is None:
            continue
        abs_id = int(abs_ref.get(qn('w:val')))
        overrides = {}
        for ovr in num.findall(qn('w:lvlOverride')):
            ilvl_str = ovr.get(qn('w:ilvl'))
            if ilvl_str is None:
                continue
            ilvl_ov = int(ilvl_str)
            lvl_el = ovr.find(qn('w:lvl'))
            if lvl_el is not None:
                nf = lvl_el.find(qn('w:numFmt'))
                lt = lvl_el.find(qn('w:lvlText'))
                overrides[ilvl_ov] = {
                    "numFmt": nf.get(qn('w:val')) if nf is not None else None,
                    "lvlText": lt.get(qn('w:val')) if lt is not None else None,
                }
        num_maps["instance"][num_id] = {"abstractNumId": abs_id, "overrides": overrides}

    return num_maps


def _container_numpr(para):
    try:
        pPr = para._element.find(qn('w:pPr'))
        if pPr is None:
            return None
        numPr = pPr.find(qn('w:numPr'))
        if numPr is None:
            return None
        numId_el = numPr.find(qn('w:numId'))
        ilvl_el = numPr.find(qn('w:ilvl'))
        numId = int(numId_el.get(qn('w:val'))) if numId_el is not None else None
        ilvl = int(ilvl_el.get(qn('w:val'))) if ilvl_el is not None else None
        if numId is not None and ilvl is not None:
            return (numId, ilvl)
    except Exception:
        pass
    return None


def _find_numbering_lvl(num_maps, num_id, ilvl):
    if not num_maps or num_id not in num_maps["instance"]:
        return None
    inst = num_maps["instance"][num_id]
    if ilvl in inst["overrides"]:
        ov = inst["overrides"][ilvl]
        if ov.get("numFmt") and ov.get("lvlText"):
            return ov
    abs_id = inst["abstractNumId"]
    if abs_id in num_maps["abstract"] and ilvl in num_maps["abstract"][abs_id]:
        base = dict(num_maps["abstract"][abs_id][ilvl])
        if ilvl in inst["overrides"]:
            ov = inst["overrides"][ilvl]
            if ov.get("numFmt"):
                base["numFmt"] = ov["numFmt"]
            if ov.get("lvlText"):
                base["lvlText"] = ov["lvlText"]
        return base
    return None
