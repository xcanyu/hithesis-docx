"""
OOXML 原子操作工具集
统一管理所有底层 Word OOXML 操作，消除 document.py 中的重复 XML 模式
"""

from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ============================================================================
# 字体
# ============================================================================

def set_font(run, cn_font, size, bold=False):
    """设置中英文字体：中文用指定字体，英文固定 Times New Roman"""
    run.font.size = Pt(size)
    run.font.bold = bold
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:eastAsia'), cn_font)
    rFonts.set(qn('w:ascii'), "Times New Roman")
    rFonts.set(qn('w:hAnsi'), "Times New Roman")


def estimate_text_width(text, font_size_pt):
    """估算英文文本宽度（pt），按字符类型分别估算宽度"""
    width = 0
    for ch in text:
        if 'A' <= ch <= 'Z' or 'a' <= ch <= 'z':
            width += 0.52 * font_size_pt
        elif ch == ' ':
            width += 0.3 * font_size_pt
        else:
            width += 0.55 * font_size_pt
    return width


# ============================================================================
# 单元格
# ============================================================================

def set_cell_vertical_alignment(cell, alignment="center"):
    """设置单元格垂直对齐方式（center / top / bottom）"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for vAlign in tcPr.findall(qn('w:vAlign')):
        tcPr.remove(vAlign)
    vAlign = OxmlElement('w:vAlign')
    vAlign.set(qn('w:val'), alignment)
    tcPr.append(vAlign)


def clear_cell_margins(cell):
    """清除单元格所有内边距"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    cellMar = tcPr.find(qn('w:tcMar'))
    if cellMar is not None:
        tcPr.remove(cellMar)
    cellMar = OxmlElement('w:tcMar')
    for margin in ('top', 'right', 'bottom', 'left'):
        m = OxmlElement(f'w:{margin}')
        m.set(qn('w:w'), '0')
        m.set(qn('w:type'), 'dxa')
        cellMar.append(m)
    tcPr.append(cellMar)


# ============================================================================
# 边框
# ============================================================================

def make_border_element(tag, val='single', sz=4, space=0, color='000000'):
    """创建单个边框元素"""
    b = OxmlElement(f'w:{tag}')
    b.set(qn('w:val'), val)
    b.set(qn('w:sz'), str(sz))
    b.set(qn('w:space'), str(space))
    b.set(qn('w:color'), color)
    return b


def make_tbl_borders_none():
    """创建无边框的 tblBorders 元素（表格级清除所有边框）"""
    tblBorders = OxmlElement('w:tblBorders')
    for name in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        tblBorders.append(make_border_element(name, 'none', 0, 0, 'auto'))
    return tblBorders


def set_paragraph_border(para, border_position, sz, space, color):
    """设置段落边框（例如下边框）
    border_position: 'top' | 'left' | 'bottom' | 'right'
    sz: 边框粗细（1/8pt 单位，18 = 2.25pt）
    space: 边框与文字的距离（pt）
    """
    pPr = para._element.get_or_add_pPr()
    pBdr = pPr.find(qn('w:pBdr'))
    if pBdr is None:
        pBdr = OxmlElement('w:pBdr')
        pPr.append(pBdr)
    old = pBdr.find(qn(f'w:{border_position}'))
    if old is not None:
        pBdr.remove(old)
    pBdr.append(make_border_element(border_position, 'single', sz, space, color))


def set_cell_bottom_border(cell, sz, color='000000'):
    """给单元格设置仅下边框"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = tcPr.find(qn('w:tcBorders'))
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)
    old = tcBorders.find(qn('w:bottom'))
    if old is not None:
        tcBorders.remove(old)
    tcBorders.append(make_border_element('bottom', 'single', sz, 0, color))


def remove_table_borders(table):
    """移除表格所有边框，替换为 nil"""
    tbl = table._tbl
    tblPr = tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)
    old = tblPr.find(qn('w:tblBorders'))
    if old is not None:
        tblPr.remove(old)
    tblPr.append(make_tbl_borders_none())


def apply_tc_borders(cell, parts: dict):
    """按指定部分设置单元格边框，如 {'top': 12, 'bottom': 8}"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    old = tcPr.find(qn('w:tcBorders'))
    if old is not None:
        tcPr.remove(old)
    tcBorders = OxmlElement('w:tcBorders')
    for position, sz in parts.items():
        tcBorders.append(make_border_element(position, 'single', sz, 0, '000000'))
    tcPr.append(tcBorders)


# ============================================================================
# 段落格式
# ============================================================================

def set_outline_level(para, level):
    """设置段落大纲级别（供 Word 目录识别，0=Heading1, 1=Heading2...）"""
    pPr = para._element.find(qn('w:pPr'))
    if pPr is None:
        pPr = OxmlElement('w:pPr')
        para._element.insert(0, pPr)
    outlineLvl = OxmlElement('w:outlineLvl')
    outlineLvl.set(qn('w:val'), str(level))
    pPr.append(outlineLvl)


def add_page_break_before(para):
    """w:pageBreakBefore — 段落从新页开始"""
    pPr = para._element.get_or_add_pPr()
    pPr.append(OxmlElement('w:pageBreakBefore'))


def apply_heading_style(para, after_pt=None):
    """一级标题样式：居中、段后18pt、1.2倍行距、与下段同页、不对齐到网格
    after_pt: 段后间距，默认取 SPACING["heading_after"]
    """
    from .config import SPACING
    if after_pt is None:
        after_pt = int(SPACING["heading_after"] / 20)
    pPr = para._element.find(qn('w:pPr'))
    if pPr is None:
        pPr = OxmlElement('w:pPr')
        para._element.insert(0, pPr)
    for tag in ['w:jc', 'w:spacing', 'w:keepNext', 'w:keepLines', 'w:snapToGrid']:
        old = pPr.find(qn(tag))
        if old is not None:
            pPr.remove(old)
    jc = OxmlElement('w:jc')
    jc.set(qn('w:val'), 'center')
    pPr.append(jc)
    spacing = OxmlElement('w:spacing')
    spacing.set(qn('w:before'), str(SPACING["heading_before"]))
    spacing.set(qn('w:after'), str(int(after_pt * 20)))
    spacing.set(qn('w:line'), str(SPACING["heading_line"]))
    spacing.set(qn('w:lineRule'), 'auto')
    pPr.append(spacing)
    pPr.append(OxmlElement('w:keepNext'))
    pPr.append(OxmlElement('w:keepLines'))
    sg = OxmlElement('w:snapToGrid')
    sg.set(qn('w:val'), '0')
    pPr.append(sg)


def add_heading_properties(para):
    """为节标题添加 keepNext/keepLines/snapToGrid（轻量版）"""
    pPr = para._element.find(qn('w:pPr'))
    if pPr is None:
        pPr = OxmlElement('w:pPr')
        para._element.insert(0, pPr)
    for tag in ['keepNext', 'keepLines']:
        if pPr.find(qn(f'w:{tag}')) is None:
            pPr.append(OxmlElement(f'w:{tag}'))
    if pPr.find(qn('w:snapToGrid')) is None:
        sg = OxmlElement('w:snapToGrid')
        sg.set(qn('w:val'), '0')
        pPr.append(sg)


def set_first_line_indent(para, twips):
    """直接设置 OOXML w:firstLine 缩进（twips）"""
    pPr = para._element.find(qn('w:pPr'))
    if pPr is None:
        pPr = OxmlElement('w:pPr')
        para._element.insert(0, pPr)
    ind = pPr.find(qn('w:ind'))
    if ind is None:
        ind = OxmlElement('w:ind')
        pPr.append(ind)
    ind.set(qn('w:firstLine'), str(twips))


def set_hanging_indent(para, twips):
    """设置悬挂缩进：首行顶格，续行缩进 twips（2字符=480twips）
    等价于 w:left=twips + w:firstLine=-twips"""
    pPr = para._element.find(qn('w:pPr'))
    if pPr is None:
        pPr = OxmlElement('w:pPr')
        para._element.insert(0, pPr)
    ind = pPr.find(qn('w:ind'))
    if ind is None:
        ind = OxmlElement('w:ind')
        pPr.append(ind)
    ind.set(qn('w:left'), str(twips))
    ind.set(qn('w:firstLine'), str(-twips))


def disable_snap_to_grid(para):
    """禁用段落对齐网格"""
    pPr = para._element.get_or_add_pPr()
    snap = OxmlElement('w:snapToGrid')
    snap.set(qn('w:val'), '0')
    pPr.append(snap)


def add_spacer(doc, height_pt):
    """添加精确高度的空白段落（space_before + snapToGrid=0）"""
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(height_pt)
    para.paragraph_format.space_after = Pt(0)
    para.paragraph_format.line_spacing = Pt(1)
    disable_snap_to_grid(para)
    run = para.add_run('​')
    set_font(run, "宋体", 1)
    return para


# ============================================================================
# 标题段落创建
# ============================================================================

def make_heading_para(doc, text, font_cn='黑体', font_size=18, bold=False):
    """创建一级标题段落（章/附录共用），复用 apply_heading_style"""
    para = doc.add_paragraph()
    add_page_break_before(para)
    apply_heading_style(para)
    set_outline_level(para, 0)
    run = para.add_run(text)
    set_font(run, font_cn, font_size, bold)
    return para


# ============================================================================
# 书签
# ============================================================================

def add_caption_with_bookmark(para, ref, text, font_cn="宋体", font_size=10.5):
    """添加带书签的题注文本（图/表/代码/子图共用）"""
    if ref:
        bm_start = OxmlElement('w:bookmarkStart')
        bm_start.set(qn('w:id'), '0')
        bm_start.set(qn('w:name'), f"cite_{ref}")
        para._element.append(bm_start)
        run = para.add_run(text)
        set_font(run, font_cn, font_size)
        bm_end = OxmlElement('w:bookmarkEnd')
        bm_end.set(qn('w:id'), '0')
        bm_end.set(qn('w:name'), f"cite_{ref}")
        para._element.append(bm_end)
    else:
        run = para.add_run(text)
        set_font(run, font_cn, font_size)


def add_bookmark(para_element, name):
    """在段落 XML 元素中添加书签（起止标记）"""
    bm_start = OxmlElement('w:bookmarkStart')
    bm_start.set(qn('w:id'), '0')
    bm_start.set(qn('w:name'), name)
    para_element.append(bm_start)
    bm_end = OxmlElement('w:bookmarkEnd')
    bm_end.set(qn('w:id'), '0')
    bm_end.set(qn('w:name'), name)
    para_element.append(bm_end)


def add_bookmark_to_paragraph(para, bookmark_name: str):
    """为文档段落添加书签"""
    bookmarkStart = OxmlElement('w:bookmarkStart')
    bookmarkStart.set(qn('w:id'), str(hash(bookmark_name) % 32767))
    bookmarkStart.set(qn('w:name'), bookmark_name)
    para._p.insert(0, bookmarkStart)
    bookmarkEnd = OxmlElement('w:bookmarkEnd')
    bookmarkEnd.set(qn('w:id'), str(hash(bookmark_name) % 32767))
    bookmarkEnd.set(qn('w:name'), bookmark_name)
    para._p.append(bookmarkEnd)


# ============================================================================
# 超链接
# ============================================================================

def add_ref_hyperlink(para, text, anchor):
    """添加参考文献超链接（上标样式）"""
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('w:anchor'), anchor)
    hyperlink.set(qn('w:history'), '1')
    run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:eastAsia'), 'Times New Roman')
    rFonts.set(qn('w:ascii'), 'Times New Roman')
    rFonts.set(qn('w:hAnsi'), 'Times New Roman')
    rPr.append(rFonts)
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), '24')
    rPr.append(sz)
    szCs = OxmlElement('w:szCs')
    szCs.set(qn('w:val'), '24')
    rPr.append(szCs)
    vert = OxmlElement('w:vertAlign')
    vert.set(qn('w:val'), 'superscript')
    rPr.append(vert)
    run.append(rPr)
    t = OxmlElement('w:t')
    t.text = text
    run.append(t)
    hyperlink.append(run)
    para._p.append(hyperlink)


def add_hyperlink_run(para, text, anchor):
    """向段落添加超链接（黑色无下划线，字体与周围一致）"""
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('w:anchor'), anchor)
    hyperlink.set(qn('w:history'), '1')
    run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:eastAsia'), '宋体')
    rFonts.set(qn('w:ascii'), 'Times New Roman')
    rFonts.set(qn('w:hAnsi'), 'Times New Roman')
    rPr.append(rFonts)
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), '24')
    rPr.append(sz)
    color = OxmlElement('w:color')
    color.set(qn('w:val'), '000000')
    rPr.append(color)
    u = OxmlElement('w:u')
    u.set(qn('w:val'), 'none')
    rPr.append(u)
    run.append(rPr)
    t = OxmlElement('w:t')
    t.text = text
    run.append(t)
    hyperlink.append(run)
    para._element.append(hyperlink)


# ============================================================================
# 签名行
# ============================================================================

def add_signature_line(para, label, width_after=114):
    """右对齐签名行：标签 + 间距 + '日期：' + '年' + '月' + '日'"""
    def _seg(text):
        r = para.add_run(text)
        set_font(r, '宋体', 12)
        return r

    def _space(w_pt, n=4):
        run = para.add_run(' ' * n)
        set_font(run, '宋体', 12)
        rPr = run._element.get_or_add_rPr()
        sp = OxmlElement('w:spacing')
        tracking = int((w_pt * 20 - n * 80) / max(n - 1, 1))
        if tracking > 0:
            sp.set(qn('w:val'), str(tracking))
            rPr.append(sp)
        return run

    _seg(label)
    _space(width_after, 6)
    _seg('日期：')
    _space(31.7, 4)
    _seg('年')
    _space(18.875, 3)
    _seg('月')
    _space(18.875, 3)
    _seg('日')


# ============================================================================
# 引用文本渲染
# ============================================================================

def render_normal_text_with_superscripts(para, text: str):
    """渲染普通文本中的 [数字] 上标"""
    import re
    sub_parts = re.split(r'\[(\d+(?:-\d+)?)\]', text)
    for j, sub_part in enumerate(sub_parts):
        if j % 2 == 1 and re.match(r'^\d+(?:-\d+)?$', sub_part):
            run = para.add_run(f'[{sub_part}]')
            run.font.superscript = True
            set_font(run, "Times New Roman", 12)
        elif sub_part:
            run = para.add_run(sub_part)
            set_font(run, "宋体", 12)


def add_ref_runs_merged(para, keys: list, reference_db=None):
    """渲染参考文献引用，合并连续序号后输出上标"""
    if not keys:
        return
    if reference_db is None:
        for key in keys:
            run = para.add_run(f'[ref:{key}]')
            set_font(run, "Times New Roman", 12)
            run.font.superscript = True
        return
    nums = [reference_db.cite(key) for key in keys]
    nums = sorted(set(nums))
    items = []
    i = 0
    while i < len(nums):
        start = nums[i]
        end = start
        j = i + 1
        while j < len(nums) and nums[j] == end + 1:
            end = nums[j]
            j += 1
        length = end - start + 1
        if length >= 3:
            items.append((start, end))
        else:
            for num in range(start, end + 1):
                items.append(num)
        i = j
    left_bracket = para.add_run('[')
    set_font(left_bracket, "Times New Roman", 12)
    left_bracket.font.superscript = True
    for idx, item in enumerate(items):
        if isinstance(item, int):
            add_ref_hyperlink(para, str(item), f'ref_{item}')
        else:
            start, end = item
            add_ref_hyperlink(para, f'{start}-{end}', f'ref_{start}')
        if idx != len(items) - 1:
            comma = para.add_run(',')
            set_font(comma, "Times New Roman", 12)
            comma.font.superscript = True
    right_bracket = para.add_run(']')
    set_font(right_bracket, "Times New Roman", 12)
    right_bracket.font.superscript = True
