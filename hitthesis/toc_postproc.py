"""
TOC 后处理
修改 docx 内部 document.xml，修正 Word 生成的目录字体和格式
- TOC1 = 黑体加粗，TOC2/TOC3 = 宋体常规
- 可选在第一章前插入空白行
- 修复本科生 cover2 的 spacer 段落和表格边框
"""

import zipfile
import os
import shutil
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

TOC_CONFIGS = {
    'TOC1': {'font': '黑体', 'bold': False},
    'TOC 1': {'font': '黑体', 'bold': False},
    'TOC 1 Hyperlink': {'font': '黑体', 'bold': False},
    'TOC2': {'font': '宋体', 'bold': False},
    'TOC 2': {'font': '宋体', 'bold': False},
    'TOC 2 Hyperlink': {'font': '宋体', 'bold': False},
    'TOC3': {'font': '宋体', 'bold': False},
    'TOC 3': {'font': '宋体', 'bold': False},
    'TOC 3 Hyperlink': {'font': '宋体', 'bold': False},
}


def _fix_paragraph_runs(para, config, runs_in_para):
    """修正单个 TOC 段落的所有 run 字体"""
    font_cn = config['font']
    is_bold = config['bold']

    # 找 fldChar separate/end 位置（识别页码 run）
    fld_sep_idx = None
    fld_end_idx = None
    for ri, r in enumerate(runs_in_para):
        for fc in r.iter(f'{{{W}}}fldChar'):
            ct = fc.get(f'{{{W}}}fldCharType')
            if ct == 'separate':
                fld_sep_idx = ri
            elif ct == 'end':
                fld_end_idx = ri

    for run in para.iter(f'{{{W}}}r'):
        rPr = run.find(f'{{{W}}}rPr')
        if rPr is None:
            rPr = etree.SubElement(run, f'{{{W}}}rPr')
            run.insert(0, rPr)

        # 判断是否为页码 run
        try:
            run_idx = runs_in_para.index(run)
        except ValueError:
            run_idx = -1
        is_page_run = (
            fld_sep_idx is not None
            and fld_end_idx is not None
            and fld_sep_idx < run_idx < fld_end_idx
        )

        # 中文字体
        rFonts = rPr.find(f'{{{W}}}rFonts')
        if rFonts is None:
            rFonts = etree.SubElement(rPr, f'{{{W}}}rFonts')
        rFonts.set(f'{{{W}}}eastAsia', font_cn)
        rFonts.set(f'{{{W}}}ascii', 'Times New Roman')
        rFonts.set(f'{{{W}}}hAnsi', 'Times New Roman')

        # 字号 12pt = 24
        sz = rPr.find(f'{{{W}}}sz')
        if sz is None:
            sz = etree.SubElement(rPr, f'{{{W}}}sz')
        sz.set(f'{{{W}}}val', '24')
        szCs = rPr.find(f'{{{W}}}szCs')
        if szCs is None:
            szCs = etree.SubElement(rPr, f'{{{W}}}szCs')
        szCs.set(f'{{{W}}}val', '24')

        # 加粗控制（页码和前导符不加粗）
        is_leader_run = bool(list(run.iter(f'{{{W}}}tab')))
        b = rPr.find(f'{{{W}}}b')
        if is_bold and not is_page_run and not is_leader_run:
            if b is None:
                b = etree.SubElement(rPr, f'{{{W}}}b')
            b.set(f'{{{W}}}val', 'true')
        else:
            if b is not None:
                rPr.remove(b)


def _insert_toc_blank_line(tree, toc_blank_line):
    """在第一章前的最后一个 TOC1 条目后插入空白行"""
    if not toc_blank_line:
        return

    paragraphs = list(tree.iter(f'{{{W}}}p'))
    last_toc1_before_chapter = None

    for p in paragraphs:
        pPr = p.find(f'{{{W}}}pPr')
        if pPr is None:
            continue
        pStyle = pPr.find(f'{{{W}}}pStyle')
        if pStyle is None:
            continue
        if pStyle.get(f'{{{W}}}val') not in ('TOC1', 'TOC 1'):
            continue
        text_parts = []
        for t in p.iter(f'{{{W}}}t'):
            if t.text:
                text_parts.append(t.text)
        para_text = ''.join(text_parts).strip()
        if '第1章' in para_text:
            break
        last_toc1_before_chapter = p

    if last_toc1_before_chapter is None:
        print('未找到目录正文开始位置，跳过插入空白行')
        return

    blank_para = etree.Element(f'{{{W}}}p')
    pPr_orig = last_toc1_before_chapter.find(f'{{{W}}}pPr')
    if pPr_orig is not None:
        pPr_new = etree.Element(f'{{{W}}}pPr')
        for elem in pPr_orig.iterchildren():
            pPr_new.append(etree.fromstring(etree.tostring(elem)))
        blank_para.append(pPr_new)

    r_elem = etree.Element(f'{{{W}}}r')
    rPr_elem = etree.Element(f'{{{W}}}rPr')
    vanish = etree.Element(f'{{{W}}}vanish')
    rPr_elem.append(vanish)
    r_elem.append(rPr_elem)
    blank_para.append(r_elem)

    body_elem = last_toc1_before_chapter.getparent()
    actual_idx = list(body_elem).index(last_toc1_before_chapter)
    body_elem.insert(actual_idx + 1, blank_para)
    # print('已在目录正文开始处（第一章前）插入空白行')


def _fix_bachelor_cover2(tree):
    """修复本科生第二封面的 spacer 段落和表格边框（Word COM 可能篡改）"""
    body_elem = tree.find(f'{{{W}}}body')
    if body_elem is None:
        return

    # 修复 spacer 段落
    paragraphs = list(body_elem.iter(f'{{{W}}}p'))
    spacer_para = None
    for i, p in enumerate(paragraphs):
        for br in p.iter(f'{{{W}}}br'):
            if br.get(f'{{{W}}}type') == 'page':
                if i + 1 < len(paragraphs):
                    spacer_para = paragraphs[i + 1]
                break

    if spacer_para is not None:
        pPr = spacer_para.find(f'{{{W}}}pPr')
        if pPr is None:
            pPr = etree.SubElement(spacer_para, f'{{{W}}}pPr')
            spacer_para.insert(0, pPr)
        spacing = pPr.find(f'{{{W}}}spacing')
        if spacing is None:
            spacing = etree.SubElement(pPr, f'{{{W}}}spacing')
        spacer_pt = 505.3 - (3.8 / 2.54 * 72) - 6.4
        spacing.set(f'{{{W}}}line', str(int(spacer_pt * 20)))
        spacing.set(f'{{{W}}}lineRule', 'exact')
        spacing.set(f'{{{W}}}before', '0')
        spacing.set(f'{{{W}}}after', '0')
        snap = pPr.find(f'{{{W}}}snapToGrid')
        if snap is None:
            snap = etree.SubElement(pPr, f'{{{W}}}snapToGrid')
        snap.set(f'{{{W}}}val', '0')

    # 修复第一个表格的 tblInd 和边框
    first_tbl = body_elem.find(f'{{{W}}}tbl')
    if first_tbl is not None:
        tblPr = first_tbl.find(f'{{{W}}}tblPr')
        if tblPr is None:
            tblPr = etree.SubElement(first_tbl, f'{{{W}}}tblPr')
            first_tbl.insert(0, tblPr)
        left_indent_twips = int((171 - 3.0 / 2.54 * 72) * 20)
        tblInd = tblPr.find(f'{{{W}}}tblInd')
        if tblInd is None:
            tblInd = etree.Element(f'{{{W}}}tblInd')
            tblPr.insert(0, tblInd)
        tblInd.set(f'{{{W}}}w', str(left_indent_twips))
        tblInd.set(f'{{{W}}}type', 'dxa')
        tblBorders = tblPr.find(f'{{{W}}}tblBorders')
        if tblBorders is not None:
            tblPr.remove(tblBorders)
        tblBorders = etree.SubElement(tblPr, f'{{{W}}}tblBorders')
        for border_name in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            b = etree.SubElement(tblBorders, f'{{{W}}}{border_name}')
            b.set(f'{{{W}}}val', 'none')
            b.set(f'{{{W}}}sz', '0')
            b.set(f'{{{W}}}space', '0')
            b.set(f'{{{W}}}color', 'auto')


def fix_toc_fonts(filename, thesis_type=None, toc_blank_line=False):
    """后处理：修正 TOC 字体，设置段后间距，可选插空白行和修复 cover2

    Args:
        filename: docx 文件路径
        thesis_type: 论文类型，"bachelor" 时修复 cover2
        toc_blank_line: 目录第一章前是否空一行
    Returns:
        int: 修改的 run 数量
    """
    shutil.copy(filename, filename + '.backup.docx')

    with zipfile.ZipFile(filename, 'r') as z:
        doc_xml = z.read('word/document.xml')

    parser = etree.XMLParser(remove_blank_text=False)
    tree = etree.fromstring(doc_xml, parser)

    modified_count = 0

    for para in tree.iter(f'{{{W}}}p'):
        pPr = para.find(f'{{{W}}}pPr')
        if pPr is None:
            continue
        pStyle = pPr.find(f'{{{W}}}pStyle')
        if pStyle is None:
            continue
        style_id = pStyle.get(f'{{{W}}}val') or ''
        if style_id not in TOC_CONFIGS:
            continue

        config = TOC_CONFIGS[style_id]

        # 段后间距
        old_sp = pPr.find(f'{{{W}}}spacing')
        if old_sp is not None:
            pPr.remove(old_sp)
        spacing = etree.SubElement(pPr, f'{{{W}}}spacing')
        spacing.set(f'{{{W}}}before', '0')
        spacing.set(f'{{{W}}}after', '113')

        runs_in_para = list(para.iter(f'{{{W}}}r'))
        _fix_paragraph_runs(para, config, runs_in_para)
        modified_count += len(runs_in_para)

    # print(f'修改了 {modified_count} 个 run 的字体')

    _insert_toc_blank_line(tree, toc_blank_line)

    if thesis_type == "bachelor":
        _fix_bachelor_cover2(tree)

    modified_xml = etree.tostring(tree, xml_declaration=True,
                                   encoding='UTF-8', standalone=True)

    with zipfile.ZipFile(filename, 'r') as zin:
        with zipfile.ZipFile(filename + '.tmp', 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == 'word/document.xml':
                    zout.writestr(item.filename, modified_xml)
                else:
                    zout.writestr(item.filename, zin.read(item.filename))

    os.replace(filename + '.tmp', filename)
    return modified_count
