"""脚注后处理：python-docx 不支持脚注，需在编译后通过 ZIP 后处理注入。

机制：
  正文：隐藏的 w:footnoteReference（1pt，Word 内部定位用）+ ① 包裹在 w:hyperlink 中（点击跳脚注）
  脚注区：w:bookmark（跳转目标）+ ①（静态文字）+ 脚注内容
  单向跳转：正文→脚注，脚注不回跳。

已知限制：脚注编号全文连续计数，不按页重置（python-docx 生成阶段无法获知页码）。
"""

import os
import zipfile
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
RELS = 'http://schemas.openxmlformats.org/package/2006/relationships'

CIRCLE = "①②③④⑤⑥⑦⑧⑨"


def fix_footnotes(filename, footnotes):
    if not footnotes:
        return

    with zipfile.ZipFile(filename, 'r') as z:
        doc_xml = z.read('word/document.xml')
        rels_xml = z.read('word/_rels/document.xml.rels')
        settings_xml = z.read('word/settings.xml')
        try:
            ct_xml = z.read('[Content_Types].xml')
        except KeyError:
            ct_xml = None

    parser = etree.XMLParser(remove_blank_text=False)
    doc_tree = etree.fromstring(doc_xml, parser)
    rels_tree = etree.fromstring(rels_xml, parser)
    settings_tree = etree.fromstring(settings_xml, parser)
    ct_tree = etree.fromstring(ct_xml, parser) if ct_xml else None

    # 1. 正文：脚注标记 → w:hyperlink（跳转到脚注）+ w:bookmark（回跳目标）
    _process_body_markers(doc_tree, footnotes)

    # 2. 生成 footnotes.xml（含回跳超链接）
    footnotes_xml = _build_footnotes_xml(footnotes)

    # 3. 修改 rels
    _add_footnote_relationship(rels_tree)

    # 4. 确保 settings.xml 有 footnotePr
    _ensure_footnote_pr(settings_tree)

    # 5. 修改 Content_Types
    if ct_tree is not None:
        _add_footnote_content_type(ct_tree)

    # 6. 写回 ZIP：替换已修改的文件，新文档可能没有 footnotes.xml 需要补写
    modified_doc = etree.tostring(doc_tree, xml_declaration=True, encoding='UTF-8', standalone=True)
    modified_rels = etree.tostring(rels_tree, xml_declaration=True, encoding='UTF-8', standalone=True)
    modified_settings = etree.tostring(settings_tree, xml_declaration=True, encoding='UTF-8', standalone=True)

    footnotes_written = False
    with zipfile.ZipFile(filename, 'r') as zin:
        with zipfile.ZipFile(filename + '.tmp', 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == 'word/document.xml':
                    zout.writestr(item.filename, modified_doc)
                elif item.filename == 'word/_rels/document.xml.rels':
                    zout.writestr(item.filename, modified_rels)
                elif item.filename == 'word/settings.xml':
                    zout.writestr(item.filename, modified_settings)
                elif item.filename == '[Content_Types].xml':
                    modified_ct = etree.tostring(ct_tree, xml_declaration=True, encoding='UTF-8', standalone=True)
                    zout.writestr(item.filename, modified_ct)
                elif item.filename == 'word/footnotes.xml':
                    zout.writestr(item.filename, footnotes_xml)
                    footnotes_written = True
                else:
                    zout.writestr(item, zin.read(item.filename))
            if not footnotes_written:
                zout.writestr('word/footnotes.xml', footnotes_xml)

    os.replace(filename + '.tmp', filename)


def _process_body_markers(doc_tree, footnotes):
    """将正文中的 ① 替换为脚注标记。

    两遍处理：
    第一遍：在 ① 前插入 bookmark（脚注区 hyperlink 的跳转目标）
    第二遍：插入隐藏 footnoteReference + 将 ① 包裹在 hyperlink 中
    分两遍是因为 lxml 的 addprevious 顺序敏感——先插 bm_start 再插 bm_end 才能保证正确顺序。
    """
    circle_map = {CIRCLE[i]: i + 1 for i in range(len(CIRCLE))}
    bookmark_id_counter = [100]  # 从100开始避免与其他书签ID冲突

    body = doc_tree.find(f'{{{W}}}body')
    for p in body.iter(f'{{{W}}}p'):
        for r in p.findall(f'{{{W}}}r'):
            t = r.find(f'{{{W}}}t')
            if t is None or t.text not in circle_map:
                continue
            fn_id = circle_map[t.text]
            r_idx = list(p).index(r)

            # a) 回跳书签（脚注中 hyperlink 的目标）
            bm_id = bookmark_id_counter[0]
            bookmark_id_counter[0] += 2
            bm_start = etree.Element(f'{{{W}}}bookmarkStart')
            bm_start.set(f'{{{W}}}id', str(bm_id))
            bm_start.set(f'{{{W}}}name', f'_ftnref{fn_id}')
            bm_end = etree.Element(f'{{{W}}}bookmarkEnd')
            bm_end.set(f'{{{W}}}id', str(bm_id))
            r.addprevious(bm_start)
            r.addprevious(bm_end)

    # 第二遍：在 ① run 前插入 footnoteReference（脚注定位），将 ① 包裹在 hyperlink 中
    for p in body.iter(f'{{{W}}}p'):
        for r in list(p.findall(f'{{{W}}}r')):
            t = r.find(f'{{{W}}}t')
            if t is None or t.text not in circle_map:
                continue
            fn_id = circle_map[t.text]

            # footnoteReference：Word 内部用的脚注锚点，设为 1pt 隐藏（不可见但不影响跳转）
            ref_run = etree.Element(f'{{{W}}}r')
            ref_rPr = etree.SubElement(ref_run, f'{{{W}}}rPr')
            ref_style = etree.SubElement(ref_rPr, f'{{{W}}}rStyle')
            ref_style.set(f'{{{W}}}val', 'FootnoteReference')
            ref_sz = etree.SubElement(ref_rPr, f'{{{W}}}sz')
            ref_sz.set(f'{{{W}}}val', '2')
            fn_ref = etree.SubElement(ref_run, f'{{{W}}}footnoteReference')
            fn_ref.set(f'{{{W}}}id', str(fn_id))
            r.addprevious(ref_run)

            # 将 ① 包裹在 w:hyperlink 中，anchor 指向脚注区的 bookmark "_ftn{id}"
            hl = etree.Element(f'{{{W}}}hyperlink')
            hl.set(f'{{{W}}}anchor', f'_ftn{fn_id}')
            hl.set(f'{{{W}}}history', '1')
            p.remove(r)
            hl.append(r)
            p.insert(list(p).index(ref_run) + 1, hl)


def _ensure_footnote_pr(settings_tree):
    if settings_tree.find(f'{{{W}}}footnotePr') is None:
        etree.SubElement(settings_tree, f'{{{W}}}footnotePr')


def _build_footnotes_xml(footnotes):
    """生成 word/footnotes.xml（单向跳转：正文→脚注）"""
    parts = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>']
    parts.append(f'<w:footnotes xmlns:w="{W}" xmlns:r="{R}">')

    parts.append('<w:footnote w:type="separator" w:id="-1">'
                 '<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>'
                 '</w:pPr><w:r><w:separator/></w:r></w:p></w:footnote>')
    parts.append('<w:footnote w:type="continuationSeparator" w:id="0">'
                 '<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>'
                 '</w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>')

    bm_id = 200
    for fn_id, text in footnotes:
        safe_text = text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        circle = CIRCLE[fn_id - 1] if fn_id <= 9 else f'[{fn_id}]'

        parts.append(
            f'<w:footnote w:id="{fn_id}">'
            f'<w:p>'
            f'<w:pPr><w:spacing w:line="270" w:lineRule="auto"/></w:pPr>'
            # 书签（正文 hyperlink 的跳转目标）
            f'<w:bookmarkStart w:id="{bm_id}" w:name="_ftn{fn_id}"/>'
            f'<w:bookmarkEnd w:id="{bm_id}"/>'
            # 编号 ①（正常字号，不上标）
            f'<w:r>'
            f'<w:rPr>'
            f'<w:sz w:val="18"/>'
            f'<w:rFonts w:eastAsia="宋体" w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
            f'</w:rPr>'
            f'<w:t>{circle}</w:t>'
            f'</w:r>'
            # 空格分隔
            f'<w:r>'
            f'<w:rPr><w:sz w:val="21"/><w:rFonts w:eastAsia="宋体"/></w:rPr>'
            f'<w:t xml:space="preserve">　</w:t>'
            f'</w:r>'
            # 脚注文字
            f'<w:r>'
            f'<w:rPr>'
            f'<w:sz w:val="18"/>'
            f'<w:rFonts w:eastAsia="宋体" w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
            f'</w:rPr>'
            f'<w:t xml:space="preserve">{safe_text}</w:t>'
            f'</w:r>'
            f'</w:p>'
            f'</w:footnote>'
        )
        bm_id += 2

    parts.append('</w:footnotes>')
    return '\n'.join(parts).encode('utf-8')


def _add_footnote_relationship(rels_tree):
    for rel in rels_tree.findall(f'{{{RELS}}}Relationship'):
        if rel.get('Target') == 'footnotes.xml':
            return
    max_id = 0
    for rel in rels_tree.findall(f'{{{RELS}}}Relationship'):
        rid = rel.get('Id', '')
        if rid.startswith('rId'):
            try:
                max_id = max(max_id, int(rid[3:]))
            except ValueError:
                pass
    new_rel = etree.SubElement(rels_tree, f'{{{RELS}}}Relationship')
    new_rel.set('Id', f'rId{max_id + 1}')
    new_rel.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes')
    new_rel.set('Target', 'footnotes.xml')


def _add_footnote_content_type(ct_tree):
    for ov in ct_tree.findall(f'{{{CT}}}Override'):
        if ov.get('PartName') == '/word/footnotes.xml':
            return
    override = etree.SubElement(ct_tree, f'{{{CT}}}Override')
    override.set('PartName', '/word/footnotes.xml')
    override.set('ContentType',
                 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml')
