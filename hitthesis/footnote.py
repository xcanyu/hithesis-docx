"""脚注后处理：在正文中插入 w:footnoteReference，生成 word/footnotes.xml。
Word 原生处理双向跳转，无需手动超链接/书签。"""

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

    _process_body_markers(doc_tree, footnotes)
    footnotes_xml = _build_footnotes_xml(footnotes)
    _add_footnote_relationship(rels_tree)
    _ensure_footnote_pr(settings_tree)
    if ct_tree is not None:
        _add_footnote_content_type(ct_tree)

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
    """将正文中的 ① 替换为 w:footnoteReference（Word 原生编号）"""
    circle_map = {CIRCLE[i]: i + 1 for i in range(len(CIRCLE))}

    body = doc_tree.find(f'{{{W}}}body')
    for p in body.iter(f'{{{W}}}p'):
        for r in list(p.findall(f'{{{W}}}r')):
            t = r.find(f'{{{W}}}t')
            if t is None or t.text not in circle_map:
                continue
            fn_id = circle_map[t.text]

            # 插入 w:footnoteReference
            ref_run = etree.Element(f'{{{W}}}r')
            ref_rPr = etree.SubElement(ref_run, f'{{{W}}}rPr')
            ref_style = etree.SubElement(ref_rPr, f'{{{W}}}rStyle')
            ref_style.set(f'{{{W}}}val', 'FootnoteReference')
            ref_sz = etree.SubElement(ref_rPr, f'{{{W}}}sz')
            ref_sz.set(f'{{{W}}}val', '20')
            fn_ref = etree.SubElement(ref_run, f'{{{W}}}footnoteReference')
            fn_ref.set(f'{{{W}}}id', str(fn_id))
            r.addprevious(ref_run)

            # 删除原始 ① run
            p.remove(r)


def _ensure_footnote_pr(settings_tree):
    if settings_tree.find(f'{{{W}}}footnotePr') is None:
        etree.SubElement(settings_tree, f'{{{W}}}footnotePr')


def _build_footnotes_xml(footnotes):
    """生成 word/footnotes.xml"""
    parts = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>']
    parts.append(f'<w:footnotes xmlns:w="{W}" xmlns:r="{R}">')

    parts.append('<w:footnote w:type="separator" w:id="-1">'
                 '<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>'
                 '</w:pPr><w:r><w:separator/></w:r></w:p></w:footnote>')
    parts.append('<w:footnote w:type="continuationSeparator" w:id="0">'
                 '<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>'
                 '</w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>')

    for fn_id, text in footnotes:
        safe_text = text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        circle = CIRCLE[fn_id - 1] if fn_id <= 9 else f'[{fn_id}]'

        parts.append(
            f'<w:footnote w:id="{fn_id}">'
            f'<w:p>'
            f'<w:pPr><w:spacing w:line="270" w:lineRule="auto"/></w:pPr>'
            # 静态上标编号（无 footnoteRef，不生成回跳链接）
            f'<w:r>'
            f'<w:rPr>'
            f'<w:sz w:val="18"/>'
            f'<w:vertAlign w:val="superscript"/>'
            f'<w:rFonts w:eastAsia="宋体" w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
            f'</w:rPr>'
            f'<w:t>{circle}</w:t>'
            f'</w:r>'
            # 脚注文字
            f'<w:r>'
            f'<w:rPr><w:sz w:val="18"/>'
            f'<w:rFonts w:eastAsia="宋体" w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
            f'</w:rPr>'
            f'<w:t xml:space="preserve">　{safe_text}</w:t>'
            f'</w:r>'
            f'</w:p>'
            f'</w:footnote>'
        )

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
