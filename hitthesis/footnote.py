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

    # 1. 生成 footnotes.xml
    # 从 document.xml 根元素复制所有命名空间声明，确保与 Word 兼容
    ns_attrs = ' '.join(
        f'{k}="{v}"' for k, v in doc_tree.nsmap.items() if k is not None
    )
    footnotes_xml = _build_footnotes_xml(footnotes, ns_attrs)

    # 2. 修改 rels
    _add_footnote_relationship(rels_tree)

    # 3. 确保 settings.xml 有 footnotePr
    _ensure_footnote_pr(settings_tree)

    # 4. 修改 Content_Types
    if ct_tree is not None:
        _add_footnote_content_type(ct_tree)

    # 5. 写回 ZIP：document.xml 不需要修改（footnoteReference 已在生成时插入）
    modified_rels = etree.tostring(rels_tree, xml_declaration=True, encoding='UTF-8', standalone=True)
    modified_settings = etree.tostring(settings_tree, xml_declaration=True, encoding='UTF-8', standalone=True)

    footnotes_written = False
    with zipfile.ZipFile(filename, 'r') as zin:
        with zipfile.ZipFile(filename + '.tmp', 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == 'word/_rels/document.xml.rels':
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




def _ensure_footnote_pr(settings_tree):
    if settings_tree.find(f'{{{W}}}footnotePr') is None:
        etree.SubElement(settings_tree, f'{{{W}}}footnotePr')


def _build_footnotes_xml(footnotes, ns_attrs=''):
    """生成 word/footnotes.xml（使用 Word 原生脚注机制）"""
    parts = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>']
    parts.append(f'<w:footnotes {ns_attrs} xmlns:w="{W}" xmlns:r="{R}">')

    parts.append('<w:footnote w:type="separator" w:id="-1">'
                 '<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>'
                 '</w:pPr><w:r><w:separator/></w:r></w:p></w:footnote>')
    parts.append('<w:footnote w:type="continuationSeparator" w:id="0">'
                 '<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>'
                 '</w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>')

    bm_id = 200
    for global_id, display_num, text in footnotes:
        safe_text = text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        circle = CIRCLE[display_num - 1] if display_num <= 9 else f'[{display_num}]'

        parts.append(
            f'<w:footnote w:id="{global_id}">'
            f'<w:p>'
            f'<w:pPr>'
            f'<w:ind w:left="400" w:firstLine="-400"/>'
            f'<w:spacing w:line="270" w:lineRule="auto"/>'
            f'</w:pPr>'
            # 书签：正文 hyperlink 的跳转目标
            f'<w:bookmarkStart w:id="{bm_id}" w:name="_ftn{global_id}"/>'
            f'<w:bookmarkEnd w:id="{bm_id}"/>'
            f'<w:r>'
            f'<w:rPr>'
            f'<w:sz w:val="18"/>'
            f'<w:rFonts w:eastAsia="宋体" w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
            f'</w:rPr>'
            f'<w:t>{circle}</w:t>'
            f'</w:r>'
            f'<w:r>'
            f'<w:rPr><w:sz w:val="21"/><w:rFonts w:eastAsia="宋体"/></w:rPr>'
            f'<w:t xml:space="preserve">　</w:t>'
            f'</w:r>'
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
