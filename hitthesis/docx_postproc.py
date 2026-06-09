"""docx 整体后处理：直接修改 word/*.xml，处理 python-docx 不支持或全局性的设置。

当前实现：
  - 注入自动断字设置（<w:autoHyphenation/> + <w:hyphenationZone/>），
    让 Word 在英文词边界按音节加连字符换行（类似 LaTeX \\hyphenation）。
  - 参考文献段落显式设置段落语言为 en-US/zh-CN，
    让 Word 知道按英文断字规则处理参考文献里的英文长词。

⚠️ 用户须知：需在 Word 开启"自动断字"（File → Options → Proofing →
   Automatically hyphenate this document when typing）才能看到效果。
"""

import os
import zipfile
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'


def enable_auto_hyphenation(filename, zone_twips=425):
    """在 word/settings.xml 中注入自动断字设置

    Args:
        filename: docx 文件路径
        zone_twips: 断字区宽度（twips，默认 425 ≈ 0.75 inch）。
                    Word 在右页边距前 N twips 的范围内才允许断字。
                    数值越小 → 断字越保守；越大 → 越激进。
    """
    with zipfile.ZipFile(filename, 'r') as z:
        settings_xml = z.read('word/settings.xml')

    parser = etree.XMLParser(remove_blank_text=False)
    tree = etree.fromstring(settings_xml, parser)

    # 移除已有断字设置（幂等）
    for tag in (f'{{{W}}}autoHyphenation', f'{{{W}}}hyphenationZone', f'{{{W}}}suppressAutoHyphens'):
        for old in tree.findall(tag):
            tree.remove(old)

    # 注入：允许自动断字
    auto_hyp = etree.Element(f'{{{W}}}autoHyphenation')
    auto_hyp.set(f'{{{W}}}val', 'true')
    tree.append(auto_hyp)

    # 注入：断字区宽度
    zone = etree.Element(f'{{{W}}}hyphenationZone')
    zone.set(f'{{{W}}}val', str(zone_twips))
    tree.append(zone)

    # 不抑制自动断字
    suppress = etree.Element(f'{{{W}}}suppressAutoHyphens')
    suppress.set(f'{{{W}}}val', 'false')
    tree.append(suppress)

    modified = etree.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True)

    with zipfile.ZipFile(filename, 'r') as zin:
        with zipfile.ZipFile(filename + '.tmp', 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == 'word/settings.xml':
                    zout.writestr(item.filename, modified)
                else:
                    zout.writestr(item, zin.read(item.filename))

    os.replace(filename + '.tmp', filename)


def set_paragraph_lang_in_references(filename, lang_western='en-US', lang_eastasia='zh-CN'):
    """在参考文献节的每个 run 上显式设置语言属性。

    为什么需要这个：settings.xml 里的 autoHyphenation 只是"全局开关"。
    Word 实际断字时还需要知道每个 run 是什么语言——它会用对应语言的断字规则。
    python-docx 生成的 run 通常没有显式 <w:lang>，Word 默认按文档 defaultLanguage，
    但混排中英文时常常"猜错"，导致英文长词不按英文规则断字。

    显式设了 <w:lang w:val="en-US" w:eastAsia="zh-CN"/> 后，
    Word 看到参考文献里的英文词就用英文断字规则处理。

    Args:
        filename: docx 文件路径
        lang_western: 西文语言代码（默认 en-US）
        lang_eastasia: 东亚语言代码（默认 zh-CN）

    Returns:
        int: 修改的 run 数量
    """
    with zipfile.ZipFile(filename, 'r') as z:
        doc_xml = z.read('word/document.xml')

    parser = etree.XMLParser(remove_blank_text=False)
    tree = etree.fromstring(doc_xml, parser)

    modified_count = 0
    in_refs_section = False

    for p in tree.iter(f'{{{W}}}p'):
        # 检查段落是否是"参考文献"标题
        texts = []
        for t in p.iter(f'{{{W}}}t'):
            if t.text:
                texts.append(t.text)
        para_text = ''.join(texts).strip()

        if para_text == '参考文献' or para_text == 'References':
            in_refs_section = True
            continue

        if not in_refs_section:
            continue

        # 给本段每个 run 设置语言
        for r in p.findall(f'{{{W}}}r'):
            rPr = r.find(f'{{{W}}}rPr')
            if rPr is None:
                rPr = etree.SubElement(r, f'{{{W}}}rPr')
                r.insert(0, rPr)

            # 移除已有 lang（幂等）
            for old in rPr.findall(f'{{{W}}}lang'):
                rPr.remove(old)

            # 设置语言
            lang = etree.SubElement(rPr, f'{{{W}}}lang')
            lang.set(f'{{{W}}}val', lang_western)
            lang.set(f'{{{W}}}eastAsia', lang_eastasia)
            modified_count += 1

    if modified_count == 0:
        return 0

    modified_doc = etree.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True)

    with zipfile.ZipFile(filename, 'r') as zin:
        with zipfile.ZipFile(filename + '.tmp', 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == 'word/document.xml':
                    zout.writestr(item.filename, modified_doc)
                else:
                    zout.writestr(item, zin.read(item.filename))

    os.replace(filename + '.tmp', filename)
    return modified_count
