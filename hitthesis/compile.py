"""
编译流水线：python-docx 生成 .docx 后的后处理步骤。

流程：
  1. 保存 _raw.docx（中间文件，可作为备份）
  2. Word COM 更新域（仅 Windows：展开 TOC、更新页码等）
  3. TOC 字体修复（ZIP 后处理：修正 Word 生成的目录字体为黑体/宋体）
  4. 脚注注入（ZIP 后处理：python-docx 不支持脚注，需手动注入 footnotes.xml）
"""

import os
import shutil
import platform
import time


def update_fields_word(filename):
    """使用 Word COM 自动更新目录、页码等所有域（仅 Windows）"""
    import win32com.client

    abs_path = os.path.abspath(filename)
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0

    try:
        doc = word.Documents.Open(abs_path)
        doc.Fields.Update()
        for toc in doc.TablesOfContents:
            toc.Update()
        doc.Repaginate()
        doc.Fields.Update()
        doc.Save()
        time.sleep(0.5)
        doc.Close(False)
    finally:
        word.Quit()


def compile_document(doc, filename, thesis_type=None, toc_blank_line=False, footnotes=None):
    """编译生成论文文档

    流程：保存原始 → Word COM 更新域（Windows）→ 覆盖 → TOC 字体修复 → 脚注后处理

    Args:
        doc: python-docx Document 对象
        filename: 输出路径
        thesis_type: 论文类型（用于 cover2 修复），如 "bachelor"
        toc_blank_line: 目录第一章前是否空一行
        footnotes: [(id, text), ...] 脚注列表
    Returns:
        str: 最终文件路径
    """
    raw = filename.replace(".docx", "_raw.docx")
    doc.save(raw)

    if platform.system() == "Windows":
        try:
            update_fields_word(raw)
        except Exception as e:
            print(f"Word更新域失败: {e}")
        shutil.copy(raw, filename)
    else:
        print("提示：macOS/Linux 下需手动更新目录。打开文档后按 Ctrl+A → F9 → 更新整个目录")
        shutil.copy(raw, filename)

    # TOC 字体修复
    from .toc_postproc import fix_toc_fonts
    fix_toc_fonts(filename, thesis_type=thesis_type, toc_blank_line=toc_blank_line)

    # 脚注后处理
    if footnotes:
        from .footnote import fix_footnotes
        fix_footnotes(filename, footnotes)

    print(f"论文生成完成: {filename}")
    return filename
