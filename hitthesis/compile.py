"""
编译流水线：python-docx 生成 .docx 后的后处理步骤。

流程：
  1. 保存 _raw.docx（中间文件，可作为备份）
  2. Word COM 更新域（仅 Windows：展开 TOC、更新页码等）
  3. TOC 字体修复（ZIP 后处理：修正 Word 生成的目录字体为黑体/宋体）
  4. 脚注注入（ZIP 后处理：python-docx 不支持脚注，需手动注入 footnotes.xml）
  5. 自动断字（settings.xml 注入，英文按音节连字符换行）
"""

import os
import shutil
import platform
import time


def update_fields_word(filename):
    """使用 Word COM 自动更新目录、页码等所有域（仅 Windows）

    失败容错：Word COM 在某些环境（pytest、特殊用户配置下）会触发
    0x800706be 等 fatal exception，无法被 Python except 捕获。
    解决方案：用 subprocess 把 Word 操作隔离到独立进程，崩溃不会影响主进程。
    """
    import subprocess
    import sys

    abs_path = os.path.abspath(filename)
    helper = os.path.join(os.path.dirname(__file__), "_word_com_helper.py")

    try:
        result = subprocess.run(
            [sys.executable, helper, abs_path],
            timeout=120,
            capture_output=True,
            text=True,
        )
        if result.returncode != 0:
            err = (result.stderr or result.stdout or "").strip()
            print(f"Word更新域失败 (rc={result.returncode}): {err}")
    except subprocess.TimeoutExpired:
        print("Word更新域超时（>120s），已跳过")
    except Exception as e:
        print(f"Word更新域失败: {e}")


def _write_word_com_helper(path):
    """写出 Word COM 子进程 helper（首次调用时，兼容旧版本）"""
    # 实际 helper 在 hitthesis/_word_com_helper.py，此函数保留仅为向后兼容
    import warnings
    warnings.warn(
        "_write_word_com_helper 已废弃，helper 已迁至 hitthesis/_word_com_helper.py",
        DeprecationWarning,
        stacklevel=2,
    )


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

    # 清理中间文件 _raw.docx
    try:
        os.remove(raw)
    except OSError:
        pass

    # TOC 字体修复
    from .toc_postproc import fix_toc_fonts
    fix_toc_fonts(filename, thesis_type=thesis_type, toc_blank_line=toc_blank_line)

    # 脚注后处理
    if footnotes:
        from .footnote import fix_footnotes
        fix_footnotes(filename, footnotes)

    # 全局后处理：注入自动断字设置（英文按音节连字符换行）
    # 注：用户需在 Word 开启"自动断字"才能看到效果
    from .docx_postproc import enable_auto_hyphenation, set_paragraph_lang_in_references
    enable_auto_hyphenation(filename)
    # 参考文献段落显式设语言为 en-US/zh-CN，让 Word 知道按英文断字规则处理
    set_paragraph_lang_in_references(filename)

    print(f"论文生成完成: {filename}")
    return filename
