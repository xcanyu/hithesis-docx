"""
上下文管理器：章节、附录
"""

import re
from docx.shared import Pt
from .ooxml_utils import make_heading_para


class ChapterContext:
    """章节上下文管理器"""

    def __init__(self, thesis, title, number=None, spacing_before=Pt(0)):
        self.thesis = thesis
        self.doc = thesis.doc
        self.title = title
        self.number = number
        self.spacing_before = spacing_before

    def __enter__(self):
        self.thesis.chapter_number = int(self.number) if self.number else self.thesis.chapter_number + 1
        self.thesis.section_counter = [0, 0, 0]
        self.thesis._eq_counter = 0
        self.thesis._tab_counter = 0
        self.thesis._fig_counter = 0
        self.thesis._code_counter = 0
        self.thesis._thm_counter = 0
        # 不重置 cite_registry：跨章引用（如章2 引用 章1 的图）需要保留
        # 也不重置 pending_cites：让跨章 cite 仍能解析
        # 脚注按章重置（每章从 ① 重新开始）
        self.thesis._footnotes = []

        title = self.title
        if len(title) == 2 and re.match(r'^[一-龥]+$', title):
            title = f"{title[0]}　{title[1]}"
        full_title = f"第{self.number}章　{title}" if self.number else title
        para = make_heading_para(self.doc, full_title)
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        return False


class AppendixContext:
    """附录上下文管理器

    附录内的图表/公式编号采用"附X-Y"风格（X = 附录号，Y = 序号），
    与章节内的"X-Y"编号区分。计数器在 __enter__ 重置。
    """

    def __init__(self, thesis, title, number, add_to_toc=True):
        self.thesis = thesis
        self.doc = thesis.doc
        self.title = title
        self.number = number
        self.add_to_toc = add_to_toc

    def __enter__(self):
        # 重置所有章节级计数器，避免延续上一章的编号
        self.thesis.chapter_number = self.number
        self.thesis.section_counter = [0, 0, 0]
        self.thesis._eq_counter = 0
        self.thesis._tab_counter = 0
        self.thesis._fig_counter = 0
        self.thesis._code_counter = 0
        self.thesis._thm_counter = 0
        self.thesis._footnotes = []  # 脚注按章重置
        # 标记当前是附录，编号生成时会加"附"前缀
        self.thesis._in_appendix = True

        title = self.title
        if len(title) == 2 and re.match(r'^[一-龥]+$', title):
            title = f"{title[0]}　{title[1]}"
        full_title = f"附录{self.number}　{title}"
        make_heading_para(self.doc, full_title)
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        # 退出附录：恢复 chapter 模式（下一个章节用 chapter 编号）
        self.thesis._in_appendix = False
        return False
