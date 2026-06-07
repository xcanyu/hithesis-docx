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
        self.thesis._cite_registry = {}
        self.thesis._pending_cites = []

        title = self.title
        if len(title) == 2 and re.match(r'^[一-龥]+$', title):
            title = f"{title[0]}　{title[1]}"
        full_title = f"第{self.number}章　{title}" if self.number else title
        para = make_heading_para(self.doc, full_title)
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        return False


class AppendixContext:
    """附录上下文管理器"""

    def __init__(self, thesis, title, number, add_to_toc=True):
        self.thesis = thesis
        self.doc = thesis.doc
        self.title = title
        self.number = number
        self.add_to_toc = add_to_toc

    def __enter__(self):
        title = self.title
        if len(title) == 2 and re.match(r'^[一-龥]+$', title):
            title = f"{title[0]}　{title[1]}"
        full_title = f"附录{self.number}　{title}"
        make_heading_para(self.doc, full_title)
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        return False
