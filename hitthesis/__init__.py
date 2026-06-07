"""
hitthesis-docx: 哈工大学位论文Word文档生成工具

使用示例:
    from hitthesis import Thesis

    doc = Thesis(type="doctor", campus="harbin")
    doc.set_info(
        title="局部多孔质气体静压轴承关键技术的研究",
        author="于冬梅",
        supervisor="某某某教授",
        subject="机械制造及其自动化",
        affil="机电工程学院"
    )
    doc.add_cover()
    doc.add_abstract_cn("摘要内容...", ["关键词1", "关键词2"])
    doc.add_toc()

    with doc.add_chapter("绪论", "1"):
        doc.add_section("1.1 研究背景", 1)
        doc.add_paragraph("研究背景内容...")

    doc.compile("thesis.docx")
"""

from .document import (
    Thesis,
    ChapterContext,
    Figure,
    Table,
)
from .ooxml_utils import set_font
from .reference_db import ReferenceDB

__version__ = "0.1.0"
__all__ = [
    "Thesis",
    "ChapterContext",
    "Figure",
    "Table",
    "set_font",
    "ReferenceDB",
]
