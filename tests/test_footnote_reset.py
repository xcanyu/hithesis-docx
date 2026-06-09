"""脚注按章重置测试：ChapterContext.__enter__ 应清空 self._footnotes。"""

import pytest

from hitthesis import Thesis


@pytest.fixture
def doc():
    return Thesis(type="bachelor", campus="harbin")


class TestFootnotePerChapterReset:
    """每章脚注应从 ① 重新开始"""

    def test_footnotes_reset_in_chapter_context(self, doc):
        with doc.add_chapter("章1", "1"):
            doc.add_paragraph("段1", footnote="章1脚注1")
            doc.add_paragraph("段2", footnote="章1脚注2")
            assert len(doc._footnotes) == 2
        with doc.add_chapter("章2", "2"):
            # 进入新章，脚注应被清空
            assert doc._footnotes == []
            doc.add_paragraph("段1", footnote="章2脚注1")
            # 章2 的脚注序号应从 1 开始
            assert doc._footnotes[0] == (1, "章2脚注1")

    def test_appendix_also_resets(self, doc):
        """附录也应重置脚注"""
        with doc.add_chapter("章1", "1"):
            doc.add_paragraph("段1", footnote="章1脚注")
            assert len(doc._footnotes) == 1
        with doc.add_appendix("附录A"):
            assert doc._footnotes == []
            doc.add_paragraph("附段1", footnote="附脚注")
            assert doc._footnotes[0] == (1, "附脚注")
