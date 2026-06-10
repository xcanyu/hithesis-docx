"""脚注全局递增测试：global_id 始终递增，display_num 可用户自定义。"""

import pytest
from hitthesis import Thesis


@pytest.fixture
def doc():
    return Thesis(type="bachelor", campus="harbin")


class TestFootnoteGlobalIncrement:
    """脚注 global_id 始终递增，不受用户 number 影响"""

    def test_footnotes_global_increment(self, doc):
        """global_id 应始终递增，不因章节切换或用户传参而重置"""
        with doc.add_chapter("章1", "1"):
            doc.add_paragraph("段1", footnote="章1脚注1")
            doc.add_paragraph("段2", footnote="章1脚注2")
            assert len(doc._footnotes) == 2
            assert doc._footnotes[0] == (1, 1, "章1脚注1")
            assert doc._footnotes[1] == (2, 2, "章1脚注2")
        with doc.add_chapter("章2", "2"):
            doc.add_paragraph("段1", footnote="章2脚注1")
            assert doc._footnotes[-1] == (3, 3, "章2脚注1")
            assert len(doc._footnotes) == 3

    def test_footnote_user_number_does_not_affect_counter(self, doc):
        """用户传 number 只影响 display_num，不影响 global_id"""
        doc.add_paragraph("段1", footnote="脚注A")
        assert doc._footnotes[0] == (1, 1, "脚注A")
        doc.add_paragraph("段2", footnote="脚注B", footnote_number=9)
        assert doc._footnotes[1] == (2, 9, "脚注B")  # global_id=2, display_num=9
        # 下一个自动编号应继续从 3 开始
        doc.add_paragraph("段3", footnote="脚注C")
        assert doc._footnotes[2] == (3, 3, "脚注C")  # 不受 number=9 影响
