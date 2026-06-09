"""跨章交叉引用测试：ChapterContext 不应清空 cite_registry。"""

import pytest

from hitthesis import Thesis


@pytest.fixture
def doc():
    return Thesis(type="bachelor", campus="harbin")


class TestCrossChapterCitation:
    """跨章引用必须能解析（章2 cite 章1 注册的标签）"""

    def test_cite_registry_persists_across_chapters(self, doc):
        with doc.add_chapter("章1", "1"):
            doc.add_figure(None, caption="图1", ref="fig1")
            assert "fig1" in doc._cite_registry
        with doc.add_chapter("章2", "2"):
            # cite_registry 应保留 fig1
            assert "fig1" in doc._cite_registry
            # 章2 引用 fig1 不应进 pending
            doc.add_paragraph("见 [cite:fig1]")
            assert len(doc._pending_cites) == 0

    def test_cite_table_across_chapters(self, doc):
        with doc.add_chapter("章1", "1"):
            doc.add_table(2, 2, caption="表1", ref="tbl1")
        with doc.add_chapter("章2", "2"):
            doc.add_paragraph("见 [cite:tbl1]")
            assert len(doc._pending_cites) == 0

    def test_unresolved_cite_still_pending(self, doc):
        """未注册的标签仍然进 pending 列表（等用户警告）"""
        with doc.add_chapter("章1", "1"):
            doc.add_paragraph("引用未注册的 [cite:unknown]")
            assert any(e['tag'] == 'unknown' for e in doc._pending_cites)

    def test_appendix_inherits_chapter_cite_registry(self, doc):
        """附录继承正文章节的 cite_registry（支持正文 → 附录的引用）"""
        with doc.add_chapter("章1", "1"):
            doc.add_figure(None, caption="图1", ref="fig1")
        with doc.add_appendix("附录A"):
            # 附录内可引用正文注册的标签
            assert "fig1" in doc._cite_registry
            doc.add_paragraph("附录内引用 [cite:fig1]")
            assert len(doc._pending_cites) == 0
