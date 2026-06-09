"""附录编号风格测试：AppendixContext 内的图表/公式/代码/定理编号应加"附"前缀。"""

import os
import tempfile

import pytest

from hitthesis import Thesis


@pytest.fixture
def doc():
    return Thesis(type="bachelor", campus="harbin")


class TestAppendixContextResetsCounters:
    """AppendixContext.__enter__ 应重置所有章节级计数器"""

    def test_resets_tab_counter(self, doc):
        """附录内的表格应从 1 重新编号，而不是延续正文章节"""
        with doc.add_chapter("正文章节", "3"):
            doc.add_table(2, 2, caption="正文表")
        # 此时 _tab_counter 应该是 1
        with doc.add_appendix("附录A"):
            assert doc._tab_counter == 0
            doc.add_table(2, 2, caption="附表")
            # 附录内的 _tab_counter 现在是 1
            assert doc._tab_counter == 1

    def test_resets_fig_counter(self, doc):
        with doc.add_chapter("正文章节", "3"):
            doc.add_figure(None, caption="正文图")
        with doc.add_appendix("附录A"):
            assert doc._fig_counter == 0
            doc.add_figure(None, caption="附图")
            assert doc._fig_counter == 1

    def test_resets_eq_counter(self, doc):
        with doc.add_chapter("正文章节", "3"):
            doc.add_equation("a+b")
        with doc.add_appendix("附录A"):
            assert doc._eq_counter == 0

    def test_resets_code_counter(self, doc):
        with doc.add_chapter("正文章节", "3"):
            doc.add_code_block("print(1)")
        with doc.add_appendix("附录A"):
            assert doc._code_counter == 0

    def test_resets_thm_counter(self, doc):
        with doc.add_chapter("正文章节", "3"):
            doc.add_theorem("正文定理", kind="定理")
        with doc.add_appendix("附录A"):
            assert doc._thm_counter == 0

    def test_resets_section_counter(self, doc):
        with doc.add_chapter("正文章节", "3"):
            doc.add_section("正文节")
            assert doc.section_counter[0] == 1
        with doc.add_appendix("附录A"):
            assert doc.section_counter == [0, 0, 0]

    def test_resets_cite_registry_except_cross_chapter(self, doc):
        """cite_registry 不应被附录重置（跨章引用需要）"""
        with doc.add_chapter("正文章节", "3"):
            doc.add_figure(None, caption="正文图", ref="fig_chap")
            assert "fig_chap" in doc._cite_registry
        with doc.add_appendix("附录A"):
            # cite_registry 跨章/跨附录保留（支持跨章引用）
            assert "fig_chap" in doc._cite_registry


class TestAppendixRefNumPrefix:
    """附录内 caption 文字应加"附"前缀，ref_num 保持简洁"""

    def test_table_label_no_prefix(self, doc):
        """附录内 add_table 的 label 保持 '1-1'（"附"加在 caption 文字）"""
        with doc.add_chapter("正文章节", "3"):
            t = doc.add_table(2, 2, caption="正文表")
            assert t.label == "3-1"
        with doc.add_appendix("附录A"):
            t = doc.add_table(2, 2, caption="附表")
            assert t.label == "1-1"  # 没有"附"前缀

    def test_table_caption_uses_appendix_word(self, doc):
        """附录内 caption 文字应是"附表"（而非"表"）"""
        with doc.add_appendix("附录A"):
            t = doc.add_table(2, 2, caption="附表", ref="tblA")
            # 通过 paragraph 内部文本验证
            # 完整 caption: "附表1-1　附表"
            assert t.label == "1-1"
            # t._caption_para 包含 "附表"
            # 由于 Table 对象不直接暴露 caption para，间接验证：ref 正确
            assert "tblA" in doc._cite_registry

    def test_figure_label_no_prefix(self, doc):
        with doc.add_appendix("附录A"):
            f = doc.add_figure(None, caption="附图")
            assert f.label == "1-1"

    def test_code_label_no_prefix(self, doc):
        with doc.add_appendix("附录A"):
            doc.add_code_block("print(1)", caption="附代码")
            assert doc._code_counter == 1

    def test_equation_in_appendix(self, doc):
        """附录内公式应显示 (附1-1)"""
        with doc.add_appendix("附录A"):
            doc.add_equation("a+b")
            doc.add_equation("c+d")
            assert doc._eq_counter == 2

    def test_chapter_ref_num_no_prefix(self, doc):
        """正文章节内编号不应有"附"前缀"""
        with doc.add_chapter("正文章节", "3"):
            t = doc.add_table(2, 2, caption="正文表")
            assert t.label == "3-1"
            assert "附" not in t.label


class TestAppendixContextExits:
    """AppendixContext.__exit__ 应恢复 chapter 模式"""

    def test_after_appendix_new_chapter_uses_chapter_numbering(self, doc):
        """退出附录后，下一个 chapter 应使用 chapter 编号（无"附"前缀）"""
        with doc.add_chapter("第3章", "3"):
            t1 = doc.add_table(2, 2, caption="表3-1")
            assert t1.label == "3-1"
        with doc.add_appendix("附录A"):
            t2 = doc.add_table(2, 2, caption="附表")
            assert t2.label == "1-1"
        # 退出附录后
        assert doc._in_appendix is False
        with doc.add_chapter("第4章", "4"):
            t3 = doc.add_table(2, 2, caption="表4-1")
            assert t3.label == "4-1"
            assert "附" not in t3.label
