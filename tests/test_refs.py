"""参考文献 P 规则测试：验证同一处合并、分别引用独立的行为。"""

import re
import zipfile

from hitthesis.ooxml_utils import group_consecutive_refs, add_ref_runs


class TestGroupConsecutiveRefs:
    """P 规则核心函数：直接相邻 = 同一处"""

    def test_single_ref(self):
        assert group_consecutive_refs("text[ref:A]more") == [["A"]]

    def test_two_consecutive(self):
        """[ref:A][ref:B] → 同一处 → 一组"""
        assert group_consecutive_refs("text[ref:A][ref:B]more") == [["A", "B"]]

    def test_three_consecutive(self):
        assert group_consecutive_refs("[ref:A][ref:B][ref:C]") == [["A", "B", "C"]]

    def test_separated_by_comma(self):
        """[ref:A], [ref:B] → 中间有字符 → 分别独立"""
        assert group_consecutive_refs("[ref:A], [ref:B]") == [["A"], ["B"]]

    def test_separated_by_chinese(self):
        """[ref:A]研究[ref:B] → 中间有汉字 → 分别独立"""
        assert group_consecutive_refs("[ref:A]研究[ref:B]") == [["A"], ["B"]]

    def test_separated_by_space(self):
        """[ref:A] [ref:B] → 中间有空格 → 分别独立（按 P 规则严格）"""
        assert group_consecutive_refs("[ref:A] [ref:B]") == [["A"], ["B"]]

    def test_mixed(self):
        """[ref:A][ref:B], [ref:C][ref:D]研究[ref:E]"""
        result = group_consecutive_refs("[ref:A][ref:B], [ref:C][ref:D]研究[ref:E]")
        assert result == [["A", "B"], ["C", "D"], ["E"]]

    def test_no_refs(self):
        assert group_consecutive_refs("plain text only") == []

    def test_only_text(self):
        assert group_consecutive_refs("") == []

    def test_with_punctuation_in_key(self):
        assert group_consecutive_refs("[ref:key_1][ref:key.2]") == [["key_1", "key.2"]]


class TestAddRefRuns:
    """add_ref_runs 合并逻辑测试"""

    def _make_para(self):
        """轻量段落对象，绕过 python-docx Document"""
        from docx import Document
        return Document().add_paragraph()

    def test_single_key(self):
        """单 key → [5]"""
        para = self._make_para()
        add_ref_runs(para, ["A"], reference_db=None)
        text = "".join(r.text for r in para.runs)
        # reference_db=None 时输出原始占位
        assert text == "[ref:A]"

    def test_empty_keys(self):
        para = self._make_para()
        add_ref_runs(para, [], reference_db=None)
        assert para.text == ""

    def test_consecutive_merge_default(self):
        """连续 keys → 合并为区间 [1-3]（用 fake db 验证输出文本）"""
        from hitthesis.reference_db import ReferenceDB

        class FakeDB:
            def __init__(self):
                self._counter = 0
                self._map = {}
            def cite(self, key):
                if key not in self._map:
                    self._counter += 1
                    self._map[key] = self._counter
                return self._map[key]

        db = FakeDB()
        para = self._make_para()
        add_ref_runs(para, ["A", "B", "C"], reference_db=db)
        # 解析段落中所有 w:t 文本
        text = "".join(t.text or "" for t in para._p.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"))
        # 合并为 1-3
        assert "1-3" in text
        assert "[" in text and "]" in text

    def test_consecutive_with_gap(self):
        """连续但有间隔 → [1,3,5]（不合并）"""
        from hitthesis.reference_db import ReferenceDB

        class FakeDB:
            def __init__(self):
                self._counter = 0
                self._map = {}
            def cite(self, key):
                if key not in self._map:
                    self._counter += 1
                    self._map[key] = self._counter
                return self._map[key]

        db = FakeDB()
        para = self._make_para()
        add_ref_runs(para, ["A", "B", "C"], reference_db=db)
        # A=1, B=2, C=3 → 全连续，合并为 1-3
        text = "".join(t.text or "" for t in para._p.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"))
        assert "1-3" in text

    def test_min_run_length_2_merges_3(self):
        """min_run_length=2 → 长度 > 2 才合并 → 3 个连续合并为 [1-3]"""
        from hitthesis.reference_db import ReferenceDB

        class FakeDB:
            def __init__(self):
                self._counter = 0
                self._map = {}
            def cite(self, key):
                if key not in self._map:
                    self._counter += 1
                    self._map[key] = self._counter
                return self._map[key]

        db = FakeDB()
        para = self._make_para()
        add_ref_runs(para, ["A", "B", "C"], reference_db=db, min_run_length=2)
        text = "".join(t.text or "" for t in para._p.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"))
        # 长度=3 > 2 → 合并为 1-3
        assert "1-3" in text

    def test_min_run_length_2_no_merge_2(self):
        """min_run_length=2 → 长度 = 2 时不合并 → [1,2]"""
        from hitthesis.reference_db import ReferenceDB

        class FakeDB:
            def __init__(self):
                self._counter = 0
                self._map = {}
            def cite(self, key):
                if key not in self._map:
                    self._counter += 1
                    self._map[key] = self._counter
                return self._map[key]

        db = FakeDB()
        para = self._make_para()
        add_ref_runs(para, ["A", "B"], reference_db=db, min_run_length=2)
        text = "".join(t.text or "" for t in para._p.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"))
        # 长度=2 不 > 2 → 不合并 → 应该有 1, 2
        assert "1-2" not in text
        assert "1" in text and "2" in text
