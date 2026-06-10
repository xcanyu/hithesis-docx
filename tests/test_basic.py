"""基础测试：验证核心功能是否正常工作"""

import os
import tempfile
import pytest
from hitthesis import Thesis


@pytest.fixture
def doc():
    """创建测试用的 Thesis 对象"""
    return Thesis(type="bachelor", campus="harbin")


@pytest.fixture
def output_dir():
    """创建临时输出目录"""
    with tempfile.TemporaryDirectory() as tmpdir:
        yield tmpdir


class TestThesisInit:
    """测试 Thesis 初始化"""

    def test_init_bachelor(self):
        doc = Thesis(type="bachelor", campus="harbin")
        assert doc is not None

    def test_init_master(self):
        doc = Thesis(type="master", campus="harbin")
        assert doc is not None

    def test_init_doctor(self):
        doc = Thesis(type="doctor", campus="harbin")
        assert doc is not None


class TestSetInfo:
    """测试 set_info 方法"""

    def test_set_info_basic(self, doc):
        doc.set_info(
            title="测试标题",
            author="测试作者",
            supervisor="测试导师"
        )
        assert doc.info["title"] == "测试标题"
        assert doc.info["author"] == "测试作者"
        assert doc.info["supervisor"] == "测试导师"

    def test_set_info_with_english_title(self, doc):
        doc.set_info(
            title="测试标题",
            english_title="TEST TITLE",
            author="测试作者",
            supervisor="测试导师"
        )
        assert doc.info["english_title"] == "TEST TITLE"


class TestAddParagraph:
    """测试 add_paragraph 方法"""

    def test_add_paragraph_basic(self, doc):
        para = doc.add_paragraph("测试段落")
        assert para is not None

    def test_add_paragraph_with_footnote(self, doc):
        doc.add_paragraph("测试段落", footnote="脚注内容")
        assert len(doc._footnotes) == 1
        assert doc._footnotes[0] == (1, 1, "脚注内容")  # global_id=1, display_num=1

    def test_add_paragraph_with_footnote_number(self, doc):
        doc.add_paragraph("测试段落", footnote="脚注内容", footnote_number=5)
        assert len(doc._footnotes) == 1
        assert doc._footnotes[0] == (1, 5, "脚注内容")  # global_id=1, display_num=5（用户传参）


class TestAddFootnote:
    """测试 add_footnote 方法"""

    def test_add_footnote_default(self, doc):
        doc.add_paragraph("测试段落")
        doc.add_footnote("脚注内容")
        assert len(doc._footnotes) == 1
        assert doc._footnotes[0] == (1, 1, "脚注内容")  # global_id=1, display_num=1

    def test_add_footnote_custom_number(self, doc):
        doc.add_paragraph("测试段落")
        doc.add_footnote("脚注内容", number=3)
        assert len(doc._footnotes) == 1
        assert doc._footnotes[0] == (1, 3, "脚注内容")  # global_id=1, display_num=3（用户传参）

    def test_add_multiple_footnotes(self, doc):
        doc.add_paragraph("测试段落")
        doc.add_footnote("第一条")
        doc.add_footnote("第二条")
        doc.add_footnote("第三条")
        assert len(doc._footnotes) == 3
        assert doc._footnotes[0] == (1, 1, "第一条")
        assert doc._footnotes[1] == (2, 2, "第二条")
        assert doc._footnotes[2] == (3, 3, "第三条")


class TestDocumentGeneration:
    """测试文档生成"""

    def test_compile_basic(self, doc, output_dir):
        doc.set_info(
            title="测试标题",
            author="测试作者",
            supervisor="测试导师"
        )
        output_path = os.path.join(output_dir, "test.docx")
        doc.compile(output_path)
        assert os.path.exists(output_path)
        assert os.path.getsize(output_path) > 0

    def test_compile_with_content(self, doc, output_dir):
        doc.set_info(
            title="测试标题",
            author="测试作者",
            supervisor="测试导师"
        )
        doc.add_cover()
        doc.add_paragraph("测试内容")
        doc.add_footnote("测试脚注")

        output_path = os.path.join(output_dir, "test_with_content.docx")
        doc.compile(output_path)
        assert os.path.exists(output_path)
        assert os.path.getsize(output_path) > 0


class TestXMLStructure:
    """测试 XML 结构"""

    def test_footnote_marker_format(self, doc):
        """测试脚注标记格式是否正确"""
        doc.add_paragraph("测试段落", footnote="脚注内容")
        # 验证脚注列表结构
        assert isinstance(doc._footnotes, list)
        assert len(doc._footnotes) == 1
        fn_id, fn_display, fn_text = doc._footnotes[0]
        assert isinstance(fn_id, int)
        assert isinstance(fn_display, int)
        assert isinstance(fn_text, str)
        assert fn_text == "脚注内容"
