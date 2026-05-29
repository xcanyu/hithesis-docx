# LaTeX 公式与化学式渲染实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 在 hithesis-docx 中支持 `$...$` 包裹的 LaTeX 公式和化学式渲染

**Architecture:** 检测文本中的 `$...$` 包裹内容，简单化学式（纯元素+数字+±^_()）直接转上下标，复杂 LaTeX 表达式调用 latex2word 库转换为 OMML 公式对象并插入 Word。

**Tech Stack:** python-docx, latex2word

---

## 文件结构

- 创建: `hitthesis/latex_render.py` — LaTeX 公式解析与渲染核心逻辑
- 修改: `hitthesis/document.py:877` — `add_paragraph()` 在文本渲染时识别 `$...$` 并转换
- 创建: `examples/test_latex.py` — 测试示例

---

## Task 1: 安装 latex2word 依赖

- [ ] **Step 1: 安装 latex2word**

```bash
pip install latex2word
```

Run: `pip show latex2word`
Expected: 显示包信息

- [ ] **Step 2: 验证 latex2word 可导入**

```python
from latex2word import LatexToWordElement
```

Run: `python -c "from latex2word import LatexToWordElement; print('OK')"`
Expected: OK

- [ ] **Step 3: 提交**

```bash
git add requirements.txt
git commit -m "chore: add latex2word dependency"
```

---

## Task 2: 创建 latex_render.py 核心模块

**Files:**
- Create: `hitthesis/latex_render.py`
- Test: `examples/test_latex_render.py`

- [ ] **Step 1: 创建 latex_render.py**

```python
"""
LaTeX 公式与化学式渲染模块
触发方式：文本中用 $...$ 包裹
处理逻辑：
  - 简单化学式（纯元素+数字+±^_()）→ 直接转上下标
  - 复杂 LaTeX 表达式 → 调用 latex2word 转 OMML
"""

import re
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# 简单化学式检测正则：只含元素符号(大写字母可选跟小写)、数字、+-^_()
_SIMPLE_CHEM_PATTERN = re.compile(r'^\$([^$]*)\$$')
_SIMPLE_CHEM_REGEX = re.compile(r'^[A-Za-z0-9+\-^_()]+$')


def _parse_simple_chemical(latex_str: str) -> list:
    """解析简单化学式，返回 [(text, style), ...]
    style: 'normal' | 'superscript' | 'subscript'
    例如 H2O -> [('H', 'normal'), ('2', 'subscript'), ('O', 'normal')]
    """
    result = []
    i = 0
    n = len(latex_str)
    
    while i < n:
        char = latex_str[i]
        
        if char == '^':
            # 上标
            if i + 1 < n:
                next_char = latex_str[i + 1]
                if next_char == '{':
                    # 找匹配的 }
                    end = latex_str.find('}', i + 1)
                    if end != -1:
                        content = latex_str[i + 2:end]
                        result.append((content, 'superscript'))
                        i = end + 1
                        continue
                else:
                    # 单字符上标
                    result.append((next_char, 'superscript'))
                    i += 2
                    continue
            i += 1
            continue
        
        if char == '_':
            # 下标
            if i + 1 < n:
                next_char = latex_str[i + 1]
                if next_char == '{':
                    end = latex_str.find('}', i + 1)
                    if end != -1:
                        content = latex_str[i + 2:end]
                        result.append((content, 'subscript'))
                        i = end + 1
                        continue
                else:
                    result.append((next_char, 'subscript'))
                    i += 2
                    continue
            i += 1
            continue
        
        # 普通字符
        result.append((char, 'normal'))
        i += 1
    
    return result


def is_simple_chemical(latex_str: str) -> bool:
    """判断是否为简单化学式（可用直接转换）"""
    if not latex_str:
        return False
    # 移除允许的字符，检查是否只包含这些字符
    cleaned = re.sub(r'[A-Za-z0-9+\-^_{}()]', '', latex_str)
    return len(cleaned) == 0


def create_omml_from_latex(latex_str: str, doc):
    """使用 latex2word 将 LaTeX 转换为 OMML 公式对象"""
    from latex2word import LatexToWordElement
    
    # 创建临时段落
    para = doc.add_paragraph()
    
    # 使用 latex2word 转换
    latex_converter = LatexToWordElement(latex_str)
    latex_converter.add_latex_to_paragraph(para)
    
    # 获取刚创建的公式元素
    return para


def add_latex_to_paragraph(para, latex_str: str, font_func):
    """向段落添加 LaTeX 公式或化学式
    
    Args:
        para: docx Paragraph 对象
        latex_str: LaTeX 字符串，不含外层 $
        font_func: 设置字体的函数，接收 (run, size, bold)
    """
    # 判断是否为简单化学式
    if is_simple_chemical(latex_str):
        parts = _parse_simple_chemical(latex_str)
        for text, style in parts:
            run = para.add_run(text)
            font_func(run, "Times New Roman", 12)
            if style == 'superscript':
                run.font.superscript = True
            elif style == 'subscript':
                run.font.subscript = True
    else:
        # 复杂 LaTeX，调用 latex2word
        from docx import Document
        # 临时创建文档用于转换
        temp_doc = Document()
        try:
            from latex2word import LatexToWordElement
            converter = LatexToWordElement(latex_str)
            converter.add_latex_to_paragraph(temp_doc.paragraphs[0])
            # 复制公式元素到目标段落
            src_para = temp_doc.paragraphs[0]
            for child in src_para._p:
                # 跳过文本节点，复制公式相关元素
                tag = child.tag
                if 'math' in tag.lower() or 'oMath' in tag:
                    new_elem = OxmlElement(child.tag)
                    new_elem.attrib.update(child.attrib)
                    para._p.append(new_elem)
                elif child.tag.endswith('}r'):  # run 元素
                    new_run = OxmlElement(child.tag)
                    for subchild in child:
                        new_sub = OxmlElement(subchild.tag)
                        new_sub.attrib.update(subchild.attrib)
                        if new_sub.tag.endswith('}t') and subchild.text:
                            new_sub.text = subchild.text
                        new_run.append(new_sub)
                    para._p.append(new_run)
        except Exception as e:
            # 转换失败，回退到纯文本
            run = para.add_run(f"${latex_str}$")
            font_func(run, "Times New Roman", 12)


def split_text_with_latex(text: str) -> list:
    """将文本分割为普通文本和 LaTeX 片段
    
    返回: [('text', '普通文本'), ('latex', 'LaTeX内容'), ...]
    """
    parts = []
    pattern = re.compile(r'\$([^$]+)\$')
    last_end = 0
    
    for m in pattern.finditer(text):
        if m.start() > last_end:
            parts.append(('text', text[last_end:m.start()]))
        parts.append(('latex', m.group(1)))
        last_end = m.end()
    
    if last_end < len(text):
        parts.append(('text', text[last_end:]))
    
    return parts
```

- [ ] **Step 2: 创建测试文件 test_latex_render.py**

```python
"""latex_render 模块测试"""

import sys
sys.path.insert(0, '..')

from hitthesis.latex_render import (
    is_simple_chemical,
    _parse_simple_chemical,
    split_text_with_latex,
)

def test_is_simple_chemical():
    assert is_simple_chemical("H2O") == True
    assert is_simple_chemical("SO4^2-") == True
    assert is_simple_chemical("Fe^{2+}") == True
    assert is_simple_chemical("x^2 + y^2") == False
    assert is_simple_chemical("\\frac{1}{2}") == False
    print("test_is_simple_chemical PASS")

def test_parse_simple_chemical():
    result = _parse_simple_chemical("H2O")
    assert result == [('H', 'normal'), ('2', 'subscript'), ('O', 'normal')]
    
    result = _parse_simple_chemical("SO4^2-")
    assert result == [('S', 'normal'), ('O', 'normal'), ('4', 'subscript'), 
                      ('2', 'superscript'), ('-', 'normal')]
    
    result = _parse_simple_chemical("Fe^{2+}")
    assert result == [('F', 'normal'), ('e', 'normal'), ('2+', 'superscript')]
    
    print("test_parse_simple_chemical PASS")

def test_split_text_with_latex():
    result = split_text_with_latex("水的化学式为 H2O，公式 $x^2$")
    assert result == [
        ('text', '水的化学式为 H2O，公式 '),
        ('latex', 'x^2'),
    ]
    
    result = split_text_with_latex("纯文本")
    assert result == [('text', '纯文本')]
    
    print("test_split_text_with_latex PASS")

if __name__ == "__main__":
    test_is_simple_chemical()
    test_parse_simple_chemical()
    test_split_text_with_latex()
    print("All tests PASS")
```

- [ ] **Step 3: 运行测试**

Run: `python examples/test_latex_render.py`
Expected: All tests PASS

- [ ] **Step 4: 提交**

```bash
git add hitthesis/latex_render.py examples/test_latex_render.py
git commit -m "feat: add latex_render module for formula and chemical notation"
```

---

## Task 3: 集成到 document.py

**Files:**
- Modify: `hitthesis/document.py:184-219` — `_add_text_with_superscripts` 方法
- Modify: `hitthesis/document.py:862-880` — `add_paragraph` 方法

- [ ] **Step 1: 添加导入**

在 `hitthesis/document.py` 顶部导入区域添加：

```python
from .latex_render import split_text_with_latex, add_latex_to_paragraph
```

在现有的 `from .ooxml_utils import ...` 行中添加导入：
```python
from .latex_render import (
    split_text_with_latex,
    add_latex_to_paragraph,
)
```

- [ ] **Step 2: 修改 `_add_text_with_superscripts` 方法**

将现有的 `_add_text_with_superscripts` 方法（约 line 184-219）替换为：

```python
def _add_text_with_superscripts(self, para, text):
    """解析文本，处理上标、交叉引用和 $...$ LaTeX 公式"""
    import re
    ref_keys = []
    last_end = 0
    positions = []
    
    # 收集所有需要特殊处理的位置
    for m in re.finditer(r'\[ref:([a-zA-Z0-9_:.-]+)\]', text):
        positions.append((m.start(), m.end(), 'ref', m.group(1)))
    for m in re.finditer(r'\[cite:([a-zA-Z0-9_:.-]+)\]', text):
        positions.append((m.start(), m.end(), 'cite', m.group(1)))
    for m in re.finditer(r'\[(\d+(?:-\d+)?)\]', text):
        positions.append((m.start(), m.end(), 'num', m.group(1)))
    for m in re.finditer(r'\$([^$]+)\$', text):
        positions.append((m.start(), m.end(), 'latex', m.group(1)))
    
    positions.sort(key=lambda x: (x[0], x[1]))
    
    for start, end, ptype, content2 in positions:
        if ptype == 'ref':
            if start > last_end:
                render_normal_text_with_superscripts(para, text[last_end:start])
            ref_keys.append(content2)
            last_end = end
        else:
            if ref_keys:
                add_ref_runs_merged(para, ref_keys, self._reference_db)
                ref_keys = []
            if start > last_end:
                render_normal_text_with_superscripts(para, text[last_end:start])
            if ptype == 'cite':
                self._add_cite_runs(para, content2)
            elif ptype == 'num':
                run = para.add_run(f'[{content2}]')
                run.font.superscript = True
                set_font(run, "Times New Roman", 12)
            elif ptype == 'latex':
                add_latex_to_paragraph(para, content2, set_font)
            last_end = end
    
    if ref_keys:
        add_ref_runs_merged(para, ref_keys, self._reference_db)
    if last_end < len(text):
        render_normal_text_with_superscripts(para, text[last_end:])
```

- [ ] **Step 3: 运行测试验证修改**

Run: `python examples/test_latex_render.py`
Expected: All tests PASS

- [ ] **Step 4: 提交**

```bash
git add hitthesis/document.py
git commit -m "feat: integrate latex rendering into add_paragraph"
```

---

## Task 4: 创建集成测试示例

**Files:**
- Create: `examples/test_latex_formula.py`

- [ ] **Step 1: 创建测试脚本**

```python
"""测试 LaTeX 公式和化学式渲染"""

import sys
sys.path.insert(0, '..')

from hitthesis import Thesis

doc = Thesis(type="bachelor", campus="harbin")
doc.set_info(
    title="测试 LaTeX 公式渲染",
    author="测试作者",
    supervisor="测试导师",
    subject="测试学科",
    affil="测试学院"
)

doc.add_cover()
doc.start_roman_section()
doc.add_abstract_cn("这是测试摘要内容。", ["测试", "LaTeX"])
doc.add_toc()
doc.start_arabic_section()

with doc.add_chapter("绪论", "1"):
    doc.add_section("化学式测试")
    # 测试简单化学式
    doc.add_paragraph("水的化学式为 $H2O$，二氧化碳为 $CO2$。")
    doc.add_paragraph("硫酸根离子为 $SO4^2-$，铁离子为 $Fe^{2+}$。")
    
    doc.add_section("数学公式测试")
    doc.add_paragraph("根据公式 $x={-b + \\sqrt{b^2-4ac}\\over 2a}$ 计算。")
    doc.add_paragraph("纯文本不含公式的段落。")

doc.compile("output/test_latex_formula.docx")
print("生成成功: output/test_latex_formula.docx")
```

- [ ] **Step 2: 运行测试**

Run: `python examples/test_latex_formula.py`
Expected: 生成 output/test_latex_formula.docx，用 Word/WPS 打开验证

- [ ] **Step 3: 提交**

```bash
git add examples/test_latex_formula.py
git commit -m "test: add latex formula rendering example"
```

---

## Task 5: 验证与文档

- [ ] **Step 1: 用 WPS/Word 打开 output/test_latex_formula.docx 验证**

手动验证：
- 简单化学式 H2O、CO2 是否显示为下标
- 复杂公式是否渲染为可编辑公式

- [ ] **Step 2: 更新 README.md 文档**

在 README.md 的 API 概览部分添加 LaTeX 公式说明：

```markdown
### LaTeX 公式

```python
# 简单化学式（自动上下标）
doc.add_paragraph("水的化学式为 $H2O$")

# 复杂 LaTeX 公式
doc.add_paragraph("根据公式 $x={-b + \\sqrt{b^2-4ac}\\over 2a}$ 计算")
```

触发方式：用 `$...$` 包裹内容。
- 简单化学式（H2O、Fe^{2+}、SO4^2-）直接转上下标
- 复杂 LaTeX 表达式调用 latex2word 转 OMML 公式
```

- [ ] **Step 3: 最终提交**

```bash
git add .
git commit -m "docs: update README with latex formula feature"
```