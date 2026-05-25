# \[赤石科技\] 基于 python-docx 的 hithesis-docx 模板

[![Python](https://img.shields.io/badge/python-3.9+-blue.svg)](https://www.python.org/downloads/)
[![License](https://img.shields.io/badge/license-CC%20BY--NC%204.0-blue.svg)](LICENSE)

用 Python 代码生成哈尔滨工业大学学位论文 Word 文档（.docx）。

> **建议：正式学位论文写作请使用 [hithesis](https://github.com/hithesis/hithesis) LaTeX 模板编译。** 本项目适用于对格式要求不严的 `.docx` 文档生成场景，如草稿、非正式提交、格式审查前的初稿等。

---

> **代码生成声明**
>
> 本项目全部代码由基于 **Minimax M2.7** / **Deepseek V4 系列** 的 **Claude Code** 生成，人工仅做需求描述和效果验证。

> **重要声明**
>
> 本项目格式参考 [hithesis/hithesis](https://github.com/hithesis/hithesis)（哈工大 LaTeX 学位论文模板），在此基础上做了**高仿**，**并非严格符合**《哈尔滨工业大学学位论文撰写规范》。因个人精力有限，目前仅实现了**本科、本部（哈尔滨）**的论文模板，硕博及其他校区未做，仅预留了 `type` 和 `campus` 参数空间。**不建议用于正式学位论文提交**，仅供学习交流。欢迎有兴趣的朋友自行扩展。

## 特性

- **参考规范**：格式参考 [hithesis](https://github.com/hithesis/hithesis) LaTeX 模板——页面边距、字体字号、行距缩进、三线表、双线页眉
- **代码驱动**：用 Python 描述文档结构，版本可控、可复用
- **学位类型**：当前仅实现本科/本部，硕博及其他校区预留参数空间

## 安装

```bash
pip install -r requirements.txt
```

> **Windows 用户**：编译阶段会自动调用 Word COM 接口更新目录域，无需额外配置。macOS / Linux 用户需手动更新目录（打开文档后 `Ctrl+A` → `F9`）。

## 快速开始

```python
from hitthesis import Thesis

doc = Thesis(type="bachelor", campus="harbin", header_text=False)  # False=无页眉；省略=默认"哈尔滨工业大学"
doc.set_info(
    title="局部多孔质气体静压轴承关键技术的研究",
    author="于冬梅",
    supervisor="某某某教授",
    subject="机械制造及其自动化",
    affil="机电工程学院",
    date="2024年6月",
    student_id="1234567890",
)

doc.add_cover()
doc.start_roman_section()

doc.add_abstract_cn(
    "本文针对局部多孔质气体静压轴承关键技术进行了深入研究……",
    ["气体轴承", "静压轴承", "多孔质材料", "有限元分析", "实验研究"],
)
doc.add_abstract_en(
    "This dissertation presents an in-depth investigation……",
    ["gas bearing", "externally pressurized bearing", "porous material"],
)
doc.add_toc()

doc.start_arabic_section()

with doc.add_chapter("绪论", "1"):
    doc.add_section("研究背景与意义")
    doc.add_paragraph("气体轴承是一种利用气体膜支承载荷的滑动轴承……")

doc.add_conclusion(content=["本文针对……", "（1）建立了……", "（2）通过……"])
doc.add_references(bib)
doc.add_acknowledgements(content=["感谢我的导师……"])
doc.add_authorization()

doc.compile("thesis.docx")
```

详见 [`examples/thesis_example.py`](examples/thesis_example.py)。

## API 概览

### 论文信息

```python
doc = Thesis(type="bachelor", campus="harbin", header_text=False)
# type: 当前仅实现 "bachelor"（预留 "master"|"doctor"）
# campus: 当前仅实现 "harbin"（预留 "shenzhen"|"weihai"）
# header_text: False|None = 无页眉, "文字" = 自定义, 省略 = "哈尔滨工业大学"

doc.set_info(
    title="论文标题",          # 必填
    author="作者姓名",        # 必填
    supervisor="导师姓名",    # 必填
    subject="学科专业",
    affil="院系名称",
    date="2024年6月",
    student_id="1234567890",  # 本科生 cover2 表格用
)
```

### 文档结构

| 方法 | 说明 |
|------|------|
| `doc.add_cover()` | 封面（本科生自动生成双封面） |
| `doc.start_roman_section()` | 罗马数字页码节（摘要/目录） |
| `doc.start_arabic_section()` | 阿拉伯数字页码节（正文） |
| `doc.add_abstract_cn(text, keywords)` | 中文摘要，关键词用 `；` 分隔 |
| `doc.add_abstract_en(text, keywords)` | 英文摘要，关键词用 `, ` 分隔 |
| `doc.add_denotation(items)` | 物理量名称及符号表（可选） |
| `doc.add_toc(blank_line_before=False)` | 目录 |
| `doc.add_conclusion(title, content)` | 结论 |
| `doc.add_acknowledgements(title, content)` | 致谢 |
| `doc.add_references(bib)` | 参考文献（支持 `ReferenceDB` 或字符串列表） |
| `doc.add_authorization()` | 授权声明页（本科生） |
| `doc.include(module_path)` | 导入外部章节模块（分文件模式） |
| `doc.compile(filename)` | 编译输出 docx |

### 章节

```python
with doc.add_chapter("绪论", "1"):
    doc.add_section("研究背景")          # 1.1　研究背景
    doc.add_subsection("国内研究")       # 1.1.1　国内研究
    doc.add_subsubsection("具体案例")    # 1.1.1.1　具体案例
    doc.add_paragraph("正文内容……")
```

支持附录：

```python
with doc.add_appendix("符号"):
    doc.add_section("主要符号")
    doc.add_paragraph("D —— 轴承直径，mm")
```

### 内容元素

```python
# 段落（自动处理 [数字] 上标和 [ref:key] 引用）
doc.add_paragraph("正文内容……", indent=True)

# 三线表
tbl = doc.add_table(rows=3, cols=2, caption="实验数据", label="1-1", ref="tab_data")
tbl.set_cell(0, 0, "参数", bold=True)
tbl.set_cell(0, 1, "数值", bold=True)

# 图片（自动编号，居中）
doc.add_figure("image.png", caption="实验结果对比", width=Cm(10), ref="fig_result")

# 公式（无边框表格，编号右对齐）
doc.add_equation(r"\frac{\partial p}{\partial x} = 0", label="2-1", ref="eq_ns")

# 脚注
doc.add_footnote("此处指公开发表的学术论文。")

# 分页
doc.add_page_break()
```

### 文献引用

```python
from hitthesis import ReferenceDB

bib = ReferenceDB("references.bib")
doc.set_reference_db(bib)

# 正文中引用
doc.add_paragraph("气体轴承理论最早由 Gross[ref:Gross1962] 系统阐述。")
doc.add_paragraph("多项研究[ref:Zhang2015][ref:Chen2019][ref:Li2018]验证了这一结论。")

# 编译时自动按引用顺序编排 [1], [2-4], [5,6]…
doc.add_references(bib)
```

连续引用会自动合并区间：`[ref:A][ref:B][ref:C]` → `[1-3]`，非连续保持独立：`[ref:A][ref:C]` → `[1,3]`。

### 交叉引用

```python
doc.add_table(3, 2, caption="实验数据", ref="tab_data")
doc.add_figure("fig.png", caption="结果对比", ref="fig_result")
doc.add_equation(r"E=mc^2", label="1-1", ref="eq_mass")

doc.add_paragraph("实验结果如表 [cite:tab_data] 和图 [cite:fig_result] 所示。")
doc.add_paragraph("由式 [cite:eq_mass] 可知……")
```

引用的文字会自动匹配格式（"1-1"、"图 1-1"、"式(1-1)"），并附带指向对应图表/公式的超链接。

### 分文件组织

大型论文可将各章拆分为独立文件，由主控文件组合生成：

```python
# thesis_main.py —— 主控文件
doc.start_arabic_section()
doc.include("body.chapter1_intro")
doc.include("body.chapter2_theory")
doc.add_conclusion(...)
doc.compile("output/thesis.docx")
```

```python
# body/chapter1_intro.py —— 章节文件
def build(doc):
    with doc.add_chapter("绪论", "1"):
        doc.add_section("研究背景")
        doc.add_paragraph("正文内容...")
```

`include()` 自动导入模块并调用其 `build(doc)` 函数，章节内部 API 与单文件写法完全一致。

详见 `examples/thesis_split/`。

## 论文类型

| 类型 | `type` | 状态 | 封面数 | cover2 | 授权声明 |
|------|--------|:----:|:------:|:------:|:--------:|
| 本科毕业设计 | `bachelor` | **已实现** | 2 | 包含 | 包含 |
| 硕士学位论文 | `master` | 预留 | 1 | - | - |
| 博士学位论文 | `doctor` | 预留 | 1 | - | - |

校区：`harbin`（哈尔滨，已实现）、`shenzhen`（深圳，预留）、`weihai`（威海，预留）。

## 格式说明

以下为当前实现的格式（参考 [hithesis](https://github.com/hithesis/hithesis) LaTeX 模板，部分调整）：

- **页面**：A4 (210×297mm)，版芯 150×236mm，上边距 38mm，左右边距 30mm，下边距 30mm
- **封面标题**：黑体二号不加粗，居中
- **章标题**：黑体 18pt 加粗，居中，段前 24pt，段后 23pt
- **一级小节**：黑体 15pt 加粗，左对齐
- **二级小节**：黑体 14pt 加粗，左对齐
- **正文**：宋体 12pt，1.3 倍行距（20.5pt），首行缩进 2 字符
- **页眉**：宋体 9pt 居中 + 双线（粗 2.25pt + 细 0.75pt）
- **页脚**：`— 页码 —`，Times New Roman 9pt
- **三线表**：顶线/底线 1.5 磅，中线 1 磅，表题在上方居中
- **图表题注**：宋体 10.5pt，居中
- **封面（本科）**：双封面设计，第二封面含学生信息表格

## 目录结构

```
hithesis-docx/
├── hitthesis/                  # 核心库
│   ├── __init__.py             # 公开 API
│   ├── config.py               # 格式常量（页面、字体、行距等）
│   ├── document.py             # Thesis 主类（文档编排）
│   ├── ooxml_utils.py          # OOXML 原子操作（字体、边框、书签等）
│   ├── cover.py                # 封面与本科生第二封面
│   ├── authorization.py        # 授权声明页
│   ├── compile.py              # 编译流水线（保存 → Word COM → 后处理）
│   ├── toc_postproc.py         # TOC 字体修复与后处理
│   └── reference_db.py         # BibTeX 文献数据库与 GB/T 7714 格式化
├── examples/
│   ├── thesis_example.py       # 单文件完整示例
│   └── thesis_split/           # 分文件组织示例
│       ├── thesis_main.py      #   主控文件
│       └── body/               #   章节文件
├── output/                     # 生成的文档输出
├── requirements.txt
├── README.md
└── LICENSE
```

## 运行示例

```bash
python examples/thesis_example.py
```

生成的文档位于 `output/thesis_example.docx`。

## 已知问题

- **空白页 Bug**：TOC 域展开后，目录与第一章之间可能出现多余空白页。这是 Word COM 域更新的固有问题，需手动删除。
- **图片**：单图可用，子图、并排图未支持。
- **公式**：仅渲染公式编号，复杂 LaTeX 公式建议先渲染为图片再插入。
- **三级小节段后间距**：标准约 8.5pt，当前为 6pt。
- **格式偏差**：部分间距、字体与标准不完全一致。

## 未实现

- 硕博封面及论文结构
- 深圳、威海校区适配
- 子图 / 公式本体渲染
- 英文目录
- 全日制/非全日制区分

## 注意事项

- **目录更新**：生成的目录是 Word TOC 域代码。Windows 下编译时自动更新；macOS / Linux 下首次打开文档后需手动更新（`Ctrl+A` → `F9` → "更新整个目录"）
- **仅限 Windows**：Word COM 接口仅在 Windows 下可用

## 致谢

- [hithesis/hithesis](https://github.com/hithesis/hithesis) — 哈工大 LaTeX 学位论文模板，本项目格式参考来源
- 哈尔滨工业大学相关论文规范
- 基于 Minimax M2.7 / Deepseek V4 系列的 Claude Code — 本项目全部代码由此生成

## License

[CC BY-NC 4.0](LICENSE) — 署名-非商业性使用 4.0 国际
