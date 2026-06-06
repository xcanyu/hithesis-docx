# \[赤石科技\] 基于 python-docx 的 hithesis-docx 模板

[![Python](https://img.shields.io/badge/python-3.9+-blue.svg)](https://www.python.org/downloads/)
[![License](https://img.shields.io/badge/license-CC%20BY--NC%204.0-blue.svg)](LICENSE)

用 Python 代码生成哈尔滨工业大学学位论文 Word 文档（.docx）。

> **建议：正式学位论文写作请使用 [hithesis](https://github.com/hithesis/hithesis) LaTeX 模板编译。** 本项目适用于对格式要求不严的 `.docx` 文档生成场景，如草稿、非正式提交、课程报告等。

---

> **代码生成声明**
>
> 本项目全部代码由基于 **Minimax M2.7** / **Deepseek V4 系列** 的 **Claude Code** 生成，人工仅做需求描述和效果验证。

> **重要声明**
>
> 本项目格式参考 [hithesis/hithesis](https://github.com/hithesis/hithesis)（哈工大 LaTeX 学位论文模板），在此基础上做了**高仿**，**并非严格符合**《哈尔滨工业大学学位论文撰写规范》。目前仅实现了**本科**，硕博（master/doctor）未做。深圳/威海校区未做，仅预留了 `campus` 参数空间。**不建议用于正式学位论文提交**，仅供学习交流。

## 特性

- **参考规范**：格式参考 [hithesis](https://github.com/hithesis/hithesis) LaTeX 模板——页面边距、字体字号、行距缩进、三线表、双线页眉
- **代码驱动**：用 Python 描述文档结构，版本可控、可复用
- **学位类型**：本科（bachelor），校区仅实现本部（harbin）

## 安装

建议在虚拟环境中安装（避免污染全局 Python）：

```bash
python -m venv .venv          # 创建虚拟环境
.venv\Scripts\activate        # Windows 激活
pip install -r requirements.txt
```

Linux/macOS 激活用 `source .venv/bin/activate`。

> **Windows 用户**：编译阶段会自动调用 Word COM 接口更新目录域，无需额外配置。macOS / Linux 用户需手动更新目录（打开文档后 `Ctrl+A` → `F9`）。

## 快速开始

```python
from hitthesis import Thesis

doc = Thesis(type="bachelor", campus="harbin")  # type: bachelor|master|doctor, campus: harbin
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
doc = Thesis(type="bachelor", campus="harbin")
# type: "bachelor"|"master"|"doctor"
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
| `doc.add_conclusion(title, content, intro)` | 结论（intro 导语不编号，content 自动 `（1）（2）...`） |
| `doc.add_acknowledgements(title, content, auto_space=True)` | 致谢 |
| `doc.add_references(bib)` | 参考文献（支持 `ReferenceDB` 或字符串列表） |
| `doc.add_authorization()` | 授权声明页（根据 type 自动选择本科/硕博版本） |
| `doc.add_publications(sections)` | 发表文章页（硕博），自动编号、悬挂缩进 |
| `doc.add_resume(title, content, auto_space=True)` | 个人简历（博士） |
| `doc.include(module_path)` | 导入外部模块（支持 `front.xxx` / `body.xxx` / `back.xxx`） |
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

# 代码块（Consolas 9pt 灰底，题注在上）
doc.add_code_block("def hello():\n    print('hello')", caption="Hello World", ref="code_hello")

# 列表（有序/无序/括号三种格式）
doc.add_list(["研究内容一", "研究内容二"], style="decimal")   # 1. 2. 3.
doc.add_list(["设备一", "设备二"], style="bullet")              # — 前缀
doc.add_list(["步骤一", "步骤二"], style="paren")               # （1）（2）

# 定理/定义/引理（标题正文同行，cite 支持 [ref:xxx] 文献引用）
doc.add_theorem("对于不可压缩流体...", kind="定理", ref="thm_1", cite="[ref:Gross1962]")

# 引用块（楷体 12pt，左右缩进）
doc.add_quote("气体轴承的概念最早由法国学者Hirn于1854年提出...")

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

### LaTeX 公式与化学式

本项目支持在文本和公式中使用 LaTeX 语法渲染上下标和复杂公式：

```python
# 简单化学式（用 $...$ 包裹）
doc.add_paragraph("水的化学式为 $H_{2}O$。")
doc.add_paragraph("硫酸根离子为 $SO_{4}^{2-}$。")
doc.add_paragraph("实验测得 $Fe^{2+}$ 浓度为 $10^{-3}$ mol/L。")

# 复杂公式（调用 latex2word 渲染）
doc.add_equation(
    r"\frac{\partial p}{\partial x} = 6\mu\frac{\partial(\rho u)}{\partial x}",
    label="2-1", ref="eq_reynolds"
)
```

**语法规则：**
- `$H_2O$` → H₂O（单个字符下标）
- `$H_{2}O$` → H₂O（多字符下标）
- `$Fe^{2+}$` → Fe²⁺（上标）
- `$10^{-3}$` → 10⁻³（复杂上标）

**已知问题：**
- 公式字体由 Word/WPS 内部管理，暂不支持强制设置为 Times New Roman

### 分文件组织

大型论文可将内容拆分为 `front/`、`body/`、`back/`，由主控文件组合生成，结构与 LaTeX 一致：

```python
# thesis_main.py —— 主控文件（只编排结构）
doc.include("front.cover")
doc.start_roman_section()
doc.include("front.abstract_cn")
doc.include("front.abstract_en")
doc.include("front.denotation")
doc.add_toc()

doc.start_arabic_section()
doc.include("body.chapter1_intro")     # 正文章节
doc.include("body.chapter2_theory")

doc.include("back.conclusion")         # 后文
doc.include("back.references")
doc.include("back.appendix")
doc.include("back.publications")
doc.include("back.authorization")
doc.include("back.acknowledgements")
doc.include("back.resume")
doc.compile("output/thesis.docx")
```

```python
# body/chapter1_intro.py —— 章节文件
def build(doc):
    with doc.add_chapter("绪论", "1"):
        doc.add_section("研究背景")
        doc.add_paragraph("正文内容...")
```

`include()` 自动导入模块并调用 `build(doc)` 函数，内部 API 与单文件写法完全一致。
前后文均可独立为文件，主控文件只负责顺序编排。

详见 `examples/thesis_split/`。

## 论文类型

| 类型 | `type` | 状态 | 封面数 | 授权声明 | 发表文章 | 简历 |
|------|--------|:----:|:------:|:--------:|:--------:|:----:|
| 本科毕业设计 | `bachelor` | ✅ 已实现 | 2 | 本科版 | - | - |
| 硕士学位论文 | `master` | ❌ 预留 | - | - | - | - |
| 博士学位论文 | `doctor` | ❌ 预留 | - | - | - | - |

校区：`harbin`（哈尔滨，已实现）、`shenzhen`（深圳，预留）、`weihai`（威海，预留）。

## 格式说明

以下为当前实现的格式（以学校 **2025 年官方本科 Word 模板**为唯一标准）：

- **页面**：A4 (210×297mm)，版芯 150×236mm，上边距 38mm，左右边距 30mm，下边距 30mm
- **章标题**：黑体 18pt 加粗，居中，段前 20pt，段后 17pt
- **一级小节**：黑体 15pt 加粗，左对齐，段前 10pt，段后 12pt
- **二级小节**：黑体 14pt 加粗，左对齐，段前 8pt，段后 12pt
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
│   ├── config.py               # 格式常量
│   ├── document.py             # Thesis 主类
│   ├── ooxml_utils.py          # OOXML 原子操作
│   ├── cover.py                # 封面
│   ├── authorization.py        # 授权声明页（本科/硕博双版本）
│   ├── compile.py              # 编译流水线
│   ├── toc_postproc.py         # TOC 字体修复
│   ├── footnote.py             # 脚注生成
│   ├── latex_render.py         # LaTeX 公式渲染
│   └── reference_db.py         # BibTeX 文献数据库
├── examples/
│   ├── thesis_example.py       # 单文件完整示例
│   └── thesis_split/           # 分文件组织示例
│       ├── thesis_main.py      #   主控文件
│       ├── front/              #   前文（封面、摘要、符号表）
│       ├── body/               #   正文章节
│       └── back/               #   后文（结论、参考文献、附录等）
├── output/                     # 生成的文档输出
├── requirements.txt
├── README.md
├── md转Word排版-skill.md       # OpenCode skill（md→docx 转换）
└── LICENSE
```

## 运行示例

```bash
# 单文件示例
python examples/thesis_example.py

# 分文件示例（front/body/back 架构）
python examples/thesis_split/thesis_main.py
```

生成的文档分别位于 `output/thesis_example.docx` 和 `examples/thesis_split/output/thesis_split.docx`。

## 已知问题

- **空白页 Bug**：TOC 域展开后，目录与第一章之间可能出现多余空白页。
- **公式字体**：OMML 公式字体由 Word 内部管理，暂不支持强制设置为 Times New Roman。
- **WPS 兼容性**：字间距（`w:spacing`）在 Word 和 WPS 下渲染效果有差异，以 Word 为准。

## 未实现

- 硕博封面（暂时用本科封面替代）
- 深圳、威海校区适配
- 中英双语图题/表题（博士要求）
- 子图（并排图、子图编号 a/b）
- 长表格（续表，跨页重复表头）
- 算法伪代码（行号+关键词着色）
- 开题/中期报告
- 英文目录
- 全日制/非全日制区分

## Skill

本项目附带 OpenCode skill（`md转Word排版-skill.md`），基于本项目的格式规范，可自动将 Markdown 转换为 hithesis 格式的 Word 文档。将其放入 `~/.claude/skills/md-to-hithesis-docx/` 或 OpenCode 对应目录即可生效。

## 注意事项

- **目录更新**：生成的目录是 Word TOC 域代码。Windows 下编译时自动更新；macOS / Linux 下首次打开文档后需手动更新（`Ctrl+A` → `F9` → "更新整个目录"）
- **仅限 Windows**：Word COM 接口仅在 Windows 下可用

## 致谢

- [hithesis/hithesis](https://github.com/hithesis/hithesis) — 哈工大 LaTeX 学位论文模板，本项目格式参考来源
- 哈尔滨工业大学相关论文规范
- 基于 Minimax M2.7 / Deepseek V4 系列的 Claude Code — 本项目全部代码由此生成

## License

[CC BY-NC 4.0](LICENSE) — 署名-非商业性使用 4.0 国际
