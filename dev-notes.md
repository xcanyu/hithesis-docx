# 开发笔记

> 最后更新：2026-05-30
>
> 本项目格式参考 [hithesis/hithesis](https://github.com/hithesis/hithesis)（哈工大 LaTeX 学位论文模板），
> 为高仿版，并非严格符合规范。目前仅实现本科/本部，不建议用于正式学位论文提交。

## 当前状态

- 已实现：**本科（bachelor）** + **硕士/博士（master/doctor）部分功能**，校区：**本部（harbin）**
- 封面：本科双封面完善，硕博暂用本科封面
- 授权声明：本科和硕博两个版本，通过 `thesis_type` 自动切换
- 预留空间：深圳/威海（`campus="shenzhen"|"weihai"`），页眉自定义（`header_text`）
- 核心库模块：`document.py` / `cover.py` / `authorization.py` / `compile.py` / `toc_postproc.py` / `ooxml_utils.py` / `reference_db.py` / `footnote.py` / `latex_render.py`
- 分文件章节组织模式已实现（`doc.include()` 方法），支持 `front/` / `body/` / `back/` 架构
- LaTeX 公式渲染已实现（latex2word + 手动上下标解析）

## 已解决

| 问题 | 说明 |
|------|------|
| 一级标题段前间距 | `before="480"` + keepNext + keepLines + snapToGrid |
| 页码格式 | 封面无页码，罗马数字→阿拉伯数字 |
| 封面格式 | 双封面（本科），学生信息表格 |
| 三线表 | 顶线/底线1.5pt，中线0.5pt |
| 页眉页脚 | 双线页眉 + `— PAGE —` 页脚 |
| 目录字体 | TOC后处理修正为黑体/宋体 |
| 公式 | 无边框1×3表格，编号右对齐 |
| 交叉引用 | `[cite:xxx]` 超链接引用 |
| 文献引用 | `[ref:key]` + BibTeX，GB/T 7714 |
| 分文件组织 | `doc.include()` 动态导入章节模块 |
| 脚注 | 上标圆圈数字 6.5pt，正文 9pt 宋体/TNR，ZIP后处理生成 footnotes.xml |
| 参考文献悬挂缩进 | `w:left=450` + `w:firstLine=-450`，续行对齐编号后文字 |
| LaTeX 公式渲染 | `$...$` 包裹文本内公式，latex2word 转 OMML，支持 `\frac`/`\partial` 等 |
| 化学式/上下标 | `^x` 单字符上标，`^ {xxx}` 多字符上标，`_x`/`_{xxx}` 下标 |
| 发表文章页 | `add_publications()`，自动编号 `[1]`，悬挂缩进 |
| 个人简历 | `add_resume()`，默认标题"个人简历" |
| 硕博授权声明 | 黑体小二不加粗，1.6cm间距，条款行字间距控制 |
| 标题自动空格 | `auto_space` 参数可关闭两汉字标题自动加全角空格 |

## 间距校准结果（2026-05-30）

通过 VBA 宏实测，以学校官方本科 Word 模板为基准，建立了 Word 特定的间距公式：

**核心发现：**
- Word 视觉间距 = space_after + line_spacing × 0.30
- 行距是间距的决定性因素，before/after 影响有限
- 不能直接搬 LaTeX bp 值，需用官方 Word 模板实测

**当前配置（已校准）：**

| 参数 | 值 | 说明 |
|------|-----|------|
| heading_before | 454tw (22.7pt) | 双线到标题间距 1cm |
| heading_after | 400tw (20pt) | 标题到正文间距 |
| heading_line | 288 (1.2×) | 标题行距 |
| section_before | 0pt | 节标题段前 |
| section_after | 8pt | 节标题段后 |
| subsection_before | 7pt | 小节标题段前 |
| subsection_after | 9pt | 小节标题段后 |
| subsubsection_before | 0pt | 子小节标题段前 |
| subsubsection_after | 3pt | 子小节标题段后 |

**页面设置（对齐官方本科模板）：**

| 参数 | 值 |
|------|-----|
| 上边距 | 3.80cm |
| 下边距 | 3.00cm |
| 左边距 | 3.00cm |
| 右边距 | 3.00cm |
| 页眉距离 | 3.00cm |
| 页脚距离 | 2.30cm |

**实测间距（与官方模板对比）：**

| 组合 | 官方模板 | 实测值 | 状态 |
|------|----------|--------|------|
| L1→L2 | 1.68cm | 1.69cm | ✓ |
| L2→L3 | 1.04cm | 1.06cm | ✓ |
| L3→L4 | 0.98cm | 1.00cm | ✓ |
| L2↔正文 | 1.04cm | 1.06-1.07cm | ✓ 统一 |
| L3↔正文 | 0.98cm | 0.95-1.00cm | ✓ 统一 |
| L4↔正文 | 0.66cm | 0.68-0.71cm | ✓ |

**L4（子小节标题）处理：**
- 不加入目录（不设置 outline_level）
- 用文本模式 `X.Y.Z.W` 识别

## 未解决

| 问题 | 说明 |
|------|------|
| 硕博封面 | 暂用本科封面，硕博封面格式不同 |
| 深圳/威海校区 | 页眉、书序、封面不同 |
| 英文论文 | `language=english` 选项 |
| 中英双语图题/表题 | 博士要求图题中文+英文两行 |
| 子图 | 分图编号 (a)(b) |
| 定理环境 | definition/theorem/lemma/proof |
| 长表格（续表） | 跨页重复表头 |
| 开题/中期报告 | 单独的报告类模板 |
| 空白页 Bug | TOC域展开后多余空白页，Word固有问题 |
| 公式字体 | OMML 公式字体由 Word 内部管理，暂不支持强制设置为 Times New Roman |

## 与 hithesis v3.1d (LaTeX) 格式差异（已知不修正）

以下为审阅确认不调整的差异，记录备查。

| 类别 | 项 | LaTeX 规范值 | 代码实际值 | 备注 |
|------|-----|-------------|-----------|------|
| 页面 | 下边距 | ≈23mm | 30mm | 对齐官方Word模板 |
| 章标题 | 段前 | 28.35bp≈28.45pt | 22.7pt | 双线到标题1cm |
| 章标题 | 段后 | 28.75bp≈28.85pt | 20pt | |
| 章标题 | 行距倍率 | 1.575×（baselineskip=28.35bp） | 1.2×（288twips） | Word行距模型不同 |
| 节标题 | 段前 | 19.84bp≈19.9pt | 0pt | |
| 节标题 | 段后 | 19.84bp≈19.9pt | 8pt | |
| 小节标题 | 段前 | 17.01bp≈17.1pt | 7pt | |
| 小节标题 | 段后 | 17.01bp≈17.1pt | 9pt | |
| 子小节标题 | 段前 | 8.50bp≈8.5pt | 0pt | |
| 子小节标题 | 段后 | 8.50bp≈8.5pt | 3pt | |
| 三线表 | 中线 | 0.5pt（sz=4） | 1pt（sz=8） | 主动选择 |
| 图题 | 序号与图名间距 | 2个半角字符 | 全角空格 | 视觉宽度相近 |
| 表题 | 序号与表名间距 | 1格 | 全角空格 | 可能略宽 |
| 授权声明 | 本科原创性声明标题字号 | 黑体小二18bp | 黑体22pt | 设计选择 |
| 图题 | 序号与图名间距 | 2个半角字符 | 全角空格 | 视觉宽度相近 |
| 表题 | 序号与表名间距 | 1格 | 全角空格 | 可能略宽 |
| 授权声明 | 本科原创性声明标题字号 | 黑体小二18bp | 黑体22pt | 设计选择 |

## 关键设计决策

- **垂直定位**：不用 `wp:anchor` 锚点，改用 `space_before` spacer（避免Word COM保存验证错误）
- **分页**：章标题用 `pageBreakBefore` 替代手动分页
- **授权声明**：5行条款用独立段落+字间距撑宽（`w:spacing`），本科和硕博两版本通过 `thesis_type` 切换
- **TOC生成**：Word TOC域代码 + COM自动更新 + `_fix_toc_fonts()` 后处理
- **分文件模式**：`include()` 使用 `importlib.import_module` + `build(doc)` 约定，保持API兼容

## 重构历史

2026-05-30 间距校准（以学校官方本科Word模板为基准）：
- 建立 Word 特定的间距公式：视觉间距 = space_after + line_spacing × 0.30
- 页面设置对齐官方模板：上3.8cm、下3.0cm、左3.0cm、右3.0cm、页眉3.0cm、页脚2.3cm
- 章标题段前：454tw (22.7pt)，确保双线到标题间距1cm
- 章标题段后：400tw (20pt)，行距1.2× (288tw)
- 节标题：段前0pt、段后8pt，不加粗（bachelor）
- 小节标题：段前7pt、段后9pt，不加粗
- 子小节标题：段前0pt、段后3pt，不加粗，不加入目录
- 统一各级标题↔正文间距：L2↔正文约1.04cm、L3↔正文约0.98cm、L4↔正文约0.66cm
- 新增 VBA 测量宏：measure_*.bas，用于精确测量各类间距

2026-05-27 硕博支持 + 分文件架构完善：
- 新增 `add_publications()`、`add_resume()` 方法
- 新增 `auto_space` 参数，可关闭两汉字标题自动空格
- 重写 `authorization.py`：本科/硕博两个版本，通过 `thesis_type` 自动切换
- 硕博授权声明：黑体小二不加粗、学位论文原创性声明间距1.6cm、条款行字间距精确控制
- 示例文件 `thesis_example.py` 切换为 `type="master"`，书序改为硕博顺序
- 示例文件 `thesis_split/` 重构为 `front/` / `body/` / `back/` 架构，与 LaTeX 结构一致
- `config.py` 过时备注同步代码实际值

2026-05-25 LaTeX 公式渲染：
- 新增 `latex_render.py` 模块，支持 `$...$` 包裹的公式和化学式
- 简单上下标（`^x`, `_{xxx}`）直接渲染
- 复杂 LaTeX 公式调用 latex2word 转 OMML
- 已知问题：公式字体无法强制设置为 Times New Roman

2026-05-14 将 `document.py`（~2641行）拆分为：
- `ooxml_utils.py` — OOXML原子操作
- `cover.py` — 封面
- `authorization.py` — 授权声明
- `compile.py` — 编译流水线
- `toc_postproc.py` — TOC后处理

2026-05-12 封面重构：页边距调整、动态类型显示、英文标题字号自动判断

2026-05-10 标题自动编号、目录两汉字空格、前导符不加粗

---

## 搁置调试：章标题段后间距（heading_after）

**日期**: 2026-05-29
**状态**: ✅ 已解决（2026-05-30）
**分支**: `fix/bachelor-harbin-format`
**关联文件**: `config.py` / `ooxml_utils.py`

### 问题

章标题底部到正文第一段的视觉间距，在 Word docx 中明显大于 LaTeX 模板 PDF。LaTeX 实测约 0.66cm，Word 实测约 0.9cm（575 twips 时）。

### 解决方案

通过 VBA 宏实测，建立了 Word 特定的间距公式：
- **Word 视觉间距 = space_after + line_spacing × 0.30**
- 行距是间距的决定性因素，before/after 影响有限
- 以学校官方本科 Word 模板为基准，校准所有间距参数

当前 heading_after = 400tw (20pt)，配合 heading_line = 288 (1.2×)，实现与官方模板一致的视觉间距。

### LaTeX 参考值

```
hithesisbook.cls: afterskip = 28.74646bp ≈ 28.86pt
实际渲染视觉间距 ≈ 0.66cm
```

### 根本原因

Word 和 LaTeX 的段落间距测量模型不同：
- **LaTeX**: baseline-to-baseline，afterskip 值直接决定间距
- **Word**: line-box bottom → next line-box top，`lineRule=auto` 时 line-box 底部到文字基线有额外空白，视觉上放大了 space_after

### 尝试过的值

| twips | 对应 pt | 测量结果 |
|-------|---------|---------|
| 575 | 28.75 | 约 0.9cm，偏大 |
| 490 | 24.5 | 仍偏大 |
| 350 | 17.5 | 无明显变化 |
| **435** | **21.75** | **当前值，未确认** |

### 待排查

- Word line-box 内空白的精确高度
- `lineRule=exact` 替代 `lineRule=auto` 是否能减少空白
- 正文段落 `space_before` / `snapToGrid` 是否干扰
- 详尽分析见 `docs/superpowers/specs/heading-after-spacing-debug.md`

---

## 本轮格式修复完整记录（2026-05-29）— 已被 2026-05-30 校准结果替代

**注意**：以下记录为 2026-05-29 的修改，已被 2026-05-30 的校准结果替代。当前配置以"间距校准结果（2026-05-30）"章节为准。

**分支**: `fix/bachelor-harbin-format`
**目标**: 严格对齐 LaTeX 模板 hithesisbook.cls 的格式参数

### 1. hitthesis/config.py — 页面布局 + 标题间距

**PAGE（页面边距）：**

| 参数 | 修改前 | 修改后 | 说明 |
|------|--------|--------|------|
| header | 0.5cm | 3.2cm | 页眉距页面顶部，对齐 LaTeX headsep=1mm |
| footer | 0.52cm | 2.78cm | 页脚距页面底部，对齐 LaTeX footsep=1mm |

**SPACING（标题间距）：**

| 参数 | 修改前 | 修改后 | LaTeX 来源 |
|------|--------|--------|-----------|
| heading_before | 480 (24pt) | 567 (28.35pt) | beforeskip={28.34646bp} |
| heading_after | 460 (23pt) | 435 (21.75pt) | afterskip={28.74646bp} 视觉等效（**搁置，仍需调试**） |
| section_before | 18pt | 19.84pt | beforeskip={19.84252bp} |
| section_after | 6pt | 19.84pt | afterskip={19.84252bp} |
| subsection_before | 12pt | 15pt | beforeskip={15bp} |
| subsection_after | 6pt | 15pt | afterskip={15bp} |

### 2. hitthesis/document.py — 页眉内容 + 标题引用修复

| 位置 | 修改前 | 修改后 |
|------|--------|--------|
| `__init__` 页眉文字 | "哈尔滨工业大学" | "哈尔滨工业大学本科毕业论文（设计）"（bachelor） |
| 页眉双线间距 | Pt(0.77) | Pt(0.75)（对齐 LaTeX `\doublerulesep`） |
| `add_subsection` space_after | 误用 `section_after` | 修正为 `subsection_after` |
| `add_subsubsection` space_after | 误用 `section_after` | 修正为 `subsubsection_after` |

### 3. hitthesis/cover.py — 封面间距

**封面第一页：**

| 参数 | 修改前 | 修改后 |
|------|--------|--------|
| "本科毕业论文（设计）" space_before | 34pt | 24.1pt |
| 论文标题 space_after | 44pt | 36pt |
| 英文标题 space_after | 76.9pt | 67pt |
| 作者姓名 space_after | 136.8pt | 126pt |

**封面第二页：**

| 参数 | 修改前 | 修改后 |
|------|--------|--------|
| 密级字号 | 14bp | 12bp |
| 密级到标题间距 | 79.7pt | 59.5pt |

### 4. hitthesis/ooxml_utils.py — 章标题样式

| 参数 | 修改前 | 修改后 |
|------|--------|--------|
| `apply_heading_style` 默认 after_pt | 23pt | 21.75pt |
