# 开发笔记

> 最后更新：2026-05-27
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
| 页面 | 下边距 | ≈23mm | 30mm | |
| 章标题 | 段前 | 28.35bp≈28.45pt | 24pt | |
| 章标题 | 段后 | 28.75bp≈28.85pt | 23pt | |
| 章标题 | 行距倍率 | 1.575×（baselineskip=28.35bp） | 1.2×（288twips） | |
| 节标题 | 段前 | 19.84bp≈19.9pt | 18pt | |
| 节标题 | 段后 | 19.84bp≈19.9pt | 6pt | 差距最大的项 |
| 小节标题 | 段前 | 17.01bp≈17.1pt | 12pt | |
| 小节标题 | 段后 | 17.01bp≈17.1pt | 6pt | |
| 子小节标题 | 段前 | 8.50bp≈8.5pt | 6pt | |
| 子小节标题 | 段后 | 8.50bp≈8.5pt | 6pt | |
| 三线表 | 中线 | 0.5pt（sz=4） | 1pt（sz=8） | 主动选择 |
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
