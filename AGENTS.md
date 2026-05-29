# hithesis-docx — 哈工大学位论文Word文档生成工具

## 项目概述
Python 库，用 python-docx 生成哈尔滨工业大学学位论文 .docx 文件。
已实现 **本科/本部（bachelor/harbin）**，硕博（master/doctor）只做了部分功能（代码已注释，留作调试）。

## 目录结构
```
hitthesis/              # 核心库
├── __init__.py         # 公开 API：Thesis, ReferenceDB, set_font
├── config.py           # 格式常量（页面、字体、行距、段落间距）
├── document.py         # Thesis 主类（文档编排核心，~1719行）
├── ooxml_utils.py      # OOXML 原子操作（字体、边框、书签、超链接等）
├── cover.py            # 封面
├── authorization.py    # 授权声明页（本科/硕博双版本）
├── compile.py          # 编译流水线（保存 → Word COM → 后处理）
├── toc_postproc.py     # TOC 字体修复与后处理
├── footnote.py         # 脚注生成
├── latex_render.py     # LaTeX 公式渲染（$...$ 语法）
└── reference_db.py     # BibTeX 文献数据库与 GB/T 7714 格式化

examples/
├── thesis_example.py              # 单文件完整示例（硕博）
├── thesis_split/                  # 分文件示例（front/body/back 架构）
│   ├── thesis_main.py
│   ├── front/                     # 前文（封面、摘要、符号表）
│   ├── body/                      # 正文章节
│   └── back/                      # 后文（结论、参考文献、附录、发表文章、授权、致谢、简历）
```

## 运行环境

项目使用虚拟环境 `.venv/`（Python 3.12），所有依赖已安装。编译前需激活：

```bash
.venv\Scripts\activate        # Windows PowerShell
source .venv/bin/activate      # Linux/macOS
```

激活后终端会有 `(.venv)` 前缀。VSCode 打开项目文件夹会自动识别。

## 架构设计
- **`Thesis`** 类对外暴露全部 API，内部调用各模块
- 文档元素通过 `python-docx` 的 OOXML 底层操作构建
- 分文件模式：`doc.include(module_path)` 用 `importlib` 动态导入章节模块，支持 `front.xxx` / `body.xxx` / `back.xxx`
- 授权声明：通过 `thesis_type` 自动选择本科或硕博版本

## 编码约定
- 类型提示：`set_info` 用断言检查类型，非 PEP 484
- 文档字符串：模块级和类级用 `"""doc"""`，方法有简短注释
- 中英夹杂：变量名/注释中英文混用，保持项目既有风格
- 导入：标准库 → 第三方 → 本地模块
- 行距/间距：用 `before/after` 控制，单位 `pt`
- OOXML 操作：用 `OxmlElement` + `qn()` 命名空间，尽量少用 python-docx 高层 API

## 关键设计决策
- 垂直定位：用 `space_before` spacer 而非 `wp:anchor` 锚点
- 分页：章标题用 `pageBreakBefore` 而非手动分页
- 交叉引用：`[cite:xxx]` 标记，编译时解析为超链接
- 文献引用：`[ref:key]` 标记，编译时按引用顺序编号 `[1], [2-4]`
- 脚注：收集后 ZIP 后处理生成 `footnotes.xml`
- 授权声明：`authorization.py` 内 `_add_bachelor` / `_add_graduate` 双版本，通过 `thesis_type` 切换
- 标题自动空格：`auto_space` 参数控制两汉字标题是否自动加全角空格
- 分文件架构：`front/`（前文）/ `body/`（正文）/ `back/`（后文）三级目录，与 LaTeX 一致

## 测试

```bash
.venv\Scripts\activate                           # 先激活环境

python examples/thesis_example.py                # 单文件示例
# 输出见 output/thesis_example.docx

python examples/thesis_split/thesis_main.py      # 分文件示例
# 输出见 output/thesis_split.docx
```

## 常见任务
- **新增论文类型**：修改 `config.py` 中对应常量，扩展 `document.py` 中的流程
- **调整格式**：修改 `config.py` 中的 `PAGE`、`FONTS`、`LINE_SPACING` 等常量
- **添加文档元素**：在 `document.py` 中新增方法，复用 `ooxml_utils.py` 的原子操作
