---
name: md-to-hithesis-docx
description: Use when converting markdown to Word .docx with university thesis formatting ("md 转 word", "编译成word"). Requires the hithesis-docx project (https://github.com).
---

> **前置依赖**：此 skill 基于 [hithesis-docx](https://github.com/xcanyu/hithesis-docx) 项目，需先安装该项目（含 `.venv`）才能使用。

# Markdown → Hithesis Docx

Convert any markdown file to a .docx with 哈工大 thesis formatting using the hithesis-docx library.

## hithesis-docx Project

The hithesis-docx project is required for this conversion. First, try to locate it:

1. Check for a `HITHESIS_PATH` variable in the user's CLAUDE.md / AGENTS.md
2. Search common locations like `~/Desktop/`, `~/Documents/`, or similar

If found, use that path. If NOT found, ask the user:

> "需要 hithesis-docx 项目才能转换。请提供项目路径，或者从 https://github.com/xcanyu/hithesis-docx 克隆一份。"

Do NOT proceed with conversion until the project path is confirmed.

The project uses a `.venv` (Python 3.9+, dependencies installed). Activate with:
```
.venv\Scripts\activate
```

## Available Formatting Parts

Import these to build documents:

```python
from hitthesis import Thesis, set_font
from hitthesis.config import SPACING
from hitthesis.ooxml_utils import set_first_line_indent, disable_snap_to_grid
```

| Part | What it does | Usage |
|------|-------------|-------|
| `set_font(run, "黑体", 18, bold)` | Set font name, size, bold | Per-run formatting |
| `set_first_line_indent(para, 480)` | First-line indent in twips | Body paragraphs |
| `disable_snap_to_grid(para)` | Disable Word grid snapping | Headings |
| `SPACING` dict | All spacing constants | `SPACING["body_line_spacing"]`, `SPACING["section_before"]`, etc. |
| `doc.add_chapter(title)` | Chapter heading (黑体18pt, centered, page break before) | `with doc.add_chapter("绪论"):` |
| `doc.add_section(title)` | Auto-numbered section heading (e.g., "1.1 标题") | Inside chapter context |
| `doc.add_paragraph(text, indent=True)` | Body paragraph (宋体12pt, 20.5pt line spacing) | Simple single-style text |
| `doc.doc.add_paragraph()` | Raw paragraph for custom styling | Rich text with mixed bold/normal |
| `doc.add_code_block(text, caption, ref)` | Code block (Consolas 9pt, gray bg) | Multi-line code strings |
| `doc.compile(filename)` | Save and finalize the docx | Always call at the end |

## Key Format Specifications

| Element | Font | Size | Notes |
|---------|------|------|-------|
| Chapter heading | 黑体 | 18pt | Centered, page break before |
| Section heading | 黑体 | 15pt | Left-aligned, auto-numbered |
| Subsection | 黑体 | 14pt | Left-aligned |
| Body text | 宋体 | 12pt | 20.5pt line spacing, first-line indent 480 twips |
| Code block | Consolas | 9pt | Gray bg #F5F5F5, left indent 0.5cm |
| Caption (table/figure/code) | 宋体 | 10.5pt | Centered, 14pt line spacing |
| Page | A4 | — | Top 3.8cm, bottom 3.0cm, left/right 3.0cm |

## Conversion Strategy

**Analyze the markdown structure first, then decide how to convert.**

### Simple documents (headings + paragraphs only)

If the md has only `##`, `###` headings and plain paragraphs (no options, no answers, no special patterns):

```python
from hitthesis import Thesis

doc = Thesis(type="bachelor", campus="harbin")

for line in open(md_path, encoding='utf-8'):
    if line.startswith('## '):
        with doc.add_chapter(line[3:].strip()):
            pass
    elif line.startswith('### '):
        doc.add_section(line[4:].strip())
    elif line.strip():
        doc.add_paragraph(line.strip(), indent=True)

doc.compile(output_path)
```

### Complex documents (exam questions, homework, options, answers)

If the md has special patterns (numbered questions, `**答案：**` markers, A/B/C/D options, case analysis):

1. Read the md file fully
2. Analyze the structure manually — identify heading levels, answer markers, option formats
3. Write a custom Python script that:
   - Uses `doc.doc.add_paragraph()` for paragraphs with mixed bold/normal text
   - Creates custom section headings with `set_font(run, "黑体", 15)` when auto-numbering is not desired
   - Parses answer markers (`**答案：**`, `**答：**`) and renders "答案：" in bold followed by normal text
   - Renders options as indented paragraphs
4. Save the script to a temp directory under `$env:TEMP\opencode` and run it from the hithesis-docx project root

**For custom section headings without auto-numbering**, use this pattern instead of `add_section()`:

```python
def add_heading_without_number(doc, title, font_name="黑体", font_size=15):
    para = doc.doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    para.paragraph_format.line_spacing = Pt(22)
    para.paragraph_format.space_before = Pt(SPACING["section_before"])
    para.paragraph_format.space_after = Pt(SPACING["section_after"])
    disable_snap_to_grid(para)
    run = para.add_run(title)
    set_font(run, font_name, font_size, False)
    return para
```

**For rich text paragraphs** (mixed bold/normal in one paragraph):

```python
def add_rich_paragraph(doc, segments, indent=True, extra_indent=0):
    """segments = [(text, is_bold), ...]"""
    para = doc.doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    para.paragraph_format.line_spacing = Pt(SPACING["body_line_spacing"])
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after = Pt(0)
    if indent:
        set_first_line_indent(para, SPACING["first_line_indent"] + extra_indent)
    disable_snap_to_grid(para)
    for text, bold in segments:
        run = para.add_run(text)
        set_font(run, "宋体", SPACING["body_font_size"], bold)
    return para
```

## Running the Conversion

Always run from the hithesis-docx project root using its venv:

```powershell
# Navigate to project first (ask user for path if needed)
Set-Location "path/to/hithesis-docx"
.venv\Scripts\activate
python convert_script.py
```


