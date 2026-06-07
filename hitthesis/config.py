"""
哈工大学位论文格式配置 - 基于《哈尔滨工业大学研究生学位论文撰写规范》
"""

# 页面设置 (单位: cm)
PAGE = {
    # A4纸张: 210mm x 297mm
    "width": 21.0,
    "height": 29.7,
    # 版芯尺寸
    "text_width": 15.0,      # 正文宽度 150mm
    "text_height": 23.6,     # 正文高度 236mm
    # 页边距（对齐官方Word模板）
    "top": 3.80,           # 上边距 38mm
    "bottom": 3.00,        # 下边距 30mm
    "left": 3.0,            # 左边距 30mm
    "right": 3.0,           # 右边距 30mm
    "header": 3.00,         # 页眉距页面顶部 30mm
    "footer": 2.30,         # 页脚距页面底部 23mm
}

# 学校名称
UNIVERSITY_NAME = "哈尔滨工业大学"

# 标题段落间距（twips，1pt = 20twips）
SPACING = {
    "heading_before": 400,      # 章标题段前（twips）= 20pt（一级标题距页顶4.50cm）
    "heading_after": 340,       # 章标题段后（twips）= 17pt（对齐本科模板 L1→正文 1.59cm）
    "heading_line": 288,        # 章标题行距（twips）= 1.2倍
    "section_before": 10,       # 一级小节段前（pt）（对齐本科模板 正文→L2 1.02cm）
    "section_after": 12,        # 一级小节段后（pt）（对齐本科模板 L2→正文 1.19cm）
    "subsection_before": 8,     # 二级小节段前（pt）（对齐本科模板 正文→L3 1.02cm）
    "subsection_after": 12,     # 二级小节段后（pt）（对齐本科模板 L3→正文 1.13cm）
    "subsubsection_before": 0,  # 三级小节段前（pt）
    "subsubsection_after": 3,   # 三级小节段后（pt）
    "first_line_indent": 480,    # 首行缩进（twips）= 2字符
    "body_line_spacing": 20.5,  # 正文行距（pt）
    "body_font_size": 12,        # 正文字号（pt）
}
