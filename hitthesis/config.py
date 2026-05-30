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

# 字体定义
FONTS = {
    "cover_title": ("宋体", "24bp", True), # 封面第一行（小一）
    "cover_info": ("宋体", "26bp"),       # 封面信息
    "cover_label": ("黑体", "26bp"),      # 封面标签(学校名称等)
    "normal": ("宋体", "12bp"),           # 正文
    "bold": ("宋体", "12bp", True),      # 粗体正文
    "chapter": ("黑体", "18bp"),          # 章标题 一级
    "section": ("黑体", "15bp"),          # 节标题 二级
    "subsection": ("黑体", "14bp", True), # 小节标题 三级
    "figure_caption": ("宋体", "10.5bp"), # 图题
    "table_caption": ("宋体", "10.5bp"),  # 表题
    "abstract_title": ("黑体", "18bp"),   # 摘要标题
    "abstract": ("宋体", "12bp"),         # 摘要正文
    "toc_title": ("黑体", "18bp"),        # 目录标题
    "header": ("宋体", "9bp"),            # 页眉 小五
    "footer": ("Times New Roman", "9bp"), # 页脚 小五
    "footnote": ("宋体", "9bp"),          # 脚注
    "reference": ("宋体", "10.5bp"),     # 参考文献
}

# 行距设置
LINE_SPACING = {
    "normal": {"line": 20.5, "lineRule": "auto", "before": 0, "after": 0},  # 1.3倍行距约20.5pt
    "tight": {"line": 18, "lineRule": "auto", "before": 0, "after": 0},     # 紧凑行距
    "relaxed": {"line": 24, "lineRule": "auto", "before": 0, "after": 6},   # 宽松行距(段后6pt)
    "chapter": {"line": 28, "lineRule": "exact", "before": 24, "after": 18}, # 章标题行距
}

# 段落间距
PARAGRAPH_SPACING = {
    "chapter_before": 24,    # 章标题段前
    "chapter_after": 18,      # 章标题段后
    "section_before": 12,    # 节标题段前
    "section_after": 6,       # 节标题段后
    "normal_before": 0,       # 正文段前
    "normal_after": 0,       # 正文段后
    "caption_before": 6,     # 题注段前
    "caption_after": 6,      # 题注段后
}

# 缩进
INDENTS = {
    "normal": 0,                      # 正文无缩进
    "first_line": 21.84,             # 首行缩进2字符(约21.84pt)
    "table_cell": 3.5,               # 表格单元格内缩进
}

# 边框
BORDERS = {
    "none": {"val": "none", "sz": 0, "space": 0, "color": "auto"},
    "single": {"val": "single", "sz": 4, "space": 0, "color": "000000"},
    "thick": {"val": "single", "sz": 8, "space": 0, "color": "000000"},
    "table_top": {"val": "single", "sz": 12, "space": 0, "color": "000000"},  # 三线表顶线
    "table_bottom": {"val": "single", "sz": 12, "space": 0, "color": "000000"}, # 三线表底线
    "table_mid": {"val": "single", "sz": 8, "space": 0, "color": "000000"},    # 三线表中线 1磅
}

# 表格样式
TABLE = {
    "col_width_auto": True,           # 自动列宽
    "cell_margin": {"top": 0, "bottom": 0, "left": 85, "right": 85},  # 单元格边距(单位: twips)
    "first_row_font": ("宋体", "10.5bp", True),  # 表头字体
    "data_font": ("宋体", "10.5bp", False),     # 表数据字体
}

# 封面信息项顺序
COVER_ITEMS = [
    "学校代码", "密级", "学号", "分类号",
    "答辩日期", "学校名称", "论文题目", "院系",
    "学科专业", "作者姓名", "导师姓名", "副导师姓名",
]

# 摘要设置
ABSTRACT = {
    "title_cn": "摘  要",
    "title_en": "Abstract",
    "keywords_cn": "关键词：",
    "keywords_en": "Key words:",
    "max_keywords": 5,
}

# 目录设置
TOC = {
    "title_cn": "目  录",
    "title_en": "Contents",
    "levels": 3,  # 显示到三级标题
}

# 论文类型（当前仅实现 bachelor，master/doctor 预留）
THESIS_TYPES = {
    "bachelor": "本科毕业设计",
    "master": "硕士学位论文（预留）",
    "doctor": "博士学位论文（预留）",
}

# 校区
CAMPUSES = {
    "harbin": "哈尔滨",
    "shenzhen": "深圳",
    "weihai": "威海",
}

# 学位（master/doctor 预留）
DEGREES = {
    "bachelor": "学士",
    "master": "硕士（预留）",
    "doctor": "博士（预留）",
}

# 封面标题（master/doctor 预留）
COVER_TITLES = {
    "bachelor": "哈尔滨工业大学本科毕业设计",
    "master": "哈尔滨工业大学硕士学位论文（预留）",
    "doctor": "哈尔滨工业大学博士学位论文（预留）",
}

# 学校名称
UNIVERSITY_NAME = "哈尔滨工业大学"

# 标题段落间距（twips，1pt = 20twips）
SPACING = {
    "heading_before": 454,      # 章标题段前（twips）= 22.7pt（双线到标题间距1cm）
    "heading_after": 400,       # 章标题段后（twips）= 20pt（精确匹配官方Word模板1.68cm）
    "heading_line": 288,        # 章标题行距（twips）= 1.2倍（测试：减小行距看效果）
    "section_before": 10,       # 一级小节段前（pt）
    "section_after": 8,         # 一级小节段后（pt）（微调使L2→正文和正文→L2统一为1.04cm）
    "subsection_before": 7,     # 二级小节段前（pt）（统一正文→L3为0.98cm）
    "subsection_after": 9,      # 二级小节段后（pt）（官方L3→L4=0.98cm，当前0.77cm，增加6pt）
    "subsubsection_before": 0,  # 三级小节段前（pt）（官方正文→四级0.66cm，当前0.83cm，需减小）
    "subsubsection_after": 3,   # 三级小节段后（pt）（官方四级→正文0.98cm，当前0.98cm，已匹配）
    "first_line_indent": 480,    # 首行缩进（twips）= 2字符
    "body_line_spacing": 20.5,  # 正文行距（pt）
    "body_font_size": 12,        # 正文字号（pt）
}

HEADING_FONTS = {
    "chapter": ("黑体", 18, True),
    "section": ("黑体", 15, True),
    "subsection": ("黑体", 14, True),
    "subsubsection": ("黑体", 14, True),
}
