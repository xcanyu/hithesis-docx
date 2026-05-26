#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# ==============================================================================
#                         hithesis-docx 分文件示例
# ==============================================================================
# 主控文件，只负责编排文档结构。
# 内容分散在 front/（前文）、body/（正文）、back/（后文）中。
#
# 运行方式（从项目根目录）：
#     python examples/thesis_split/thesis_main.py
#
# 生成位置：output/thesis_split.docx（项目根目录下）
# ==============================================================================

import sys
sys.path.insert(0, ".")
sys.path.insert(0, "examples/thesis_split")

from hitthesis import Thesis, ReferenceDB

# ==============================================================================
# 文档类选项
#
# 必填参数：
#   type=doctor|master|bachelor|postdoc
#   campus=harbin|shenzhen|weihai
#
# 选填参数（缺省值已满足多数需求）：
#   header_text=False|None|"文字"  页眉：False/None=无，省略=默认"哈尔滨工业大学"
#   toc_blank_line=True|False      目录第一章前是否空一行
#   ...（更多参数参考 hitthesis 文档）
# ==============================================================================
doc = Thesis(type="bachelor", campus="harbin")
# doc = Thesis(type="master", campus="harbin")  # 硕博（预留调试）
# doc = Thesis(type="bachelor", campus="harbin", header_text="哈工大机电")  # 本科自定义页眉

# 参考文献数据库
doc.bib = ReferenceDB("examples/references.bib")
doc.set_reference_db(doc.bib)

# ==============================================================================
# 前文（frontmatter）
#   封面 → 中文摘要 → 英文摘要 → 符号表 → 目录
# ==============================================================================
doc.include("front.cover")          # 封面（含 set_info 元信息）
doc.start_roman_section()           # 开始罗马数字页码 I, II, III...
doc.include("front.abstract_cn")    # 中文摘要（可选）
doc.include("front.abstract_en")    # 英文摘要（可选）
doc.include("front.denotation")     # 物理量符号表（可选）
doc.add_toc(blank_line_before=True) # 目录（blank_line_before：第一章前空一行，本科生要求）

# ==============================================================================
# 正文（mainmatter）
#   阿拉伯数字页码 1, 2, 3...
# ==============================================================================
doc.start_arabic_section()
doc.include("body.chapter1_intro")   # 第1章 绪论
doc.include("body.chapter2_theory")  # 第2章 理论基础与数学模型
doc.include("body.chapter3_fem")     # 第3章 有限元分析与优化设计

# ==============================================================================
# ==============================================================================
# 以下按学位类型选择书序，不需要的整段注释/取消注释
# ==============================================================================
# ==============================================================================


# ==============================================================================
# 本科书序（哈尔滨、深圳校区）
#   结论 → 参考文献 → 授权 → 致谢 → 附录
# ==============================================================================
doc.include("back.conclusion")      # 结论
doc.include("back.references")      # 参考文献
doc.include("back.authorization")   # 授权声明
doc.include("back.acknowledgements")# 致谢
doc.include("back.appendix")        # 附录

# ==============================================================================
# 本科书序（威海校区）
#   结论 → 授权 → 参考文献 → 致谢 → 附录
# ==============================================================================
# doc.include("back.conclusion")
# doc.include("back.authorization")
# doc.include("back.references")
# doc.include("back.acknowledgements")
# doc.include("back.appendix")

# ==============================================================================
# 硕博书序
#   结论 → 参考文献 → 附录 → 发表文章 → 索引 → 授权 → 致谢 → 简历
# ==============================================================================
# doc.include("back.conclusion")
# doc.include("back.references")
# doc.include("back.appendix")
# doc.include("back.publications")   # 发表文章页
# # doc.include("back.index")       # 索引（预留）
# doc.include("back.authorization")
# doc.include("back.acknowledgements")
# doc.include("back.resume")         # 个人简历（博士）

# ==============================================================================
# 博后书序
#   结论 → 参考文献 → 致谢 → 发表文章 → 简历 → 通信地址
# ==============================================================================
# doc.include("back.conclusion")
# doc.include("back.references")
# doc.include("back.acknowledgements")
# doc.include("back.publications")
# doc.include("back.resume")
# # doc.include("back.correspondence_addr")  # 通信地址（预留）

# ==============================================================================
# 编译输出
# ==============================================================================
doc.compile("output/thesis_split.docx")
# print("分文件示例生成完成：output/thesis_split.docx")
