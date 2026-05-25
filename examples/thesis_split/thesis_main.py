#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# ==============================================================================
#                         hithesis-docx 分文件示例
# ==============================================================================
# 本文件是学位论文的"主控文件"，负责编排文档结构、导入各章节模块。
# 与 LaTeX 的 thesis.tex 作用相同：\include、\input → doc.include()
#
# 运行方式（从项目根目录）：
#     python examples/thesis_split/thesis_main.py
#
# 生成位置：examples/thesis_split/output/thesis_split.docx
# ==============================================================================

import sys
sys.path.insert(0, ".")                       # 项目根目录（导入 hitthesis）
sys.path.insert(0, "examples/thesis_split")   # 本目录（导入 body.*）

from hitthesis import Thesis, ReferenceDB


# ==============================================================================
# 主函数
# ==============================================================================
def main():
    # ================================================================
    # 1. 创建文档
    # ================================================================
    # Thesis() 必填参数：
    #   type="bachelor"        学位类型（当前仅实现 bachelor，master/doctor 预留）
    #   campus="harbin"        校区（当前仅实现 harbin，shenzhen/weihai 预留）
    #
    # Thesis() 选填参数：
    #   header_text=False      页眉：False|None=无页眉，"文字"=自定义，省略=默认"哈尔滨工业大学"
    # ================================================================
    doc = Thesis(type="bachelor", campus="harbin")
    # doc = Thesis(type="bachelor", campus="harbin", header_text=False)       # 无页眉
    # doc = Thesis(type="bachelor", campus="harbin", header_text="哈工大机电") # 自定义页眉

    # ================================================================
    # 2. 论文元信息（必填：title, author, supervisor, date）
    # ================================================================
    doc.set_info(
        title="局部多孔质气体静压轴承关键技术的研究",  # 必填
        author="于冬梅",                                # 必填
        supervisor="某某某教授",                         # 必填
        subject="机械制造及其自动化",                    # 学科专业（封面用）
        affil="机电工程学院",                           # 院系（封面用）
        date="2024年6月12日",                            # 日期（封面用）
        student_id="1234567890",                        # 学号（本科cover2表格用）
    )

    # ================================================================
    # 3. 参考文献数据库（可选——不需要文献引用则可删除整段）
    # ================================================================
    bib = ReferenceDB("examples/references.bib")
    doc.set_reference_db(bib)

    # ================================================================
    # 4. 封面（本科自动生成双封面：封面+学生信息页）
    #    -- 不需要封面则注释掉下面这句
    # ================================================================
    doc.add_cover()

    # ================================================================
    # 5. 前文（罗马数字页码 I, II, III...）
    #    以下各项不需要哪个就注释/删除哪个
    # ================================================================
    doc.start_roman_section()                   # 开始罗马数字页码节

    # ---- 中文摘要（可选） ----
    doc.add_abstract_cn(
        """本文针对局部多孔质气体静压轴承关键技术进行了深入研究。气体静压轴承以其无接触、高精度、高转速等优点，在精密机床、超高速离心机等高端装备中得到了广泛应用。然而，传统的气体静压轴承存在承载能力低、刚度不足等问题，限制了其在重载条件下的应用。

本文首先分析了局部多孔质气体静压轴承的工作原理，建立了完整的理论模型，通过数值模拟研究了孔穴分布、供气压力等关键参数对轴承性能的影响规律。在此基础上，设计并加工了实验样机，开展了系统的静态和动态性能测试，验证了理论分析的正确性。

研究表明，局部多孔质结构能够有效提高轴承的承载能力和刚度，同时保持较低的功耗和温升。研究成果对于推动气体静压轴承在高精度机床和航空航天领域的应用具有重要的理论价值和工程意义。""",
        ["气体轴承", "静压轴承", "多孔质材料", "有限元分析", "实验研究"]
    )

    # ---- 英文摘要（可选） ----
    doc.add_abstract_en(
        """This dissertation presents an in-depth investigation on the key technologies of partial porous externally pressurized gas bearing. Gas bearings, with their advantages of non-contact operation, high precision, and high speed, have been widely applied in precision machine tools and ultra-high-speed centrifugal equipment.

However, traditional gas bearings suffer from low load capacity and insufficient stiffness, which limit their applications under heavy loads. This research first analyzes the working principle of partial porous gas bearings and establishes a comprehensive theoretical model.

Based on the theoretical analysis, an experimental prototype is designed and fabricated. Systematic static and dynamic performance tests are carried out to verify the theoretical predictions.""",
        ["gas bearing", "externally pressurized bearing", "porous material", "FEM analysis", "experimental research"]
    )

    # ---- 物理量符号表（可选） ----
    doc.add_denotation(
        items=[
            ("D", "轴承直径，mm"),
            ("B", "轴承宽度，mm"),
            ("h", "气膜厚度，μm"),
            ("p", "气体压力，Pa"),
            ("W", "承载力，N"),
            ("K", "刚度，N/μm"),
            ("μ", "气体动力粘度，Pa·s"),
            ("ρ", "气体密度，kg/m³"),
        ],
        add_to_toc=True,
        title="表1　主要物理量符号及说明"
    )

    # ---- 目录（可选） ----
    # blank_line_before=True：目录中第一章前空一行（本科生要求）
    doc.add_toc(blank_line_before=True)

    # ================================================================
    # 6. 正文（阿拉伯数字页码 1, 2, 3...）
    #    章节拆分为独立 Python 文件，由 doc.include() 导入
    #
    #    章节文件约定：
    #      - 文件名任意，模块路径用点号分隔（body.xxx）
    #      - 必须导出 build(doc) 函数，接收 Thesis 实例
    #      - 内部 API 与单文件写法完全一致
    #
    #    要增减章节：增减 include() 调用 + 对应的章节文件
    # ================================================================
    doc.start_arabic_section()                  # 开始阿拉伯数字页码节

    doc.include("body.chapter1_intro")          # 第1章 绪论
    doc.include("body.chapter2_theory")         # 第2章 理论基础与数学模型
    doc.include("body.chapter3_fem")            # 第3章 有限元分析与优化设计

    # ================================================================
    # 7. 结论（可选——不需要则注释掉）  标题两汉字自动加全角空格"结　论"
    # ================================================================
    doc.add_conclusion(
        title="结论",
        content=[
            "本文针对局部多孔质气体静压轴承的关键技术问题[ref:Zhang2015][ref:Chen2019]，开展了系统的理论分析、数值模拟和实验研究工作，取得了以下主要成果：",
            "（1）建立了局部多孔质气体静压轴承的完整理论模型，揭示了多孔质结构对轴承静动态性能的影响规律。",
            "（2）通过有限元分析和优化设计，确定了最优的多孔质分布方案，显著提高了轴承的承载能力和刚度[ref:Blaschke2005][ref:Jones2012]。",
            "（3）设计并加工了实验样机，完成了静态和动态性能测试，验证了理论分析的正确性[ref:Wang2015book]。",
            "研究成果对于推动气体静压轴承在高精度机床和航空航天领域的应用具有重要的理论价值和工程意义[ref:Zhao2016][ref:Gross1962][ref:Zhang2015][ref:Wang2010][ref:Blaschke2005][ref:Li2018][ref:Jones2012][ref:Wang2020][ref:Smith2010]。"
        ]
    )

    # ================================================================
    # 8. 参考文献（需要 bib 数据库 + 正文中有 [ref:key] 引用）
    #    -- 不需要则注释掉
    # ================================================================
    doc.add_references(bib)

    # ================================================================
    # 9. 授权声明（本科专属——硕博删除）  标题两汉字自动加全角空格"授　权"
    # ================================================================
    doc.add_authorization()

    # ================================================================
    # 10. 致谢（可选——不需要则注释掉）  标题两汉字自动加全角空格"致　谢"
    # ================================================================
    doc.add_acknowledgements(
        content=[
            "在本文即将完成之际，我首先要衷心感谢我的导师某某某教授。在三年的研究生学习期间，导师在学术上给予我悉心指导，在生活上给予我热忱关怀。",
            "感谢实验室的各位师兄弟在科研工作中给予的帮助和支持。感谢所有关心和帮助过我的老师和同学。",
            "最后，特别感谢我的家人对我学业的理解和支持。"
        ]
    )

    # ================================================================
    # 11. 附录（可选——不需要则注释掉）
    # ================================================================
    with doc.add_appendix("符号"):
        doc.add_section("主要符号")
        doc.add_paragraph("D —— 轴承直径，mm")
        doc.add_paragraph("B —— 轴承宽度，mm")
        doc.add_paragraph("h —— 气膜厚度，μm")
        doc.add_paragraph("p —— 气体压力，Pa")
        doc.add_paragraph("W —— 承载力，N")
        doc.add_paragraph("K —— 刚度，N/μm")

    with doc.add_appendix("推导"):
        doc.add_section("雷诺方程推导")
        doc.add_paragraph("从 Navier-Stokes 方程出发，推导气体润滑的雷诺方程...")

    # ================================================================
    # 12. 编译输出
    # ================================================================
    doc.compile("output/thesis_split.docx")
    print("分文件示例生成完成：output/thesis_split.docx")


# ==============================================================================
# ==============================================================================
# 其他学位类型书序参考（当前仅实现本科本部，以下为预留规划）
# ---- 本科（哈尔滨/深圳校区）：封面 → 摘要 → 目录 → ... → 结论 → 参考文献 → 授权 → 致谢 → 附录
# ---- 本科（威海校区）：      封面 → 摘要 → 目录 → ... → 结论 → 授权 → 参考文献 → 致谢 → 附录
# ---- 硕士（未实现）：        封面 → 摘要 → 目录 → ... → 结论 → 参考文献 → 附录 → 授权 → 致谢
# ---- 博士（未实现）：        封面 → 摘要 → 目录 → ... → 结论 → 参考文献 → 附录 → 授权 → 致谢 → 简历
# ==============================================================================

if __name__ == "__main__":
    main()
