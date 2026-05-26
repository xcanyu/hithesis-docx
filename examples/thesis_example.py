#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# ==============================================================================
#                     hithesis-docx 单文件完整示例
# ==============================================================================
# 本文件在单个 Python 文件中展示所有 API 用法。
# 适合：快速上手、小论文、不想分文件的场景。
# 大型论文建议使用 examples/thesis_split/ 分文件模式。
#
# 运行方式（从项目根目录）：
#     python examples/thesis_example.py
#
# 生成位置：output/thesis_example.docx
# ==============================================================================

import sys
sys.path.insert(0, ".")         # 项目根目录（导入 hitthesis）

from hitthesis import Thesis, ReferenceDB


# ==============================================================================
# 主函数
# ==============================================================================
def main():
    # ================================================================
    # 1. 创建文档
    # ================================================================
    # Thesis() 必填参数：
    #   type="master"          学位类型（当前临时用本科封面，master/doctor 待完善）
    #   campus="harbin"        校区（当前仅实现 harbin，shenzhen/weihai 预留）
    #
    # Thesis() 选填参数：
    #   header_text=False      页眉：False|None=无页眉，"文字"=自定义，省略=默认"哈尔滨工业大学"
    # ================================================================
    doc = Thesis(type="bachelor", campus="harbin")
    # doc = Thesis(type="master", campus="harbin")                     # 硕博（预留调试）
    # doc = Thesis(type="bachelor", campus="harbin", header_text="哈工大机电") # 本科自定义页眉

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
    # 前文（frontmatter）——罗马数字页码 I, II, III...
    # ================================================================

    # ================================================================
    # 4. 封面（当前临时使用本科封面，硕博封面待完善）
    #    -- 不需要封面则注释掉下面这句
    # ================================================================
    doc.add_cover()

    # ================================================================
    # 5. 前文
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
            ("ρ", "气体密度，kg/m$^3$"),
        ],
        add_to_toc=True,
        title="表1　主要物理量符号及说明"
    )

    # ---- 目录（可选） ----
    # blank_line_before=True：目录中第一章前空一行（本科生要求）
    doc.add_toc(blank_line_before=True)

    # ================================================================
    # 正文（mainmatter）——阿拉伯数字页码 1, 2, 3...
    # ================================================================

    # ================================================================
    # 6. 正文
    # ================================================================
    doc.start_arabic_section()                  # 开始阿拉伯数字页码节

    # ------- 第1章 绪论 -------
    with doc.add_chapter("绪论", "1"):
        doc.add_section("研究背景与意义")
        doc.add_paragraph(
            "气体轴承是一种利用气体膜支承载荷的滑动轴承[ref:Gross1962]。"
            "与液体润滑轴承相比，气体轴承具有以下显著优点：无磨损、寿命长、"
            "精度保持性好；摩擦功耗极低，适用于高速运转[ref:Lee2012]。"
        )
        doc.add_paragraph(
            "然而，气体静压轴承也存在一些固有缺陷：由于气体的可压缩性和低粘度，"
            "承载能力和刚度较低是制约其广泛应用的主要瓶颈[ref:Zhang2015]。"
        )

        doc.add_section("国内外研究现状")
        doc.add_paragraph(
            "关于气体静压轴承的研究可以追溯到20世纪50年代。"
            "1952年，Gross等人首次系统研究了静压气体轴承的基本理论[ref:Gross1962]。"
        )

        doc.add_subsection("国外研究进展")
        doc.add_paragraph(
            "Blaschke等人在2005年研究了多孔质气体静压轴承的静动态特性，"
            "指出多孔质材料能够显著提高轴承的承载能力[ref:Blaschke2005]。"
        )

        doc.add_subsection("国内研究进展")
        doc.add_paragraph(
            "国内学者在气体轴承领域也开展了大量研究工作[ref:Li2018][ref:Chen2019]，"
            "取得了丰硕的成果。"
        )

        doc.add_section("本文主要研究内容")
        doc.add_paragraph("本文针对局部多孔质气体静压轴承关键技术问题，开展以下研究工作：")

        # 三线表：顶线1.5磅、中线0.5磅、底线1.5磅
        tbl = doc.add_table(4, 2, caption="本文主要研究内容", ref="tab_research")
        tbl.set_cell(0, 0, "章节", bold=False)
        tbl.set_cell(0, 1, "研究内容", bold=False)
        tbl.set_cell(1, 0, "第2章", bold=False)
        tbl.set_cell(1, 1, "理论基础与数学模型", bold=False)
        tbl.set_cell(2, 0, "第3章", bold=False)
        tbl.set_cell(2, 1, "有限元分析与优化设计", bold=False)
        tbl.set_cell(3, 0, "第4章", bold=False)
        tbl.set_cell(3, 1, "实验验证与分析", bold=False)

        # 交叉引用表格：[cite:tab_research] → "表 1-1"
        doc.add_paragraph("本文主要研究内容见如表 [cite:tab_research] 所示。")

    # ------- 第2章 理论基础 -------
    with doc.add_chapter("理论基础与数学模型", "2"):
        doc.add_section("气体润滑基本方程")
        doc.add_paragraph(
            "气体润滑的基本方程是修正的雷诺方程[ref:Gross1962]。"
            "对于可压缩气体润滑，稳态下的雷诺方程可以表示为："
        )

        # 公式：label 决定编号"2-1"，ref 决定引用名
        # 支持完整 LaTeX 公式语法，如 \frac, \partial, \rho 等
        doc.add_equation(
            r"\frac{\partial}{\partial x}\left(\frac{\rho h^3}{12\mu}\frac{\partial p}{\partial x}\right)"
            r" + \frac{\partial}{\partial y}\left(\frac{\rho h^3}{12\mu}\frac{\partial p}{\partial y}\right)"
            r" = 6\mu\frac{\partial(\rho u)}{\partial x} + 6\mu\frac{\partial(\rho v)}{\partial y}",
            label="2-1", ref="eq_reynolds"
        )

        doc.add_paragraph(
            "其中，p为气体压力，h为气膜厚度，μ为气体动力粘度，ρ为气体密度[ref:Wang2010]。"
            "由式 [cite:eq_reynolds] 可知，..."
        )

        doc.add_section("化学式与浓度表示")
        doc.add_paragraph(
            "实验中使用的硫酸铜溶液含有 $Cu^{2+}$ 和 $SO_{4}^{2-}$ 离子，"
            "浓度为 $10^{-3}$ mol/L。水的化学式为 $H_{2}O$。"
        )
        doc.add_paragraph(
            "多孔质材料选用青铜基粉末冶金材料，孔隙率为15%~25%[ref:Jones2012]。"
        )

        doc.add_section("边界条件")
        doc.add_paragraph(
            "气膜边界条件包括供气孔处的压力边界条件和气膜出口处的压力边界条件[ref:Zhang2015]。"
        )

        # 连续引用 [ref:Gross1962][ref:Lee2012][ref:Zhang2015] → [1-3]
        doc.add_paragraph(
            "关于供气压力对轴承性能的影响，已有大量研究"
            "[ref:Gross1962][ref:Lee2012][ref:Zhang2015]。"
        )
        # 非连续引用 [ref:Zhang2015][ref:Chen2019] → [1,2]
        doc.add_paragraph(
            "多孔质材料的孔径分布对轴承刚度有显著影响[ref:Zhang2015][ref:Chen2019]。"
        )

        doc.add_paragraph(
            "关于多孔质材料，国内外已有大量公开研究。",
            footnote="此处指公开发表的学术论文，不包括内部技术报告。"
        )

    # ------- 第3章 有限元分析 -------
    with doc.add_chapter("有限元分析与优化设计", "3"):
        doc.add_section("几何模型建立")
        doc.add_paragraph(
            "本文采用三维建模软件建立了局部多孔质气体静压轴承的几何模型[ref:Chen2019]。"
        )

        # 图片：自动编号"图3-1"，ref 决定引用名
        doc.add_figure("examples/fig_sample.png", caption="实验结果对比", ref="fig1")
        doc.add_paragraph("有限元分析结果如图 [cite:fig1] 所示。")

        doc.add_subsection("轴承基本参数")
        doc.add_paragraph(
            "轴承的主要结构参数如下：节流孔直径d=0.3mm，节流孔数n=8，"
            "轴承直径D=50mm，轴承宽度B=40mm，气膜间隙h0=0.02mm[ref:Li2018]。"
        )

        doc.add_subsection("材料选择")
        doc.add_paragraph(
            "多孔质材料选用青铜基粉末冶金材料，孔隙率为15%~25%[ref:Jones2012]。"
        )
        # 混合引用：连续+非连续 → [1,2,8,9]
        doc.add_paragraph(
            "数值模拟结果与理论预测一致"
            "[ref:Gross1962][ref:Lee2012][ref:Smith2010][ref:Jones2012]，"
            "验证了模型的准确性。"
        )

        doc.add_section("网格划分与求解")
        doc.add_paragraph(
            "采用ANSYS Workbench进行网格划分和有限元求解[ref:Wang2020]。"
        )

        doc.add_subsection("网格质量评价")
        doc.add_paragraph(
            "网格划分完成后进行了质量检查，确保计算精度[ref:Smith2010]。"
        )

    # ================================================================
    # 后文（backmatter）——按学位类型选择书序
    # ================================================================


    # ================================================================
    # 本科书序（哈尔滨、深圳校区）
    #   7. 结论（可选） → 8. 参考文献 → 11. 授权 → 12. 致谢 → 9. 附录
    # ================================================================

    # ================================================================
    # 7. 结论（可选）  标题两汉字自动加全角空格"结　论"
    # ================================================================
    doc.add_conclusion(
        title="结论",
        content=[
            "本文针对局部多孔质气体静压轴承的关键技术问题[ref:Zhang2015][ref:Chen2019]，"
            "开展了系统的理论分析、数值模拟和实验研究工作，取得了以下主要成果：",
            "（1）建立了局部多孔质气体静压轴承的完整理论模型，揭示了多孔质结构对轴承静动态性能的影响规律。",
            "（2）通过有限元分析和优化设计，确定了最优的多孔质分布方案，显著提高了轴承的承载能力和刚度"
            "[ref:Blaschke2005][ref:Jones2012]。",
            "（3）设计并加工了实验样机，完成了静态和动态性能测试，验证了理论分析的正确性[ref:Wang2015book]。",
            "研究成果对于推动气体静压轴承在高精度机床和航空航天领域的应用具有重要的理论价值和工程意义"
            "[ref:Zhao2016][ref:Gross1962][ref:Zhang2015][ref:Wang2010][ref:Blaschke2005]"
            "[ref:Li2018][ref:Jones2012][ref:Wang2020][ref:Smith2010]。"
        ]
    )

    # ================================================================
    # 8. 参考文献（需要 bib 数据库 + 正文中有 [ref:key] 引用）
    # ================================================================
    doc.add_references(bib)

    # ================================================================
    # 11. 授权声明（当前为本科格式，硕博待完善）
    # ================================================================
    doc.add_authorization()

    # ================================================================
    # 12. 致谢（可选）
    # ================================================================
    doc.add_acknowledgements(
        content=[
            "在本文即将完成之际，我首先要衷心感谢我的导师某某某教授。"
            "在三年的研究生学习期间，导师在学术上给予我悉心指导，在生活上给予我热忱关怀。",
            "感谢实验室的各位师兄弟在科研工作中给予的帮助和支持。"
            "感谢所有关心和帮助过我的老师和同学。",
            "最后，特别感谢我的家人对我学业的理解和支持。"
        ]
    )

    # ================================================================
    # 9. 附录（可选）
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
    # 本科书序（威海校区）
    #   7. 结论（可选） → 11. 授权 → 8. 参考文献 → 12. 致谢 → 9. 附录
    # ================================================================
    # doc.add_conclusion(...)
    # doc.add_authorization()
    # doc.add_references(bib)
    # doc.add_acknowledgements(...)
    # with doc.add_appendix("符号"): ...
    # with doc.add_appendix("推导"): ...


    # ================================================================
    # 硕博书序
    #   7. 结论（可选） → 8. 参考文献 → 9. 附录 → 10. 发表文章 → 索引 → 11. 授权 → 12. 致谢 → 13. 简历
    # ================================================================
    # doc.add_conclusion(...)
    # doc.add_references(bib)
    # with doc.add_appendix("符号"): ...
    # with doc.add_appendix("推导"): ...
    # doc.add_publications(...)      # 发表文章页（可选——硕士/博士）
    # # doc.add_index(...)           # 索引（预留）
    # doc.add_authorization()
    # doc.add_acknowledgements(...)
    # doc.add_resume(...)            # 个人简历（可选——博士）


    # ================================================================
    # 博后书序
    #   7. 结论（可选） → 8. 参考文献 → 12. 致谢 → 10. 发表文章 → 13. 简历 → 通信地址
    # ================================================================
    # doc.add_conclusion(...)
    # doc.add_references(bib)
    # doc.add_acknowledgements(...)
    # doc.add_publications(...)
    # doc.add_resume(...)
    # # doc.add_correspondence_addr(...)  # 通信地址（预留）


    # ================================================================
    # 14. 编译输出
    # ================================================================
    doc.compile("output/thesis_example.docx")
    # print("=" * 50)
    # print("示例文档已生成：output/thesis_example.docx")
    # print("=" * 50)


# ==============================================================================
# ==============================================================================
# 书序参考
# ---- 本科（哈尔滨/深圳）：封面 → 摘要 → 目录 → ... → 结论 → 参考文献 → 授权 → 致谢 → 附录
# ---- 本科（威海）：       封面 → 摘要 → 目录 → ... → 结论 → 授权 → 参考文献 → 致谢 → 附录
# ---- 硕士/博士：          封面 → 摘要 → 目录 → ... → 结论 → 参考文献 → 附录 → 发表文章 → 索引 → 授权 → 致谢 → 简历
# ==============================================================================

if __name__ == "__main__":
    main()
