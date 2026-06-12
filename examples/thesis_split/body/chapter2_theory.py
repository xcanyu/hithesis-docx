"""第2章 理论基础与数学模型"""

def build(doc):
    with doc.add_chapter("理论基础与数学模型", "2"):
        doc.add_section("气体润滑基本方程")
        doc.add_paragraph("气体润滑的基本方程是修正的雷诺方程[ref:Gross1962]。对于可压缩气体润滑，稳态下的雷诺方程可以表示为：")

        # 公式：支持完整 LaTeX 公式语法
        doc.add_equation(r"\frac{\partial}{\partial x}\left(\frac{\rho h^3}{12\mu}\frac{\partial p}{\partial x}\right) + \frac{\partial}{\partial y}\left(\frac{\rho h^3}{12\mu}\frac{\partial p}{\partial y}\right) = 6\mu\frac{\partial(\rho u)}{\partial x} + 6\mu\frac{\partial(\rho v)}{\partial y}", label="2-1", ref="eq_reynolds")

        doc.add_paragraph("其中，p为气体压力，h为气膜厚度，μ为气体动力粘度，ρ为气体密度[ref:Wang2010]。由式 [cite:eq_reynolds] 可知，...")

        doc.add_section("化学式与浓度表示")
        doc.add_paragraph("实验中使用的硫酸铜溶液含有 $Cu^{2+}$ 和 $SO_{4}^{2-}$ 离子，浓度为 $10^{-3}$ mol/L。水的化学式为 $H_{2}O$。")

        doc.add_subsection("国外研究进展")
        doc.add_paragraph(
            "Blaschke等人在2005年研究了多孔质气体静压轴承的静动态特性，"
            "指出多孔质材料能够显著提高轴承的承载能力[ref:Blaschke2005]。"
        )

        doc.add_subsection("国内研究进展")
        doc.add_paragraph("国内学者在气体轴承领域也开展了大量研究工作[ref:Li2018][ref:Chen2019]，取得了丰硕的成果。")

        doc.add_section("定理与引理")
        # 定理环境：add_theorem()，支持定理、定义、推论等类型
        doc.add_theorem(
            "气体静压轴承的稳态承载力等于气膜压力在承载面上的积分。"
            "对于不可压缩流体，承载力与供气压力呈线性关系。",
            kind="定理", ref="thm_bearing", cite="[ref:Gross1962]"
        )
        doc.add_theorem(
            "气体静压轴承是依靠外部供气压力在轴承间隙中形成承载气膜的"
            "一种滑动轴承，其工作原理基于流体动压效应与外部静压效应的耦合。",
            kind="定义", ref="def_bearing"
        )
        doc.add_theorem(
            "由定理 [cite:thm_bearing] 可知，提高供气压力是增大轴承承载力的有效途径。"
            "当供气压力从 0.3 MPa 提升至 0.5 MPa 时，理论承载力可提高约 40%。",
            kind="推论"
        )
        doc.add_paragraph("定理 [cite:thm_bearing] 和定义 [cite:def_bearing] 为本文的理论分析奠定了基础。")

        doc.add_section("边界条件")
        doc.add_paragraph("气膜边界条件包括供气孔处的压力边界条件和气膜出口处的压力边界条件[ref:Zhang2015]。")

        doc.add_paragraph("关于供气压力对轴承性能的影响，已有大量研究[ref:Gross1962][ref:Lee2012][ref:Zhang2015]。")
        doc.add_paragraph(
            "多孔质材料的孔径分布对轴承刚度有显著影响[ref:Zhang2015][ref:Chen2019]。"
        )

        doc.add_paragraph(
            "关于多孔质材料，国内外已有大量公开研究。",
            footnote="此处指公开发表的学术论文，不包括内部技术报告。"
        )
