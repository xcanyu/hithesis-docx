"""第1章 绪论"""

def build(doc):
    with doc.add_chapter("绪论", "1"):
        doc.add_section("研究背景与意义")
        doc.add_paragraph("气体轴承是一种利用气体膜支承载荷的滑动轴承[ref:Gross1962]。与液体润滑轴承相比，气体轴承具有以下显著优点：无磨损、寿命长、精度保持性好；摩擦功耗极低，适用于高速运转[ref:Lee2012]。")
        doc.add_paragraph("然而，气体静压轴承也存在一些固有缺陷：由于气体的可压缩性和低粘度，承载能力和刚度较低是制约其广泛应用的主要瓶颈[ref:Zhang2015]。")

        doc.add_section("国内外研究现状")
        doc.add_paragraph("关于气体静压轴承的研究可以追溯到20世纪50年代。1952年，Gross等人首次系统研究了静压气体轴承的基本理论[ref:Gross1962]。")

        doc.add_subsection("国外研究进展")
        doc.add_paragraph("Blaschke等人在2005年研究了多孔质气体静压轴承的静动态特性，指出多孔质材料能够显著提高轴承的承载能力[ref:Blaschke2005]。")

        doc.add_subsection("国内研究进展")
        doc.add_paragraph("国内学者在气体轴承领域也开展了大量研究工作[ref:Li2018][ref:Chen2019]，取得了丰硕的成果。")

        doc.add_section("本文主要研究内容")
        doc.add_paragraph("本文针对局部多孔质气体静压轴承关键技术问题，开展以下研究工作：")

        tbl = doc.add_table(4, 2, caption="本文主要研究内容", ref="tab_research")
        tbl.set_cell(0, 0, "章节", bold=False)
        tbl.set_cell(0, 1, "研究内容", bold=False)
        tbl.set_cell(1, 0, "第2章", bold=False)
        tbl.set_cell(1, 1, "理论基础与数学模型", bold=False)
        tbl.set_cell(2, 0, "第3章", bold=False)
        tbl.set_cell(2, 1, "有限元分析与优化设计", bold=False)
        tbl.set_cell(3, 0, "第4章", bold=False)
        tbl.set_cell(3, 1, "实验验证与分析", bold=False)

        doc.add_paragraph("本文主要研究内容见如表 [cite:tab_research] 所示。")
