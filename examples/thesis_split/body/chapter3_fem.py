"""第3章 有限元分析与优化设计"""

def build(doc):
    with doc.add_chapter("有限元分析与优化设计", "3"):
        doc.add_section("几何模型建立")
        doc.add_paragraph("本文采用三维建模软件建立了局部多孔质气体静压轴承的几何模型[ref:Chen2019]。")

        doc.add_figure(
            "examples/fig/fig_sample.jpg",
            caption="砂箱建模图",
            ref="fig1"
        )

        doc.add_paragraph("砂箱结构如图 [cite:fig1] 所示。")

        doc.add_subsection("轴承基本参数")
        doc.add_paragraph("轴承的主要结构参数如下：节流孔直径d=0.3mm，节流孔数n=8，轴承直径D=50mm，轴承宽度B=40mm，气膜间隙h0=0.02mm[ref:Li2018]。")

        doc.add_subsection("材料选择")
        doc.add_paragraph("多孔质材料选用青铜基粉末冶金材料，孔隙率为15%~25%[ref:Jones2012]。")
        doc.add_paragraph("数值模拟结果与理论预测一致[ref:Gross1962][ref:Lee2012][ref:Smith2010][ref:Jones2012]，验证了模型的准确性。")

        doc.add_section("网格划分与求解")
        doc.add_paragraph("采用ANSYS Workbench进行网格划分和有限元求解[ref:Wang2020]。")

        doc.add_subsection("网格质量评价")
        doc.add_paragraph("网格划分完成后进行了质量检查，确保计算精度[ref:Smith2010]。")
