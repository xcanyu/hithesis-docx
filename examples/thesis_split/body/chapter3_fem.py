"""第3章 有限元分析与优化设计"""

def build(doc):
    with doc.add_chapter("有限元分析与优化设计", "3"):
        doc.add_section("几何模型建立")
        doc.add_paragraph("本文采用三维建模软件建立了局部多孔质气体静压轴承的几何模型[ref:Chen2019]。这里顺便截了一张图，有点丑。")

        # 图片：自动编号"图3-1"，ref 决定引用名
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

        doc.add_section("子图示例")
        doc.add_paragraph(
            "子图功能仅供测试，效果有一定局限性，不具备跨页检测功能，子图功能使用 add_subfigure() 方法，支持并排和网格布局。"
            "每个子图有 (a)(b)(c) 标号，总图题在表格下方居中。"
        )
        # 子图：add_subfigure()，并排布局（默认 cols=None）
        doc.add_subfigure(
            subfigures=[
                ("examples/fig/fig_sample.jpg", "局部多孔质结构"),
                ("examples/fig/fig_sample.jpg", "气体流动示意"),
            ],
            caption="子图并排布局示例",
            ref="fig_sub1"
        )
        doc.add_paragraph("子图并排布局如图 [cite:fig_sub1] 所示。")

        # 子图：网格布局（cols=2）
        doc.add_subfigure(
            subfigures=[
                ("examples/fig/fig_sample.jpg", "压力分布"),
                ("examples/fig/fig_sample.jpg", "温度分布"),
                ("examples/fig/fig_sample.jpg", "位移分布"),
                ("examples/fig/fig_sample.jpg", "应力分布"),
            ],
            caption="子图网格布局示例（2×2）",
            ref="fig_sub2",
            cols=2
        )
        doc.add_paragraph("子图网格布局如图 [cite:fig_sub2] 所示。")

        doc.add_section("代码块示例")
        doc.add_paragraph(
            "代码块功能使用 add_code_block() 方法，采用 Consolas 9pt 灰底样式。"
            "注：作者不懂代码块的标准格式，这部分也没有规范格式，加入这个只是测试功能。"
        )
        # 代码块：add_code_block()
        doc.add_code_block(
            "import numpy as np\n"
            "\n"
            "def calculate_bearing_capacity(pressure, area):\n"
            "    # 计算轴承承载力\n"
            "    return pressure * area\n"
            "\n"
            "# 参数设置\n"
            "p_supply = 0.5e6  # 供气压力 0.5 MPa\n"
            "D = 50e-3         # 轴承直径 50 mm\n"
            "area = np.pi * (D/2)**2\n"
            "\n"
            "# 计算承载力\n"
            "W = calculate_bearing_capacity(p_supply, area)\n"
            "print(f'承载力: {W:.2f} N')",
            caption="Python 计算脚本示例",
            ref="code1"
        )
        doc.add_paragraph("代码块示例如代码 [cite:code1] 所示。")
