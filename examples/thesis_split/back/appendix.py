"""附录"""

def build(doc):
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
