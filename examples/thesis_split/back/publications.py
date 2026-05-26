"""发表文章页"""

def build(doc):
    doc.add_publications(sections=[
        ("（一）发表的学术论文", [
            "Yu D M, Zhang X X. Static Oxidation Model of Al-Mg/C Dissipation Thermal Protection Materials[J]. Rare Metal Materials and Engineering, 2010, 39(Suppl. 1): 520-524.（SCI收录，IDS号为669JS，IF=0.16）",
            "Yu D M, Li W. 精密超声振动切削单晶铜的计算机仿真研究[J]. 系统仿真学报, 2007, 19(4): 738-741, 753.（EI收录号：20071310514841）",
            "Yu D M, Wang H. 局部多孔质气体静压轴向轴承静态特性的数值求解[J]. 摩擦学学报, 2007(1): 68-72.（EI收录号：20071510544816）",
        ]),
        ("（二）申请及已获得的专利", [
            "Yu D M, Zhang X X. 一种温热外敷药制备方案：中国，88105607.3[P]. 1989-07-26.",
        ]),
    ])
