"""物理量符号表"""

def build(doc):
    doc.add_denotation(
        items=[
            ("D", "轴承直径，mm"), ("B", "轴承宽度，mm"),
            ("h", "气膜厚度，μm"), ("p", "气体压力，Pa"),
            ("W", "承载力，N"), ("K", "刚度，N/μm"),
            ("μ", "气体动力粘度，Pa·s"), ("ρ", "气体密度，kg/m$^3$"),
        ],
        add_to_toc=True,
        title="表1　主要物理量符号及说明"
    )
