"""测试 LaTeX 公式和化学式渲染"""

import sys
sys.path.insert(0, ".")         # 项目根目录（导入 hitthesis）

from hitthesis import Thesis

doc = Thesis(type="bachelor", campus="harbin")
doc.set_info(
    title="测试 LaTeX 公式渲染",
    author="测试作者",
    supervisor="测试导师",
    subject="测试学科",
    affil="测试学院"
)

doc.add_cover()
doc.start_roman_section()
doc.add_abstract_cn("这是测试摘要内容。", ["测试", "LaTeX"])
doc.add_toc()
doc.start_arabic_section()

with doc.add_chapter("绪论", "1"):
    doc.add_section("化学式测试")
    doc.add_paragraph("水的化学式为 $H2O$。")
    doc.add_paragraph("硫酸根离子为 $SO4^2-$。")
    doc.add_paragraph("铁离子为 $Fe^{2+}$。")
    doc.add_paragraph("钙离子为 $Ca^{2+}$。")

    doc.add_section("混合测试")
    doc.add_paragraph("实验测得 $Fe^{2+}$ 浓度为 $10^{-3}$ mol/L。")
    doc.add_paragraph("纯文本不含公式的段落。")
    doc.add_paragraph("含有 $SO4^{2-}$ 和 $CO3^{2-}$ 的溶液。")

doc.compile("output/test_latex_formula.docx")
print("生成成功: output/test_latex_formula.docx")