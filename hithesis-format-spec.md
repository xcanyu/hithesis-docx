# 哈工大学位论文格式规范深度解析

> 源码分析来源：[hithesis](https://github.com/hithesis/hithesis) v3.1d（hithesisbook.cls, 2025/03/03）

---

## 一、页面设置

### 1.1 纸张与版芯尺寸

| 参数 | 值（默认 newgeometry=two） | 说明 |
|------|--------------------------|------|
| 纸张 | A4 210mm × 297mm | |
| 正文宽度 | 150mm | |
| 正文高度 | 236mm | |
| 左边距 | 30mm | |
| 页眉高度 | 5mm | |
| 页眉与正文间距（headsep） | 2mm | |
| 页脚下边缘到版芯底边（footskip） | 0mm | |
| 页脚高度 | 5.2mm | |

```
版芯布局示意图（从上到下）:
┌──────────────────────────────────────┐
│         页眉上边距（距页面顶 30.5mm）  │  ← header_distance
│──────────────────────────────────────│  ← 粗线 2.276208pt
│  gap 1.19pt                          │
│──────────────────────────────────────│  ← 细线 0.75pt
│  headsep = 2mm                       │
│──────────────────────────────────────│
│                                      │
│         150mm 正文宽度               │
│                                      │
│──────────────────────────────────────│
│  footskip=0mm（无额外间距）           │
│  页脚文字底部到版芯下边距 = 5.2mm      │
└──────────────────────────────────────┘
```

**三种版芯配置（通过 `newgeometry` 选项切换）：**

| 选项 | 正文宽度 | 正文高度 | top | left | head | headsep | footskip | foot |
|------|---------|---------|-----|------|------|--------|---------|------|
| `newgeometry=two`（默认） | 150mm | 236mm | — | 30mm | 5mm | 2mm | 0mm | 5.2mm |
| `newgeometry=one` | 150mm | 240mm | — | 30mm | 5mm | 0mm | 0mm | 0mm |
| `newgeometry=no`（旧） | 150mm | 224mm | 35.5mm | 30mm | 5mm | 2.5mm | — | 8.5mm |

**本科生哈尔滨专版（额外调整）：**
```
top=36.5mm, headsep=1mm, bottom=28.8mm
```

### 1.2 页眉格式

```latex
页眉组成（从上到下）:
1. 学校名称（黑体/宋体小五 9bp）+ 粗线下边框 sz=18（2.25pt）
2. 间距段落（line=0.77pt，含细线）   ← 控制粗线与细线间距 0.75pt
3. 细线下边框 sz=6（0.75pt）

总占用高度 = 粗线 + gap + 细线 ≈ 14pt
```

代码实现（hithesisbook.cls:447-451）:
```latex
\vskip 1.190132pt        % 粗细线之间间距 1.19pt（≈细线宽度0.75pt + 余量）
\hrule\@height 2.276208pt  % 粗线 2.276208pt
\vskip 0.75pt           % 细线与粗线间距
\hrule\@height 0.75pt     % 细线 0.75pt
```

**页眉内容（按论文类型）：**

| 类型 | 页眉内容 |
|------|---------|
| 博士-哈尔滨 | `哈尔滨工业大学 博 学位论文`（左）+ `哈尔滨工业大学博 学位论文`（右） |
| 博士-深圳 | `哈尔滨工业大学 深圳校区 博士 学位论文`（右） |
| 硕士-哈尔滨 | `哈尔滨工业大学 硕士 学位论文`（居中） |
| 本科-哈尔滨 | `哈尔滨工业大学 本科 毕业设计`（居中） |

### 1.3 页脚格式

```latex
格式: - PAGE -
字体: Times New Roman 小五 9bp
```

本科哈尔滨专版额外控制短横线间距:
```latex
\setlength\hit@pagenumber@sep{.3333333em}  % 哈尔滨本科: 1/3 em
% 其他: .1666667em
```

---

## 二、字体系统

### 2.1 字号定义（hithesisbook.cls:407-429）

| 名称 | 英文名 | 字号 | 代码 |
|------|--------|------|------|
| 特大 | 大初 | 58bp | `\dachu{58bp}` |
| 初号 | 大初 | 42bp | `\chuhao{42bp}` |
| 小初 | 小初 | 36bp | `\xiaochu{36bp}` |
| 一号 | 一号 | 26bp | `\yihao{26bp}` |
| 小一 | 小一 | 24bp | `\xiaoyi{24bp}` |
| 二号 | 二号 | 22bp | `\erhao{22bp}` |
| 小二 | 小二 | 18bp | `\xiaoer{18bp}` |
| 三号 | 三号 | 16bp | `\sanhao{16bp}` |
| 小三 | 小三 | 15.9bp | `\xsanhao{15.9bp}` |
| 四号 | 小三 | 15bp | `\xiaosan{15bp}` |
| **小四** | **四号** | **14bp** | `\sihao{14bp}` |
| 五号 | 五号 | 10.5bp | `\wuhao{10.5bp}` |
| 小五 | 小五 | 9bp | `\xiaowu{9bp}` |
| 六号 | 六号 | 7.5bp | `\liuhao{7.5bp}` |
| 小六 | 小六 | 6.5bp | `\xiaoliu{6.5bp}` |

### 2.2 正文字号

```latex
\normalsize 配置（hithesisbook.cls:401-406）:
  字号: 12bp
  行距: \ifhit@glue 20.50394bp @plus 2.834646bp @minus 0bp \else 20.50394bp \fi
       ≈ 1.3倍行距（约 20.5pt）

行距换算:
  12bp × 1.3 = 15.6pt，但实际使用 20.50394bp ≈ 17.1pt
  （LaTeX 中 bp ≠ pt，1bp ≈ 1.00374pt）
```

---

## 三、章节标题

### 3.1 章（一级）

```latex
格式（hithesisbook.cls:833-851）:
  字体: 黑体 二号 18bp 加粗（\xiaoer[1.57481]）
  对齐: 居中（\centering）
  序号格式: "第×章 "（哈尔滨硕士和深圳/哈尔滨本科用 \quad，其他用 \enspace）
  标题: 悬挂缩进（hangindent = 序号宽度）
  fixskip: true（去除默认间距）

  beforeskip = 28.34646bp （约 1 行，18bp×1.57481）
  afterskip  = 28.74646bp  （约 0.8 行 + 0.57481×18）

  其他属性: keepNext, keepLines（自动，由 ctexbook 控制）
```

**不同论文类型的 afterskip 差异：**
- 威海本科: `beforeskip=49bp, afterskip=46bp`（额外多一行）
- 其他: `beforeskip=28.34646bp, afterskip=28.74646bp`

### 3.2 节（二级）

```latex
格式（hithesisbook.cls:853-867）:
  字体: 黑体 小三 15bp 加粗（\fontsize{15bp}{21bp}）
  序号后空格: \quad（哈尔滨硕士和深圳/哈尔滨本科），\enspace（其他）
  标题悬挂缩进: aftername=\enspace

  beforeskip = afterskip = 19.84252bp（约 0.5 行）
  fixskip = true, break = {}

  代码:
  beforeskip={\ifhit@glue 19.84252bp \@plus 2.834646bp \@minus 0bp
               \else 19.84252bp \fi},
```

### 3.3 小节（三级）

```latex
格式（hithesisbook.cls:869-883）:
  字体: 黑体 四号 14bp 加粗（\fontsize{14bp}{18bp}）
  beforeskip = afterskip = 17.00787bp（约 0.5 行）

  代码:
  beforeskip={\ifhit@glue 17.00787bp \@plus 2.834646bp \@minus 0bp
               \else 17.00787bp \fi},
```

### 3.4 细节（四级）

```latex
格式（hithesisbook.cls:885-899）:
  字体: 黑体 四号 14bp 加粗（同三级）
  beforeskip = afterskip = 8.503937bp（约 0.25 行）
  afterindent = true（段后缩进）
```

### 3.5 章节间距总结

| 级别 | 字号 | 字体 | beforeskip | afterskip | 说明 |
|------|------|------|-----------|-----------|------|
| 章 | 二号 18bp | 黑体 | 28.35bp | 28.75bp | 居中，悬挂缩进 |
| 节 | 小三 15bp | 黑体 | 19.84bp | 19.84bp | 左对齐，悬挂缩进 |
| 小节 | 四号 14bp | 黑体 | 17.01bp | 17.01bp | 左对齐 |
| 细节 | 四号 14bp | 黑体 | 8.50bp | 8.50bp | 左对齐 |

> 注: 1bp(big point) ≈ 1.00374pt，LaTeX 中 1in = 72bp = 72.27pt

---

## 四、目录

### 4.1 目录结构

```latex
目录标题（hithesisbook.cfg:83）:
  名称: 目 \ccwd 录（目 + 1个ccwd + 录，ccwd ≈ 一个汉字宽度）
  字体: 黑体 二号 18bp（通过 \chapter*{contentsname} 自动应用章标题样式）

  beforeskip: 28.34646bp
  afterskip:  28.74646bp
  fixskip: true
```

### 4.2 目录条目格式

```latex
TOC 条目样式（hithesisbook.cls:1676-1698）:
  TOC1 (章): \l@chapter → 黑体加粗，无编号，用 leader dotted 引导线
  TOC2 (节): \@dottedtocline{1}{1em}{2em}
  TOC3 (小节): \@dottedtocline{2}{2em}{3em}
  TOC4 (细节): \@dottedtocline{3}{3\ccwd}{3.1em}

  页码宽度（提前悬挂）: \@pnumwidth = 4em
  引导线点间距: \@dotsep = 1（中文）/ 0.75（英文）

  宽度配置:
  - numberline宽度（目录序号区）: 4em
  - 缩进: 1em / 2em / 3\ccwd（三级）
  - 标签宽度: 2em / 3em / 3.1em
```

### 4.3 目录层次缩进

| 级别 | indent | numwidth | 样式 |
|------|--------|---------|------|
| TOC1（章） | 0em | 4em | 黑体加粗 |
| TOC2（节） | 1em | 2em | 宋体/黑体（取决于 chapterbold 选项） |
| TOC3（小节） | 2em | 3em | 宋体 |
| TOC4（细节） | 3\ccwd | 3.1em | 宋体 |

### 4.4 其他目录设置

```latex
tocdepth: 2（默认显示到三级标题）
tocblank: true（第一章前自动加一行空白，由 \addtocontents{toc}{\vspace{\baselineskip}} 实现）
chapterhang: true（章标题悬挂居中，规范要求章标题<15字）
```

---

## 五、摘要

### 5.1 中文摘要

```latex
标题:
  文本: "摘  要"（中间两个ccwd空格）
  字体: 黑体 二号 18bp 加粗
  样式: 章节样式（居中、beforeskip=28.35bp）

正文:
  字体: 宋体/正文 12bp
  行距: 20.50394bp（约1.3倍）
  首行缩进: 2字符（\parindent=2em）

关键词:
  标签: "关键词：" 黑体 14bp 加粗
  内容: 宋体 12bp，关键词之间用"；"分隔
  换行: \vskip12bp（关键词前空 12bp）
```

### 5.2 英文摘要

```latex
标题:
  文本: "Abstract"
  字体: Times New Roman 二号 18bp 加粗

正文:
  字体: Times New Roman 12bp
  首行缩进: 2字符

关键词:
  标签: "Key words: " Times New Roman 12bp 加粗
  内容: Times New Roman 12bp，关键词之间用","分隔
```

---

## 六、三线表

```latex
booktabs 参数（hithesisbook.cls:243-245）:
  \heavyrulewidth = 1.5pt   % 顶线/底线 = 1.5磅 = sz=12
  \lightrulewidth = 1pt      % 未使用
  \cmidrulewidth = 0.5pt    % 中线 = 0.5磅 = sz=4

表题:
  字体: 五号 10.5bp 宋体（博士中英双语，硕士仅中文）
  位置: 表格上方
  格式: "表 ×-× 名称"（序号与表名之间空1格）

表头:
  字体: 五号 10.5bp
  样式: 加粗

表内容:
  字体: 五号 10.5bp
  对齐: 居中（表格内容整体居中）
  续表: 表题后加"（续表）"，重复表头
```

---

## 七、公式

```latex
公式样式（hithesisbook.cls:572-585）:
  序号格式: (章号-序号)，如 (1-1)
  字体:amsmath默认（数学斜体）
  对齐: 居中
  标签右端对齐

正文引用格式: "见式(1-1)" 或 "由公式(1-1)"
编号字体: 正文字号（12bp）

代码:
  \renewcommand{\theequation}{\ifnum \c@chapter>\z@ \thechapter-\fi\@arabic\c@equation}
  \def\tagform@#1{\maketag@@@{(\ignorespaces #1\unskip\@@italiccorr)}}

公式前空4个半角字符（若前有文字如"解"）
公式注释（物理量说明）:
  使用 tabularx 环境，格式: 式中 | 符号 | 注释——
  三列: @{}l@{\quad}r@{———}X@{}
```

---

## 八、图片

```latex
图题样式（hithesisbook.cls:602-623）:
  字体: 五号 10.5bp
  位置: 图下方
  格式: "图 ×-× 名称"（图序与图名之间空2个半角字符）
  对齐: 居中
  博士: 中英双语（中文在上，英文用 Times New Roman）
  硕士: 仅中文

分图:
  分图号: (a)、(b) 等
  位置: 分图题在分图之下或图题之下

图注:
  位置: 图题之上
  字体: 五号

图内文字:
  字体: 宋体/Times New Roman
  字号: 五号（字数多时可用小五）
  物理量符号: 斜体

插图与正文间距: 上下各空一行（约一个 baselineskip）
```

---

## 九、页码

```latex
章节分节（hithesisbook.cls:366-400）:
  \frontmatter: 罗马数字 I, II, III...（从1开始）
  \mainmatter: 阿拉伯数字 1, 2, 3...（从1开始）

页码位置: 版心下边线之下居中（Word 中页脚居中）
页码格式:
  前导: "-"（破折号）
  数字: "PAGE" 域或 \thepage
  结尾: "-"
  字体: Times New Roman 小五 9bp

封面（无页码）:
  \frontmatter 之前使用 \pagestyle{hit@empty}
```

---

## 十、封面

### 10.1 第一封面（哈尔滨硕士/博士）

```latex
布局（从上到下）:
  1. 空行
  2. 论文类型: "博 士 学 位 论 文" 或 "硕 士 学 位 论 文"
     - 字体: 宋体 小一 24bp 加粗
     - 若非全日制: 加 "(学术学位论文)" 等
  3. 空多行
  4. 中文题目: 二号 22bp 黑体加粗
     - 若有副标题: 副标题前加"——"
  5. 英文题目（大写）: 二号 Times New Roman
  6. 空行
  7. 作者姓名: 小二 宋体 加粗
  8. 学校名称（哈尔滨）: 小二 楷体 加粗
  9. 日期: 小二 宋体 加粗
```

### 10.2 第二封面

```latex
顶部信息（仅哈尔滨硕士）:
  左: 国内图书分类号 / 国际图书分类号
  右: 学校代码 10213 / 密级

论文类型: 小二 宋体加粗
  哈尔滨硕士/博士: 硕/博 + 士 + 学位论文
  深圳硕士: "Dissertation for the Master Degree"

表格信息:
  - 学科/类别（哈尔滨硕士）: 学科
  - 申请学位: 学位
  - 指导教师 / 副导师 / 联合导师
  - 所在单位（院系）
  - 答辩日期
  - 哈尔滨工业大学
```

---

## 十一、授权声明

```latex
标题: "学位论文原创性声明" 黑体 小二 18bp
正文: 小四 12bp 宋体，首行缩进 2字符

格式（hithesisbook.cls:1796-1810）:
  声明文本（2-3段）
  空一行
  作者签名 + 日期（右对齐）
  空 2\baselineskip
  "学位论文使用权限" 标题
  使用权限文本
  空 2\baselineskip
  作者签名 + 日期
  空 2\baselineskip
  导师签名 + 日期

字体:
  标题: 黑体 小三 16bp 加粗
  正文: 宋体 小四 12bp
  签名标签: 宋体 小四
```

---

## 十二、其他格式

### 12.1 脚注

```latex
脚注序号: \textcircled{小六 6.5bp}（圆圈数字，1-9）
脚注正文: 小五 9bp
脚注分隔线: \hrule\@width 0.3\@tempdima \@height 0.4\p@（0.4pt 细线）
间距: \vskip -3\p@ + \vskip 2.6\p@（净高度 -0.4pt）

行距: \xiaowu[1.5]（小五 1.5倍行距）
左边界: 1.5em
```

### 12.2 正文行距与缩进

```latex
正文:
  字号: 12bp
  行距: 20.50394bp（约 1.3倍，但实际约 17.1pt）
  段前: 0
  段后: 0
  首行缩进: 2em（两字符）

列表项:
  itemsep = 0em
  parsep = 0em
  topsep = 0em
  partopsep = 0em
  leftmargin = 0em
  itemindent = 3em（三级节）/ 3.5em（节）
```

### 12.3 结论

```latex
格式:
  标题: "结    论" 黑体 小二 18bp（居中，不加章序号）
  正文: 宋体 小四 12bp，首行缩进 2字符
  条目编号: （1）、（2）、（3）...（用全角括号）

  注意: 结论不加章标题序号，不标注引用文献
```

### 12.4 致谢

```latex
标题: "致    谢" 黑体 小二 18bp（居中，不加章序号）
正文: 宋体 小四 12bp，首行缩进 2字符
```

---

## 十三、Word 模板对应关系

| LaTeX 属性 | Word OOXML 等效 |
|------------|----------------|
| 12bp 正文 | 12pt = 24 half-pts |
| 二号 18bp | 18pt = 36 half-pts |
| 小三 15bp | 15pt = 30 half-pts |
| 四号 14bp | 14pt = 28 half-pts |
| 五号 10.5bp | 10.5pt = 21 half-pts |
| 小五 9bp | 9pt = 18 half-pts |
| 1.5pt 粗线 | `w:sz="12"` (1.5×8) |
| 0.5pt 细线 | `w:sz="4"` (0.5×8) |
| 行距 20.50394bp | 约 20.5pt |
| 28.35bp 段前 | `w:before="360"` (360 twips ≈ 28pt) |
| 28.75bp 段后 | `w:after="368"` (368 twips ≈ 29pt) |
| 19.84bp 段前/后 | `w:before/after="254"` (254 twips) |
| 首行缩进 2em | `w:firstLine="437"` (约 21.85pt) |

**换算参考：**
- 1bp ≈ 1.00374pt（LaTeX 特有）
- 1pt = 20 twips
- 1cm ≈ 567 twips
- 1bp ≈ 567 × 1.00374 / 72 ≈ 7.9 twips

---

## 十四、关键代码索引

| 功能 | 文件 | 关键位置 |
|------|------|---------|
| 页面版芯 | hithesisbook.cls | 272-310 |
| 页眉页脚 | hithesisbook.cls | 441-540 |
| 章节标题样式 | hithesisbook.cls | 832-903 |
| 目录格式 | hithesisbook.cls | 1646-1716 |
| 三线表边框 | hithesisbook.cls | 243-245 |
| 公式编号 | hithesisbook.cls | 572-585 |
| 封面布局 | hithesisbook.cls | 1106-1590 |
| 中文配置 | hithesisbook.cfg | 全文 |
| 摘要格式 | hithesisbook.cls | 1598-1616 |
| 授权声明 | hithesisbook.cls | 1796-1869 |
| 正文行距 | hithesisbook.cls | 401-406 |
| 脚注格式 | hithesisbook.cls | 561-568 |