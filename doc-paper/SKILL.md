---
name: doc-paper
description: 把 Markdown 源稿生成为排版规范的学位论文 Word（.docx）——封面、中英文摘要、关键词、自动目录、分节符与罗马/阿拉伯页码、按章编号的图3-2/表4-1、公式右对齐编号、GB/T 7714 参考文献悬挂缩进。支持论文正式版（六部分齐全，严格校验）与日常版（省略摘要、关键词、参考文献）。触发词：论文、学位论文、毕业设计、论文排版、封面、摘要、关键词、参考文献、罗马数字页码、分节符、GB/T 7714、Markdown 转 Word。仅生成中文论文体例的 Word 文档；中文报告/调研/组会材料请用 doc-format。
---

# doc-paper —— Markdown 生成规范学位论文 Word

**核心主张：排版问题在生成时就解决，不要生成完再回头修。**

体例照学位论文通用排版规范实现：封面 / 中英文摘要 / 关键词 / 自动目录 / 分节符与
罗马·阿拉伯页码 / 按章编号的图表与公式 / GB/T 7714 参考文献著录。

## 两种模式（`--mode` 必填，没有默认）

| 模式 | 内容 | 什么时候用 |
|---|---|---|
| `paper` 论文正式版 | **封面、摘要、关键词、目录、正文、参考文献六部分齐全**；缺哪部分就报错、且不落盘 | 要交的学位论文、投稿、送审 |
| `daily` 日常版 | 省略**摘要、关键词、参考文献**三部分，其余（封面、目录、正文、致谢、附录）与正式版完全一致 | 日常给导师看的稿子、阶段性材料 |

两种模式的排版**完全一致**（同一份 spec），差别只在省略三部分。
用户说「正式版 / 交稿版 / 完整版」选 paper，说「日常 / 先给我看一眼 / 不要摘要」选 daily；
**没说就问一句** —— 日常版会把源稿里的摘要、关键词、参考文献**整段丢掉**，选错了是丢内容。

## 用法

```bash
python "${CLAUDE_SKILL_DIR}/scripts/md2paper.py" 论文.md --mode paper
python "${CLAUDE_SKILL_DIR}/scripts/md2paper.py" 论文.md --mode daily
python "${CLAUDE_SKILL_DIR}/scripts/md2paper.py" 论文.md --mode paper --caption-mode auto
python "${CLAUDE_SKILL_DIR}/scripts/md2paper.py" 论文.md --mode paper --formula text
```

## 公式（LaTeX → Word 真公式）

源稿照写 LaTeX，文档里得到的是 **Word 公式对象**：分数上下堆叠、Σ 带上下限、∫ 带上下限、
`\left( \right)` 的括号自动放大，双击可进公式编辑器改。

支持的子集：`\frac \dfrac \sqrt[n]{}`、`^ _`、`\sum \prod \int \oint \iint \bigcup \bigcap`、
`\left…\right`、`\hat \bar \vec \tilde \dot \overline \underline`、希腊字母与常用运算符、
`\sin \log \max \lim…`、`\text{} \mathrm{} \mathbf{}`、`\, \; \quad`、`\% \& \_ #`。

**超出子集的（`\begin{matrix}`、自定义宏、没基底的 `_`）整条回退成纯文本并给警告** ——
宁可原样显示 LaTeX 也不要排出个错公式；要插图就换成图片。校验时这种回退会算作
「公式没排成公式对象」而报 `[!!]`，逼你知道而不是悄悄放过。

生成后**必须**跑校验：

```bash
python "${CLAUDE_SKILL_DIR}/scripts/verify_paper.py" 论文.docx --md 论文.md
```

`verify_paper.py` 断言 60 余项，其中两条最要紧：
**分节与页码**（封面节无页脚 → 前置 upperRoman → 正文 decimal 从 1 起）与
**内容无损**（源 md 与产物的全文剥格式后逐字符一致）。改排版改掉了字，比排版丑严重得多。

## Markdown 约定

标准方言之外，本 skill 认这几个：

| 写法 | 含义 |
|---|---|
| 开头 `---` 包起来的一段 | **front matter = 封面元数据**：`题名：…` 渲染成封面大字，其余行原样居中成一行（`作者：…`、`单位：…`、`指导教师：…`、`完成时间：…`） |
| `# 摘要` / `# Abstract` | 摘要标题（一级标题样式，进目录） |
| `关键词：甲；乙；丙` / `Keywords: a; b` | 关键词行，标签加粗、不缩进；分号由源稿自带，不改写 |
| `# 目录` | 想自己定目录位置（比如像某些模板那样排在摘要之前）才写；不写就在正文前自动插 |
| `# 1 绪论` / `# 第一章 绪论` / `# 一、绪论` | 章标题 —— 章号从标题文本里读，供 `图3-2`/`表4-1`/`(2-1)` 编号用 |
| `表：说明` / `图：说明` | 题注，**自动编号**成 `表 1-1 说明`（第 1 章第 1 表）。表题在表上方、图题在图下方 |
| `表 4-1 说明` | 手写编号，沿用你的号（其后自动编号顺延） |
| `$$ … $$` | 公式：**排成 Word 公式对象**（LaTeX 语法照写，Word 里是能双击编辑的真公式）＋右对齐自动编号 `(章-序)`；写 `$$! … $$` 则不编号 |
| `# 参考文献` | 其中的 `[1] …`、`1. …`、`- …` 都按条目排（五号、悬挂缩进 2 字符；`1.` 自动转 `[N]`） |
| `# 致谢` / `# 附录 A …` | 不推进章号；正式版不查、日常版保留 |
| `\pagebreak` | 另起一页（分页符，不改变页码分区） |
| 三线表、`![]()` 插图、代码围栏、列表、引用、超链接 | 与 doc-format 同一套约定，见其 SKILL.md |

**有 front matter 时，首个 `#` 不再当文档标题**（题名在封面上），`# 摘要` 就是摘要标题；
没有 front matter 时（日常版常见），首个 `#` 仍是文档标题，与报告体例一致。

**章号必须写在标题文本里**（`# 2 方法设计`）。本 skill 绝不自动给标题加编号 ——
那是往正文里塞源稿没有的字，内容无损校验会当场抓住。

## 生成的规范

- **页码分区**（md 规范「必须使用分节符」）：封面节**无页码** → 前置节（摘要/Abstract/关键词/目录）
  **Ⅰ、Ⅱ、Ⅲ**（upperRoman）→ 正文起（含参考文献/附录/致谢）**1、2、3**（decimal，`w:start=1` 重新数）
- **标题**：一级 黑体三号(16pt) 居中 / 二级 黑体四号(14pt) 左对齐 / 三级 楷体小四(12pt)加粗 /
  四级 宋体小四(12pt)加粗；真 Heading 样式 + outlineLvl，目录靠它
- **正文**：宋体(中文) + Times New Roman(西文) 小四 12pt，1.5 倍行距，首行缩进 2 字符，两端对齐
- **封面**：题名 二号黑体居中 + 字段行 四号居中；自建样式（内置 `Title` 自带蓝色下边框，不能复用）
- **目录**：TOC 域（`\o "1-3"`，最多三级）；缓存文本是「按 F9 更新」提示，打开时 Word 自动重算。
  **目录单独占一页** —— 前面排过摘要/关键词就分页隔开，后面接正文靠分节符隔开；
  源稿把 `# 目录` 写在摘要之前时，摘要也会被分页推到下一页
- **表格**：三线表（顶/底 1.5pt、表头下 0.75pt），表内 五号宋体，单元格显式禁首行缩进
- **公式**：LaTeX → **Word 公式对象**（OMML），Word 里双击可进公式编辑器改；
  版式是 1×3 无框表格（空栏｜公式居中｜编号右对齐），公式真居中、编号贴右边。
  `--formula text` 可退回「LaTeX 源文本」的老做法
- **参考文献**：五号宋体/TNR，悬挂缩进 2 字符（`w:left/w:hanging` + 显式清零 `firstLineChars`）
- A4，页边距 上 3.0 / 下 2.5 / 左 3.0 / 右 2.5 cm（在 md 给的区间内取整）

全部参数在 `scripts/spec_paper.py`，改体例只改那一个文件。完整规范与 OOXML 对照见
`references/spec-paper.md`；可直接填空的 md 骨架见 `references/论文模板.md`。

## 与 doc-format 的关系（重要）

**本 skill 不自带渲染引擎。** 中文字体槽、szCs、firstLineChars、三线表、schema 子元素顺序、
域、文本流比对这些只有一份实现，在 `doc-format/scripts/`，由 `_engine.py` 跨目录 import：

```
doc-format/scripts/   md_parse.py  docx_builder.py  md2docx.py  verify.py  spec.py   ← 引擎
doc-paper/scripts/    _engine.py  md_parse_paper.py  thesis_xml.py  md2paper.py  verify_paper.py
```

- **两个 skill 必须同时安装**；doc-paper 单独拷走会报「找不到 doc-format 的引擎目录」
  （可用环境变量 `DOC_FORMAT_SCRIPTS` 指过去）。
- 报告体例（doc-format）的行为**不因本 skill 而改变**；改共享引擎时两个 skill 都要回归。
- 两边的差异是**体例数据**，不是机制：论文参数在 `spec_paper.py`，报告参数在 doc-format 的 `spec.py`；
  共用值（页边距、三线表边框、表格字号）只定义一份，在 doc-format 的 `spec.py`。

## 硬约束

1. **只用 python-docx 1.2.0**，不引入新依赖（不装 pyyaml / pandoc / LaTeX）。
2. **改写完必须跑 `verify_paper.py`**，尤其是「内容无损」那一项。
3. **不要用 `paragraph.text = ...` 做文本改写** —— 会清空所有 run、丢光 run 级格式。
4. **不要自动给标题加编号、不要发明封面内容** —— md 里没有的，Word 里也不会出现。
5. **正式版缺件宁可不落盘**（`validate()` 在渲染前跑）—— 出一份缺摘要的「正式版」比不出文件更糟。
6. **改页码格式只改 `w:pgNumType`**，不要在 PAGE 域上写 `\* ROMAN` 开关；`pgNumType` 在 `w:sectPr`
   里必须排在 `cols` 之前（顺序错了 Word 报「文件内容有问题」）。
7. **`w:ind` 是逐属性合并的**：写悬挂缩进时必须显式写 `firstLine/firstLineChars="0"`，
   否则 Normal 的 2 字符首行缩进会顺着继承链回来。

## 排查

- **目录是空的** → 正常。TOC 是域，在 Word 里按 F9 或确认更新即可（生成时已设 `updateFields`）
- **封面出现页码** → 有代码碰了 `sections[0].footer`。封面节全程不能碰页脚（「不碰 = 没有」）
- **正文页码从 Ⅱ 或 4 接着数** → 缺 `w:start="1"`：正文节必须显式写 `w:start`，否则会接着前置节数
- **参考文献首行也缩进** → `w:ind` 里的 `firstLineChars` 没清零，见硬约束 7
- **公式显示成小号、Σ 上下限挤在右侧** → 公式被当成行内 oMath 排了。Word 只在
  `m:oMathPara` 里用显示式排版，而 oMathPara 与普通 run 同段就失效 ——
  所以必须是「表格中栏独占 oMathPara」，别把 `add_formula` 改回制表位版式
- **公式回退成了文本**（生成时有 ⚠）→ 用了子集外的语法；`latex_omml.py` 文件头上列了支持范围，
  要么改写法（`\begin{matrix}` → 拆成单行公式或换图片），要么 `--formula text` 认了
- **中文变成方框或西文字体** → `docx_builder.set_rfonts` 的 eastAsia 槽没写对
- **正式版报「缺封面/摘要/…」** → 按提示补源稿，或改用 `--mode daily`；提示里逐个写了该怎么写
- **`表 2-1` 被拆成 `表 2` + `-1 …`** → 共享的 `CAP_RE` 只认 `表 4`，`retag()` 里那段接回来；
  改正则时别把 `CAP_TAIL_RE` 删了