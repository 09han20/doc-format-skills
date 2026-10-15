---
name: doc-format
description: 把 Markdown 源稿生成排版规范的中文 Word 文档（.docx）——真标题样式、三线表、表题自动编号、页脚页码、首行缩进 2 字符、宋体/Times New Roman。当用户要写中文报告/调研报告/组会材料并要 Word 交付，或抱怨已有 Word 排版乱、表格没有边框、目录生成不了、表号断号时使用。触发词：Word、docx、报告排版、三线表、表题、页码、目录、Markdown 转 Word。仅用于生成中文 Word 文档，不处理 PDF / PPT / Excel。
---

# doc-format —— Markdown 生成规范中文 Word

**核心主张：排版问题在生成时就解决，不要生成完再回头修。**

事后修补的坑（标题层级只能猜、单元格边框已被覆盖、样式缩进漏进表格）在生成路径上全都不存在，
因为 XML 是自己写的。所以这个 skill 只做生成，不做"美化已有文档"。

## 何时用

- 用户要写中文报告 / 调研 / 组会材料，最终要 `.docx`
- 用户抱怨 Word 排版乱、表格没边框、目录点不动、表号断号
- 已有 `.md` 源稿要出 Word

**不要用于**：PDF、PPT、Excel；也不要用于"把一份别人给的乱排版 docx 修好"——
那个方向这个 skill 不覆盖（见文末「不做什么」）。

## 用法

```bash
python "${CLAUDE_SKILL_DIR}/scripts/md2docx.py" 报告.md                      # → 报告.docx
python "${CLAUDE_SKILL_DIR}/scripts/md2docx.py" 报告.md --toc                # 附目录域
python "${CLAUDE_SKILL_DIR}/scripts/md2docx.py" 报告.md --body-size 五号      # 改正文字号
python "${CLAUDE_SKILL_DIR}/scripts/md2docx.py" 报告.md --caption-mode auto  # 表题一律重编号
```

生成后**必须**跑校验：

```bash
python "${CLAUDE_SKILL_DIR}/scripts/verify.py" 报告.docx --md 报告.md
```

`verify.py` 断言 27 项，其中最关键的一条是**内容无损**：源 md 与产物的全文
（剥离格式后）必须逐字符一致。改排版改掉了字，比排版丑严重得多。

## Markdown 约定

标准方言之外，本 skill 认这几个：

| 写法 | 含义 |
|---|---|
| 第一个 `#` | 文档标题（居中，不参与编号） |
| 后续 `#` | 章标题（Heading 1） |
| `表：说明文字` | 紧跟其下那张表的题注，**自动编号** |
| `表 7 说明文字` | 手写编号，沿用你的号并校验连续性 |
| `表注：…`（表格下一行） | 表注，原文渲染 |
| `\pagebreak` | 另起一页 |
| ` ``` ` 代码围栏 | 等宽 + 底纹 |
| `![说明](图片路径)` 独占一行 | 插图（居中、超栏宽时等比缩放到栏宽）。相对路径按 md 所在目录解析 |
| `图：说明文字`（写在图片下一行） | 图题，**渲染在图片下方**、自动编号（与表题方向相反） |
| `[文本](url)` | 真超链接 |
| `**粗体**` | **剥掉标记、按正文排**——不加粗。保留标记只是让源稿在 GitHub/Typora 里可读 |

两种编号混用时的行为（默认 `--caption-mode both`）：有号用号并同步计数器，无号顺延。
所以 `表 7` 后面跟一个 `表：…` 会得到 `表 8`，不会退回 `表 1`。

## 生成的规范

- 正文 宋体(eastAsia) + Times New Roman(ascii)，小四 12pt，1.5 倍行距，首行缩进 2 字符
- 正文**一律不加粗**：`**…**` 只剥标记、不写 `w:b`。中文报告正文不用加粗强调，
  加粗只留给体例（黑体标题、表题、三线表表头行）
- 标题 黑体，H1 15 / H2 14 / H3 12pt，**真 Heading 样式 + outlineLvl**（目录和导航窗格靠它）
- 封面标题 黑体 小二 18pt 居中，**自建 `文档标题` 样式，无边框**（内置 `Title` 有蓝线）
- 表格 **三线表**：顶/底 1.5pt、表头下 0.75pt，其余无框线；表头跨页重复
- 表题 小五 9pt 黑体居中，在表格上方，带 `keepNext` 防与表分页断开
- A4，页边距 3.0/2.5/3.0/2.5 cm，页脚居中页码
- 表格单元格**显式禁首行缩进**（Normal 的 2 字符缩进会继承进去毁掉整张表）

全部参数在 `scripts/spec.py`，改体例只改那一个文件。完整规范与 OOXML 对照见
`references/spec.md`。

## 硬约束

1. **只用 python-docx 1.2.0，不引入新依赖。** 不要用 pandoc / LibreOffice / docx-js。
2. **改写完必须跑 `verify.py`。** 尤其是内容无损那一项。
3. **不要用 `paragraph.text = ...` 做文本改写** —— 它会清空所有 run、丢光 run 级格式。
4. **标题类一律用自建样式，不复用内置 `Title` / `Caption`。**
   内置 `Title` 自带 `w:pPr/w:pBdr` 蓝色 accent1（4F81BD）下边框——封面标题下面那道
   蓝线就是它；内置 `Caption` 是蓝色 accent1 + semiHidden。二者都会毁掉中文报告版面。
   封面标题用自建的 `文档标题`，表题用自建的 `表题`（样式名见 `spec.STYLE_TITLE` /
   `spec.STYLE_CAPTION`）。`install_styles()` 另用 `strip_para_border()` 对内置样式
   做一次兜底清除。
5. **不要硬编码 styleId**（`Heading2` vs `21` 因文档而异），走样式名。
6. **正文不加粗。** 行内 `**…**` 必须剥标记且不设 `w:b`；只有标题（黑体）、表题、
   三线表表头行允许加粗——那是体例，不是强调。改 `md_parse.INLINE` 时别把 `'b'` 放回来。

## 不做什么

- **不改写已有 docx**。用户若要求"把我这份排版乱的 Word 修好"，说明本 skill 只做生成，
  建议把内容导回 Markdown 再重新生成（那样更干净），或另找工具。
- 图片只支持「独占一行的 `![]()`」，不支持行内图、图文混排、多图并排。
  图片来源必须是本地文件（不走网络）；超栏宽自动缩到 15.5 cm，不做放大。

## 同门 skill：doc-paper（论文体例）

学位论文（封面 / 中英文摘要 / 关键词 / 目录 / 参考文献、罗马数字页码、按章编号的图3-2 / 表4-1、
公式右对齐编号）用另一个 skill `doc-paper`，本 skill 只管报告体例。

**它不自带引擎** —— `doc-paper/scripts/_engine.py` 反向 import 本目录的
`md_parse.py` / `docx_builder.py` / `md2docx.py` / `verify.py` / `spec.py`。所以：

- 动这几个文件，两个 skill 都要回归：报告跑 `md2docx.py` + `verify.py`，论文跑
  `md2paper.py` + `verify_paper.py`；
- `spec.py` 里的页边距 / 三线表边框 / 表格字号是**两份体例共用**的，改了论文那边也跟着变；
- 论文专属的体例参数在 `doc-paper/scripts/spec_paper.py`，两边不要互相搬参数。

## 排查

生成结果不对时先看这三处：

- **表格没有题注** → md 里表格上一行缺 `表：说明文字`；生成器会打印无题注表格的序号
- **目录是空的** → 正常。TOC 是域，需在 Word 里按 F9 或打开时确认更新
- **中文变成方框/西文字体** → 检查 `spec.py` 里的 `FONT_BODY_CJK`；eastAsia 槽没写对会这样
- **正文里出现加粗** → 检查 `md_parse.iter_inline` 是否又把 `**…**` 归成了 `'b'`；
  正常情况它只产出 `'t' / 'c' / 'a'`
- **封面标题下有一道蓝线** → 段落引用了 Word 内置 `Title` 样式（自带 accent1 下边框）。
  改用 `spec.STYLE_TITLE`（`文档标题`）；`verify.py` 会断言内置 `Title`/`Caption` 无 `pBdr`
  且文档正文未引用它们