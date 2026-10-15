# 中文科技报告排版规范 + OOXML 对照

## 一、单位换算（最常出错的地方）

| 目标 | OOXML 属性 | 换算 | 例 |
|---|---|---|---|
| 字号 | `w:sz` / `w:szCs` | **half-point = pt × 2** | 12pt → `24` |
| 边框粗细 | `w:tblBorders/*/@w:sz` | **1/8 point = pt × 8** | 1.5pt → `12`；0.75pt → `6` |
| 缩进 | `w:ind/@w:firstLine` | twips = pt × 20 | 2 字符 @12pt = 24pt → `480` |
| 段前段后 | `w:spacing/@w:before` | twips = pt × 20 | 12pt → `240` |
| 行距 | `w:spacing/@w:line` | 240 × 倍数 | 1.5 倍 → `360` |
| 页面/边距 | — | python-docx 用 `Cm()` | 3.0cm |

## 二、中文字号对照

| 中文名 | pt | `w:sz` |
|---|---|---|
| 初号 | 42 | 84 |
| 小初 | 36 | 72 |
| 一号 | 26 | 52 |
| 小一 | 24 | 48 |
| 二号 | 22 | 44 |
| 小二 | 18 | 36 |
| 三号 | 16 | 32 |
| 小三 | 15 | 30 |
| 四号 | 14 | 28 |
| **小四** | **12** | **24** ← 本文正文 |
| 五号 | 10.5 | 21 |
| 小五 | 9 | 18 |
| 六号 | 7.5 | 15 |

`--body-size` 接受中文名或数字，映射表在 `spec.CN_SIZE`。

## 三、本规范各样式精确参数

| 样式 | eastAsia / ascii | pt → `sz` | 行距 | 首行缩进 | 段前/段后 | 其它 |
|---|---|---|---|---|---|---|
| Normal（正文） | 宋体 / Times New Roman | 12 → 24 | 1.5 → 360 | **200 chars / 480** | 0 / 0 | 两端对齐 |
| Heading 1 | 黑体 / TNR | 15 → 30 | 1.5 | **0 / 0** | 24 / 12 | outlineLvl 0，keepNext |
| Heading 2 | 黑体 / TNR | 14 → 28 | 1.5 | 0 / 0 | 18 / 8 | outlineLvl 1，keepNext |
| Heading 3 | 黑体 / TNR | 12 → 24 | 1.5 | 0 / 0 | 12 / 6 | outlineLvl 2，keepNext |
| Heading 4 | 黑体 / TNR | 12 → 24 | 1.5 | 0 / 0 | 12 / 6 | outlineLvl 3，keepNext |
| Title（文档标题） | 黑体 / TNR | 18 → 36 | 1.5 | 0 / 0 | 6 / 18 | 居中 |
| 表题 | 黑体 / TNR | 9 → 18 | 1.0 → 240 | 0 / 0 | 6 / 3 | 居中，keepNext |
| 表注 | 宋体 / TNR | 9 → 18 | 1.0 | 0 / 0 | 3 / 6 | 左对齐 |
| 代码 | Consolas（中西同） | 9 → 18 | 1.0 | 0 / 0 | 2 / 2 | 底纹 `F2F2F2` |
| 表格文字 | 宋体 / TNR | 10.5 → 21 | 1.0 | **0 / 0** | 1.5 / 1.5 | ← 防缩进漏进表格 |

**标题一律 `firstLineChars=0` 且不能只删属性** —— `w:ind` 会沿 `basedOn` 链继承回来，
必须显式写双 0（`firstLine` 和 `firstLineChars` 都设 0）。

**正文一律不加粗** —— 行内 `**…**` 在解析层就剥成普通文本（`md_parse.iter_inline`
不再产出 `'b'`），run 上不写 `w:b`。中文报告正文不做加粗强调；加粗只出现在体例
要求的位置：黑体标题、表题、三线表的表头行。所以 `w:b` 在本文档里只有三个来源，
排查加粗问题时逐一对照即可。

## 四、三线表

```xml
<w:tblPr>
  <w:tblBorders>            <!-- 六边必须全部显式写出 -->
    <w:top     w:val="single" w:sz="12" w:space="0" w:color="000000"/>
    <w:bottom  w:val="single" w:sz="12" w:space="0" w:color="000000"/>
    <w:left    w:val="none"/>  <w:right   w:val="none"/>
    <w:insideH w:val="none"/>  <w:insideV w:val="none"/>
  </w:tblBorders>
</w:tblPr>
```

表头行下细线走**单元格**级（`w:tc/tcPr/tcBorders/w:bottom`，`sz=6`），不用
`w:tblStylePr`（python-docx 零支持，且要求保留表样式）。

**优先级**：直接格式 `tcBorders` > 表样式。所以必须先把 `tbl.style` 换成
`Normal Table` 这类无边框基座，否则 `Table Grid` 的全网格线会盖住 `tblBorders`。

### 表格对齐（`spec.py`）

| 参数 | 默认 | 说明 |
|---|---|---|
| `TABLE_HEADER_JC` | `center` | 表头行水平对齐 |
| `TABLE_CELL_JC` | `center` | 表体水平对齐 |
| `TABLE_CELL_VALIGN` | `center` | 单元格垂直对齐（水平居中时配垂直居中最协调） |
| `TABLE_LONG_TEXT_LEFT` | `False` | 开启后，某列只要存在超过 `TABLE_LONG_TEXT_THRESHOLD`（默认 24）字符的单元格，该列整体改左对齐 —— 整句话居中会让右边缘参差 |

单元格对齐是**逐格写 `w:jc`**（直接格式），不靠 `表格文字` 样式继承 —— 因为要支持
"长文本列单独左对齐"这种逐列差异。垂直对齐走 `w:tcPr/w:vAlign`。

## 五、schema 子元素顺序（插错 Word 报「内容有问题」）

- `w:tblPr`：tblStyle, tblpPr, tblOverlap, bidiVisual, tblStyleRowBandSize,
  tblStyleColBandSize, tblW, jc, tblCellSpacing, tblInd, **tblBorders**, shd,
  tblLayout, tblCellMar, tblLook, tblCaption, tblDescription, tblPrChange
- `w:tcPr`：cnfStyle, tcW, gridSpan, hMerge, vMerge, **tcBorders**, shd, noWrap,
  tcMar, textDirection, tcFitText, vAlign, hideMark, headers, … , tcPrChange
- `w:pPr`：pStyle, keepNext, keepLines, pageBreakBefore, …, pBdr, shd, tabs, …,
  spacing, **ind**, contextualSpacing, …, jc, …, **outlineLvl**, …, rPr
- `w:rPr`：rStyle, **rFonts**, b, bCs, i, iCs, …, color, …, **sz**, **szCs**,
  highlight, u, …, lang

`w:tc` 若无 `tcPr`，`tcPr` 必须排在 `w:p` 之前。
代码里统一走 `docx_builder.insert_in_order()`，不要手写 `append`。

## 六、python-docx 1.2.0 能力边界（实测）

这些**原生不支持**，必须手写 XML：

| 需求 | 原生情况 |
|---|---|
| 中文字体 | `CT_Fonts` 只暴露 `@w:ascii` / `@w:hAnsi`，**没有 `@w:eastAsia`** |
| `w:szCs` | `font.size` 只写 `w:sz`，从不写 `szCs` |
| 首行缩进 2 字符 | 只有绝对长度 `first_line_indent`，无 `firstLineChars` |
| 表格边框 | `w:tblBorders` / `w:tcBorders` 在 `_tag_seq` 里但被 `del`，无 accessor |
| 域 | 全库无 `fldChar` / `instrText` / `fldSimple` |
| `updateFields` | 在 schema 里但无 accessor，走 `doc.settings.element` |

**回读校验的陷阱**：绝不能用 `style.font.name` 断言中文字体 —— 它只读 `@w:ascii`，
而中文由 `@w:eastAsia` 决定，会假绿。必须直读 `rFonts.get(qn('w:eastAsia'))`。

**styleId 因文档而异**：同一份模板里可能是 `Heading2` 也可能是 `21`。
一律走样式名（`doc.styles['Heading 2']`），不要硬编码 id。

## 七、容易翻车的地方

1. **Normal 的 2 字符缩进会继承进表格** —— 逐格显式写 `firstLineChars=0`。
2. **`Paragraph.runs` 不遍历 `w:hyperlink`** —— 超链接里的 run 会被整体漏掉。
3. **`w:ind` 只删属性无效** —— 会从 `basedOn` 链继承回来，必须显式双 0。
4. **`w:tblBorders` 六边必须全写** —— 省略的边会从样式继承，可能继承回非 none。
5. **`paragraph.text = ...` 会清空所有 run** —— 文本改写必须就地改 `w:t`，
   且首尾有空格的节点要补 `xml:space="preserve"`。
6. **读 `section.footer` 有副作用** —— 对 `is_linked_to_previous=True` 的节读取，
   会静默创建 footer part。体检要判断有无页脚请直读 zip 里的 `word/footer*.xml`。
7. **`updateFields` 要 find-then-set** —— 无脑 insert 会造出两个同名元素，Word 可能报损坏。
8. **中文数字正则别嵌套字符类** —— `CN='[一二三…]'` 再嵌进 `f'^{CN}[、]'` 会让 `re`
   静默改语义（只给 `FutureWarning`），曾导致全表 0 命中。