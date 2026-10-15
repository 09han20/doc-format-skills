# -*- coding: utf-8 -*-
"""论文专属 OOXML 原语：分节与页码格式、封面、公式行、关键词、参考文献悬挂缩进。

报告体例用不到、而 python-docx 1.2.0 又没有 API 的，全在这里：
  · 分节符（分节＝页码分区）与 w:pgNumType（罗马数字 / 从 1 重数）
  · 封面题名与字段行（自建样式，绕开内置 Title 的蓝线）
  · 公式行：居中制表位 + 右对齐制表位，一行里放下公式和编号
  · 参考文献悬挂缩进（w:ind 是逐属性合并的，见 set_hanging 的注释）

其余一律复用 doc-format 的 docx_builder / md2docx —— 不要在这里重造三线表或字体槽。
"""
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.shared import Cm

from _engine import B, M, spec as S
import latex_omml as L
import spec_paper as SP

W = B.W

# CT_SectPr 的子元素顺序。python-docx 在类体末尾把 _tag_seq del 掉了（运行时
# hasattr 为 False），只能自己写死。pgNumType 坐在 lnNumType 与 cols 之间 ——
# 插错位置 Word 会报「文件内容有问题」。
SECTPR_SEQ = [W(t) for t in (
    'headerReference', 'footerReference', 'footnotePr', 'endnotePr', 'type', 'pgSz',
    'pgMar', 'paperSrc', 'pgBorders', 'lnNumType', 'pgNumType', 'cols', 'formProt',
    'vAlign', 'noEndnote', 'titlePg', 'textDirection', 'bidi', 'rtlGutter', 'docGrid',
    'printerSettings', 'sectPrChange')]


# ---------------------------------------------------------------- 分节与页码
def add_section(doc):
    """另起一节（分节符＝下一页）。新节继承上一节的页面设置。

    注意 add_section 会把当前 body 末尾的 sectPr 克隆进一个段落作为**上一节**，
    并剥掉 sentinel 上的页眉页脚引用 —— 所以「封面不给页码」只需要全程别碰
    sections[0].footer，什么都不用写。
    """
    return doc.add_section(WD_SECTION.NEW_PAGE)


def set_pgnum(section, fmt=None, start=None):
    """本节页码格式与起始值：w:sectPr/w:pgNumType。

    Word 的 PAGE 域本身不带 \\* ROMAN 之类开关 —— 显示成 Ⅰ 还是 1 由这里的
    fmt（upperRoman / decimal）决定，start 只影响本节的第一个页码。
    """
    el = B.find_or_add(section._sectPr, W('pgNumType'), SECTPR_SEQ)
    if fmt:
        el.set(W('fmt'), fmt)
    if start is not None:
        el.set(W('start'), str(start))
    return el


def number_zone(section, fmt=None, start=None):
    """把一个节变成独立的页码区：格式 + 起始值 + 页脚页码域。"""
    set_pgnum(section, fmt, start)
    B.add_page_number(section)
    return section


# ---------------------------------------------------------------- 样式
def _ind_of(obj):
    """段落或样式对象上的 w:ind（按 schema 顺序插入）。"""
    return B.find_or_add(obj._element.get_or_add_pPr(), W('ind'), B.PPR_SEQ)


def set_hanging(obj, chars=SP.REF_HANG_CHARS, twips=None):
    """悬挂缩进：左缩进 = 悬挂量，首行回到 0。段落或样式都能用。

    三个坑，缺一个就白写：
      ① w:ind 是**逐属性**合并的 —— 不写 firstLine/firstLineChars，Normal 的
         firstLineChars=200 会顺着 basedOn 链继承进来，首行照样缩进；
      ② firstLineChars 的优先级高于 firstLine，只写 w:firstLine="0" 没用；
      ③ 悬挂是「首行相对左缩进回退」，left 不写就没有悬挂可言。
    """
    tw = twips if twips is not None else S.first_line_twips(SP.REF_PT)
    el = _ind_of(obj)
    el.set(W('left'), str(tw))
    el.set(W('leftChars'), str(chars))
    el.set(W('hanging'), str(tw))
    el.set(W('hangingChars'), str(chars))
    el.set(W('firstLine'), '0')
    el.set(W('firstLineChars'), '0')
    return el


def install_styles(doc):
    """注册并配置论文体例的全部样式。

    标题必须是真 Heading 样式 + outlineLvl，否则目录生成不了、导航窗格是空的。
    报告体例的 install_styles 不能直接复用 —— 那份的一级标题是左对齐黑体小三，
    论文体例是一级黑体三号居中、三级楷体加粗、四级宋体加粗。
    """
    St = doc.styles

    # --- 正文：宋体小四 / 1.5 倍 / 首行缩进 2 字符 / 两端对齐 ---
    B.style_run(St['Normal'], S.BODY_PT, S.FONT_BODY_CJK, S.FONT_BODY_LAT)
    B.style_para(St['Normal'], spacing=S.LINE_SPACING,
                 ind_chars=S.FIRST_LINE_CHARS,
                 ind_twips=S.first_line_twips(S.BODY_PT),
                 before=0, after=0, jc='both')

    # --- 标题 1–4：黑体三号居中 / 黑体四号 / 楷体小四加粗 / 宋体小四加粗 ---
    for lvl in (1, 2, 3, 4):
        st = St['Heading %d' % lvl]
        B.style_run(st, SP.HEAD_PT[lvl], SP.FONT_H[lvl], SP.FONT_LAT,
                    bold=SP.HEAD_BOLD[lvl])
        B.style_para(st, spacing=S.LINE_SPACING,
                     before=SP.HEAD_SPACE_BEFORE[lvl],
                     after=SP.HEAD_SPACE_AFTER[lvl],
                     outline_lvl=lvl - 1, keep_next=True, clear_indent=True,
                     jc=SP.HEAD_JC[lvl])

    # --- 封面题名 / 封面字段 ---
    # 自建样式而非复用内置 Title：内置 Title 自带 accent1 蓝色下边框，会在题名下画一道蓝线
    title = B._ensure_custom_style(doc, SP.STYLE_COVER_TITLE)
    B.style_run(title, SP.COVER_TITLE_PT, SP.FONT_COVER_TITLE, S.FONT_HEAD_LAT, bold=True)
    B.style_para(title, spacing=S.LINE_SPACING, before=SP.COVER_TITLE_BEFORE,
                 after=SP.COVER_TITLE_AFTER, clear_indent=True, jc='center')
    B.strip_para_border(title)
    B.strip_para_border(St['Title'])      # 内置 Title 一并清，防止别处把蓝线带回来

    field = B._ensure_custom_style(doc, SP.STYLE_COVER_FIELD)
    B.style_run(field, SP.COVER_FIELD_PT, SP.FONT_COVER_FIELD, S.FONT_BODY_LAT)
    B.style_para(field, spacing=S.LINE_SPACING, after=SP.COVER_FIELD_AFTER,
                 clear_indent=True, jc='center')

    # --- 无封面时的文档标题（日常版首个 # 仍作文档标题）---
    doc_title = B._ensure_custom_style(doc, SP.STYLE_TITLE)
    B.style_run(doc_title, SP.DOC_TITLE_PT, S.FONT_HEAD_CJK, S.FONT_HEAD_LAT, bold=True)
    B.style_para(doc_title, spacing=S.LINE_SPACING, before=6, after=18,
                 clear_indent=True, jc='center')
    B.strip_para_border(doc_title)

    # --- 「目录」二字：不写 outlineLvl，否则它自己会出现在目录里 ---
    toc = B._ensure_custom_style(doc, SP.STYLE_TOC_TITLE)
    B.style_run(toc, SP.TOC_TITLE_PT, S.FONT_HEAD_CJK, S.FONT_HEAD_LAT, bold=True)
    B.style_para(toc, spacing=S.LINE_SPACING, before=SP.TOC_TITLE_BEFORE,
                 after=SP.TOC_TITLE_AFTER, clear_indent=True, jc='center')

    # --- 关键词：标签黑体加粗（标签的加粗在 run 上写），整行不缩进 ---
    kw = B._ensure_custom_style(doc, SP.STYLE_KW)
    B.style_run(kw, SP.KW_PT, S.FONT_BODY_CJK, S.FONT_BODY_LAT)
    B.style_para(kw, spacing=S.LINE_SPACING, before=SP.KW_SPACE_BEFORE,
                 after=SP.KW_SPACE_AFTER, clear_indent=True, jc='both')

    # --- 公式：左对齐 + 制表位（不要两端对齐，制表位会跑偏）---
    eq = B._ensure_custom_style(doc, SP.STYLE_EQ)
    B.style_run(eq, SP.EQ_PT, SP.FONT_EQ, S.FONT_BODY_LAT)
    B.style_para(eq, spacing=S.LINE_SPACING, before=SP.EQ_SPACE_BEFORE,
                 after=SP.EQ_SPACE_AFTER, clear_indent=True, jc='left')

    # --- 参考文献：五号 + 悬挂缩进 2 字符 ---
    # 基于 Normal 会继承它的首行缩进，set_hanging 已显式清零；这里再把 basedOn
    # 摘掉做双保险 —— 缩进规则一劳永逸，不用每次改样式都想着继承链。
    ref = B._ensure_custom_style(doc, SP.STYLE_REF)
    try:
        ref.base_style = None
    except Exception:
        pass
    B.style_run(ref, SP.REF_PT, S.FONT_BODY_CJK, S.FONT_BODY_LAT)
    B.style_para(ref, spacing=SP.REF_LINE_SPACING, before=SP.REF_SPACE_BEFORE,
                 after=SP.REF_SPACE_AFTER, clear_indent=True, jc='both')
    set_hanging(ref, SP.REF_HANG_CHARS, S.first_line_twips(SP.REF_PT))

    # --- 表题（与报告体例同名，M.add_table 等按名取用）/ 表注 / 代码 / 表格文字 ---
    cap = B._ensure_custom_style(doc, SP.STYLE_CAPTION)
    B.style_run(cap, SP.CAP_PT, SP.FONT_CAP, S.FONT_CAP_LAT, bold=True)
    B.style_para(cap, spacing=1.0, before=SP.CAP_SPACE_BEFORE,
                 after=SP.CAP_SPACE_AFTER, clear_indent=True, jc='center', keep_next=True)

    note = B._ensure_custom_style(doc, SP.STYLE_NOTE)
    B.style_run(note, SP.NOTE_PT, S.FONT_BODY_CJK, S.FONT_BODY_LAT)
    B.style_para(note, spacing=1.0, before=SP.NOTE_SPACE_BEFORE,
                 after=SP.NOTE_SPACE_AFTER, clear_indent=True, jc='left')

    code = B._ensure_custom_style(doc, SP.STYLE_CODE)
    B.style_run(code, S.CODE_PT, S.FONT_MONO, S.FONT_MONO)
    B.style_para(code, spacing=1.0, before=2, after=2, clear_indent=True, jc='left')

    tbl = B._ensure_custom_style(doc, SP.STYLE_TBL)
    B.style_run(tbl, S.TABLE_PT, S.FONT_BODY_CJK, S.FONT_BODY_LAT)
    B.style_para(tbl, spacing=S.TABLE_CELL_LINE_SPACING,
                 before=S.TABLE_CELL_SPACE_BEFORE, after=S.TABLE_CELL_SPACE_AFTER,
                 clear_indent=True, jc=S.TABLE_CELL_JC)


# ---------------------------------------------------------------- 各类块
def build_cover(doc, title, entries):
    """封面：顶部留白 + 题名（二号黑体居中）+ 各字段行（四号居中，整行照排）。

    封面这节的页脚全程不碰 —— 不碰就是没有页码。
    """
    for _ in range(SP.COVER_TOP_BLANK):
        M.plain_para(doc)
    if title:
        p = M.plain_para(doc, style=SP.STYLE_COVER_TITLE, align=WD_ALIGN_PARAGRAPH.CENTER)
        B.add_run(p, title, SP.COVER_TITLE_PT, SP.FONT_COVER_TITLE, S.FONT_HEAD_LAT,
                  bold=True)
    for ln in entries:
        p = M.plain_para(doc, style=SP.STYLE_COVER_FIELD, align=WD_ALIGN_PARAGRAPH.CENTER)
        B.add_run(p, ln, SP.COVER_FIELD_PT, SP.FONT_COVER_FIELD, S.FONT_BODY_LAT)


def add_kw(doc, head, text):
    """关键词行：`关键词：` 标签黑体加粗，后面照排（分号由源稿自带，不改写）。"""
    p = doc.add_paragraph(style=SP.STYLE_KW)
    B.add_run(p, head, SP.KW_PT, SP.FONT_KW_LABEL, S.FONT_BODY_LAT, bold=True)
    M.emit_runs(p, text, SP.KW_PT, S.FONT_BODY_CJK, S.FONT_BODY_LAT)
    return p


def add_ref(doc, num, text):
    """参考文献条目：[N] + 正文，五号、悬挂缩进 2 字符。"""
    p = doc.add_paragraph(style=SP.STYLE_REF)
    set_hanging(p, SP.REF_HANG_CHARS, S.first_line_twips(SP.REF_PT))
    B.add_run(p, '[%d] ' % num, SP.REF_PT, S.FONT_BODY_CJK, S.FONT_BODY_LAT)
    M.emit_runs(p, text, SP.REF_PT, S.FONT_BODY_CJK, S.FONT_BODY_LAT)
    return p


def mathpara(omath_xml):
    """`<m:oMath>` 包成 `<m:oMathPara>` —— 这是「显示式」排版的关键。

    Word 只在 oMathPara 里按显示式排公式（分数满号、Σ 的上下限画在上下）；
    行内 oMath 一律按文本式排（分式压成小号、上下限挤到右侧）。实测过：
    oMathPara 与普通 run 同段也不行 —— 所以版式必须是「表格中栏独占 oMathPara」。
    """
    return '<m:oMathPara %s>%s</m:oMathPara>' % (L._NSDECL, omath_xml)


def add_formula(doc, text, number=None, omml=True):
    """编号公式：1×3 无框表格 —— 空栏 | 公式（Word 公式对象，居中） | 编号（右对齐）。

    这是 Word 自己「插入带编号公式」模板的排法，也是唯一能让公式真居中的做法：
    用制表位凑的话，公式左端随公式宽度浮动，编号位置倒是对的、公式并不居中。

    omml=True 时把 LaTeX 转成 OMML —— Word 里就是**真公式**，双击可进编辑器改；
    转不了（超出子集）整条回退成纯文本。返回是否排成了真公式。
    """
    xml = L.to_omml(text) if omml else None
    widths = (SP.EQ_COL_PAD_CM, SP.EQ_COL_BODY_CM, SP.EQ_COL_PAD_CM)
    tbl = doc.add_table(rows=1, cols=3)
    tbl.style = 'Normal Table'               # 无框线基座（与三线表同一招）
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    B.set_fixed_layout(tbl)
    B.set_table_width_pct(tbl, 5000)
    twips_per_cm = 566.9291
    grid = tbl._tbl.find(W('tblGrid'))
    if grid is not None:
        for gc, cm in zip(grid.findall(W('gridCol')), widths):
            gc.set(W('w'), str(int(cm * twips_per_cm)))
    for cell, cm in zip(tbl.rows[0].cells, widths):
        cell.width = Cm(cm)
        B.set_cell_valign(cell, 'center')
        B.set_para_indent(cell.paragraphs[0], clear=True)   # 空栏也别继承正文缩进

    body = tbl.cell(0, 1).paragraphs[0]
    body.style = doc.styles[SP.STYLE_EQ]
    body.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if xml:
        body._p.append(parse_xml(mathpara(xml)))
    else:
        M.emit_runs(body, text, SP.EQ_PT, SP.FONT_EQ, S.FONT_BODY_LAT)

    if number:
        pnum = tbl.cell(0, 2).paragraphs[0]
        pnum.style = doc.styles[SP.STYLE_EQ]
        pnum.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        B.add_run(pnum, number, SP.EQ_PT, SP.FONT_EQ, S.FONT_BODY_LAT)
    return bool(xml)


def add_caption(doc, text, keep_next=True, before=None, after=None):
    """表题/图题。参数取论文体例（五号黑体），不是报告体例的小五。"""
    p = doc.add_paragraph(style=SP.STYLE_CAPTION)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    B.add_run(p, text, SP.CAP_PT, SP.FONT_CAP, S.FONT_CAP_LAT, bold=True)
    if not keep_next:
        B.set_keep_next(p, False)
    if before is not None or after is not None:
        B.set_para_spacing(p, line=1.0, before=before, after=after)
    return p


def add_toc(doc, title=None, anchor=None):
    """「目录」二字 + TOC 域（≤3 级）。返回 (标题段, 域段)。

    title 给源稿里显式写的 `# 目录`（可能是「目 录」「Contents」）——
    照它的字排，别换成我们自己的两个字，否则内容无损校验两侧对不上。
    TOC 是域，缓存文本只是提示；打开 Word 时 settings 里的 updateFields 会重算。
    """
    p = doc.add_paragraph(style=SP.STYLE_TOC_TITLE)
    B.add_run(p, title or SP.TOC_TITLE, SP.TOC_TITLE_PT, S.FONT_HEAD_CJK,
              S.FONT_HEAD_LAT, bold=True)
    field = B.add_toc_field(doc, anchor or p, levels=SP.TOC_LEVELS)
    # 域段落自身清掉缩进/行距：Word 重算目录时会沿用这个段落的格式
    B.set_para_indent(field, clear=True)
    B.set_para_spacing(field, line=S.LINE_SPACING, before=0, after=0)
    return p, field