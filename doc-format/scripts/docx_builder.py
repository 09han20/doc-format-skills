# -*- coding: utf-8 -*-
"""Word 原语层 —— 唯一直接操作 lxml 的模块。

python-docx 1.2.0 在这些地方是零支持或不完整的，全部手写 XML：
  · 中文字体：CT_Fonts 只暴露 @w:ascii / @w:hAnsi，没有 @w:eastAsia
  · 字号：font.size 只写 w:sz，从不写 w:szCs
  · 首行缩进「2 字符」：只有绝对长度的 first_line_indent，没有 w:firstLineChars
  · 表格边框：w:tblBorders / w:tcBorders 无高层 API
  · 域：全库无 fldChar / instrText / fldSimple 支持

所有插入都走 insert_in_order()，因为 OOXML 对子元素顺序是 schema 强制的，
插错位置 Word 会报「文件内容有问题」。
"""
import os

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import RGBColor, Pt, Cm

import spec

W = lambda t: qn('w:' + t)

# ---------------------------------------------------------------- schema 顺序
# 只列出我们可能插入的元素，未列出的按「排最后」处理。
PPR_SEQ = [W(t) for t in (
    'pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl',
    'numPr', 'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens',
    'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN',
    'bidi', 'adjustRightInd', 'snapToGrid', 'spacing', 'ind', 'contextualSpacing',
    'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection', 'textAlignment',
    'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange')]

RPR_SEQ = [W(t) for t in (
    'rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike',
    'outline', 'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish',
    'webHidden', 'color', 'spacing', 'w', 'kern', 'position', 'sz', 'szCs', 'highlight',
    'u', 'effect', 'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang',
    'eastAsianLayout', 'specVanish', 'oMath')]

TBLPR_SEQ = [W(t) for t in (
    'tblStyle', 'tblpPr', 'tblOverlap', 'bidiVisual', 'tblStyleRowBandSize',
    'tblStyleColBandSize', 'tblW', 'jc', 'tblCellSpacing', 'tblInd', 'tblBorders',
    'shd', 'tblLayout', 'tblCellMar', 'tblLook', 'tblCaption', 'tblDescription',
    'tblPrChange')]

TCPR_SEQ = [W(t) for t in (
    'cnfStyle', 'tcW', 'gridSpan', 'hMerge', 'vMerge', 'tcBorders', 'shd', 'noWrap',
    'tcMar', 'textDirection', 'tcFitText', 'vAlign', 'hideMark', 'headers', 'cellIns',
    'cellDel', 'cellMerge', 'tcPrChange')]

TRPR_SEQ = [W(t) for t in (
    'cnfStyle', 'divId', 'gridBefore', 'gridAfter', 'wBefore', 'wAfter', 'cantSplit',
    'trHeight', 'tblHeader', 'tblCellSpacing', 'jc', 'hidden', 'ins', 'del', 'trPrChange')]


def insert_in_order(parent, child, seq):
    """按 schema 顺序把 child 插进 parent。已存在同等或更后序的元素则插其前面。"""
    tag = child.tag
    idx = seq.index(tag) if tag in seq else len(seq)
    for existing in parent.iterchildren():
        if existing.tag in seq and seq.index(existing.tag) > idx:
            existing.addprevious(child)
            return child
    parent.append(child)
    return child


def find_or_add(parent, tag, seq):
    """tag 是 Clark name（{ns}local）；OxmlElement 要的是 'w:local'，故补前缀。"""
    el = parent.find(tag)
    if el is None:
        local = tag.split('}')[1] if '}' in tag else tag
        el = OxmlElement(local if ':' in local else 'w:' + local)
        insert_in_order(parent, el, seq)
    return el


# ---------------------------------------------------------------- run 属性
def set_rfonts(rpr, cjk, lat, cs=None):
    """写全 rFonts 四槽并清除主题引用（主题会盖掉显式字体）。"""
    rf = find_or_add(rpr, W('rFonts'), RPR_SEQ)
    rf.set(W('ascii'), lat)
    rf.set(W('hAnsi'), lat)
    rf.set(W('eastAsia'), cjk)
    rf.set(W('cs'), cs or lat)
    for dead in ('asciiTheme', 'hAnsiTheme', 'eastAsiaTheme', 'cstheme'):
        if rf.get(W(dead)) is not None:
            del rf.attrib[W(dead)]
    return rf


def set_sz(rpr, pt):
    """w:sz 与 w:szCs 必须同时写；只写前者会让中文在某些视图下不缩放。"""
    for tag in ('sz', 'szCs'):
        el = find_or_add(rpr, W(tag), RPR_SEQ)
        el.set(W('val'), str(spec.hp(pt)))


def set_bold(rpr, on=True):
    for tag in ('b', 'bCs'):
        el = rpr.find(W(tag))
        if on:
            if el is None:
                el = OxmlElement('w:' + tag)
                insert_in_order(rpr, el, RPR_SEQ)
            el.set(W('val'), '1')
        elif el is not None:
            rpr.remove(el)


def set_color(rpr, val='000000'):
    el = find_or_add(rpr, W('color'), RPR_SEQ)
    el.set(W('val'), val)
    for dead in ('themeColor', 'themeTint', 'themeShade'):
        if el.get(W(dead)) is not None:
            del el.attrib[W(dead)]


def style_run(style, pt, cjk, lat, bold=False, color='000000'):
    """配置一个段落样式层面的 rPr。"""
    rpr = style.element.get_or_add_rPr()
    set_rfonts(rpr, cjk, lat)
    set_sz(rpr, pt)
    set_bold(rpr, bold)
    if color:
        set_color(rpr, color)
    return rpr


def strip_para_border(style):
    """剥离样式自带的段落边框（w:pPr/w:pBdr）。

    Word 内置 Title 样式带一条 accent1（4F81BD）下边框 —— 中文报告封面标题
    下面那道蓝线就是它。内置 Caption 同理。自建样式虽不含边框，但内置样式
    一旦被引用就会带出来，所以这里做一次显式清除。
    """
    ppr = style.element.get_or_add_pPr()
    bdr = ppr.find(W('pBdr'))
    if bdr is not None:
        ppr.remove(bdr)
        return True
    return False


def style_para(style, *, spacing=None, ind_chars=None, ind_twips=None,
               before=None, after=None, jc=None, outline_lvl=None,
               keep_next=False, clear_indent=False):
    """配置一个段落样式层面的 pPr。"""
    ppr = style.element.get_or_add_pPr()
    if spacing is not None:
        el = find_or_add(ppr, W('spacing'), PPR_SEQ)
        el.set(W('line'), str(int(round(spec.hp(spacing) * 240 / 2))))
        el.set(W('lineRule'), 'auto')
        el.set(W('before'), str(int(round((before or 0) * 20))))
        el.set(W('after'), str(int(round((after or 0) * 20))))
    if clear_indent:
        # 显式双 0：不能删属性，w:ind 会走 basedOn 链继承回来
        el = find_or_add(ppr, W('ind'), PPR_SEQ)
        for k in ('firstLine', 'firstLineChars', 'hanging', 'hangingChars', 'leftChars'):
            el.set(W(k), '0')
        el.set(W('left'), '0')
    elif ind_chars is not None:
        el = find_or_add(ppr, W('ind'), PPR_SEQ)
        el.set(W('firstLineChars'), str(ind_chars))
        el.set(W('firstLine'), str(ind_twips if ind_twips is not None
                                    else spec.first_line_twips()))
    if jc is not None:
        find_or_add(ppr, W('jc'), PPR_SEQ).set(W('val'), jc)
    if outline_lvl is not None:
        find_or_add(ppr, W('outlineLvl'), PPR_SEQ).set(W('val'), str(outline_lvl))
    if keep_next:
        el = ppr.find(W('keepNext'))
        if el is None:
            insert_in_order(ppr, OxmlElement('w:keepNext'), PPR_SEQ)
    return ppr


# ---------------------------------------------------------------- 段落级
def set_para_indent(p, chars=None, twips=None, clear=False):
    """段落级首行缩进。clear=True 写显式双 0（表格单元格必须用）。"""
    ppr = p._p.get_or_add_pPr()
    el = find_or_add(ppr, W('ind'), PPR_SEQ)
    if clear:
        for k in ('firstLine', 'firstLineChars', 'hanging', 'hangingChars'):
            el.set(W(k), '0')
    else:
        el.set(W('firstLineChars'), str(chars))
        el.set(W('firstLine'), str(twips if twips is not None
                                   else spec.first_line_twips()))
    return el


def set_para_spacing(p, line=None, before=None, after=None):
    ppr = p._p.get_or_add_pPr()
    el = find_or_add(ppr, W('spacing'), PPR_SEQ)
    if line is not None:
        el.set(W('line'), str(int(round(line * 240))))
        el.set(W('lineRule'), 'auto')
    if before is not None:
        el.set(W('before'), str(int(round(before * 20))))
    if after is not None:
        el.set(W('after'), str(int(round(after * 20))))
    return el


def set_keep_next(p, on=True):
    """与下段同页。on=False 显式写 val=0 —— 样式层（表题）带 keepNext，
    段落要关掉它只能显式写 0，删元素会从样式继承回来。"""
    ppr = p._p.get_or_add_pPr()
    el = ppr.find(W('keepNext'))
    if el is None:
        el = OxmlElement('w:keepNext')
        insert_in_order(ppr, el, PPR_SEQ)
    el.set(W('val'), '1' if on else '0')


def shade(p, fill):
    ppr = p._p.get_or_add_pPr()
    shd = find_or_add(ppr, W('shd'), PPR_SEQ)
    shd.set(W('val'), 'clear')
    shd.set(W('color'), 'auto')
    shd.set(W('fill'), fill)


# ---------------------------------------------------------------- 样式注册
def _ensure_custom_style(doc, name, base='Normal'):
    """取或建一个段落样式（按 UI 名）。"""
    from docx.enum.style import WD_STYLE_TYPE
    styles = doc.styles
    try:
        return styles[name]
    except KeyError:
        st = styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        st.base_style = styles[base]
        return st


def install_styles(doc, body_pt=None):
    """注册/配置规范要求的全部样式。返回一个 name → style 的字典。

    这是整个生成器的根治点：标题必须是真正的 Word 标题样式，
    否则目录生成不了、导航窗格空的、大纲里什么都没有。
    """
    body_pt = body_pt or spec.BODY_PT
    S = doc.styles

    # --- 正文 ---
    style_run(S['Normal'], body_pt, spec.FONT_BODY_CJK, spec.FONT_BODY_LAT)
    style_para(S['Normal'], spacing=spec.LINE_SPACING,
               ind_chars=spec.FIRST_LINE_CHARS,
               ind_twips=spec.first_line_twips(body_pt),
               before=spec.BODY_SPACE_BEFORE, after=spec.BODY_SPACE_AFTER,
               jc='both')

    # --- 标题 1-4（真样式 + outlineLvl，TOC 靠它） ---
    for lvl in (1, 2, 3, 4):
        st = S['Heading %d' % lvl]
        style_run(st, spec.HEAD_PT[lvl], spec.FONT_HEAD_CJK, spec.FONT_HEAD_LAT, bold=True)
        style_para(st, spacing=spec.LINE_SPACING,
                   before=spec.HEAD_SPACE_BEFORE[lvl], after=spec.HEAD_SPACE_AFTER[lvl],
                   outline_lvl=lvl - 1, keep_next=True,
                   clear_indent=True,
                   jc='center' if lvl == 1 else 'left')
    # Heading 1 兼作章标题，不居中；文档标题另用 Title 样式
    style_para(S['Heading 1'], spacing=spec.LINE_SPACING,
               before=spec.HEAD_SPACE_BEFORE[1], after=spec.HEAD_SPACE_AFTER[1],
               outline_lvl=0, keep_next=True, clear_indent=True, jc='left')

    # --- 文档标题（封面）---
    # 自建样式而非复用内置 Title：内置 Title 自带 w:pBdr 蓝色 accent1 下边框，
    # 会把封面标题下面画出一道蓝线。自建样式基于 Normal，无边框。
    title = _ensure_custom_style(doc, spec.STYLE_TITLE)
    style_run(title, spec.TITLE_PT, spec.FONT_HEAD_CJK, spec.FONT_HEAD_LAT, bold=True)
    style_para(title, spacing=spec.LINE_SPACING,
               before=spec.TITLE_SPACE_BEFORE, after=spec.TITLE_SPACE_AFTER,
               clear_indent=True, jc='center')
    strip_para_border(title)
    # 内置 Title 一并清边框：防止别处（或样式继承）把蓝线带回来
    strip_para_border(S['Title'])

    # --- 表题 / 表注 / 代码 ---
    cap = _ensure_custom_style(doc, spec.STYLE_CAPTION)
    style_run(cap, spec.CAP_PT, spec.FONT_CAP_CJK, spec.FONT_CAP_LAT, bold=True)
    style_para(cap, spacing=1.0, before=spec.CAP_SPACE_BEFORE, after=spec.CAP_SPACE_AFTER,
               clear_indent=True, jc='center', keep_next=True)

    note = _ensure_custom_style(doc, '表注')
    style_run(note, spec.NOTE_PT, spec.FONT_BODY_CJK, spec.FONT_BODY_LAT)
    style_para(note, spacing=1.0, before=spec.NOTE_SPACE_BEFORE, after=spec.NOTE_SPACE_AFTER,
               clear_indent=True, jc='left')

    code = _ensure_custom_style(doc, '代码')
    style_run(code, spec.CODE_PT, spec.FONT_MONO, spec.FONT_MONO)
    style_para(code, spacing=1.0, before=2, after=2, clear_indent=True, jc='left')

    # --- 表格内文字：显式禁缩进，否则会从 Normal 继承出 2 字符缩进 ---
    tbltxt = _ensure_custom_style(doc, '表格文字')
    style_run(tbltxt, spec.TABLE_PT, spec.FONT_BODY_CJK, spec.FONT_BODY_LAT)
    style_para(tbltxt, spacing=spec.TABLE_CELL_LINE_SPACING,
               before=spec.TABLE_CELL_SPACE_BEFORE, after=spec.TABLE_CELL_SPACE_AFTER,
               clear_indent=True, jc=spec.TABLE_CELL_JC)

    return {s.name: s for s in doc.styles}


# ---------------------------------------------------------------- 三线表
def three_line_table(tbl, header_rows=1):
    """把表格改成三线表：顶线粗、底线粗、表头行下细线，其余无框线。

    注意：必须先把 tblStyle 换成无边框基座。直接格式(tcBorders)优先级最高，
    样式边框(tblBorders)次之 —— 若沿用 Table Grid，全网格线会盖住我们写的值。
    """
    tbl.style = 'Normal Table'
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER

    tblPr = tbl._tbl.tblPr
    old = tblPr.find(W('tblBorders'))
    if old is not None:
        tblPr.remove(old)
    borders = OxmlElement('w:tblBorders')
    # 六边必须全部显式写出：省略的边会从样式继承，可能继承回非 none
    for edge, val, sz in (('top', 'single', spec.RULE_THICK),
                          ('bottom', 'single', spec.RULE_THICK),
                          ('left', 'none', 0),
                          ('right', 'none', 0),
                          ('insideH', 'none', 0),
                          ('insideV', 'none', 0)):
        e = OxmlElement('w:' + edge)
        e.set(W('val'), val)
        e.set(W('sz'), str(sz))
        e.set(W('space'), '0')
        e.set(W('color'), spec.RULE_COLOR)
        borders.append(e)
    insert_in_order(tblPr, borders, TBLPR_SEQ)

    for r in range(min(header_rows, len(tbl.rows))):
        _bottom_border(tbl.rows[r], spec.RULE_THIN)


def _bottom_border(row, sz):
    for cell in row.cells:
        tcPr = cell._tc.get_or_add_tcPr()
        old = tcPr.find(W('tcBorders'))
        if old is not None:
            tcPr.remove(old)
        tb = OxmlElement('w:tcBorders')
        bot = OxmlElement('w:bottom')
        bot.set(W('val'), 'single')
        bot.set(W('sz'), str(sz))
        bot.set(W('space'), '0')
        bot.set(W('color'), spec.RULE_COLOR)
        tb.append(bot)
        insert_in_order(tcPr, tb, TCPR_SEQ)


def set_cell_valign(cell, val='center'):
    """单元格垂直对齐（w:tcPr/w:vAlign）。表体水平居中时配垂直居中最协调。"""
    tcPr = cell._tc.get_or_add_tcPr()
    find_or_add(tcPr, W('vAlign'), TCPR_SEQ).set(W('val'), val)


def set_row_header(row, repeat=True):
    trPr = row._tr.get_or_add_trPr()
    if repeat and trPr.find(W('tblHeader')) is None:
        insert_in_order(trPr, OxmlElement('w:tblHeader'), TRPR_SEQ)


def set_fixed_layout(tbl):
    tblPr = tbl._tbl.tblPr
    lay = find_or_add(tblPr, W('tblLayout'), TBLPR_SEQ)
    lay.set(W('type'), 'fixed')


def set_table_width_pct(tbl, pct=5000):
    tblPr = tbl._tbl.tblPr
    w = find_or_add(tblPr, W('tblW'), TBLPR_SEQ)
    w.set(W('type'), 'pct')
    w.set(W('w'), str(pct))


# ---------------------------------------------------------------- 图片
def natural_cm(path):
    """图片的原始物理尺寸（cm）。读不出 dpi 时返回 None，交给调用方兜底。"""
    from docx.image.image import Image as DocxImage
    try:
        im = DocxImage.from_file(path)
    except Exception:
        return None
    return im.width / 360000.0, im.height / 360000.0      # EMU → cm


def add_picture_para(doc, path, max_w_cm=None, align=WD_ALIGN_PARAGRAPH.CENTER):
    """插图：居中、显式清零首行缩进、超过正文栏宽时等比缩放到栏宽。

    段落级清零缩进不能省 —— Normal 有 2 字符首行缩进，图片会被推得偏右。
    """
    if not os.path.isfile(path):
        raise FileNotFoundError('插图文件不存在：%s' % path)
    max_w_cm = max_w_cm or spec.CONTENT_W_CM
    p = doc.add_paragraph()
    p.alignment = align
    set_para_indent(p, clear=True)
    set_para_spacing(p, line=1.0, before=spec.PIC_SPACE_BEFORE, after=spec.PIC_SPACE_AFTER)
    nat = natural_cm(path)
    if nat and nat[0] > max_w_cm:
        p.add_run().add_picture(path, width=Cm(max_w_cm))     # 只给宽，等比缩放
    else:
        p.add_run().add_picture(path)
    return p


# ---------------------------------------------------------------- 域
def add_page_number(section, pt=None, align=WD_ALIGN_PARAGRAPH.CENTER):
    """页脚居中页码。用 fldSimple —— 比 fldChar 三件套短，Word 同样认。"""
    section.footer.is_linked_to_previous = False
    footer = section.footer
    p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    p.alignment = align
    for r in list(p.runs):
        r._element.getparent().remove(r._element)
    fld = OxmlElement('w:fldSimple')
    fld.set(W('instr'), 'PAGE  \\* MERGEFORMAT')
    run = OxmlElement('w:r')
    rpr = OxmlElement('w:rPr')
    set_rfonts(rpr, spec.FONT_BODY_CJK, spec.FONT_BODY_LAT)
    set_sz(rpr, pt or spec.FOOT_PT)
    run.append(rpr)
    t = OxmlElement('w:t')
    t.text = '1'
    run.append(t)
    fld.append(run)
    p._p.append(fld)
    return p


def add_toc_field(doc, after_p, levels='1-3'):
    """插入 TOC 域。after_p 之后插入；返回新建的段落。

    begin 上带 w:dirty="true"，Word 打开时会重算这个域；settings 里的
    updateFields 是双保险。必须带 separate + 缓存文本，否则打开时域区是空的。
    """
    from docx.text.paragraph import Paragraph
    xml = (
        '<w:p %s>'
        '<w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>'
        '<w:r><w:instrText xml:space="preserve"> TOC \\o "%s" \\h \\z \\u </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        '<w:r><w:t>请在 Word 中按 Ctrl+A 后 F9 更新目录。</w:t></w:r>'
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
        '</w:p>' % (_nsdecls(), levels)
    )
    from docx.oxml import parse_xml
    p = Paragraph(parse_xml(xml), after_p._parent)
    after_p._p.addnext(p._p)
    return p


def _nsdecls():
    from docx.oxml.ns import nsdecls
    return nsdecls('w')


SETTINGS_SEQ = [W(t) for t in (
    'writeProtection', 'view', 'zoom', 'removePersonalInformation',
    'removeDateAndTime', 'doNotDisplayPageBoundaries', 'displayBackgroundShape',
    'printPostScriptOverText', 'printFractionalCharacterWidth', 'printFormsData',
    'embedTrueTypeFonts', 'embedSystemFonts', 'saveSubsetFonts', 'saveFormsData',
    'mirrorMargins', 'alignBordersAndEdges', 'bordersDoNotSurroundHeader',
    'bordersDoNotSurroundFooter', 'gutterAtTop', 'hideSpellingErrors',
    'hideGrammaticalErrors', 'activeWritingStyle', 'proofState', 'formsDesign',
    'attachedTemplate', 'linkStyles', 'stylePaneFormatFilter', 'stylePaneSortMethod',
    'documentType', 'mailMerge', 'revisionView', 'trackChanges', 'doNotTrackMoves',
    'doNotTrackFormatting', 'documentProtection', 'autoFormatOverride',
    'styleLockTheme', 'styleLockQFSet', 'defaultTabStop', 'autoHyphenation',
    'consecutiveHyphenLimit', 'hyphenationZone', 'doNotHyphenateCaps', 'showEnvelope',
    'summaryLength', 'clickAndTypeStyle', 'defaultTableStyle', 'evenAndOddHeaders',
    'bookFoldRevPrinting', 'bookFoldPrinting', 'bookFoldPrintingSheets',
    'drawingGridHorizontalSpacing', 'drawingGridVerticalSpacing', 'displayHorizontalDrawingGridEvery',
    'displayVerticalDrawingGridEvery', 'doNotUseMarginsForDrawingGridOrigin',
    'drawingGridHorizontalOrigin', 'drawingGridVerticalOrigin', 'doNotShadeFormData',
    'noPunctuationKerning', 'characterSpacingControl', 'printTwoOnOne', 'strictFirstAndLastChars',
    'noLineBreaksAfter', 'noLineBreaksBefore', 'savePreviewPicture',
    'doNotValidateAgainstSchema', 'saveInvalidXml', 'ignoreMixedContent',
    'alwaysShowPlaceholderText', 'doNotDemarcateInvalidXml', 'saveXmlDataOnly',
    'useXSLTWhenSaving', 'saveThroughXslt', 'showXMLTags', 'alwaysMergeEmptyNamespace',
    'updateFields', 'hdrShapeDefaults', 'footnotePr', 'endnotePr', 'compat',
    'docVars', 'rsids')]


def ensure_update_fields(doc, on=True):
    """settings.xml 里设 updateFields —— 幂等（已有则改值，不重复插）。"""
    sel = doc.settings.element
    uf = sel.find(W('updateFields'))
    if uf is None:
        uf = OxmlElement('w:updateFields')
        insert_in_order(sel, uf, SETTINGS_SEQ)
    uf.set(W('val'), 'true' if on else 'false')
    return uf


# ---------------------------------------------------------------- 行内
def add_hyperlink(paragraph, url, text, cjk=None, lat=None, pt=None):
    """真超链接（带外部关系），不是纯文本。"""
    part = paragraph.part
    r_id = part.relate_to(url, RT.HYPERLINK, is_external=True)
    hl = OxmlElement('w:hyperlink')
    hl.set(qn('r:id'), r_id)
    run = OxmlElement('w:r')
    rpr = OxmlElement('w:rPr')
    set_rfonts(rpr, cjk or spec.FONT_BODY_CJK, lat or spec.FONT_BODY_LAT)
    set_sz(rpr, pt or spec.BODY_PT)
    set_color(rpr, '0563C1')
    u = OxmlElement('w:u')
    u.set(W('val'), 'single')
    insert_in_order(rpr, u, RPR_SEQ)
    run.append(rpr)
    t = OxmlElement('w:t')
    t.text = text
    run.append(t)
    hl.append(run)
    paragraph._p.append(hl)
    return hl


def add_run(p, text, pt, cjk, lat, bold=False, mono=False, color=None):
    r = p.add_run(text)
    rpr = r._element.get_or_add_rPr()
    set_rfonts(rpr, cjk, lat)
    set_sz(rpr, pt)
    if bold:
        set_bold(rpr, True)
    if color:
        set_color(rpr, color)
    if mono:
        r.font.name = lat
    return r