# -*- coding: utf-8 -*-
"""Markdown → 规范中文科技报告 Word 文档。

    python md2docx.py 报告.md                 # 输出 报告.docx
    python md2docx.py 报告.md -o out.docx --toc
    python md2docx.py 报告.md --caption-mode auto --body-size 五号

标题层级：第一个 h1 作文档标题（居中、不进目录编号），其余 h1 作章。
表题：`表：说明文字` 自动编号；`表 7 说明文字` 沿用你写的号并校验连续性。
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# 控制台是 GBK(cp936) 时，末尾的 ⚠ / ✅ 会让 print 直接抛 UnicodeEncodeError ——
# 实测中文 Windows 默认控制台下必崩（文件已存好，但退出码变 1，自动化校验全废）。
# 统一改走 UTF-8 输出：既不崩，也不会把中文显示成乱码。
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding='utf-8')
    except Exception:
        pass

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

import docx_builder as B
import md_parse as P
import spec


# ---------------------------------------------------------------- 行内
def emit_runs(p, text, pt, cjk, lat, bold=False):
    for kind, payload in P.iter_inline(text):
        if kind == 't':
            if payload:
                B.add_run(p, payload, pt, cjk, lat, bold=bold)
        elif kind == 'b':
            # 上游已不再产出 'b'（见 md_parse.iter_inline）。留着只为兜底，
            # 万一还有 'b' 也只按正文排——正文不加粗，且绝不能因漏分支丢字。
            B.add_run(p, payload, pt, cjk, lat, bold=bold)
        elif kind == 'c':
            B.add_run(p, payload, pt, spec.FONT_MONO, spec.FONT_MONO, bold=bold)
        elif kind == 'a':
            label, url = payload
            B.add_hyperlink(p, url, label or url, cjk, lat, pt)
    return p


def body_para(doc, align=WD_ALIGN_PARAGRAPH.JUSTIFY, indent=True):
    p = doc.add_paragraph()
    p.alignment = align
    if not indent:
        # 裸 <w:ind/> 清不掉缩进：OOXML 是「逐属性合并」，缺 firstLineChars 时会从
        # Normal 继承回 200。必须像 set_para_indent(clear=True) 那样显式写 0。
        B.set_para_indent(p, clear=True)
    return p


def plain_para(doc, style=None, align=WD_ALIGN_PARAGRAPH.LEFT, indent=False):
    """建段落；indent=False 时显式清零首行缩进（不能靠删属性，会从样式继承）。"""
    p = doc.add_paragraph(style=style) if style else doc.add_paragraph()
    p.alignment = align
    if not indent:
        B.set_para_indent(p, clear=True)
    return p


# ---------------------------------------------------------------- 各类块
def add_heading(doc, level, text):
    p = doc.add_paragraph(style='Heading %d' % level)
    emit_runs(p, text, spec.HEAD_PT[level], spec.FONT_HEAD_CJK, spec.FONT_HEAD_LAT,
              bold=True)
    return p


def add_code(doc, text):
    lines = text.split('\n') or ['']
    for ln in lines:
        p = plain_para(doc, style='代码')
        B.set_para_spacing(p, line=1.0, before=0, after=0)
        B.shade(p, spec.CODE_SHADING)
        B.add_run(p, ln if ln else ' ', spec.CODE_PT, spec.FONT_MONO, spec.FONT_MONO)
    return None


def add_quote(doc, text):
    p = plain_para(doc, align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    p.paragraph_format.left_indent = Pt(18)
    ppr = p._p.get_or_add_pPr()
    pbdr = OxmlElement('w:pBdr')
    left = OxmlElement('w:left')
    left.set(qn('w:val'), 'single')
    left.set(qn('w:sz'), str(spec.QUOTE_BAR_SZ))
    left.set(qn('w:space'), '8')
    left.set(qn('w:color'), spec.QUOTE_BAR_COLOR)
    pbdr.append(left)
    B.insert_in_order(ppr, pbdr, B.PPR_SEQ)
    emit_runs(p, text, spec.BODY_PT, spec.FONT_BODY_CJK, spec.FONT_BODY_LAT)
    return p


def add_hr(doc):
    p = plain_para(doc)
    B.set_para_spacing(p, line=1.0, before=6, after=6)
    ppr = p._p.get_or_add_pPr()
    pbdr = OxmlElement('w:pBdr')
    bot = OxmlElement('w:bottom')
    bot.set(qn('w:val'), 'single')
    bot.set(qn('w:sz'), str(spec.RULE_THIN))
    bot.set(qn('w:space'), '1')
    bot.set(qn('w:color'), spec.HR_COLOR)
    pbdr.append(bot)
    B.insert_in_order(ppr, pbdr, B.PPR_SEQ)
    return p


def add_pagebreak(doc):
    p = doc.add_paragraph()
    p.add_run().add_break(WD_BREAK.PAGE)
    return p


def add_list_item(doc, text, marker, hang=Pt(20)):
    p = plain_para(doc, align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    B.set_para_indent(p, clear=True)
    p.paragraph_format.left_indent = hang
    p.paragraph_format.first_line_indent = Pt(-18)
    B.add_run(p, marker + ' ', spec.BODY_PT, spec.FONT_BODY_CJK, spec.FONT_BODY_LAT)
    emit_runs(p, text, spec.BODY_PT, spec.FONT_BODY_CJK, spec.FONT_BODY_LAT)
    return p


def add_caption(doc, text, keep_next=True, before=None, after=None):
    """表题 / 图题。图题在图下方：keep_next=False（样式层默认开，要显式关掉），
    改由图片段落 keepNext 把两者绑在一起。"""
    p = plain_para(doc, style='表题', align=WD_ALIGN_PARAGRAPH.CENTER)
    B.add_run(p, text, spec.CAP_PT, spec.FONT_CAP_CJK, spec.FONT_CAP_LAT, bold=True)
    if not keep_next:
        B.set_keep_next(p, False)
    if before is not None or after is not None:
        B.set_para_spacing(p, line=1.0, before=before, after=after)
    return p


def add_note(doc, text):
    """表注按原文渲染，不补也不改前缀（改写会破坏内容无损）。"""
    p = plain_para(doc, style='表注')
    B.add_run(p, text, spec.NOTE_PT, spec.FONT_BODY_CJK, spec.FONT_BODY_LAT)
    return p


def add_table(doc, rows):
    ncol = max(len(r) for r in rows)
    tbl = doc.add_table(rows=len(rows), cols=ncol)
    B.three_line_table(tbl)
    B.set_table_width_pct(tbl, 5000)
    B.set_fixed_layout(tbl)
    B.set_row_header(tbl.rows[0])

    widths = P.col_widths(rows, floor=spec.TABLE_COL_FLOOR,
                          first_floor=spec.TABLE_COL_FIRST_FLOOR)
    twips_per_cm = 566.9291
    grid = tbl._tbl.find(qn('w:tblGrid'))
    if grid is not None:
        for gc, frac in zip(grid.findall(qn('w:gridCol')), widths):
            gc.set(qn('w:w'), str(int(spec.CONTENT_W_CM * twips_per_cm * frac)))

    ALIGN = {'center': WD_ALIGN_PARAGRAPH.CENTER,
             'left': WD_ALIGN_PARAGRAPH.LEFT,
             'right': WD_ALIGN_PARAGRAPH.RIGHT}

    # 长文本列改左对齐（可选）：整句话居中会让右边缘参差
    left_cols = set()
    if spec.TABLE_LONG_TEXT_LEFT:
        for j in range(ncol):
            longest = max((P.plain_len(r[j]) for r in rows if j < len(r)), default=0)
            if longest > spec.TABLE_LONG_TEXT_THRESHOLD:
                left_cols.add(j)

    for i, row in enumerate(rows):
        for j in range(ncol):
            cell = tbl.cell(i, j)
            cell.width = Cm(spec.CONTENT_W_CM * widths[j])
            txt = row[j] if j < len(row) else ''
            p = cell.paragraphs[0]
            p.style = doc.styles['表格文字']
            if i == 0:
                p.alignment = ALIGN[spec.TABLE_HEADER_JC]
            elif j in left_cols:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            else:
                p.alignment = ALIGN[spec.TABLE_CELL_JC]
            B.set_cell_valign(cell, spec.TABLE_CELL_VALIGN)
            # 显式清零：Normal 的 2 字符缩进会继承进来，实测能毁掉整张表
            B.set_para_indent(p, clear=True)
            B.set_para_spacing(p, line=spec.TABLE_CELL_LINE_SPACING,
                               before=spec.TABLE_CELL_SPACE_BEFORE,
                               after=spec.TABLE_CELL_SPACE_AFTER)
            emit_runs(p, txt, spec.TABLE_PT, spec.FONT_BODY_CJK, spec.FONT_BODY_LAT,
                      bold=(i == 0))
    return tbl


# ---------------------------------------------------------------- 主流程
def _next_caption_num(cap, counters, caption_mode):
    """题注编号：手写号优先（同步计数器），无号顺延；auto 模式一律重编。"""
    k, num = cap['kind'], cap['num']
    counters.setdefault(k, 0)
    if num is None or caption_mode == 'auto':
        counters[k] += 1
        num = counters[k]
    else:
        counters[k] = num
    return '%s %d %s' % (k, num, cap['text'])


def _resolve_img(src, base_dir):
    """图片路径按 md 文件所在目录解析 —— 脚本在别处运行也能找到图。"""
    return src if os.path.isabs(src) else os.path.normpath(os.path.join(base_dir, src))


def build(src, out=None, toc=False, caption_mode='both', body_pt=None,
          page_number=True, title=None):
    body_pt = body_pt or spec.BODY_PT
    blocks = P.parse_blocks(src, is_path=True)
    base_dir = os.path.dirname(os.path.abspath(src))

    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(spec.PAGE_W_CM), Cm(spec.PAGE_H_CM)
    sec.top_margin, sec.bottom_margin = Cm(spec.MARGIN_TOP_CM), Cm(spec.MARGIN_BOTTOM_CM)
    sec.left_margin, sec.right_margin = Cm(spec.MARGIN_LEFT_CM), Cm(spec.MARGIN_RIGHT_CM)
    if page_number:
        B.add_page_number(sec)

    B.install_styles(doc, body_pt)

    counters = {'表': 0, '图': 0, '代码': 0}
    stats = {'tables': 0, 'images': 0, 'captioned': 0, 'headings': 0,
             'uncaptioned': []}
    toc_anchor = None

    for kind, payload in blocks:
        if kind == 'title':
            p = plain_para(doc, style=spec.STYLE_TITLE, align=WD_ALIGN_PARAGRAPH.CENTER)
            emit_runs(p, title or payload, spec.TITLE_PT,
                      spec.FONT_HEAD_CJK, spec.FONT_HEAD_LAT, bold=True)
            toc_anchor = p

        elif kind == 'h':
            level, text = payload
            stats['headings'] += 1
            add_heading(doc, level, text)

        elif kind == 'p':
            p = body_para(doc)
            emit_runs(p, payload, body_pt, spec.FONT_BODY_CJK, spec.FONT_BODY_LAT)

        elif kind == 'quote':
            add_quote(doc, payload)

        elif kind == 'hr':
            add_hr(doc)

        elif kind == 'pb':
            add_pagebreak(doc)

        elif kind == 'todo':
            add_list_item(doc, payload, spec.TODO_BOX)

        elif kind == 'ul':
            add_list_item(doc, payload, spec.BULLET)

        elif kind == 'ol':
            add_list_item(doc, payload[1], payload[0] + '.')

        elif kind == 'code':
            add_code(doc, payload[1])

        elif kind == 'cap':
            # 没被表格认领的题注（多半是正文里提到「见表 3」）→ 按原文渲染，
            # 不能丢手写编号，也不能当表格题注编号
            p = body_para(doc)
            emit_runs(p, payload.get('raw') or payload['text'], body_pt,
                      spec.FONT_BODY_CJK, spec.FONT_BODY_LAT)

        elif kind == 'note':
            add_note(doc, payload)

        elif kind == 'img':
            stats['images'] += 1
            pic = B.add_picture_para(doc, _resolve_img(payload['src'], base_dir))
            cap = payload.get('cap')
            if cap:
                stats['captioned'] += 1
                # 图题在图下方：图片 keepNext 把两者绑住，图题自己不再 keepNext
                B.set_keep_next(pic)
                add_caption(doc, _next_caption_num(cap, counters, caption_mode),
                            keep_next=False, before=spec.CAP_FIG_SPACE_BEFORE,
                            after=spec.CAP_FIG_SPACE_AFTER)
            if payload.get('note'):
                add_note(doc, payload['note'])

        elif kind == 'table':
            rows = payload['rows']
            stats['tables'] += 1
            cap = payload.get('cap')
            if cap:
                stats['captioned'] += 1
                add_caption(doc, _next_caption_num(cap, counters, caption_mode))
            else:
                stats['uncaptioned'].append(stats['tables'])
            add_table(doc, rows)
            if payload.get('note'):
                add_note(doc, payload['note'])

    if toc:
        if toc_anchor is not None:
            B.add_toc_field(doc, toc_anchor)
        B.ensure_update_fields(doc)

    out = out or os.path.splitext(src)[0] + '.docx'
    _ensure_writable(out)
    doc.save(out)
    return out, stats, blocks


def _ensure_writable(path):
    """目标文件被 Word 打开时，open(...,'r+b') 会抛 PermissionError。
    提前给出可读的提示，比最后抛原始 OSError 好。"""
    if not os.path.exists(path):
        return
    try:
        with open(path, 'r+b'):
            pass
    except PermissionError:
        raise SystemExit(
            '目标文件正被占用，无法写入：\n  %s\n'
            '多半是它正在 Word 里开着。请先关闭该文档再重试。' % path)


def parse_body_size(v):
    if v is None:
        return None
    v = str(v).strip()
    if v in spec.CN_SIZE:
        return spec.CN_SIZE[v]
    return float(v)


def main(argv=None):
    ap = argparse.ArgumentParser(description='Markdown → 规范中文报告 Word 文档')
    ap.add_argument('src', help='输入 .md 文件')
    ap.add_argument('-o', '--out', help='输出 .docx（默认同名）')
    ap.add_argument('--toc', action='store_true', help='在标题后插入目录域（默认不插）')
    ap.add_argument('--caption-mode', choices=['auto', 'manual', 'both'], default='both',
                    help='表题编号：auto=一律重编 manual=沿用你写的号 both=有号用号、无号顺延')
    ap.add_argument('--body-size', help='正文字号，如 五号 / 12 / 10.5（默认小四 12pt）')
    ap.add_argument('--no-page-number', action='store_true', help='不插页脚页码')
    ap.add_argument('--title', help='覆盖文档标题')
    a = ap.parse_args(argv)

    if not os.path.isfile(a.src):
        print('找不到文件：%s' % a.src, file=sys.stderr)
        return 1

    out, stats, blocks = build(
        a.src, a.out, toc=a.toc, caption_mode=a.caption_mode,
        body_pt=parse_body_size(a.body_size),
        page_number=not a.no_page_number, title=a.title)

    print('已生成 %s' % out)
    print('  段落块 %d | 标题 %d | 表格 %d | 插图 %d（有题注 %d）'
          % (len(blocks), stats['headings'], stats['tables'], stats['images'],
             stats['captioned']))
    if stats['uncaptioned']:
        print('  ⚠ 无题注的表格（第 %s 张）—— 在表格上一行写「表：说明文字」即可自动编号'
              % ', '.join(map(str, stats['uncaptioned'][:12])))
    return 0


if __name__ == '__main__':
    sys.exit(main())