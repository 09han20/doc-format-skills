# -*- coding: utf-8 -*-
"""Markdown → 规范学位论文 Word 文档（论文正式版 / 日常版）。

    python md2paper.py 论文.md --mode paper      # 论文正式版：六部分齐全，严格校验
    python md2paper.py 论文.md --mode daily      # 日常版：省略摘要、关键词、参考文献

体例照学位论文通用排版规范（封面/页码分区/标题层级/图表编号/GB-T 7714）。
两种模式的排版完全一致，
差别只有一处：日常版把摘要、关键词、参考文献整段省略。

生成后必须跑：
    python verify_paper.py 论文.docx --md 论文.md
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# 中文 Windows 控制台是 GBK，⚠/✅ 会让末尾的 print 抛 UnicodeEncodeError
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding='utf-8')
    except Exception:
        pass

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm

from _engine import B, M, P, spec as S
import md_parse_paper as PP
import spec_paper as SP
import thesis_xml as X

CORE_MODE_TAG = 'doc-paper-mode:%s'      # 写进 docx 核心属性，verify_paper 自动读
PART_HINT = {
    SP.PART_COVER: '在文件开头写 front matter（`---` 之间写「题名：…」等）',
    SP.PART_ABSTRACT: '写 `# 摘要` 标题 + 摘要正文段',
    SP.PART_KEYWORDS: '摘要末尾写一行 `关键词：甲；乙；丙`',
    SP.PART_TOC: '（正式版会自动生成，不用你写）',
    SP.PART_BODY: '正文至少要有章节标题，如 `# 1 绪论`',
    SP.PART_REFS: '写 `# 参考文献` 章节 + `[1] …` 条目',
}


# ---------------------------------------------------------------- 渲染
class Renderer:
    """按块顺序渲染。章号在遇到一级标题时更新，题注/公式编号按章顺延。"""

    def __init__(self, doc, base_dir, mode, caption_mode='both', formula='omml'):
        self.doc = doc
        self.base_dir = base_dir
        self.mode = mode
        self.caption_mode = caption_mode
        self.formula = formula          # 'omml' 排成 Word 公式对象 / 'text' 原样文本
        self.chapter = 0
        self.counters = {'表': 0, '图': 0, '公式': 0}
        self.stats = {'headings': 0, 'tables': 0, 'images': 0, 'formulas': 0,
                      'refs': 0, 'kw': 0, 'uncaptioned': [], 'eq_omml': 0,
                      'eq_text': []}

    # --- 编号 ---
    def _caption_num(self, cap):
        """题注编号：`表 4-1 说明`（按章）。手写编号原样沿用，无号顺延。"""
        kind, num = cap['kind'], cap.get('num')
        self.counters.setdefault(kind, 0)
        if num is None or self.caption_mode == 'auto':
            self.counters[kind] += 1
            num = '%d-%d' % (self.chapter or 1, self.counters[kind])
        else:
            num = str(num)
            tail = num.rsplit('-', 1)[-1]
            if tail.isdigit():
                self.counters[kind] = int(tail)
        return '%s %s %s' % (kind, num, cap['text'])

    def _eq_num(self):
        self.counters['公式'] += 1
        return SP.eq_number(self.chapter, self.counters['公式'])

    # --- 分区 ---
    def render_zone(self, blocks, zone, omit=frozenset()):
        """渲染一个分区；返回实际排进去的块数（调用方据此判断目录要不要另起一页）。"""
        if not zone:
            return 0
        todo = [i for i in range(zone[0], zone[1]) if i not in omit]
        for pos, idx in enumerate(todo):
            k, v = blocks[idx]
            if k == 'h' and PP.special_kind(v[1]) == 'toc':
                # 源稿里显式写了 `# 目录`：就地放「目录」二字 + TOC 域。
                # 目录单独占一页 —— 上面有内容就先另起一页，下面还有内容也同样隔开。
                if pos > 0:
                    M.add_pagebreak(self.doc)
                X.add_toc(self.doc, title=v[1])
                if pos < len(todo) - 1:
                    M.add_pagebreak(self.doc)
                continue
            self.render_block(k, v)
        return len(todo)

    def render_block(self, k, v):
        doc = self.doc
        if k == 'title':
            p = M.plain_para(doc, style=SP.STYLE_TITLE, align=WD_ALIGN_PARAGRAPH.CENTER)
            M.emit_runs(p, v, SP.DOC_TITLE_PT, S.FONT_HEAD_CJK, S.FONT_HEAD_LAT,
                        bold=True)

        elif k == 'h':
            self.render_heading(*v)

        elif k == 'p':
            p = M.body_para(doc)
            M.emit_runs(p, v, S.BODY_PT, S.FONT_BODY_CJK, S.FONT_BODY_LAT)

        elif k == 'kw':
            self.stats['kw'] += 1
            X.add_kw(doc, v['head'], v['text'])

        elif k == 'ref':
            self.stats['refs'] += 1
            X.add_ref(doc, v['num'], v['text'])

        elif k == 'formula':
            self.stats['formulas'] += 1
            num = self._eq_num() if v['numbered'] else None
            if X.add_formula(doc, v['text'], num, omml=(self.formula == 'omml')):
                self.stats['eq_omml'] += 1
            else:
                self.stats['eq_text'].append(v['text'])

        elif k == 'quote':
            M.add_quote(doc, v)
        elif k == 'hr':
            M.add_hr(doc)
        elif k == 'pb':
            M.add_pagebreak(doc)
        elif k == 'todo':
            M.add_list_item(doc, v, S.TODO_BOX)
        elif k == 'ul':
            M.add_list_item(doc, v, S.BULLET)
        elif k == 'ol':
            M.add_list_item(doc, v[1], v[0] + '.')
        elif k == 'code':
            M.add_code(doc, v[1])
        elif k == 'note':
            M.add_note(doc, v)

        elif k == 'cap':
            # 没被表格/图片认领的题注（多半是正文里提到「见表 3」）→ 按原文渲染
            p = M.body_para(doc)
            M.emit_runs(p, v.get('raw') or v['text'], S.BODY_PT,
                        S.FONT_BODY_CJK, S.FONT_BODY_LAT)

        elif k == 'img':
            self.stats['images'] += 1
            pic = B.add_picture_para(doc, self._img_path(v['src']))
            if v.get('cap'):
                B.set_keep_next(pic)
                X.add_caption(doc, self._caption_num(v['cap']), keep_next=False,
                              before=S.CAP_FIG_SPACE_BEFORE,
                              after=S.CAP_FIG_SPACE_AFTER)
            if v.get('note'):
                M.add_note(doc, v['note'])

        elif k == 'table':
            self.stats['tables'] += 1
            if v.get('cap'):
                X.add_caption(doc, self._caption_num(v['cap']))
            else:
                self.stats['uncaptioned'].append(self.stats['tables'])
            M.add_table(doc, v['rows'])
            if v.get('note'):
                M.add_note(doc, v['note'])

    def render_heading(self, level, text):
        kind = PP.special_kind(text)
        if level == 1 and kind is None:
            n = PP.chapter_num_of(text) or (self.chapter + 1)
            if n != self.chapter:
                self.chapter = n
                self.counters = {'表': 0, '图': 0, '公式': 0}
        self.stats['headings'] += 1
        p = self.doc.add_paragraph(style='Heading %d' % level)
        M.emit_runs(p, text, SP.HEAD_PT[level], SP.FONT_H[level], SP.FONT_LAT,
                    bold=SP.HEAD_BOLD[level])
        return p

    def _img_path(self, src):
        return src if os.path.isabs(src) else os.path.normpath(
            os.path.join(self.base_dir, src))


# ---------------------------------------------------------------- 校验
def validate(st, mode):
    """→ (problems, warnings)。problems 非空就不落盘 —— 宁可不出文件，也不要出一份缺件的。"""
    problems, warnings = [], []
    have = {
        SP.PART_COVER: bool(st['has_meta'] and st['meta_title']),
        SP.PART_ABSTRACT: st['has_cn_abstract'],
        SP.PART_KEYWORDS: st['has_cn_kw'],
        SP.PART_TOC: True,          # 正式版自动生成，不用源稿写
        SP.PART_BODY: st['n_chapters'] >= 1,
        SP.PART_REFS: st['n_refs'] >= 1,
    }
    if mode == SP.MODE_PAPER:
        problems = [p for p in SP.REQUIRED_PARTS if not have[p]]
    else:
        for p in (SP.PART_COVER, SP.PART_BODY):
            if not have[p]:
                warnings.append('没有%s' % p)

    if st['has_meta'] and not st['meta_title']:
        warnings.append('front matter 里没写「题名：」——封面会缺题目')
    if st['has_cn_abstract'] and not st['has_en_abstract']:
        warnings.append('有中文摘要但没有 Abstract——md 规范要求英文摘要与中文对应')
    if st['has_cn_kw'] and not st['has_en_kw']:
        warnings.append('有关键词但没有 Keywords——中英文关键词应一一对应')
    for n in st['kw_counts']:
        if not (SP.KW_MIN <= n <= SP.KW_MAX):
            warnings.append('关键词 %d 个，md 规范要求 3–5 个' % n)

    if problems:
        lines = ['论文正式版要求六个部分齐全，源稿里缺了：']
        for p in problems:
            lines.append('  · %s —— %s' % (p, PART_HINT[p]))
        lines.append('想要日常版（省略摘要、关键词、参考文献）就加 --mode daily。')
        raise SystemExit('\n'.join(lines))
    return warnings


# ---------------------------------------------------------------- 组装
def _setup_page(sec):
    sec.page_width, sec.page_height = Cm(S.PAGE_W_CM), Cm(S.PAGE_H_CM)
    sec.top_margin, sec.bottom_margin = Cm(S.MARGIN_TOP_CM), Cm(S.MARGIN_BOTTOM_CM)
    sec.left_margin, sec.right_margin = Cm(S.MARGIN_LEFT_CM), Cm(S.MARGIN_RIGHT_CM)


def build(src, out=None, mode=None, caption_mode='both', formula='omml'):
    if mode not in SP.MODES:
        raise SystemExit('必须指定 --mode paper（正式版）或 --mode daily（日常版）')

    with open(src, encoding='utf-8') as f:
        text = f.read()
    base_dir = os.path.dirname(os.path.abspath(src))

    meta_lines, body = PP.split_front_matter(text)
    blocks = PP.retag(P.parse_blocks(body, is_path=False), bool(meta_lines))
    st = PP.annotate(blocks, meta_lines)
    zones = PP.plan_zones(blocks, st, mode)
    warnings = validate(st, mode)

    doc = Document()
    _setup_page(doc.sections[0])
    X.install_styles(doc)
    R = Renderer(doc, base_dir, mode, caption_mode, formula)

    # --- 封面节：只排内容，全程不碰页脚（不碰 = 没有页码）---
    if st['has_meta']:
        X.build_cover(doc, st['meta_title'], st['cover_entries'])
        R.render_zone(blocks, zones['cover'], zones['omit'])

    sec = doc.sections[-1]
    has_front = zones['front'] is not None
    if st['has_meta']:
        sec = X.add_section(doc)

    # --- 前置节（摘要 / Abstract / 关键词 / 目录）：罗马数字页码 ---
    if has_front:
        X.number_zone(sec, SP.PAGE_FMT_FRONT, SP.PAGE_START)
        n_front = R.render_zone(blocks, zones['front'], zones['omit'])
        if PP.auto_toc(st):
            if n_front:
                # 目录单独占一页：前面排过摘要/关键词，就先另起一页
                M.add_pagebreak(doc)
            X.add_toc(doc)
        sec = X.add_section(doc)

    # --- 正文节（含参考文献 / 致谢 / 附录）：阿拉伯数字，从 1 重新起数 ---
    X.number_zone(sec, SP.PAGE_FMT_MAIN, SP.PAGE_START)
    R.render_zone(blocks, zones['main'], zones['omit'])

    if PP.toc_index(st) is not None or has_front:
        B.ensure_update_fields(doc)

    doc.core_properties.title = st['meta_title'] or ''
    doc.core_properties.comments = '%s formula:%s' % (CORE_MODE_TAG % mode, formula)

    out = out or os.path.splitext(src)[0] + '.docx'
    M._ensure_writable(out)
    doc.save(out)
    return out, st, zones, R.stats, warnings


def main(argv=None):
    ap = argparse.ArgumentParser(description='Markdown → 规范学位论文 Word 文档')
    ap.add_argument('src', help='输入 .md 文件')
    ap.add_argument('-o', '--out', help='输出 .docx（默认同名）')
    ap.add_argument('--mode', required=True, choices=list(SP.MODES),
                    help='paper=论文正式版（封面/摘要/关键词/目录/正文/参考文献齐全）；'
                         'daily=日常版（省略摘要、关键词、参考文献）')
    ap.add_argument('--caption-mode', choices=['auto', 'manual', 'both'], default='both',
                    help='题注编号：auto=一律重编 manual=沿用你写的号 both=有号用号、无号顺延')
    ap.add_argument('--formula', choices=['omml', 'text'], default='omml',
                    help='公式：omml=排成 Word 公式对象（默认，双击可编辑）；'
                         'text=保留 LaTeX 源文本')
    a = ap.parse_args(argv)

    if not os.path.isfile(a.src):
        print('找不到文件：%s' % a.src, file=sys.stderr)
        return 1

    out, st, zones, stats, warnings = build(
        a.src, a.out, mode=a.mode, caption_mode=a.caption_mode, formula=a.formula)

    print('已生成 %s' % out)
    print('  模式：%s | 标题 %d | 表格 %d | 插图 %d | 公式 %d（Word 公式对象 %d）| 参考文献 %d'
          % ('论文正式版' if a.mode == SP.MODE_PAPER else '日常版',
             stats['headings'], stats['tables'], stats['images'],
             stats['formulas'], stats['eq_omml'], stats['refs']))
    if stats['eq_text']:
        print('  ⚠ %d 条公式超出 LaTeX 子集，已按原样文本排（要插图就换成图片）：'
              % len(stats['eq_text']))
        for t in stats['eq_text'][:5]:
            print('      %s' % t[:64].replace('\n', ' '))
    if zones['omit']:
        print('  已按日常版省略：摘要 / 关键词 / 参考文献（源稿里有，共 %d 块，未输出）'
              % len(zones['omit']))
    if stats['uncaptioned']:
        print('  ⚠ 无题注的表格（第 %s 张）—— 在表格上一行写「表：说明文字」即可自动编号'
              % ', '.join(map(str, stats['uncaptioned'][:12])))
    for w in warnings:
        print('  ⚠ %s' % w)
    return 0


if __name__ == '__main__':
    sys.exit(main())