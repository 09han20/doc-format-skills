# -*- coding: utf-8 -*-
"""论文方言：在 doc-format 的 parse_blocks 之上做一层后处理。

共享解析器（`_engine.P`）保持不动 —— 报告体例的输出因此逐字节不变。
这里只把 md 规范要求、而报告方言没有的东西认出来：

  · front matter（首个非空行是 --- 才算）→ 封面元数据；有它时首个 # 不再作文档标题
  · 「关键词：…」/「Keywords: …」行            → 关键词块（标签加粗、不缩进）
  · `# 参考文献` 章节里的 `[N] …` / `1. …` / `- …` → 参考文献条目（悬挂缩进）
  · `$$ … $$`（可写 `$$!` 表示不编号）          → 公式块（居中 + 右对齐编号）
  · `表 4-1 说明` 这类带连字符的手写题注        → 题注块（共享正则只认 `表 4`）
  · 标题文本里的章号：`第1章` / `1` / `一、`    → 供 图3-2 / 表4-1 / (2-1) 编号用

并据此算出分区（封面 / 前置 / 正文）与「日常版要省略哪些块」。
"""
import re

from _engine import P
import latex_omml as L
import spec_paper as SP

# 行内标记先剥掉再判断（`**关键词**：…` 也得认）
STRIP = P.strip_inline

# ---------------------------------------------------------------- 正则
KW_RE = re.compile(
    # 冒号后的空白并进 head：`Keywords: catenary` 渲染出来要留那个空格
    r'^\s*\*{0,2}\s*(关键词|关键字|Keywords?|KEY\s*WORDS?)\s*\*{0,2}\s*([:：]\s*)\s*(.*)$',
    re.I)
EQ_RE = re.compile(r'^\s*\$\$\s*(!?)\s*(.*?)\s*\$\$\s*$', re.S)
REF_HEAD_RE = re.compile(r'^(参考文献|引用文献|References?)$', re.I)
REF_ITEM_RE = re.compile(r'^\s*\[(\d{1,3})\]\s*(.*)$')
CAP_DASH_RE = re.compile(r'^\s*(表|图|公式|式|算法|代码)\s*(\d{1,2}-\d{1,3})\s*[:：、.．]?\s*(.+)$')
# 共享 CAP_RE 会把 `表 4-1 说明` 拆成 num=4 + text='-1 说明'（它只认 `表 4`）。
# 这里把被拆走的后半截接回来，还原成 4-1。
CAP_TAIL_RE = re.compile(r'^\s*[-–—－]\s*(\d{1,3})\s*[:：、.．]?\s*(.*)$')

SPECIAL_RES = (
    ('abstract_cn', re.compile(r'^(中文)?\s*摘\s*要$')),
    ('abstract_en', re.compile(r'^(英文\s*摘\s*要|Abstract|ABSTRACT)$', re.I)),
    ('toc', re.compile(r'^(目\s*录|目\s*次|Contents)$', re.I)),
    ('refs', REF_HEAD_RE),
    ('thanks', re.compile(r'^(致\s*谢|谢\s*辞|Acknowledge?ments?)$', re.I)),
    ('appendix', re.compile(r'^(附\s*录|Appendix)')),
)

META_TITLE_KEYS = ('题名', '题目', '标题', 'title')
CN_NUM = {'一': 1, '二': 2, '三': 3, '四': 4, '五': 5,
          '六': 6, '七': 7, '八': 8, '九': 9}

CH_RE = re.compile(r'^\s*第\s*([0-9]{1,2}|[一二三四五六七八九十]{1,3})\s*[章节篇讲]\s*[:：.、]?\s*')
NUM_RE = re.compile(r'^\s*([0-9]{1,2})\s*(?=[\s、.．:：]|$)')
CN_RE = re.compile(r'^\s*([一二三四五六七八九十]{1,3})\s*[、.．:：]\s*')

FRONT_KINDS = ('abstract_cn', 'abstract_en', 'toc')       # 算「前置部分」的标题
SPECIAL_NO_CHAPTER = ('abstract_cn', 'abstract_en', 'toc', 'refs', 'thanks', 'appendix')


# ---------------------------------------------------------------- front matter
def split_front_matter(text):
    """首个非空行是 `---` 才算 front matter（`...` 或 `---` 收尾）。
    没闭合就当没有 —— 宁可把 `---` 当分隔线，也不要吃掉误写的正文。"""
    lines = text.replace('\r\n', '\n').replace('\r', '\n').split('\n')
    i = 0
    while i < len(lines) and not lines[i].strip():
        i += 1
    if i >= len(lines) or lines[i].strip() != '---':
        return [], text
    j = i + 1
    while j < len(lines) and lines[j].strip() not in ('---', '...'):
        j += 1
    if j >= len(lines):
        return [], text
    meta = lines[i + 1:j]
    body = '\n'.join(lines[:i] + lines[j + 1:])
    return meta, body


def parse_meta(meta_lines):
    """→ (题名, 其余行原样)。题名单独拎出来放大居中；其余整行照排（含 `作者：张三` 的标签）。"""
    title, entries = None, []
    for ln in meta_lines:
        s = ln.strip()
        if not s:
            continue
        key, value = _split_kv(s)
        if key and value and key.strip().lower() in META_TITLE_KEYS:
            title = value
        else:
            entries.append(s)
    return title, entries


def _split_kv(s):
    for sep in ('：', ':'):
        if sep in s:
            k, _, v = s.partition(sep)
            return k, v.strip()
    return None, s


# ---------------------------------------------------------------- 章号
def _cn_to_int(s):
    """一 → 1，十 → 10，十一 → 11，二十三 → 23。认不出返回 None。"""
    if not s:
        return None
    if s == '十':
        return 10
    if s.startswith('十'):
        return 10 + CN_NUM.get(s[1:], 0)
    if len(s) == 1:
        return CN_NUM.get(s)
    if len(s) >= 2 and s[1] == '十':
        rest = CN_NUM.get(s[2:], 0) if len(s) > 2 else 0
        return CN_NUM.get(s[0], 0) * 10 + rest
    return None


def chapter_num_of(text):
    """从标题文本里读章号：`第1章 绪论` / `第一章 绪论` / `1 绪论` / `1. 引言` / `一、本周任务`。

    读不出返回 None —— 绝不自动给标题加编号，那是往正文里塞源稿没有的字。
    """
    s = STRIP(text).strip()
    m = CH_RE.match(s)
    if m:
        g = m.group(1)
        return int(g) if g.isdigit() else _cn_to_int(g)
    m = NUM_RE.match(s)
    if m:
        return int(m.group(1))
    m = CN_RE.match(s)
    if m:
        return _cn_to_int(m.group(1))
    return None


def special_kind(text):
    """标题属于哪个特殊部分；不是特殊标题返回 None。"""
    s = STRIP(text).strip()
    for kind, rx in SPECIAL_RES:
        if rx.match(s):
            return kind
    return None


# ---------------------------------------------------------------- 后处理
def retag(blocks, has_meta):
    """就地重标注块类型。返回新的块列表（共享解析器不动）。"""
    if has_meta:
        # 有 front matter 时题名在封面，首个 # 是章节/摘要，不是文档标题
        blocks = [('h', (1, v)) if k == 'title' else (k, v) for k, v in blocks]

    out = []
    ref_level = None          # 进入参考文献章节后的标题层级
    for k, v in blocks:
        if k == 'h':
            kind = special_kind(v[1])
            if kind == 'refs':
                ref_level = v[0]
            elif ref_level is not None and v[0] <= ref_level:
                ref_level = None
            out.append((k, v))
            continue

        # 参考文献章节里的条目
        if ref_level is not None:
            if k == 'ol':
                out.append(('ref', {'num': int(v[0]), 'text': v[1], 'raw': '%s. %s' % v}))
                continue
            if k == 'ul':
                out.append(('ref', {'num': None, 'text': v, 'raw': '- ' + v}))
                continue
            if k == 'p':
                m = REF_ITEM_RE.match(v)
                if m:
                    out.append(('ref', {'num': int(m.group(1)), 'text': m.group(2),
                                        'raw': v}))
                    continue

        if k == 'p':
            m = KW_RE.match(v)
            if m:
                en = bool(re.match(r'(?i)keyword|key\s*words', m.group(1)))
                out.append(('kw', {'head': m.group(1) + m.group(2), 'text': m.group(3),
                                   'en': en, 'raw': v}))
                continue
            m = EQ_RE.match(v)
            if m:
                out.append(('formula', {'text': m.group(2), 'numbered': not m.group(1),
                                        'raw': v}))
                continue
            m = CAP_DASH_RE.match(v)
            if m:
                # 共享 CAP_RE 只认 `表 4`；`表 4-1 说明` 走这里
                out.append(('cap', {'text': m.group(3).strip(), 'kind': m.group(1),
                                    'num': m.group(2), 'raw': v}))
                continue

        out.append((k, v))

    # 新造出来的题注要重新并入相邻表格/图片（函数本身可重复执行，安全）
    out = P.attach_captions(out)

    # 最后一遍：`表 4-1 说明` 会被共享 CAP_RE 拆成 num=4 + text='-1 说明'，
    # 而且多半已经在 parse_blocks 里并进 table 块了 —— 两种形态都要接回来。
    fixed = []
    for k, v in out:
        if k in ('table', 'img') and v.get('cap'):
            v = dict(v, cap=_fix_dash_cap(v['cap']))
        elif k == 'cap':
            v = _fix_dash_cap(v)
        fixed.append((k, v))
    return fixed


def _fix_dash_cap(cap):
    """num=4 + text='-1 说明' → num='4-1' + text='说明'。"""
    if not cap or cap.get('num') is None:
        return cap
    m = CAP_TAIL_RE.match(cap.get('text') or '')
    if not m:
        return cap
    return dict(cap, num='%s-%s' % (cap['num'], m.group(1)), text=m.group(2).strip())


def annotate(blocks, meta):
    """扫描一遍，算出分区与各部分的齐全情况。

    返回的字典是渲染与校验共同的依据（单一事实来源，别在两处各算一遍）。
    """
    meta_title, cover_entries = parse_meta(meta)
    st = {
        'meta_title': meta_title,
        'cover_entries': cover_entries,
        'has_meta': bool(meta),
        'special': {},        # idx → kind（摘要/Abstract/目录/参考文献/致谢/附录）
        'kw_idx': [],
        'ref_idx': [],
        'chapter_of': {},     # idx → 章号
        'main_start': None,   # 第一个非特殊一级标题 = 正文起点
        'front_start': None,  # 第一个前置块（摘要/Abstract/目录/关键词）
        'has_cn_abstract': False, 'has_en_abstract': False,
        'has_cn_kw': False, 'has_en_kw': False,
        'n_chapters': 0, 'n_refs': 0, 'kw_counts': [],
    }
    chapter = 0
    for idx, (k, v) in enumerate(blocks):
        if k == 'h' and v[0] == 1:
            kind = special_kind(v[1])
            if kind:
                st['special'][idx] = kind
                if kind == 'abstract_cn':
                    st['has_cn_abstract'] = True
                elif kind == 'abstract_en':
                    st['has_en_abstract'] = True
            else:
                n = chapter_num_of(v[1])
                chapter = n if n else chapter + 1
                st['chapter_of'][idx] = chapter
                st['n_chapters'] += 1
                if st['main_start'] is None:
                    st['main_start'] = idx
        elif k == 'kw':
            st['kw_idx'].append(idx)
            if v['en']:
                st['has_en_kw'] = True
            else:
                st['has_cn_kw'] = True
            parts = [x for x in re.split(r'[；;]', v['text']) if x.strip()]
            st['kw_counts'].append(len(parts))
        elif k == 'ref':
            st['ref_idx'].append(idx)
            st['n_refs'] += 1

    # 前置区起点：摘要 / Abstract / 目录 / 关键词
    for idx, (k, v) in enumerate(blocks):
        if k == 'h' and st['special'].get(idx) in FRONT_KINDS:
            st['front_start'] = idx
            break
        if k == 'kw' and (st['main_start'] is None or idx < st['main_start']):
            st['front_start'] = idx
            break

    # 参考文献条目补号：没写号的（`- …` 或裸段落）按位置顺延
    seq = 0
    for idx in st['ref_idx']:
        b = blocks[idx][1]
        if b.get('num') is None:
            seq += 1
            b['num'] = seq
        else:
            seq = b['num']

    return st


def omitted_indices(blocks, st, mode):
    """日常版要省略哪些块：摘要、关键词、参考文献三部分，整段剔除。

    目录、致谢、附录不在省略之列 —— 用户原话是「省略摘要、关键词和参考文献三部分」。
    """
    if mode != SP.MODE_DAILY:
        return set()
    out = set(st['kw_idx'])
    for idx, kind in st['special'].items():
        if kind not in ('abstract_cn', 'abstract_en', 'refs'):
            continue
        lvl = blocks[idx][1][0]
        j = idx
        while j < len(blocks):
            k, v = blocks[j]
            if j > idx and k == 'h' and v[0] <= lvl:
                break
            out.add(j)
            j += 1
    return out


def toc_index(st):
    """源稿里显式写了 `# 目录` 的位置；没有则 None（那就由生成器自动插）。"""
    return next((i for i, k in st['special'].items() if k == 'toc'), None)


def auto_toc(st):
    """要不要自动插目录：有章节可列、且源稿没自己写 `# 目录`。

    正式版与日常版同一条规则 —— 用户原话是「日常版省略摘要、关键词和参考文献
    三部分，其余内容与格式要求仍遵循 md 规定」，而 md 规定目录是自动生成的。
    """
    return st['n_chapters'] >= 1 and toc_index(st) is None


def plan_zones(blocks, st, mode):
    """→ {'cover': (a,b)|None, 'front': (a,b)|None, 'main': (a,b), 'omit': set}

    封面区 / 前置区（罗马页码）/ 正文区（阿拉伯页码）。分节符落在两个边界上。
    正式版必须有目录，所以前置区一定存在；日常版按内容实际决定（可能没有前置区，
    那就只有一节，页脚从 1 起 —— 与报告体例一致）。
    """
    n = len(blocks)
    omit = omitted_indices(blocks, st, mode)
    live = [i for i in range(n) if i not in omit]

    def first_live(pred):
        for i in live:
            if pred(*blocks[i], i):
                return i
        return None

    # 分区边界只看**没被省略**的块：日常版把摘要整段剔掉后，前置区可能就空了 ——
    # 那时候还要硬留一个罗马页码节，就会多出一整页空白页。
    main_start = first_live(lambda b, v, i: b == 'h' and v[0] == 1
                            and not st['special'].get(i))
    front_start = first_live(lambda b, v, i: (b == 'h' and st['special'].get(i) in FRONT_KINDS)
                             or b == 'kw')
    has_meta = st['has_meta']
    front_zone = (main_start is not None and front_start is not None
                  and front_start < main_start)
    if auto_toc(st) and not front_zone:
        # 目录要有地方放：没有前置块就单开一个只装目录的前置节
        front_zone = True
        front_start = main_start if main_start is not None else 0

    if front_zone:
        f0 = front_start if has_meta else 0
        front = (f0, main_start if main_start is not None else n)
        cover = (0, front_start) if has_meta else None
        main = (front[1], n)
    else:
        body_start = main_start if main_start is not None else 0
        cover = (0, body_start) if has_meta else None
        front = None
        main = (body_start, n)

    return {'cover': cover, 'front': front, 'main': main,
            'omit': omit, 'front_zone': front_zone}


def _emit_block(k, v, out, formula_mode='omml'):
    """单个块的可见文本（内容无损校验的 md 侧）。"""
    if k == 'table':
        if v.get('cap'):
            out.append('%s %s %s' % (v['cap']['kind'], v['cap']['num'] or '',
                                     v['cap']['text']))
        out += [STRIP(c) for r in v['rows'] for c in r]
        if v.get('note'):
            out.append(v['note'])
    elif k == 'img':
        if v.get('cap'):
            out.append('%s %s %s' % (v['cap']['kind'], v['cap']['num'] or '',
                                     v['cap']['text']))
        if v.get('note'):
            out.append(v['note'])
    elif k == 'code':
        out.extend(v[1].split('\n'))
    elif k in ('h', 'ol'):
        out.append(STRIP(v[1]))
    elif k == 'title':
        out.append(v)
    elif k == 'cap':
        out.append(v.get('raw') or ('%s %s %s' % (v['kind'], v['num'] or '', v['text'])))
    elif k == 'note':
        out.append(v)
    elif k == 'kw':
        out.append(v['head'] + v['text'])
    elif k == 'ref':
        out.append('[%s] %s' % (v['num'], STRIP(v['text'])))
    elif k == 'formula':
        # 排成 Word 公式对象时，文档里显示的是**可见字符**（\frac{1}{N} 变成一上一下的
        # 1 和 N），不是 LaTeX 源文本 —— 这一侧要用同一条规则算出应有的字符才算对得上
        out.append(L.canon_text(v['text'])
                   if (formula_mode == 'omml' and L.to_omml(v['text']))
                   else STRIP(v['text']))
    elif k in ('p', 'quote', 'todo', 'ul'):
        out.append(STRIP(v))


def stream_lines(blocks, meta, mode, st=None, formula_mode='omml'):
    """md 侧的可见文本流（内容无损校验用），顺序必须与渲染顺序一致。

    比 doc-format 的 md_stream 多三样：封面元数据行、自动插入的「目录」二字
    （正式版才有，位置在前置区末尾）、以及日常版被省略的块（两侧都要剔除）。
    """
    meta_title, cover_entries = parse_meta(meta)
    if st is None:
        st = annotate(retag(blocks, bool(meta)), meta)
    omit = omitted_indices(blocks, st, mode)
    zones = plan_zones(blocks, st, mode)

    out = []
    if meta_title:
        out.append(meta_title)
    out.extend(cover_entries)

    def emit(zone):
        if not zone:
            return
        for idx in range(zone[0], zone[1]):
            if idx not in omit:
                _emit_block(*blocks[idx], out, formula_mode)

    # 目录：自动插的话「目录」二字在源稿里没有，得由这里补进文本流（与渲染同序）
    need_toc = auto_toc(st)

    emit(zones['cover'])
    emit(zones['front'])
    if need_toc:
        out.append(SP.TOC_TITLE)
    emit(zones['main'])
    return out