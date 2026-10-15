# -*- coding: utf-8 -*-
"""Markdown → 块序列。

方言基于用户现有 parse_blocks()，补齐了它缺的四样：
  · 代码围栏 ``` —— 原解析器不认，代码会被当普通段落（实测语料里 20 处）
  · 链接 [文本](url) —— 原解析器不认，会原样输出方括号（实测 23 处）
  · h4 以上不再截断成 h3（实测 15 处）
  · 表题/表注约定（原语料里一条题注都没有，全靠脚本手写编号）

块类型：
  ('title', text)         第一个 h1，作文档标题（居中、不参与编号）
  ('h', (level, text))    1..4
  ('p', text)
  ('quote', text)
  ('hr', None)            分隔线
  ('pb', None)            \\pagebreak 另起一页
  ('todo', text)          - [ ] / - [x]
  ('ul', text)
  ('ol', (num, text))
  ('code', (lang, text))
  ('cap', text)           题注行：表：说明  或  表 7 说明
  ('note', text)          表注行：表注：...
  ('table', rows)         解析阶段；attach_captions() 之后变成
  ('table', {'rows':…, 'cap':…, 'note':…, 'kind':'表'|'图'})
  ('img', {'alt':…, 'src':…, 'cap':…, 'note':…})   独占一行的 ![](路径)

图片的题注约定与表格相反：中文规范里**图题在图下方**，所以
`图：说明` 写在图片下一行（写成上一行也认）。图片不产出文本，
「内容无损」校验两侧都跳过它，改由 verify.py 单独断言张数与文件名。
"""
import re

# 行内标记。**粗体** 只用来在源稿（GitHub/Typora）里保持可读，
# 生成 Word 时**剥掉标记但不加粗** —— 中文报告正文不做加粗强调。
# 顺序重要：** 要在 * 之前，链接要在纯文本之前。
INLINE = re.compile(r'\*\*(.+?)\*\*|`([^`]+)`|\[([^\]]*)\]\(([^)]+)\)')

H_RE = re.compile(r'^(#{1,6})\s+(.*)$')
HR_RE = re.compile(r'^\s*(-{3,}|\*{3,}|_{3,})\s*$')
TODO_RE = re.compile(r'^[-*]\s+\[([ xX])\]\s+(.*)$')
UL_RE = re.compile(r'^[-*]\s+(.*)$')
OL_RE = re.compile(r'^(\d+)[.、]\s+(.*)$')
FENCE_RE = re.compile(r'^\s*(```+|~~~+)\s*(\S*)\s*$')
IMG_RE = re.compile(r'^!\[([^\]]*)\]\(\s*(<[^>]+>|[^)\s]+)(?:\s+"[^"]*")?\s*\)\s*$')
SEP_CELL_RE = re.compile(r'^:?-{2,}:?$')

# 题注：表：说明 / 表:说明 / 表 7 说明 / 图 3 说明 / 代码 1 说明
CAP_RE = re.compile(r'^\s*(表|图|代码|算法)\s*(?:[:：]\s*(.*)|(\d+)\s*[:：、.．]?\s*(.*))$')
NOTE_RE = re.compile(r'^\s*(表注|图注|注)\s*[:：]\s*(.*)$')
PAGEBREAK_RE = re.compile(r'^\s*\\pagebreak\s*$')

MAX_LEVEL = 4


def split_row(line):
    line = line.strip()
    if line.startswith('|'):
        line = line[1:]
    if line.endswith('|'):
        line = line[:-1]
    return [c.strip() for c in line.split('|')]


def _is_sep_row(cells):
    return bool(cells) and all(SEP_CELL_RE.match(c) for c in cells if c)


def parse_blocks(text_or_path, is_path=None):
    """解析成块序列。传路径或字符串都行。"""
    if is_path is None:
        is_path = '\n' not in text_or_path and len(text_or_path) < 400
    if is_path:
        with open(text_or_path, encoding='utf-8') as f:
            text = f.read()
    else:
        text = text_or_path

    lines = text.replace('\r\n', '\n').replace('\r', '\n').split('\n')
    blocks, i, n = [], 0, len(lines)
    seen_h1 = False

    while i < n:
        s = lines[i].strip()

        if not s:
            i += 1
            continue

        if PAGEBREAK_RE.match(s):
            blocks.append(('pb', None))
            i += 1
            continue

        m = FENCE_RE.match(s)
        if m:
            fence, lang = m.group(1)[:3], m.group(2)
            i += 1
            buf = []
            while i < n and not lines[i].strip().startswith(fence):
                buf.append(lines[i])
                i += 1
            if i < n:
                i += 1                      # 跳过收尾围栏
            blocks.append(('code', (lang, '\n'.join(buf).rstrip())))
            continue

        if HR_RE.match(s):
            blocks.append(('hr', None))
            i += 1
            continue

        m = H_RE.match(s)
        if m:
            level = min(len(m.group(1)), MAX_LEVEL)
            title = m.group(2).strip()
            if level == 1 and not seen_h1:
                seen_h1 = True
                blocks.append(('title', title))
            else:
                blocks.append(('h', (level, title)))
            i += 1
            continue

        m = IMG_RE.match(s)
        if m:
            src_path = m.group(2)
            if src_path.startswith('<') and src_path.endswith('>'):
                src_path = src_path[1:-1]
            blocks.append(('img', {'alt': m.group(1).strip(), 'src': src_path,
                                   'cap': None, 'note': None}))
            i += 1
            continue

        if s.startswith('|'):
            rowlines = []
            while i < n and lines[i].strip().startswith('|'):
                rowlines.append(lines[i])
                i += 1
            rows = []
            for b in rowlines:
                cells = split_row(b)
                if _is_sep_row(cells):
                    continue
                rows.append(cells)
            if rows:
                blocks.append(('table', {'rows': rows, 'cap': None,
                                         'note': None, 'kind': '表'}))
            continue

        m = NOTE_RE.match(s)
        if m:
            # 原样保留整行（含「表注：」前缀）—— 擅自改写前缀会破坏内容无损
            blocks.append(('note', s))
            i += 1
            continue

        m = CAP_RE.match(s)
        if m:
            kind = m.group(1)
            cap = (m.group(2) if m.group(2) is not None else m.group(4)) or ''
            manual = m.group(3)
            blocks.append(('cap', {'text': cap.strip(), 'kind': kind,
                                   'num': int(manual) if manual else None,
                                   'raw': s}))
            i += 1
            continue

        if s.startswith('>'):
            buf = []
            while i < n and lines[i].strip().startswith('>'):
                buf.append(lines[i].strip().lstrip('>').strip())
                i += 1
            blocks.append(('quote', ' '.join(x for x in buf if x)))
            continue

        m = TODO_RE.match(s)
        if m:
            blocks.append(('todo', m.group(2).strip()))
            i += 1
            continue

        m = OL_RE.match(s)
        if m:
            blocks.append(('ol', (m.group(1), m.group(2).strip())))
            i += 1
            continue

        m = UL_RE.match(s)
        if m:
            blocks.append(('ul', m.group(1).strip()))
            i += 1
            continue

        # 普通段落：合并到下一个空行（软换行接成一段）
        buf = [s]
        i += 1
        while i < n:
            nxt = lines[i].strip()
            if not nxt or _starts_block(nxt):
                break
            buf.append(nxt)
            i += 1
        blocks.append(('p', ' '.join(buf)))

    return attach_captions(blocks)


def _starts_block(s):
    return bool(H_RE.match(s) or HR_RE.match(s) or FENCE_RE.match(s)
                or IMG_RE.match(s)
                or s.startswith('|') or s.startswith('>')
                or TODO_RE.match(s) or OL_RE.match(s) or UL_RE.match(s)
                or CAP_RE.match(s) or NOTE_RE.match(s) or PAGEBREAK_RE.match(s))


def attach_captions(blocks):
    """把独立的 cap / note 块并入相邻的 table 块。

    · cap  紧挨在表格**前一行** → 该表的题注（中文规范：表题在表格上方）
    · cap  紧挨在图片**后一行** → 该图的图题（中文规范：图题在图片下方）；
      写在图片前一行也认，但不推荐 —— 排版时图题始终渲染在图下方
    · note 紧挨在表格/图片**后一行** → 该对象的表注（表注在表格下方）
    没有归属的 cap / note 原样保留 —— 多半是正文里提到「见表 3」的句子，
    这时题注要按原文渲染，不能丢编号也不能当表格题注。
    """
    stage, i = [], 0
    while i < len(blocks):
        k, v = blocks[i]
        if k == 'cap':
            if v.get('kind') == '图' and stage and stage[-1][0] == 'img':
                stage[-1][1]['cap'] = v      # 图题在图片下一行
                i += 1
                continue
            j = i + 1
            while j < len(blocks) and blocks[j][0] == 'cap':
                j += 1                       # 允许题注与表格之间夹注释行
            if j < len(blocks) and blocks[j][0] == 'table':
                blocks[j][1]['cap'] = v
                i += 1
                continue
            if v.get('kind') == '图' and j < len(blocks) and blocks[j][0] == 'img':
                blocks[j][1]['cap'] = v      # 容忍写在图片上一行
                i += 1
                continue
        stage.append(blocks[i])
        i += 1

    out, i = [], 0
    while i < len(stage):
        k, v = stage[i]
        if k in ('table', 'img') and i + 1 < len(stage) and stage[i + 1][0] == 'note':
            v['note'] = stage[i + 1][1]
            out.append(stage[i])
            i += 2
            continue
        out.append(stage[i])
        i += 1
    return out


def plain_len(t):
    """去掉行内标记后的可见字符数，用于估算列宽。"""
    t = re.sub(r'\*\*(.+?)\*\*', r'\1', t)
    t = re.sub(r'`([^`]+)`', r'\1', t)
    t = re.sub(r'\[([^\]]*)\]\(([^)]+)\)', r'\1', t)
    return len(t.strip())


def col_widths(rows, floor=0.10, first_floor=0.13):
    """按各列最长内容估比例。平方根压缩，避免长文本列吃光宽度把首列压成竖排。"""
    ncol = max(len(r) for r in rows)
    m = []
    for j in range(ncol):
        best = 1
        for r in rows:
            if j < len(r):
                best = max(best, plain_len(r[j]))
        m.append(best ** 0.5)
    s = sum(m) or 1
    p = [max(x / s, first_floor if j == 0 else floor) for j, x in enumerate(m)]
    s2 = sum(p) or 1
    return [x / s2 for x in p]


def iter_inline(text):
    """产出 (kind, payload)：('t',文本) ('c',代码) ('a',(文本,链接))

    刻意不再产出 'b'：**…** 只剥标记、按正文排。正文不加粗是本规范的硬约束。
    """
    pos = 0
    for m in INLINE.finditer(text):
        if m.start() > pos:
            yield ('t', text[pos:m.start()])
        if m.group(1) is not None:
            yield ('t', m.group(1))
        elif m.group(2) is not None:
            yield ('c', m.group(2))
        else:
            yield ('a', (m.group(3), m.group(4)))
        pos = m.end()
    if pos < len(text):
        yield ('t', text[pos:])


def strip_inline(text):
    """剥掉行内标记，返回纯可见文本（用于「内容无损」校验）。"""
    out = []
    for kind, payload in iter_inline(text):
        out.append(payload[0] if kind == 'a' else payload)
    return ''.join(out)


def plain_text(blocks):
    """所有块的可见文本，用于「内容无损」校验。

    'img' 不产出文本：图片本身不是文本，题注虽在图下方渲染，但它由
    图的 cap 字段承载（见 verify.md_stream），此处不重复计入。
    """
    out = []
    for b in blocks:
        k, v = b
        if k == 'img':
            continue
        if k == 'table':
            for r in v['rows']:
                out.extend(r)
            if v.get('cap'):
                out.append(v['cap']['text'])
            if v.get('note'):
                out.append(v['note'])
        elif k == 'code':
            out.append(v[1])
        elif k == 'h':
            out.append(v[1])
        elif k == 'ol':
            out.append(v[1])
        elif k in ('title', 'p', 'quote', 'todo', 'ul', 'note'):
            out.append(v)
        elif k == 'cap':
            out.append(v['text'])
    return out