# -*- coding: utf-8 -*-
"""LaTeX 数学子集 → OMML（Word 原生公式）。

源稿里照常写 LaTeX：

    $$ mAP = \\frac{1}{N}\\sum_{i=1}^{N} AP_i $$

Word 里得到的是**真公式对象**：分数上下堆叠、Σ 带上下限，双击可进公式编辑器改。
.docx 里存的是 OMML（Office Math Markup Language）—— Word **不会**回头解析 LaTeX，
所以这层转换只能在生成时做完。全部手写 XML：不引第三方依赖，也不借 Word COM。

支持的子集（够学位论文用）：
  分式 \\frac \\dfrac \\tfrac；上下标 `^` `_`；根式 \\sqrt[n]{}
  大型算符 \\sum \\prod \\int \\oint \\iint \\bigcup \\bigcap …（Σ 的上下限在上下、∫ 的在右侧）
  定界符 \\left( … \\right]、重音 \\hat \\bar \\vec \\tilde \\dot、
  上/下划线 \\overline \\underline、正体 \\text{} \\mathrm{} \\operatorname{}、粗体 \\mathbf{}
  希腊字母与常用运算符、函数名 \\sin \\log \\max…、间距 \\, \\; \\! \\quad、转义 \\% \\& \\_ \\#

**面对不认识的东西（\\begin{matrix}、自定义宏、没基底的 `_`）整条公式回退成纯文本**并给警告 ——
宁可原样显示 LaTeX，也不要半懂不懂地排出个错公式。

关于内容无损：`canon_text()` 给出「这条公式在 Word 里应显示成哪些字符」，
与从 .docx 里抠出来的 `m:t` 拼串逐字符比对 —— 转换环节丢字，校验会当场抓住。
"""
import re
from xml.sax.saxutils import escape as _esc

M_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
_NSDECL = 'xmlns:m="%s" xmlns:w="%s"' % (M_NS, W_NS)
MATH_FONT = 'Cambria Math'      # 数学字形只有它有（∑ ∫ 之类），必须显式指定，
CJK_FONT = '宋体'               # 否则会顺着公式样式的宋体/TNR 渲染成方框


class Unsupported(Exception):
    """超出子集 —— 调用方据此整条回退成纯文本。"""


# ---------------------------------------------------------------- 符号表
SYMBOLS = {
    # 希腊字母
    'alpha': 'α', 'beta': 'β', 'gamma': 'γ', 'delta': 'δ', 'epsilon': 'ϵ',
    'varepsilon': 'ε', 'zeta': 'ζ', 'eta': 'η', 'theta': 'θ', 'vartheta': 'ϑ',
    'iota': 'ι', 'kappa': 'κ', 'lambda': 'λ', 'mu': 'μ', 'nu': 'ν', 'xi': 'ξ',
    'pi': 'π', 'varpi': 'ϖ', 'rho': 'ρ', 'varrho': 'ϱ', 'sigma': 'σ',
    'varsigma': 'ς', 'tau': 'τ', 'upsilon': 'υ', 'phi': 'ϕ', 'varphi': 'φ',
    'chi': 'χ', 'psi': 'ψ', 'omega': 'ω',
    'Gamma': 'Γ', 'Delta': 'Δ', 'Theta': 'Θ', 'Lambda': 'Λ', 'Xi': 'Ξ',
    'Pi': 'Π', 'Sigma': 'Σ', 'Upsilon': 'Υ', 'Phi': 'Φ', 'Psi': 'Ψ', 'Omega': 'Ω',
    # 运算符 / 关系
    'times': '×', 'div': '÷', 'pm': '±', 'mp': '∓', 'cdot': '⋅', 'ast': '∗',
    'star': '⋆', 'circ': '∘', 'bullet': '∙', 'oplus': '⊕', 'ominus': '⊖',
    'otimes': '⊗', 'oslash': '⊘', 'odot': '⊙',
    'le': '≤', 'leq': '≤', 'ge': '≥', 'geq': '≥', 'ne': '≠', 'neq': '≠',
    'approx': '≈', 'equiv': '≡', 'sim': '∼', 'simeq': '≃', 'cong': '≅',
    'propto': '∝', 'll': '≪', 'gg': '≫', 'asymp': '≍',
    'infty': '∞', 'partial': '∂', 'nabla': '∇', 'forall': '∀', 'exists': '∃',
    'nexists': '∄', 'in': '∈', 'notin': '∉', 'ni': '∋', 'subset': '⊂',
    'subseteq': '⊆', 'supset': '⊃', 'supseteq': '⊇', 'cup': '∪', 'cap': '∩',
    'setminus': '∖', 'emptyset': '∅', 'varnothing': '∅', 'complement': '∁',
    'to': '→', 'rightarrow': '→', 'leftarrow': '←', 'leftrightarrow': '↔',
    'Rightarrow': '⇒', 'Leftarrow': '⇐', 'Leftrightarrow': '⇔', 'mapsto': '↦',
    'uparrow': '↑', 'downarrow': '↓', 'longrightarrow': '⟶',
    'angle': '∠', 'perp': '⊥', 'parallel': '∥', 'triangle': '△', 'square': '□',
    'therefore': '∴', 'because': '∵', 'ldots': '…', 'dots': '…', 'cdots': '⋯',
    'vdots': '⋮', 'ddots': '⋱', 'prime': '′', 'degree': '°', 'neg': '¬',
    'land': '∧', 'wedge': '∧', 'lor': '∨', 'vee': '∨', 'bigcirc': '○',
    'mid': '∣', 'colon': ':', 'langle': '⟨', 'rangle': '⟩', 'vert': '|',
    'Vert': '‖', 'lVert': '‖', 'rVert': '‖', 'lvert': '|', 'rvert': '|',
    'iff': '⇔', 'implies': '⟹', 'hbar': 'ℏ', 'ell': 'ℓ', 'Re': 'ℜ',
    'Im': 'ℑ', 'aleph': 'ℵ', 'wp': '℘', 'dagger': '†', 'S': '§',
    'clubsuit': '♣', 'diamondsuit': '♢', 'heartsuit': '♡', 'spadesuit': '♠',
    # 间距与转义
    'quad': ' ', 'qquad': '  ', 'enspace': ' ',
    ',': ' ', ';': ' ', ':': ' ', '!': '', ' ': ' ', '~': ' ',
    '%': '%', '&': '&', '#': '#', '$': '$', '{': '{', '}': '}', '_': '_',
}

# 大型算符：字符 + 上下限位置（∑ 在上下，∫ 在右侧）
NARY = {
    'sum': ('∑', 'undOvr'), 'prod': ('∏', 'undOvr'), 'coprod': ('∐', 'undOvr'),
    'bigcup': ('⋃', 'undOvr'), 'bigcap': ('⋂', 'undOvr'),
    'bigoplus': ('⨁', 'undOvr'), 'bigotimes': ('⨂', 'undOvr'),
    'bigwedge': ('⋀', 'undOvr'), 'bigvee': ('⋁', 'undOvr'),
    'int': ('∫', 'subSup'), 'iint': ('∬', 'subSup'), 'iiint': ('∭', 'subSup'),
    'oint': ('∮', 'subSup'), 'oiint': ('∯', 'subSup'),
}

# 函数名（正体）
FUNCS = ('sin', 'cos', 'tan', 'cot', 'sec', 'csc', 'arcsin', 'arccos', 'arctan',
         'sinh', 'cosh', 'tanh', 'coth', 'log', 'ln', 'lg', 'exp', 'lim', 'max',
         'min', 'arg', 'deg', 'det', 'dim', 'gcd', 'lcm', 'mod', 'ker', 'hom',
         'sup', 'inf', 'tr', 'rank', 'diag')

# 重音：\\hat{x} → 头上一个 ^。用 Word 认的那几个字符
ACCENTS = {'hat': '^', 'widehat': '^', 'check': 'ˇ', 'breve': '˘',
           'tilde': '~', 'widetilde': '~', 'acute': '´', 'grave': '`',
           'dot': '˙', 'ddot': '¨', 'vec': '⃗', 'bar': '‾',
           'overrightarrow': '⃗', 'overleftarrow': '⃖'}

STYLE_CMDS = {'mathbf': 'b', 'boldsymbol': 'b', 'bm': 'b', 'mathbb': None,
              'mathcal': None, 'mathit': None, 'mathsf': None, 'mathtt': None}
UPRIGHT_CMDS = ('text', 'textrm', 'mathrm', 'operatorname')

# 大型算符「吃」操作数时的排除项：空格、二元运算符、关系符不当操作数
_OPER_CHARS = set('+-*/=<>±×÷≤≥≠≈→←↔∈∉∪∩∅∞∝⋅∘≪≫'
                  '，。；：、,.;:!?()[]{}')


# ---------------------------------------------------------------- 解析
class _P:
    """字符级递归下降。每个「原子」是单字符或一个结构节点。"""

    def __init__(self, s):
        self.s = s
        self.i = 0

    # --- 基础 ---
    def eof(self):
        return self.i >= len(self.s)

    def peek(self):
        return self.s[self.i] if self.i < len(self.s) else ''

    def expect(self, c):
        if self.peek() != c:
            raise Unsupported('期待 %r，实际 %r' % (c, self.peek()))
        self.i += 1

    def parse_all(self):
        atoms = self.parse_seq()
        if not self.eof():
            raise Unsupported('解析在 %d 处停下' % self.i)
        return atoms

    def parse_seq(self, stop=None):
        out = []
        while not self.eof():
            if stop and self.peek() == stop:
                break
            if self.peek() == '}':
                raise Unsupported('多余的 }')
            atom = self.parse_scripts(self.parse_atom())
            # 大型算符吃下紧跟着的那个「像操作数」的原子，做成 nary 的 e
            if (out and out[-1][0] == 'nary' and out[-1][5] is None
                    and _is_operand(atom)):
                out[-1] = out[-1][:5] + ([atom],)
                continue
            out.append(atom)
        return out

    def parse_atom(self):
        c = self.peek()
        if c == '{':
            self.i += 1
            body = self.parse_seq('}')
            self.expect('}')
            return ('grp', body)
        if c == '\\':
            return self.parse_command()
        if c in '_^':
            raise Unsupported('裸的上下标（前面没有基底）：%s' % c)
        self.i += 1
        return ('ch', c)

    def parse_scripts(self, base):
        sub = sup = None
        while True:
            c = self.peek()
            if c == '_' and sub is None:
                self.i += 1
                sub = self.parse_operand()
            elif c == '^' and sup is None:
                self.i += 1
                sup = self.parse_operand()
            else:
                break
        if sub is None and sup is None:
            return base
        if base[0] == 'narysym':
            return ('nary', base[1], base[2], sub, sup, None)
        if sub is not None and sup is not None:
            return ('subsup', base, sub, sup)
        if sub is not None:
            return ('sub', base, sub)
        return ('sup', base, sup)

    def parse_operand(self):
        """`_` / `^` 后面跟的一个原子（单个字符或 {...}），**包成 list**。

        统一成「原子表」是为了让发射/收集两头不必分情况 —— 这里曾经一半是
        原子、一半是列表，结果 `x_i^2` 把下标整个吞掉，还静默不报错。
        """
        if self.eof():
            raise Unsupported('上下标后面没有内容')
        return [self.parse_atom()]

    def parse_group(self):
        self.expect('{')
        body = self.parse_seq('}')
        self.expect('}')
        return body

    def raw_group(self, until='}'):
        """\\text{...} 的内容：原样取到收尾括号，不按数学解析。"""
        self.expect('{')
        j = self.i
        depth = 1
        while j < len(self.s):
            if self.s[j] == '\\':
                j += 2
                continue
            if self.s[j] == '{':
                depth += 1
            elif self.s[j] == '}':
                depth -= 1
                if depth == 0:
                    break
            j += 1
        if j >= len(self.s):
            raise Unsupported('缺少收尾的 }')
        text = self.s[self.i:j]
        self.i = j + 1
        return text

    # --- 命令 ---
    def parse_command(self):
        self.expect('\\')
        if self.eof():
            raise Unsupported('孤立的反斜杠')
        if self.peek().isalpha():
            j = self.i
            while j < len(self.s) and self.s[j].isalpha():
                j += 1
            name = self.s[self.i:j]
            self.i = j
        else:
            name = self.peek()
            self.i += 1
        return self.build_command(name)

    def build_command(self, name):
        if name in NARY:
            return ('narysym', NARY[name][0], NARY[name][1])
        if name in ('frac', 'dfrac', 'tfrac', 'cfrac'):
            return ('frac', self.parse_group(), self.parse_group())
        if name == 'sqrt':
            deg = self.parse_bracket() if self.peek() == '[' else None
            return ('sqrt', deg, self.parse_group())
        if name in ('overline', 'underline'):
            pos = 'top' if name == 'overline' else 'bot'
            return ('bar', pos, self.parse_group())
        if name in ACCENTS:
            # 包成 list：与 bar 统一成「体是原子表」，发射/收集两头就不用分情况
            return ('acc', ACCENTS[name], [self.parse_atom()])
        if name in UPRIGHT_CMDS:
            return ('txt', self.raw_group() if name != 'operatorname'
                    else self.raw_group())
        if name in STYLE_CMDS:
            sty = STYLE_CMDS[name]
            body = self.parse_group()
            return ('sty', sty, body) if sty else ('grp', body)
        if name in FUNCS:
            return ('func', name)
        if name in ('limits', 'nolimits', 'displaystyle', 'textstyle'):
            return ('sp', '')          # 无输出，只是排版提示
        if name == 'left':
            beg = self.parse_delim()
            body = self.parse_until_right()
            end = self.parse_delim()
            return ('delim', beg, end, body)
        if name in ('big', 'Big', 'bigg', 'Bigg', 'middle'):
            # 只放大定界符：忽略掉反而安全（定界符照常排，只是不放大），
            # 比整条公式回退成文本强得多
            return ('sp', '')
        if name == 'right':
            raise Unsupported('孤立的 \\right')
        if name in SYMBOLS:
            return ('sym', SYMBOLS[name])
        if name in ('begin', 'end'):
            raise Unsupported('不支持环境（\\begin{…}），请改用单行公式或插图')
        raise Unsupported('不认识的命令 \\%s' % name)

    def parse_bracket(self):
        self.expect('[')
        body = self.parse_seq(']')
        self.expect(']')
        return body

    def parse_delim(self):
        if self.eof():
            raise Unsupported('\\left/\\right 后面缺少定界符')
        c = self.peek()
        if c == '\\':
            self.i += 1
            if self.peek().isalpha():
                j = self.i
                while j < len(self.s) and self.s[j].isalpha():
                    j += 1
                name = self.s[self.i:j]
                self.i = j
            else:
                name = self.peek()
                self.i += 1
            return {'{': '{', '}': '}', '|': '‖', '.': '', 'langle': '⟨',
                    'rangle': '⟩', 'lvert': '|', 'rvert': '|', 'vert': '|',
                    'Vert': '‖', 'lVert': '‖', 'rVert': '‖'}.get(name, '')
        self.i += 1
        return '' if c == '.' else c

    def parse_until_right(self):
        out = []
        while not self.eof():
            if self.s.startswith('\\right', self.i):
                self.i += len('\\right')
                return out
            out.append(self.parse_scripts(self.parse_atom()))
        raise Unsupported('\\left 没有配对的 \\right')


def _is_operand(atom):
    """够不够格当大型算符的操作数：数字/字母/结构节点算，运算符与空格不算。"""
    k = atom[0]
    if k == 'ch':
        return not (atom[1].isspace() or atom[1] in _OPER_CHARS)
    return k in ('grp', 'frac', 'sqrt', 'sub', 'sup', 'subsup', 'func', 'txt',
                 'delim', 'bar', 'acc', 'sty', 'sym')


# ---------------------------------------------------------------- 发射 OMML
def _r(text, upright=False, sty=None):
    """一个数学 run。sty='p' 正体（函数名/\\text），'b' 粗体；不给就交给 Word
    的数学斜体规则（字母斜、数字正），与 LaTeX 的默认一致。"""
    m_rpr = ''
    if sty:
        m_rpr = '<m:rPr><m:sty m:val="%s"/></m:rPr>' % sty
    elif upright:
        m_rpr = '<m:rPr><m:sty m:val="p"/></m:rPr>'
    return ('<m:r>%s<w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:cs="%s" '
            'w:eastAsia="%s"/></w:rPr><m:t xml:space="preserve">%s</m:t></m:r>'
            % (m_rpr, MATH_FONT, MATH_FONT, MATH_FONT, CJK_FONT, _esc(text)))


def _emit(atoms, sty=None):
    return ''.join(_emit_atom(a, sty) for a in atoms)


def _emit_atom(a, sty=None):
    k = a[0]
    if k == 'ch':
        return _r(a[1], sty=sty)
    if k == 'sym':
        return _r(a[1], sty=sty)
    if k == 'sp':
        return _r(a[1]) if a[1] else ''
    if k == 'grp':
        return _emit(a[1], sty)
    if k == 'txt':
        return _r(a[1], upright=True)
    if k == 'func':
        return _r(a[1], upright=True)
    if k == 'sty':
        return _emit(a[2], a[1])
    if k == 'frac':
        return ('<m:f><m:num>%s</m:num><m:den>%s</m:den></m:f>'
                % (_emit(a[1], sty), _emit(a[2], sty)))
    if k == 'sqrt':
        pr = '' if a[1] else '<m:radPr><m:degHide m:val="1"/></m:radPr>'
        deg = '<m:deg>%s</m:deg>' % _emit(a[1], sty) if a[1] else '<m:deg/>'
        return '<m:rad>%s%s<m:e>%s</m:e></m:rad>' % (pr, deg, _emit(a[2], sty))
    if k == 'sub':
        return ('<m:sSub><m:e>%s</m:e><m:sub>%s</m:sub></m:sSub>'
                % (_emit_atom(a[1], sty), _emit(a[2])))
    if k == 'sup':
        return ('<m:sSup><m:e>%s</m:e><m:sup>%s</m:sup></m:sSup>'
                % (_emit_atom(a[1], sty), _emit(a[2])))
    if k == 'subsup':
        return ('<m:sSubSup><m:e>%s</m:e><m:sub>%s</m:sub><m:sup>%s</m:sup>'
                '</m:sSubSup>'
                % (_emit_atom(a[1], sty), _emit(a[2]), _emit(a[3])))
    if k == 'nary':
        _, ch, loc, sub, sup, body = a
        pr = ('<m:naryPr><m:chr m:val="%s"/><m:limLoc m:val="%s"/>'
              '<m:subHide m:val="0"/><m:supHide m:val="0"/></m:naryPr>'
              % (_esc(ch), loc))
        return ('<m:nary>%s<m:sub>%s</m:sub><m:sup>%s</m:sup><m:e>%s</m:e>'
                '</m:nary>'
                % (pr, _emit(sub or []), _emit(sup or []), _emit(body or [], sty)))
    if k == 'delim':
        pr = ('<m:dPr><m:begChr m:val="%s"/><m:endChr m:val="%s"/></m:dPr>'
              % (_esc(a[1]), _esc(a[2])))
        return '<m:d>%s<m:e>%s</m:e></m:d>' % (pr, _emit(a[3], sty))
    if k == 'acc':
        pr = '<m:accPr><m:chr m:val="%s"/></m:accPr>' % _esc(a[1])
        return '<m:acc>%s<m:e>%s</m:e></m:acc>' % (pr, _emit(a[2], sty))
    if k == 'bar':
        pr = '<m:barPr><m:pos m:val="%s"/></m:barPr>' % a[1]
        return '<m:bar>%s<m:e>%s</m:e></m:bar>' % (pr, _emit(a[2], sty))
    raise Unsupported('内部节点 %s' % k)


# ---------------------------------------------------------------- 对外
def to_omml(latex):
    """LaTeX → `<m:oMath>` XML 字符串；认不出来返回 None（调用方回退成纯文本）。"""
    try:
        atoms = _P(latex.strip()).parse_all()
        inner = _emit(atoms)
    except Unsupported:
        return None
    except RecursionError:
        return None
    if not inner.strip():
        return None
    return '<m:oMath %s>%s</m:oMath>' % (_NSDECL, inner)


def canon_text(latex):
    """这条公式在 Word 里应当显示成哪些字符（内容无损校验用）。

    与从 .docx 里抠出来的 `m:t` 拼串对齐：定界符、大型算符、重音这些存在
    OMML **属性**里、不出现在 m:t 中的东西，两边都不计入。
    """
    try:
        out = []
        _collect(_P(latex.strip()).parse_all(), out)
        return ''.join(out)
    except (Unsupported, RecursionError):
        return re.sub(r'[\\{}^_$]', '', latex)


def _collect(atoms, out):
    for a in atoms:
        k = a[0]
        if k in ('ch', 'sym', 'sp'):
            out.append(a[1])
        elif k in ('txt', 'func'):
            out.append(a[-1])
        elif k == 'grp':
            _collect(a[1], out)
        elif k == 'sty':
            _collect(a[2], out)
        elif k == 'frac':
            _collect(a[1], out)
            _collect(a[2], out)
        elif k == 'sqrt':
            if a[1]:
                _collect(a[1], out)
            _collect(a[2], out)
        elif k == 'sub':
            _collect([a[1]], out)
            _collect(a[2], out)
        elif k == 'sup':
            _collect([a[1]], out)
            _collect(a[2], out)
        elif k == 'subsup':
            _collect([a[1]], out)
            _collect(a[2], out)
            _collect(a[3], out)
        elif k == 'nary':
            _collect(a[3] or [], out)          # 上下限
            _collect(a[4] or [], out)
            _collect(a[5] or [], out)          # 操作数
        elif k == 'delim':
            _collect(a[3], out)                # 定界符存在属性里，不计
        elif k in ('acc', 'bar'):
            _collect(a[2], out)                # 重音字符存在属性里，不计