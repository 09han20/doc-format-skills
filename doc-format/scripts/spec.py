# -*- coding: utf-8 -*-
"""中文科技报告排版规范 —— 全部可调参数集中在这里。

改体例只改这个文件，其余模块只读不写。单位约定：
  · 字号一律用 point（pt）
  · 页面/边距用厘米（cm）
  · OOXML 的 w:sz 是「half-point」= pt × 2（半角），w:szCs 同
  · OOXML 边框 w:sz 是「1/8 point」= pt × 8
  · 缩进 w:firstLine 用 twips = pt × 20
"""
from docx.shared import Pt, Cm

# ------------------------------------------------------------------ 字体
# 关键：中文字形由 w:rFonts/@w:eastAsia 决定，西文由 @w:ascii/@w:hAnsi 决定。
# 四个槽位都要写，否则 Word 会用主题字体兜底，中文可能被渲染成西文字体。
FONT_BODY_CJK = '宋体'
FONT_BODY_LAT = 'Times New Roman'
FONT_HEAD_CJK = '黑体'
FONT_HEAD_LAT = 'Times New Roman'
FONT_MONO = 'Consolas'          # 代码，中西文同一字体
FONT_CAP_CJK = '黑体'           # 表题/表注用黑体，与正文区分
FONT_CAP_LAT = 'Times New Roman'

# 样式名。标题类一律自建，不复用内置样式 —— 内置 Title 自带蓝色 accent1 下边框，
# 内置 Caption 是蓝色 accent1 + semiHidden，两者都会毁掉中文报告版面。
STYLE_TITLE = '文档标题'        # 封面标题（替代内置 Title）
STYLE_CAPTION = '表题'          # 表题（替代内置 Caption）

# ------------------------------------------------------------------ 字号
BODY_PT = 12.0                  # 正文 小四
HEAD_PT = {1: 15.0, 2: 14.0, 3: 12.0, 4: 12.0}   # 标题 黑体：小三 / 四号 / 小四
TITLE_PT = HEAD_PT[1] + 3.0     # 封面标题 小二 18pt
CAP_PT = 9.0                    # 表题 小五
NOTE_PT = 9.0                   # 表注 小五
CODE_PT = 9.0                   # 代码块
TABLE_PT = 10.5                 # 表格内文字 五号
FOOT_PT = 10.5                  # 页码

# 中文字号对照（供人工查阅 / 命令行 --body-size 换算）
CN_SIZE = {
    '初号': 42.0, '小初': 36.0, '一号': 26.0, '小一': 24.0,
    '二号': 22.0, '小二': 18.0, '三号': 16.0, '小三': 15.0,
    '四号': 14.0, '小四': 12.0, '五号': 10.5, '小五': 9.0,
    '六号': 7.5, '小六': 6.5, '七号': 5.5, '八号': 5.0,
}

# ------------------------------------------------------------------ 段落
LINE_SPACING = 1.5              # 倍行距 → w:line=360 lineRule=auto
FIRST_LINE_CHARS = 200          # 首行缩进 2 字符 → w:ind/@w:firstLineChars="200"
                                # 双写降级值：Word UI 设「2 字符」时也会同时写 w:firstLine

def first_line_twips(body_pt=None):
    """2 字符对应的 twips 降级值。12pt × 2 = 24pt = 480 twips。"""
    return int(round((body_pt or BODY_PT) * 2 * 20))

HEAD_SPACE_BEFORE = {1: 24.0, 2: 18.0, 3: 12.0, 4: 12.0}   # 段前 pt
HEAD_SPACE_AFTER = {1: 12.0, 2: 8.0, 3: 6.0, 4: 6.0}       # 段后 pt
BODY_SPACE_BEFORE = 0.0
BODY_SPACE_AFTER = 0.0
TITLE_SPACE_BEFORE = 6.0        # 封面标题段前
TITLE_SPACE_AFTER = 18.0        # 封面标题段后（标题与署名行之间留白）
CAP_SPACE_BEFORE = 6.0          # 表题在表格上方
CAP_SPACE_AFTER = 3.0
NOTE_SPACE_BEFORE = 3.0
NOTE_SPACE_AFTER = 6.0

# ------------------------------------------------------------------ 插图
PIC_SPACE_BEFORE = 6.0          # 图片段前
PIC_SPACE_AFTER = 3.0           # 图片段后（与下方图题之间）
CAP_FIG_SPACE_BEFORE = 0.0      # 图题段前（图题紧贴图片，靠 keepNext 绑定）
CAP_FIG_SPACE_AFTER = 9.0       # 图题段后（与下一段正文拉开）

# ------------------------------------------------------------------ 页面
PAGE_W_CM = 21.0                # A4
PAGE_H_CM = 29.7
MARGIN_TOP_CM = 3.0
MARGIN_BOTTOM_CM = 2.5
MARGIN_LEFT_CM = 3.0
MARGIN_RIGHT_CM = 2.5
CONTENT_W_CM = PAGE_W_CM - MARGIN_LEFT_CM - MARGIN_RIGHT_CM   # 15.5cm

# ------------------------------------------------------------------ 三线表
# OOXML 边框 w:sz 单位是 1/8 pt：12 → 1.5pt（顶/底线），6 → 0.75pt（表头下细线）
RULE_THICK = 12
RULE_THIN = 6
RULE_COLOR = '000000'
TABLE_ROW_HEIGHT_PT = None      # None = 自动行高
TABLE_CELL_LINE_SPACING = 1.0   # 表格内单倍行距
TABLE_CELL_SPACE_BEFORE = 1.5
TABLE_CELL_SPACE_AFTER = 1.5
TABLE_COL_FLOOR = 0.10          # 列宽下限（防首列被压成竖排）
TABLE_COL_FIRST_FLOOR = 0.13    # 首列另给更高下限

# 表格对齐。可选 'center' / 'left' / 'right'
TABLE_HEADER_JC = 'center'      # 表头行水平对齐
TABLE_CELL_JC = 'center'        # 表体水平对齐
TABLE_CELL_VALIGN = 'center'    # 单元格垂直对齐（'top' / 'center'）

# 长文本列自动左对齐：整句话居中会让右边缘参差。默认关。
# 开启后，只要某列存在超过阈值的单元格，该列全部改左对齐。
TABLE_LONG_TEXT_LEFT = False
TABLE_LONG_TEXT_THRESHOLD = 24  # 字符数

# ------------------------------------------------------------------ 其它
CODE_SHADING = 'F2F2F2'         # 代码块底纹
QUOTE_BAR_COLOR = '808080'      # 引用块左侧竖线
QUOTE_BAR_SZ = 12
HR_COLOR = '999999'             # 分隔线
TODO_BOX = '□'                  # 待办方框（宋体自带，打印安全）
BULLET = '·'

# ------------------------------------------------------------------ 派生
def hp(pt):
    """point → half-point（w:sz / w:szCs 用）。"""
    return int(round(pt * 2))

def eighth(pt):
    """point → 1/8 point（边框 w:sz 用）。"""
    return int(round(pt * 8))

def cm(v):
    return Cm(v)

def pt(v):
    return Pt(v)