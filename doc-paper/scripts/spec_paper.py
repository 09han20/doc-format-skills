# -*- coding: utf-8 -*-
"""学位论文体例参数 —— 全部可调项集中在这个文件。

规范来源：学位论文通用排版规范（封面/页码分区/标题层级/图表编号/GB-T 7714）。
与 doc-format 的《中文科技报告》体例是**两份独立参数**：报告体例的参数在
`doc-format/scripts/spec.py`，本文件不重复定义两边本就相同的值（页边距、
三线表边框、表格字号、正文缩进语义），直接沿用那份 —— 见 `_engine.spec`。

改体例只改这个文件。单位约定同 doc-format：字号一律 point，
页面用厘米，OOXML 的换算是 docx_builder 的事。
"""
# ---------------------------------------------------------------- 模式
MODE_PAPER = 'paper'            # 论文正式版：md 规范六部分齐全，严格校验
MODE_DAILY = 'daily'            # 日常版：省略摘要、关键词、参考文献，其余照规范
MODES = (MODE_PAPER, MODE_DAILY)

# md 规范「一、论文整体结构」的 1–6（第 7 项附录可选，不计入必查）
PART_COVER = '封面'
PART_ABSTRACT = '摘要'
PART_KEYWORDS = '关键词'
PART_TOC = '目录'
PART_BODY = '正文'
PART_REFS = '参考文献'
REQUIRED_PARTS = (PART_COVER, PART_ABSTRACT, PART_KEYWORDS, PART_TOC,
                  PART_BODY, PART_REFS)
OMIT_IN_DAILY = (PART_ABSTRACT, PART_KEYWORDS, PART_REFS)

# ---------------------------------------------------------------- 字体
# md 表格：一级黑体三号/二号、二级黑体四号、三级楷体小四加粗、四级宋体小四加粗
FONT_H = {1: '黑体', 2: '黑体', 3: '楷体', 4: '宋体'}
FONT_LAT = 'Times New Roman'
FONT_COVER_TITLE = '黑体'
FONT_COVER_FIELD = '宋体'
FONT_KW_LABEL = '黑体'          # 「关键词：」标签
FONT_CAP = '黑体'               # 表题/图题
FONT_EQ = '宋体'                # 公式里的中文兜底；西文走 FONT_LAT

# ---------------------------------------------------------------- 字号
HEAD_PT = {1: 16.0, 2: 14.0, 3: 12.0, 4: 12.0}   # 三号 / 四号 / 小四 / 小四
HEAD_BOLD = {1: True, 2: True, 3: True, 4: True}
HEAD_JC = {1: 'center', 2: 'left', 3: 'left', 4: 'left'}   # 一级居中，其余左对齐
HEAD_SPACE_BEFORE = {1: 24.0, 2: 18.0, 3: 12.0, 4: 12.0}
HEAD_SPACE_AFTER = {1: 12.0, 2: 8.0, 3: 6.0, 4: 6.0}

COVER_TITLE_PT = 22.0           # 封面题名 二号
COVER_FIELD_PT = 14.0           # 封面字段行 四号
DOC_TITLE_PT = 18.0             # 无 front matter 时的文档标题 小二（日常版沿用）
COVER_TOP_BLANK = 4             # 封面顶部空行数（把题名压到上半页偏上）
COVER_TITLE_BEFORE = 24.0
COVER_TITLE_AFTER = 48.0
COVER_FIELD_AFTER = 12.0

KW_PT = 12.0                    # 关键词 小四
KW_MIN, KW_MAX = 3, 5           # md：3–5 个，分号隔开（越界只警告，不算错）
KW_SPACE_BEFORE = 6.0
KW_SPACE_AFTER = 12.0

TOC_TITLE = '目录'
TOC_LEVELS = '1-3'              # md：自动生成，最多显示到三级标题
TOC_TITLE_PT = 16.0             # 「目录」二字 黑体三号居中
TOC_TITLE_BEFORE = 0.0
TOC_TITLE_AFTER = 18.0
TOC_HINT = '在 Word 中按 Ctrl+A 后 F9 更新目录。'

# ---------------------------------------------------------------- 页码
# md：前置部分用罗马数字（Ⅰ、Ⅱ、Ⅲ），正文开始阿拉伯数字从 1 起，必须用分节符
PAGE_FMT_FRONT = 'upperRoman'
PAGE_FMT_MAIN = 'decimal'
PAGE_START = 1                  # 前置与正文都显式从 1 起（否则会接着上一节数）

# ---------------------------------------------------------------- 公式
# md：公式居中，编号右对齐，如 (2-1)
EQ_PT = 12.0                    # 小四
EQ_NUM_TMPL = '({c}-{k})'       # 按章编号：(2-1) = 第 2 章第 1 式
# 版式是 1×3 无框表格：空栏 | 公式（居中） | 编号（右对齐）。
# 两边留白取一样宽，中栏中心正好落在页心（1.75 + 12/2 = 7.75cm）—— 公式才是真居中
EQ_COL_PAD_CM = 1.75
EQ_COL_BODY_CM = 12.0           # 1.75 + 12 + 1.75 = 15.5cm = 正文栏宽
EQ_SPACE_BEFORE = 6.0
EQ_SPACE_AFTER = 6.0

# ---------------------------------------------------------------- 参考文献
# GB/T 7714-2015 著录（md 未规定字号，取五号最通用）
REF_PT = 10.5                   # 五号
REF_HANG_CHARS = 200            # 悬挂缩进 2 字符
REF_LINE_SPACING = 1.5
REF_SPACE_BEFORE = 0.0
REF_SPACE_AFTER = 0.0

# ---------------------------------------------------------------- 题注 / 表注
# 表格/图片：md 要求「图表中文字一般用五号或小五号宋体」
CAP_PT = 10.5                   # 表题/图题 五号（与表格内文字同号）
CAP_SPACE_BEFORE = 6.0
CAP_SPACE_AFTER = 3.0
NOTE_PT = 9.0                   # 表注 小五
NOTE_SPACE_BEFORE = 3.0
NOTE_SPACE_AFTER = 6.0

# ---------------------------------------------------------------- 样式名
STYLE_COVER_TITLE = '封面题名'
STYLE_COVER_FIELD = '封面字段'
STYLE_TOC_TITLE = '目录标题'    # 不进目录（不写 outlineLvl）
STYLE_KW = '关键词'
STYLE_EQ = '公式'
STYLE_REF = '参考文献'
STYLE_CAPTION = '表题'          # 沿用 doc-format 的名字：M.add_table 等按名取用
STYLE_NOTE = '表注'
STYLE_TBL = '表格文字'
STYLE_CODE = '代码'
STYLE_TITLE = '文档标题'        # 无 front matter 时首个 # 仍作文档标题（日常版）


def eq_number(chapter, index):
    """公式编号文本，如 (2-1)。"""
    return EQ_NUM_TMPL.format(c=chapter if chapter else 1, k=index)