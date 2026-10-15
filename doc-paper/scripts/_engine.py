# -*- coding: utf-8 -*-
"""把 doc-format 的共享引擎挂上 sys.path，并 import 出来。

doc-paper（学位论文体例）**不自带渲染引擎**：把 Markdown 变成 Word 的那台机器
只有一份，在 doc-format skill 里 —— 中文字体槽、szCs、firstLineChars、三线表、
schema 子元素顺序、域、文本流比对，全部复用那边。本 skill 只写论文专属的东西
（体例参数、分节符与页码、封面、公式行、参考文献、结构校验）。

两个 skill 必须同时存在：

    <skills>/doc-format/scripts/{md_parse,docx_builder,md2docx,verify,spec}.py
    <skills>/doc-paper/scripts/   ← 本文件所在处

查找顺序（第一个能用的即采用）：
  1. 环境变量 DOC_FORMAT_SCRIPTS 指向的目录
  2. 同级 skill 目录 <skills>/doc-format/scripts（正常情况）
  3. skills 上一层里的 doc-format/scripts（两个 skill 一起被拷到别处时）

找不到时抛的是人话，而不是一句 `No module named 'docx_builder'`。
"""
import os
import sys

_HERE = os.path.dirname(os.path.abspath(__file__))
_SKILLS = os.path.dirname(os.path.dirname(_HERE))          # ...\.claude\skills


def _candidates():
    env = os.environ.get('DOC_FORMAT_SCRIPTS')
    if env:
        yield env
    yield os.path.join(_SKILLS, 'doc-format', 'scripts')
    yield os.path.join(os.path.dirname(_SKILLS), 'doc-format', 'scripts')


def _find():
    for c in _candidates():
        if os.path.isfile(os.path.join(c, 'docx_builder.py')):
            return c
    raise ImportError(
        '找不到 doc-format 的引擎目录（里面应有 docx_builder.py / md_parse.py）。\n'
        'doc-paper 复用 doc-format 的渲染引擎，两个 skill 必须同时安装 ——\n'
        '不要为了「独立」把引擎复制一份，那样两份会各自漂移。\n'
        '已查找：\n  ' + '\n  '.join(_candidates()) +
        '\n也可用环境变量 DOC_FORMAT_SCRIPTS 指定 doc-format/scripts 的路径。')


ENGINE_DIR = _find()
if ENGINE_DIR not in sys.path:
    sys.path.insert(0, ENGINE_DIR)

import docx_builder as B    # noqa: E402  原语层：rFonts 四槽 / szCs / 三线表 / schema 顺序 / 域
import md_parse as P        # noqa: E402  Markdown → 块序列
import md2docx as M         # noqa: E402  通用块渲染：表格/列表/代码/引用/插图/分页/超链接
import spec                 # noqa: E402  doc-format 的共享常量（页边距 / 三线表 / 表格字号）
import verify as V          # noqa: E402  文本流比对与断言基座

__all__ = ['B', 'P', 'M', 'V', 'spec', 'ENGINE_DIR']