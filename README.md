# doc-format-skills

把 **Markdown 源稿生成排版规范的中文 Word 文档（.docx）**的两个 Claude Code skill。
排版在生成时就写进 XML，不做「先生成再回头修」——真标题样式、三线表、分节符、
页码分区、公式对象，都是直接写进 `.docx` 的。

| skill | 体例 | 产出 |
|---|---|---|
| **doc-format** | 中文科技报告 / 调研 / 组会材料 | 宋体小四、三线表、表题自动编号、页脚页码 |
| **doc-paper** | 学位论文 | 封面、中英文摘要、关键词、自动目录、分节符与罗马/阿拉伯页码、按章编号的图 3-2 / 表 4-1、**LaTeX → Word 公式对象**、GB/T 7714 参考文献 |

## 安装

复制到 Claude Code 的用户级 skill 目录（两个都要，doc-paper 复用 doc-format 的引擎）：

```bash
git clone https://github.com/<you>/doc-format-skills.git /tmp/doc-format-skills
cp -r /tmp/doc-format-skills/doc-format /tmp/doc-format-skills/doc-paper ~/.claude/skills/
```

依赖只有 **python-docx 1.2.0**（无 pandoc / LibreOffice / LaTeX / Word COM）：

```bash
pip install python-docx==1.2.0
```

## 用法

```bash
# 报告体例
python ~/.claude/skills/doc-format/scripts/md2docx.py   报告.md
python ~/.claude/skills/doc-format/scripts/verify.py    报告.docx --md 报告.md

# 论文体例（--mode 必填：paper=六部分齐全 / daily=日常版）
python ~/.claude/skills/doc-paper/scripts/md2paper.py   论文.md --mode paper
python ~/.claude/skills/doc-paper/scripts/verify_paper.py 论文.docx --md 论文.md
```

在 Claude Code 里直接说「把 论文.md 转成 Word，论文正式版」即可，skill 会被自动调用。

## 生成后一定要跑校验

两个 `verify*.py` 都会断言 60 余项，其中最要紧的一条是**内容无损**：
源 md 与产物的全文（剥掉格式后）必须逐字符一致 —— 改排版改掉了字，比排版丑严重得多。

## Markdown 约定（摘）

| 写法 | 含义 |
|---|---|
| 第一个 `#` | 文档标题（论文体例里有 front matter 时改为章节） |
| `表：说明` / `图：说明` | 题注，自动编号（论文体例按章：`表 4-1`） |
| `$$ … $$` | 公式；论文体例下排成 Word 公式对象，`$$!` 表示不编号 |
| 开头 `---` 包一段 | front matter = 封面元数据（`题名：…` 等），论文体例专用 |
| `# 摘要` / `关键词：…` / `# 参考文献` | 论文体例识别的部分；缺了就报错（正式版） |

完整约定见各自的 `SKILL.md`：`doc-format/SKILL.md`、`doc-paper/SKILL.md`。

## 结构

```
doc-format/    引擎在这里 —— md_parse / docx_builder / md2docx / verify / spec
doc-paper/     论文体例；_engine.py 反向 import 上面那套（两个 skill 必须同时在）
```

`docx_builder.py` 是唯一直接操作 OOXML 的地方（中文字体槽、`w:szCs`、`firstLineChars`、
三线表边框、域、schema 子元素顺序），其余模块都用它的原语。

## 已知边界

- 公式支持 LaTeX 子集（分式/上下标/根式/大型算符/定界符/重音/希腊字母/函数名/`\text{}`）；
  超出子集的（如 `\begin{matrix}`）整条回退成纯文本并给警告，不会排出个错公式。
- 图片只支持独占一行的 `![](...)`，且必须是本地文件。
- 只做生成，不修复别人给的乱排版 docx —— 建议把内容导回 Markdown 重新生成。

## 许可证

MIT，见 `LICENSE`。