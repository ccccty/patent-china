# Patent-china

把论文、实验材料、技术方案或已有交底，直接写成**能递交的**中国发明专利申请文件成稿。

输出单份 Word，排版复刻国内代理机构通稿样式：A4、宋体四号正文、首行缩进 2 字符、五节独立页眉带下边框、页脚页码、公式可编辑 OMML、附图集中文末。查新与对比证据剥离为独立附件，**正文零元信息**。

> 只做**发明**。实用新型、外观设计不在本技能范围。

## 产物形态

一份 docx，五节，对应国知局申请文件的五个部分：

```
说明书摘要 → 摘要附图 → 权利要求书 → 说明书 → 说明书附图
```

说明书内部章标题固定五个：技术领域 / 背景技术 / 发明内容 / 附图说明 / 具体实施方式。

外带两份旁路附件（不进 Word）：

- `查新与区别定位_{ts}.md` —— 公开号、对比文件、区别特征
- `核实清单_{ts}.md` —— 所有待人工确认项

## 为什么单独做一件事

专利正文和论文是两种文体。论文要论证新颖性，专利只要**把技术方案讲到本领域技术人员能实现**；论文可以引文献、标 URL，专利正文里出现这些就是元信息污染。

本技能的硬规矩就是围绕这条线：

| 规矩 | 含义 |
|---|---|
| 正文零元信息 | 不出现 `待核实`、公开号、URL、`[占位]`、TODO |
| 背景技术是散文 | 不点名公开号、不列对比文件清单、不做检索说明 |
| 不评价授权前景 | 不写「创新性强」「填补空白」，效果必须对得上手段 |
| 不编数据 | 材料没有的数值、对比、结论一律不造 |
| 图集中文末 | 正文只写「如图1所示」，图与图题都在第五节 |

## 安装

放进你的 agent 技能目录即可。以 DSH 为例：

```bash
git clone https://github.com/ccccty/patent-china "$HOME/.dsh/skills/patent-china"
```

## 环境

需要一个装了 `python-docx`、`latex2mathml`、`Pillow` 的 Python 解释器。`playwright` 可选，只有 mermaid 渲染需要。

```bash
pip install python-docx latex2mathml Pillow
# 可选（mermaid → PNG）
pip install playwright && playwright install chromium
```

用环境变量指定路径，技能内不写死绝对路径：

```bash
export PYTHON=/path/to/python
export SKILL_ROOT=/path/to/patent-china
export PATENT_OUT_DIR=/path/to/outputs     # 可选，默认 ./outputs
```

## 用法

把技术材料交给 agent，说「用 patent-china 写专利」即可。也可以只要权利要求书或只要说明书。

脚本可以单独调：

```bash
"$PYTHON" "$SKILL_ROOT/tools/selftest.py"                      # 一键回归，12 项
"$PYTHON" "$SKILL_ROOT/tools/assemble_application.py" --dir 案件目录 --name "一种…方法"
"$PYTHON" "$SKILL_ROOT/tools/emit_patent_docx.py" -i 稿.md -o 稿.docx
```

## 五道闸

出稿前必须全过，前一道不过不进下一道：

| 闸 | 脚本 | 查什么 |
|---|---|---|
| 一 | `check_clean_body.py` | 正文零元信息 + 五条固定套话逐字比对 |
| 二 | `check_patent_docx.py` | Word 版式：节数、字体字号、页眉页脚、公式字体、表格框线 |
| 三 | `check_consistency.py` | 图数/图号/公式编号/步骤同步/摘要长度自洽 |
| 四 | 人工 | 在 Word 里翻页眉、页码、公式可编辑、图题位置 |
| 五 | `check_render_layout.py` | **渲染后**逐页查墨迹越界、页眉、页脚、空页 |

### 第五道闸为什么必要

前四道读的都是 docx 的 XML。**XML 合法不等于印出来不越界**——渲染器把内容画到版心外时，XML 里找不到任何毛病。

实际踩过的三个压线源：

1. **长公式撑破版面**。OMML 公式是不可断对象，`A=…,\quad B=…,\quad C=…` 这种并列式，四号字下源码过百字符就会把内容推出右边距 **1.7 cm**，而且所有 XML 校验都通过。现已按 `\quad` 自动拆行，同一编号只出一次。
2. **满宽图片右偏**。`wp:inline` 不写 `distL/distR` 时，渲染器套用 0.32 cm 默认环绕距离，把图整体右移约 18 px。图宽上限因此从 14.66 cm 退到 **14.5 cm**。
3. **满宽表格外溢**。表宽正好等于版心时，边框描边与网格舍入会压出边界。表宽取版心 − 40 dxa。

## 目录结构

```
SKILL.md                  技能主体：工序总览、硬规矩、执行前核对
prompts/                  11 个工序提示词（intake → delivery）
references/
  template_spec.md        排版真源：页面、样式、页眉页脚、表格、图片
  markdown_dialect.md     合并稿的 ::: 五节语法契约
tools/
  emit_patent_docx.py     唯一 Word 生成器
  check_*.py              四道自动化闸
  render_figures.py       mermaid/材料图 → 黑白 PNG
  math_to_omml.py         LaTeX → OMML
tests/smoke.md            合规样本，selftest 拿它跑全链路
```

## 排版真源

一切版面以 `references/template_spec.md` 为准。`tools/emit_patent_docx.py` 是唯一 Word 生成器，改版式必须与 `check_patent_docx.py` 同步，否则生成与校验口径不一致。

## License

MIT