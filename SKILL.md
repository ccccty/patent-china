---
name: patent-china
description: "中国发明专利一条龙：把论文、实验材料、技术方案或已有交底，直接写成符合国知局申请文件格式的成稿——说明书摘要、摘要附图、权利要求书、说明书（技术领域/背景技术/发明内容/附图说明/具体实施方式）、说明书附图，输出单份 Word，排版复刻国内代理机构通稿样式（A4、宋体四号正文、首行缩进两字符、五节独立页眉、页脚页码、公式可编辑 OMML、图集中文末）。查新与对比证据剥离为独立附件，正文零元信息。| China invention patent filings end to end: turn a paper, experiment materials, or an existing disclosure into a filing-ready specification set (abstract, abstract figure, claims, description, drawings) as one formatted Word document matching Chinese agency house style, with prior-art evidence kept in a separate companion file."
version: "1.0.0"
user-invocable: true
argument-hint: "[可选：论文/项目路径 · 交底书 · 技术方案描述 · 公式或实验材料 · 只要权利要求书 · 只要说明书]"
allowed-tools: Read, Write, Edit, Grep, Glob, WebSearch, Bash
---

# Patent-china · 中国发明专利成稿

用户给的是**技术内容**（论文、方案、实验、已有交底），本技能交付的是**能直接递交的申请文件形态成稿**：五节 Word + 两份旁路附件。

**先读「工序总览」定入口，再看「硬规矩」。**

## 工序总览

| 步 | 做什么 | 读哪个 prompt | 产物 |
|----|--------|--------------|------|
| S0 | 收材料、定发明名称、定技术领域 | `prompts/intake.md` | 案件工作目录 |
| S1 | 扫材料：图、公式、表格、数据、代码 | `prompts/materials_scan.md` | `材料清单.md` |
| S2 | 抽专利点：哪些是必要特征，哪些是优选 | `prompts/patent_points.md` | 专利点清单（F1…Fk） |
| S3 | 查新与区别定位（**只进附件**） | `prompts/prior_art.md` | `查新与区别定位_{ts}.md` |
| S4 | 权项策略：独权 + 从属树 | `prompts/claim_strategy.md` | `claim_tree` |
| S5 | 写权利要求书 | `prompts/claims_builder.md` | `权利要求书.md` |
| S6 | 定图型、出附图 | `prompts/figures.md` | `assets/*.png` |
| S7 | 写说明书五部分 | `prompts/specification_builder.md` | `说明书.md` |
| S8 | 写摘要与选摘要附图 | `prompts/abstract_builder.md` | `摘要.md` |
| S9 | 合并成单份 Markdown | `prompts/assemble.md` + `tools/assemble_application.py` | `{案件名}_{ts}.md` |
| S10 | 出 Word 并五道闸自检 | `prompts/self_check.md` | `{案件名}_{ts}.docx` |
| S11 | 写核实清单、交付 | `prompts/delivery.md` | `核实清单_{ts}.md` |

**只有 S3 和 S11 的附件里可以出现公开号、URL、检索过程。S5–S10 的主交付物里一个字都不能有。**

## 输入判定

| 用户给的是 | 从哪步进 |
|------------|----------|
| 论文（PDF / LaTeX / docx / md） | S1 全文扫 |
| 已有专利交底书 | S2 起，跳过 S1 的大部分 |
| 只有口头方案描述 | S0 起，先问清技术问题与手段 |
| 一堆实验数据/图表 | S1 起，直接奔 S5 |
| 只要权利要求书 | S4 → S5，出 `权利要求书.md` 即可 |
| 只要说明书 | 前提是先有权要，S7 起 |

材料不足时：**能写的先写，缺口进核实清单**，不要停在提问。只有「完全不知道要保护什么」才必须问。

## 硬规矩

1. **正文零元信息**。说明书、权利要求书、摘要里不得出现：`待核实`、`据检索`、`经检索`、`CN11xxxxxxA`、`https://`、`[占位]`、`TODO`、`草稿`。不确定的东西一律进《核实清单》。
2. **五节固定**。摘要 → 摘要附图 → 权利要求书 → 说明书 → 说明书附图，顺序与页眉文字都写死在 `tools/emit_patent_docx.py`，不得增删节。
3. **章标题固定五个**。技术领域 / 背景技术 / 发明内容 / 附图说明 / 具体实施方式。不写「第一章」「1.1」「3.4.1」。
4. **背景技术是散文**。不点名公开号、不列对比文件清单、不出现检索说明、不做分组综述。具体对比全部去附件。
5. **图集中在文末**。正文只写「如图1所示」，图本身和「图1」图题都在第五节。图题在图**下方**居中；表题在表**上方**居中。
6. **公式独立成段、可编辑**。`$$ … $$ (1)` 带编号，正文用「式 (1)」引用；转 OMML，失败要报告，不要贴图。
7. **权项三条腿**。方法独权 1 条 + 系统（装置）独权 1 条 + 若干从属。序号手打，不用 Word 自动编号。
8. **不评价授权前景**。不写「创新性强」「填补空白」「显著优于现有技术」这类空洞评价。效果必须对得上手段。
9. **产出不覆盖**。新文件名带时间戳，不覆盖旧目录。
10. **不编数据**。材料没有的数值、对比、实验结论一律不造；表里没有的行列不加。
11. **过闸要看渲染结果**。前几道闸只读 XML，查不出渲染器把内容画到版心外。交稿前必须 `render` 一遍并跑 `check_render_layout.py`，要 `ok=1 issues=0`。长公式、满宽图片、满宽表格是三个已知的压线源。

## 排版真源

一切版面以 `references/template_spec.md` 为准（A4、上下 2.54cm、左右 3.17cm、宋体四号正文、首行缩进 2 字符、五节页眉黑体加粗带下边框、页脚居中页码、**表格全框线（双层写法）**）。`tools/emit_patent_docx.py` 是唯一 Word 生成器，改版式必须两边同步。

Markdown 写法只认 `references/markdown_dialect.md`，违反会缺章缺图。

## 产出物

**产出根目录取环境变量 `PATENT_OUT_DIR`**；未设置时默认 `./outputs/`（相对当前工作目录）。两者都按「每案一个带时间戳的子目录」组织：

```
{PATENT_OUT_DIR}/{案件标识}_{YYYYMMDDHHmmss}/
```

| 文件 | 是否进 Word | 说明 |
|------|-------------|------|
| `{案件名}_{ts}.docx` | — | **主交付**：五节成稿 |
| `{案件名}_{ts}.md` | — | 主交付的 Markdown 源 |
| `权利要求书.md` `说明书.md` `摘要.md` | 合并进主稿 | 分段稿，便于单独改 |
| `assets/*.png` | 嵌入 | 附图 |
| `查新与区别定位_{ts}.md` | 否 | 旁路：公开号、对比、区别特征 |
| `核实清单_{ts}.md` | 否 | 旁路：所有待人工确认项 |

新目录带时间戳，不覆盖既有案件。

## 交付回复

末块固定标题 **`## 交付后请确认`**，列出：主稿路径、页数、权项条数、图数、核实清单条目数，并请用户先看核实清单再看正文。

不得在交付里把「申请」「检索」「专利地图」列为下一步——本技能到成稿为止。

## 工具

全在 `tools/`，用上面「环境」里的 `$PYTHON` 跑：

| 脚本 | 干什么 |
|------|--------|
| `assemble_application.py` | 五份分段稿 → 单份带 `:::` 标记的合并稿 |
| `emit_patent_docx.py` | 合并稿 → 五节 Word（唯一生成器） |
| `check_patent_docx.py` | Word 版式校验 |
| `check_clean_body.py` | 正文零元信息体检 + 五条固定套话逐字比对 |
| `check_consistency.py` | 图数/图号/公式编号/步骤同步/摘要长度自洽 |
| `check_render_layout.py` | **渲染后**版面体检：逐页查墨迹越界/页眉/页脚/空页 |
| `render_figures.py` | mermaid → 黑白 PNG；材料图二值化（灰底图自动降级为灰阶提白，不硬压黑） |
| `mermaid_render.py` | 通用 mermaid 批量渲染（render_figures 的底层） |
| `math_to_omml.py` | LaTeX → OMML |
| `latex_delimiters.py` | 行内公式定界符规范化 |
| `stdio_utf8.py` | 控制台 UTF-8 兜底 |
| `selftest.py` | 一键回归：`"$PYTHON" "$SKILL_ROOT/tools/selftest.py"` |

`tests/smoke.md` 是合规样本，`selftest.py` 拿它跑完整链路，当前 **ok_n=12 fail=0**（含渲染体检）。改完任何脚本先跑自测。改版式要同时改 `emit_patent_docx.py` 与 `check_patent_docx.py`，两边口径必须一致。

**XML 级校验看不见渲染结果。** 上面几个 checker 都只读 docx 的 XML，查不出「渲染器把内容画到版心外」。交稿前必须补一道渲染体检。

`check_render_layout.py` 支持两条路径：

```bash
# 甲：自己已有渲染好的 PNG 目录
"$PYTHON" "$SKILL_ROOT/tools/check_render_layout.py" --render-dir <PNG目录>

# 乙：交给工具自动渲染（需本地有 LibreOffice Kit 或 LibreOffice）
"$PYTHON" "$SKILL_ROOT/tools/check_render_layout.py" --docx 稿.docx
```

自动渲染靠两个环境变量定位：`DSH_NODE`（node 可执行文件）与 `DSH_LIBREOFFICE_CLI`（LibreOffice Kit 的 `cli.js`）。没有这两个变量时会去查常见安装位置；都找不到则报 `reason=renderer_unavailable`。

渲染命令本身（甲路径的前置步骤）：

```bash
"$NODE" "$LIBREOFFICE_CLI" render --input 稿.docx --output-dir <尚不存在的目录>
```

要 `RENDER_LAYOUT: ok=1 issues=0`。两点注意：

- `--output-dir` 必须由渲染器自己创建，先建好会报 `EEXIST`。
- 无渲染器时默认**算失败**（防止闸门空转）；确实没有渲染环境才加 `--allow-skip` 显式跳过，自测里会显示为 `SKIP` 而非通过。

## 执行前核对

```
□ 已 Read 对应 prompts/，未凭记忆写
□ 五节标记齐全（::: 摘要 / 摘要附图 / 权利要求书 / 说明书 / 说明书附图）
□ check_clean_body.py 报 CLEAN_BODY ok=1
□ check_consistency.py 报 CONSISTENCY ok=1
□ emit_patent_docx.py 报 PATENT_DOCX ok=1
□ check_patent_docx.py 报 PATENT_LAYOUT ok=1
□ check_render_layout.py 报 RENDER_LAYOUT ok=1 issues=0
□ 方框图是黑白纯二值（灰度占比 0）
□ 核实清单已写，且未混入正文
□ 五节页眉文字与页码重编已在 Word 里看过
```

## 环境

需要一个装了 `python-docx`、`latex2mathml`、`Pillow` 的 Python 解释器（`playwright` 可选，仅 mermaid 渲染需要）。用 `PYTHON` 环境变量指定，未设置时用 PATH 上的 `python`。

技能内脚本一律按「相对技能根目录」调用，不写死绝对路径：

```bash
"$PYTHON" "$SKILL_ROOT/tools/emit_patent_docx.py" -i 稿.md -o 稿.docx
```

本机（Windows）配置示例，可按自己的环境替换：

```bat
set PYTHON=D:\devtools\patent-skill\bin\python.cmd
set SKILL_ROOT=D:\资产_产出\工作流\专利skill
set PATENT_OUT_DIR=D:\资产_产出\工作流\专利产出
```

> 若 PATH 上的 `python` 是别的 venv（缺 `python-docx`），务必显式设 `PYTHON`，否则脚本一跑就报 `ModuleNotFoundError`。

## 本机安装注记（DSH 本地，非上游）

- 建于 2026-11-02，替代 `patent-disclosure-skill` 的发明交底路径。
- 只做**发明**。实用新型、外观设计不在本技能范围。
- 版式逆向自用户提供的《一种基于多尺度非对称融合网络的水稻分布智能制图方法 250630(1).docx》。