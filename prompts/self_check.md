# S10 · 出稿与自检

出 Word 前后各跑一遍下面的检查。**任何一项 fail 都不许交付。**

## 出稿

```powershell
python tools/assemble_application.py --dir 案件目录 --name "一种……方法"
python tools/emit_patent_docx.py -i 合并稿.md -o 合并稿.docx
```

`assemble_application.py` 自己加五个 `:::` 标记、自己取时间戳、拒绝覆盖。若分段稿被编辑器存成带 BOM，它会把 BOM 清掉——BOM 跑到行首会让标题与首图解析失败。

## 五道闸

按顺序放行，前一道不过不进下一道。

### 闸一：正文零元信息

```powershell
python tools/check_clean_body.py -i 合并稿.md
```

看 `CLEAN_BODY: ok=1 fail=0`。它在查：

| 规则 | 拦什么 |
|------|--------|
| `publication_no` | `CN118…` 这类公开号 |
| `url` / `doi` | 链接与文献标识 |
| `search_talk` | 「经检索」「据检索」「检索到」 |
| `todo` | 「待核实」「占位」「TODO」「【…】」 |
| `draft` | 「草稿」「此处应」 |
| `meta_skill` | 「本技能」「prompt」「skill」 |
| `chapter_num` | 「第X章」「1.1」章节号 |
| `latex_bare_paren` | 用普通括号包 LaTeX 的行内公式 |

warn 级（`hype` 空洞评价、`citation_bracket` 方括号数字）可以留，但要人工看一眼是不是真有必要。

### 闸二：版式

```powershell
python tools/check_patent_docx.py -i 合并稿.docx
```

看 `PATENT_LAYOUT: ok=1 fail=0`。它核对：五节齐全、正文宋体四号、五个章标题齐、摘要附图在第 2 节、表格单元格有框线且表头加粗、公式字体为 Cambria Math 四号、页码按节重编、无杂字体残留。

`warn` 里允许出现的是模板自带的瑕疵（图5/图6 楷体残留、表头部分加粗、表宽 8696 超版心）。这些是**用户模板本身**的问题，不要复制进新稿：新稿出现同样的 warn 说明你抄了瑕疵。

### 闸三：内在一致性

```powershell
python tools/check_consistency.py -i 合并稿.md
```

看 `CONSISTENCY: ok=1 fail=0`。机器能查的八项：

| 项 | 查什么 |
|----|--------|
| `figure_list` | 附图说明列了几条 |
| `figure_count_match` | 说明书附图张数 = 附图说明条数 |
| `figure_ref_valid` | 每处「如图N所示」的 N 都有定义 |
| `figure_numbering` | 图号从 1 连续 |
| `equation_ref_valid` | 每个「式 (N)」都有对应公式，无悬空引用 |
| `equation_numbering` | 公式编号从 1 连续 |
| `step_sync` | 权项的 `S1…Sn` 与实施例的 `步骤1…步骤n` 同数 |
| `abstract_len` | 摘要 1–300 字 |

剩下机器查不了的，人工核：

- [ ] 每条从属都能在说明书「优选地，…」里找到对应段
- [ ] 有益效果（1）（2）（3）每条都能对应到具体手段，不是空喊
- [ ] 摘要提到独权的核心手段
- [ ] 背景技术四段散文，落点与发明内容的问题句一致

### 闸四：打开看

用 Word 打开，翻这五处：

1. 第 1 节页眉写「说明书摘要」，第 3 节写「权利要求书」，第 5 节写「说明书附图」
2. 每节页码都从 1 开始
3. 公式是能双击编辑的，不是图片
4. 图都在第 5 节，且图题在图下方居中
5. 表格有框线、表头加粗

### 闸五：渲染体检（XML 查不出的那一层）

前三道闸读的都是 docx 的 XML。**XML 合法不等于印出来不越界**——渲染器把内容画到版心外时，XML 里找不到任何毛病。所以必须渲染一遍再测：

```bash
node <libreoffice-kit>/lib/cli.js render --input 稿.docx --output-dir <尚不存在的目录>
python tools/check_render_layout.py --render-dir <那个目录>
```

要 `RENDER_LAYOUT: ok=1 issues=0`。它逐页查：左/右/下越界、缺页眉、缺页眉下横线、缺页脚、空页。

`--output-dir` 必须是不存在的目录，由渲染器自己创建；先 `mkdir` 好再跑会报 `EEXIST: file already exists`。

## 出问题怎么改

| 现象 | 原因 | 改法 |
|------|------|------|
| `SECTIONS: missing` | 少了 `:::` 行 | 补标记，顺序固定 |
| 图片数量不对 | 合并稿带 BOM 或路径错 | 重跑 assemble，检查 `assets/` |
| 标题没认出来 | `# ` 后有多余空格或用了 `## ` | 标题必须 `# 名称` |
| 公式变灰底文本 | `\(` `\)` 写成了普通括号 | 改 delimiters |
| `MATH: fallback>0` | 用了 `\tilde` `\hat` `\bar` | 换符号 |
| 表格没线 | 只写了表级 borders | 用默认出稿（单元格级才有线） |
| 章标题多出编号 | 用了 Markdown 有序列表 | 编号手打成文字 |
| 公式编号是裸数字 | 旧版 `_emit_math` 丢括号 | 已修；跑 `math_number_paren` 确认 |
| 套话被报不一致 | 抄成了「描述**得**比较简单」 | 模板原文是「描述**的**比较简单」，一字不能差 |
| `RENDER_LAYOUT` 报某页右越界 | 长公式不可断，整行撑出版心 | 让它含 `\quad` 且超 105 字符，`_split_math` 会自动拆行 |
| `RENDER_LAYOUT` 报含图页右越界 | `wp:inline` 缺 `distL/distR`，被套了 0.32cm 默认环绕 | 已由 `_zero_wrap_dist` 显式写 0，图宽退到 14.5cm |
| 表格右侧压出边界 | 表宽正好等于版心，描边与网格舍入外溢 | 表宽取版心 − 40 dxa（8272） |

## 已踩过的坑（照着避）

1. **BOM**：编辑器存成 UTF-8 with BOM，BOM 跑到行首会让 `# 标题` 和首张图解析失败。`assemble_application.py` 已自动清；自己手搓合并稿时要防。
2. **表格边框双层**：表级 `tblBorders=none` + 单元格级 `tcBorders=single`。只写一层，要么全没线，要么全是默认粗线。
3. **公式字体**：每个 `m:r` 都要 `Cambria Math` + `sz=28`，否则公式比正文小一圈。
4. **公式编号要带括号**：正文引用「式 (1)」，公式行末也必须出 `(1)`，不是裸 `1`。
5. **灰底图别硬二值化**：黑占比从 3% 暴涨到 21% 说明浅灰底被压成黑块。`render_figures.py` 已自动降级为灰阶提白，别加 `--force-bw` 硬来。
6. **长公式撑破版面**：`A=…,\quad B=…,\quad C=…` 这种并列式，四号字下源码过百字符就会把内容推出右边距 1.7cm。XML 一切正常，只有渲染才看得见。现在按 `\quad` 拆行、同一编号只出一次。
7. **图宽取满版心必然越界**：`wp:inline` 的默认环绕距离 0.32cm 不会写进 XML，但渲染器会算。图宽上限 14.5cm + 显式 `dist=0`。

## 不要做

- 不要为了让检查通过去改检查器
- 不要把 warn 当 fail 反复折腾（模板瑕疵本就存在）
- 不要在 `check` 没过时先交付再补
- 不要手工去 Word 里改格式（下次生成就丢了）