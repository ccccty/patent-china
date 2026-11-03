# S9 · 合并成单份交底稿

把 `摘要.md`、`权利要求书.md`、`说明书.md` 与 `assets/` 拼成一份带 `:::` 标记的 Markdown；只做拼接、编号与路径归一，不改写文字。

## 一、骨架

```markdown
::: 摘要
（摘要正文，一段，不超过300字）
::: 摘要附图
![摘要附图](assets/abstract.png)
::: 权利要求书
1.一种……，其特征在于，包括如下步骤：
S1、……；S2、……；
2.根据权利要求1所述的……，其特征在于，……。
::: 说明书
# 一种……方法
## 技术领域
## 背景技术
## 发明内容
## 附图说明
## 具体实施方式
::: 说明书附图
![图1](assets/fig1.png)
```

`:::` 五个标记顶格独占一行、顺序固定，缺一即 `PATENT_DOCX_SECTIONS: missing`；节标题由脚本写进页眉，正文不再重复；图题不手写，脚本按图序生成。示例为省行略去段间空行，实际稿每段之间空一行。

## 二、命名与图片

- Markdown 用 `{规范化案件名}_{YYYYMMDDHHmmss}.md`，Word 同名同戳只换扩展名；案件名取发明名称去掉「一种」的主干，20 字内不留空格与标点。
- 时间戳实时取，同一秒撞名加 `_1` 后缀，**不得覆盖**已有 `.md` 或 `.docx`。
- 图片一律 `assets/figN.png`，N 与图序一致；摘要附图用 `assets/abstract.png`；路径相对 Markdown 所在目录。

## 三、自检顺序（逐级放行）

1. **章节齐全**：`--check` 看 `PATENT_DOCX_SECTIONS: ok found=5/5`，出现 `missing` 先补 `:::` 再往下走。
2. **公式**：看 `PATENT_DOCX_MATH` 的 `fallback`，非 0 则简化该式或改写为文字关系。
3. **图**：看 `PATENT_DOCX_IMAGES` 的 `missing`，非 0 则补图或改路径。
4. **出 Word**：前三步全通过才执行，出后核对五节页眉与页码连续。

## 四、出 Word

用合并脚本一次完成拼接与命名（它自己加五个 `:::` 标记，并拒绝覆盖）：

```bash
"$PYTHON" "$SKILL_ROOT/tools/assemble_application.py" --dir 案件目录 --name "一种……方法"
"$PYTHON" "$SKILL_ROOT/tools/emit_patent_docx.py" -i <合并稿>.md -o <合并稿>.docx
"$PYTHON" "$SKILL_ROOT/tools/check_patent_docx.py" -i <合并稿>.docx
"$PYTHON" "$SKILL_ROOT/tools/check_clean_body.py" -i <合并稿>.md
"$PYTHON" "$SKILL_ROOT/tools/check_consistency.py" -i <合并稿>.md
"$PYTHON" "$SKILL_ROOT/tools/check_render_layout.py" --docx <合并稿>.docx
```

`emit_patent_docx.py` 参数：`-i/--input` 必填；`-o/--output` 非 `--check` 时必填；`--check` 只校验不出 Word；`--no-table-borders` 关掉表格框线（**默认加框线**，与模板一致）；`--continuous-page` 页码全文连续（默认每节从 1 重编）。