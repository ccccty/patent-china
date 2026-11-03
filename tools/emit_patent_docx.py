#!/usr/bin/env python
"""专利 Markdown → 申请文件 Word（五节、复刻模板排版）。

唯一 Word 生成器。排版依据 `references/template_spec.md`，二者须同步改动。

Markdown 方言见 `references/markdown_dialect.md`。要点：

  ::: 摘要            开启一节（摘要 / 摘要附图 / 权利要求书 / 说明书 / 说明书附图）
  # 发明名称          说明书首行标题（居中加粗三号）
  ## 技术领域         章标题（宋体四号加粗）
  ### 一级小节        可选，用于实施例分组（宋体四号加粗）
  正文段落            宋体四号、首行缩进 2 字符
  $$ ... $$           独立成段公式（OMML），编号由正文手写 (1)
       其中，…        符号解释段（其中开头的段落不缩进时仍按正文处理）
  | 表 |              表题写在表上方，单独一行「表1」
  ![图N](path)        说明书附图节内嵌图，图题「图N」单独一行在图下方
  <!-- pagebreak -->  强制分页

用法：
  python tools/emit_patent_docx.py -i 申请文件.md -o 申请文件.docx
  python tools/emit_patent_docx.py -i 申请文件.md --check      # 只做校验不出 Word
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, Twips

try:
    from math_to_omml import try_latex_to_omml
except ImportError:  # 允许无 latex2mathml 环境降级
    try_latex_to_omml = None

# ---------------------------------------------------------------------------
# 版式常量（与 references/template_spec.md 一一对应）
# ---------------------------------------------------------------------------

PAGE_W_TWIPS = 11906
PAGE_H_TWIPS = 16838
MARGIN_TB = 1440          # 2.54 cm
MARGIN_LR = 1797          # 3.17 cm
HEADER_DIST = 851
FOOTER_DIST = 992

FONT_ZH = "宋体"
FONT_EN = "宋体"
FONT_HEADER = "黑体"

SZ_BODY = 28              # 四号 14 pt
SZ_TITLE = 32             # 三号 16 pt
SZ_TABLE = 24             # 小四 12 pt
SZ_NORMAL = 21            # 五号 10.5 pt

FIRST_LINE = 560          # 2 字符 @ 四号
FIRST_LINE_CHARS = 200

MAX_IMG_W_CM = 14.5     # 版心 14.66cm，留 0.16cm 余量防舍入压线
MAX_IMG_H_CM = 20.0

MATH_NS_W_RPR = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr"

SECTION_TITLES = {
    "摘要": "说明书摘要",
    "摘要附图": "摘要附图",
    "权利要求书": "权利要求书",
    "说明书": "说明书",
    "说明书附图": "说明书附图",
}
SECTION_ORDER = ["摘要", "摘要附图", "权利要求书", "说明书", "说明书附图"]


# ---------------------------------------------------------------------------
# 低层 XML 助手
# ---------------------------------------------------------------------------

def _sub(parent, tag: str, **attrs):
    el = OxmlElement(tag)
    for k, v in attrs.items():
        el.set(qn(f"w:{k}"), str(v))
    parent.append(el)
    return el


def _set_run_font(run, *, size: int = SZ_BODY, bold: bool = False,
                  name_zh: str = FONT_ZH, name_en: str = FONT_EN) -> None:
    run.font.size = Pt(size / 2)
    run.font.bold = bold
    run.font.name = name_en
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), name_en)
    rFonts.set(qn("w:hAnsi"), name_en)
    rFonts.set(qn("w:eastAsia"), name_zh)
    rFonts.set(qn("w:cs"), name_en)
    if bold:
        bCs = OxmlElement("w:bCs")
        rPr.append(bCs)


def _mark_rpr(paragraph, *, size: int = SZ_BODY, bold: bool = False) -> None:
    """把 rPr 写进段落标记，与样板一致（空段落也带字号）。"""
    pPr = paragraph._p.get_or_add_pPr()
    rPr = pPr.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        pPr.append(rPr)
    rFonts = OxmlElement("w:rFonts")
    rFonts.set(qn("w:ascii"), FONT_EN)
    rFonts.set(qn("w:hAnsi"), FONT_EN)
    rFonts.set(qn("w:cs"), FONT_EN)
    rPr.append(rFonts)
    if bold:
        rPr.append(OxmlElement("w:b"))
        rPr.append(OxmlElement("w:bCs"))
    for tag in ("w:sz", "w:szCs"):
        el = OxmlElement(tag)
        el.set(qn("w:val"), str(size))
        rPr.append(el)


def _first_line_indent(paragraph, chars: int = FIRST_LINE_CHARS) -> None:
    pPr = paragraph._p.get_or_add_pPr()
    ind = pPr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        pPr.append(ind)
    ind.set(qn("w:firstLine"), str(FIRST_LINE))
    ind.set(qn("w:firstLineChars"), str(chars))


def _no_indent(paragraph) -> None:
    pPr = paragraph._p.get_or_add_pPr()
    ind = pPr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        pPr.append(ind)
    ind.set(qn("w:firstLine"), "0")
    ind.set(qn("w:firstLineChars"), "0")


# ---------------------------------------------------------------------------
# 页面与节
# ---------------------------------------------------------------------------

def _setup_first_section(section, *, restart_page: bool = False) -> None:
    section.page_width = Twips(PAGE_W_TWIPS)
    section.page_height = Twips(PAGE_H_TWIPS)
    section.top_margin = Twips(MARGIN_TB)
    section.bottom_margin = Twips(MARGIN_TB)
    section.left_margin = Twips(MARGIN_LR)
    section.right_margin = Twips(MARGIN_LR)
    section.header_distance = Twips(HEADER_DIST)
    section.footer_distance = Twips(FOOTER_DIST)
    section.gutter = Twips(0)

    sectPr = section._sectPr
    # 页面边框：显式关闭
    for old in sectPr.findall(qn("w:pgBorders")):
        sectPr.remove(old)
    borders = OxmlElement("w:pgBorders")
    for side in ("top", "left", "bottom", "right"):
        _sub(borders, f"w:{side}", val="none", sz="0", space="0")
    sectPr.append(borders)

    # 页码按节重编（模板：第 2 节起 pgNumType start=1）
    for old in sectPr.findall(qn("w:pgNumType")):
        sectPr.remove(old)
    if restart_page:
        _sub(sectPr, "w:pgNumType", start="1")

    # 分栏
    for old in sectPr.findall(qn("w:cols")):
        sectPr.remove(old)
    _sub(sectPr, "w:cols", space="720", num="1")
    # 文档网格
    for old in sectPr.findall(qn("w:docGrid")):
        sectPr.remove(old)
    _sub(sectPr, "w:docGrid", type="linesAndChars", linePitch="312", charSpace="0")


def _add_page_field(paragraph) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run()
    fld = OxmlElement("w:fldChar")
    fld.set(qn("w:fldCharType"), "begin")
    run._element.append(fld)

    run2 = paragraph.add_run()
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE \\* MERGEFORMAT "
    run2._element.append(instr)

    run3 = paragraph.add_run()
    fld3 = OxmlElement("w:fldChar")
    fld3.set(qn("w:fldCharType"), "separate")
    run3._element.append(fld3)

    run4 = paragraph.add_run("1")

    run5 = paragraph.add_run()
    fld5 = OxmlElement("w:fldChar")
    fld5.set(qn("w:fldCharType"), "end")
    run5._element.append(fld5)


def _style_header(section, text: str) -> None:
    section.header.is_linked_to_previous = False
    header = section.header
    p = header.paragraphs[0]
    for r in list(p.runs):
        r._element.getparent().remove(r._element)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT

    pPr = p._p.get_or_add_pPr()
    # 下边框
    for old in pPr.findall(qn("w:pBdr")):
        pPr.remove(old)
    pBdr = OxmlElement("w:pBdr")
    _sub(pBdr, "w:bottom", val="single", color="auto", sz="6", space="0")
    pPr.append(pBdr)
    # 单倍行距
    for old in pPr.findall(qn("w:spacing")):
        pPr.remove(old)
    _sub(pPr, "w:spacing", line="240", lineRule="auto")

    run = p.add_run(text)
    _set_run_font(run, size=SZ_BODY, bold=True,
                  name_zh=FONT_HEADER, name_en=FONT_HEADER)


def _style_footer(section) -> None:
    section.footer.is_linked_to_previous = False
    footer = section.footer
    p = footer.paragraphs[0]
    for r in list(p.runs):
        r._element.getparent().remove(r._element)
    _add_page_field(p)


# ---------------------------------------------------------------------------
# 解析
# ---------------------------------------------------------------------------

RE_SECTION = re.compile(r"^:::\s*(摘要|摘要附图|权利要求书|说明书|说明书附图)\s*$")
RE_H2 = re.compile(r"^##\s+(.+?)\s*$")
RE_H3 = re.compile(r"^###\s+(.+?)\s*$")
RE_H1 = re.compile(r"^#\s+(.+?)\s*$")
RE_IMG = re.compile(r"^!\[(?P<alt>[^\]]*)\]\((?P<src>[^)]+)\)\s*$")
RE_TABLE_SEP = re.compile(r"^\|[\s:|-]+\|$")
RE_PAGEBREAK = re.compile(r"^<!--\s*pagebreak\s*-->$")
RE_MATH_CLOSE = re.compile(r"^\$\$\s*(?:\(\s*(?P<num>\d+)\s*\))?\s*$")
RE_INLINE_MATH = re.compile(r"\\\((.+?)\\\)|\$(?!\$)(.+?)\$")


def parse_markdown(text: str) -> list[dict]:
    # BOM 会让首行标题、首张图片识别失败，先清掉
    text = text.replace("\ufeff", "")
    """切成块序列。块类型：section / h1 / h2 / h3 / p / table / img / math / pagebreak。"""
    lines = text.splitlines()
    blocks: list[dict] = []
    i = 0
    n = len(lines)

    while i < n:
        raw = lines[i]
        line = raw.rstrip()
        stripped = line.strip()

        if not stripped:
            i += 1
            continue

        m = RE_SECTION.match(stripped)
        if m:
            blocks.append({"t": "section", "v": m.group(1)})
            i += 1
            continue

        # 摘要附图 / 说明书附图节内的图必须优先于其它判断
        m = RE_PAGEBREAK.match(stripped)
        if m:
            blocks.append({"t": "pagebreak"})
            i += 1
            continue

        m = RE_H1.match(stripped)
        if m and not stripped.startswith("##"):
            blocks.append({"t": "h1", "v": m.group(1)})
            i += 1
            continue

        m = RE_H2.match(stripped)
        if m:
            blocks.append({"t": "h2", "v": m.group(1)})
            i += 1
            continue

        m = RE_H3.match(stripped)
        if m:
            blocks.append({"t": "h3", "v": m.group(1)})
            i += 1
            continue

        # 块级公式：结束的 $$ 行可带编号 $$ (1)
        if stripped == "$$":
            i += 1
            body = []
            num = None
            while i < n:
                cur = lines[i].strip()
                close = RE_MATH_CLOSE.match(cur)
                if close:
                    num = close.group("num")
                    i += 1
                    break
                body.append(lines[i])
                i += 1
            blocks.append({"t": "math", "v": "\n".join(body).strip(), "num": num})
            continue

        m = re.match(r"^\$\$(.+?)\$\$\s*(?:\(\s*(\d+)\s*\))?\s*$", stripped)
        if m:
            blocks.append({"t": "math", "v": m.group(1).strip(), "num": m.group(2)})
            i += 1
            continue

        m = RE_IMG.match(stripped)
        if m:
            blocks.append({"t": "img", "alt": m.group("alt"), "src": m.group("src")})
            i += 1
            continue

        # 表格
        if stripped.startswith("|") and stripped.endswith("|"):
            rows = []
            while i < n and lines[i].strip().startswith("|"):
                cur = lines[i].strip()
                if not RE_TABLE_SEP.match(cur):
                    cells = [c.strip() for c in cur.strip("|").split("|")]
                    rows.append(cells)
                i += 1
            if rows:
                blocks.append({"t": "table", "rows": rows})
            continue

        # 普通段落：连续非空行合并
        buf = [stripped]
        i += 1
        while i < n:
            nxt = lines[i].strip()
            if (not nxt or RE_MATH_CLOSE.match(nxt) or RE_SECTION.match(nxt)
                    or RE_H1.match(nxt) or RE_H2.match(nxt) or RE_H3.match(nxt)
                    or RE_IMG.match(nxt) or nxt.startswith("|")
                    or RE_PAGEBREAK.match(nxt)):
                break
            buf.append(nxt)
            i += 1
        blocks.append({"t": "p", "v": "".join(buf)})

    return blocks


# ---------------------------------------------------------------------------
# 富文本渲染（行内公式 \\() 与 **加粗**）
# ---------------------------------------------------------------------------

RE_BOLD = re.compile(r"\*\*(.+?)\*\*")


def _emit_inline(paragraph, text: str, *, size: int = SZ_BODY,
                 bold: bool = False) -> None:
    """把含 \\(…\\) 行内公式与 **加粗** 的文本写进段落。"""
    pos = 0
    tokens: list[tuple[str, str, bool]] = []  # (kind, payload, bold)

    for m in RE_INLINE_MATH.finditer(text):
        if m.start() > pos:
            tokens.append(("text", text[pos:m.start()], bold))
        latex = (m.group(1) or m.group(2) or "").strip()
        tokens.append(("math", latex, bold))
        pos = m.end()
    if pos < len(text):
        tokens.append(("text", text[pos:], bold))

    for kind, payload, is_bold in tokens:
        if kind == "math":
            if try_latex_to_omml is not None:
                el = try_latex_to_omml(payload, display=False)
                if el is not None:
                    _style_omml(el)
                    paragraph._p.append(el)
                    continue
            run = paragraph.add_run(payload)
            _set_run_font(run, size=size, bold=is_bold)
            continue

        # text：再切 **加粗**
        seg_pos = 0
        for bm in RE_BOLD.finditer(payload):
            if bm.start() > seg_pos:
                run = paragraph.add_run(payload[seg_pos:bm.start()])
                _set_run_font(run, size=size, bold=is_bold)
            run = paragraph.add_run(bm.group(1))
            _set_run_font(run, size=size, bold=True)
            seg_pos = bm.end()
        if seg_pos < len(payload):
            run = paragraph.add_run(payload[seg_pos:])
            _set_run_font(run, size=size, bold=is_bold)


# ---------------------------------------------------------------------------
# 块渲染
# ---------------------------------------------------------------------------

def _emit_body(doc, text: str, *, indent: bool = True) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if indent:
        _first_line_indent(p)
    else:
        _no_indent(p)
    _mark_rpr(p)
    _emit_inline(p, text)


def _emit_title(doc, text: str) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _no_indent(p)
    _mark_rpr(p, size=SZ_TITLE, bold=True)
    run = p.add_run(text)
    _set_run_font(run, size=SZ_TITLE, bold=True)


def _emit_heading(doc, text: str) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    _no_indent(p)
    _mark_rpr(p, bold=True)
    run = p.add_run(text)
    _set_run_font(run, size=SZ_BODY, bold=True)


TEXT_W_TWIPS = PAGE_W_TWIPS - 2 * MARGIN_LR   # 正文净宽 8312 twips


MATH_FONT = "Cambria Math"


def _style_omml(math_el, size: int = SZ_BODY) -> None:
    """给 OMML 内每个 m:r 的 w:rPr 补 Cambria Math + 指定字号。

    模板里每个 m:r 都显式写了 Cambria Math 与 sz=28；缺了这些，
    公式会比正文小一圈，且正文中文字体可能串进公式。
    超宽公式由调用方传更小的 size 以免撑出版心。
    """
    M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"

    def _apply(rPr):
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is None:
            rFonts = OxmlElement("w:rFonts")
            rPr.insert(0, rFonts)
        rFonts.set(qn("w:ascii"), MATH_FONT)
        rFonts.set(qn("w:hAnsi"), MATH_FONT)
        rFonts.set(qn("w:cs"), MATH_FONT)
        for tag in ("w:sz", "w:szCs"):
            for old in rPr.findall(qn(tag)):
                rPr.remove(old)
            el = OxmlElement(tag)
            el.set(qn("w:val"), str(size))
            rPr.append(el)

    # 1) 所有已存在的 w:rPr（m:r 内、m:ctrlPr 内）
    for node in list(math_el.iter(MATH_NS_W_RPR)):
        _apply(node)

    # 2) 没有 w:rPr 的 m:r / m:ctrlPr，补一个
    for tag in ("r", "ctrlPr"):
        for node in math_el.iter(M + tag):
            if node.find(MATH_NS_W_RPR) is None:
                rPr = OxmlElement("w:rPr")
                _apply(rPr)
                m_rPr = node.find(M + "rPr")
                if m_rPr is not None:
                    m_rPr.addnext(rPr)
                else:
                    node.append(rPr)


def _math_font_size(latex: str) -> int:
    """按 LaTeX 源码体积估算公式宽度，超宽时降字号。

    仅为兜底：真正超宽的长行由 `_split_math` 拆成多行，
    因为 OMML 公式是一个不可断对象，降字号只能救回一点点。
    """
    n = len(latex)
    if n <= 130:
        return SZ_BODY
    if n <= 170:
        return SZ_BODY - 2      # 13 pt
    return SZ_BODY - 4          # 12 pt


MATH_BREAK_CHARS = 105        # 超过此长度且含 \quad 的公式拆成多行


def _split_math(latex: str) -> list[str]:
    """把 `A,\\quad B,\\quad C` 这类并列长公式拆成多行。

    实测：四号字下单个公式宽度超过版心（14.66cm）时，渲染器不会换行，
    整行右压出边界 1.7cm。拆成多行是唯一稳妥解法。
    """
    if len(latex) <= MATH_BREAK_CHARS:
        return [latex]
    parts = [p.strip() for p in re.split(r"\\qquad|\\quad", latex)]
    parts = [p for p in parts if p]
    if len(parts) < 2:
        return [latex]
    return parts


def _emit_math(doc, latex: str, number: str | None = None) -> bool:
    """公式段。短公式单行居中；含 \\quad 的长公式拆成多行同一编号。"""
    p = doc.add_paragraph()
    _no_indent(p)
    _mark_rpr(p)
    parts = _split_math(latex)
    multiline = len(parts) > 1

    if number:
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pPr = p._p.get_or_add_pPr()
        tabs = OxmlElement("w:tabs")
        if not multiline:
            _sub(tabs, "w:tab", val="center", pos=str(TEXT_W_TWIPS // 2))
        _sub(tabs, "w:tab", val="right", pos=str(TEXT_W_TWIPS))
        pPr.append(tabs)
        if not multiline:
            lead = p.add_run()
            lead._element.append(OxmlElement("w:tab"))
    else:
        p.alignment = (WD_ALIGN_PARAGRAPH.LEFT if multiline
                       else WD_ALIGN_PARAGRAPH.CENTER)

    ok = True
    for idx, part in enumerate(parts):
        if idx:
            brk = p.add_run()
            brk._element.append(OxmlElement("w:br"))
        el = None
        if try_latex_to_omml is not None:
            el = try_latex_to_omml(part, display=True)
        if el is not None:
            _style_omml(el, size=_math_font_size(part))
            p._p.append(el)
        else:
            ok = False
            run = p.add_run(part)
            _set_run_font(run)

    if number:
        tail = p.add_run()
        tail._element.append(OxmlElement("w:tab"))
        # 编号必须带圆括号：正文引用写作「式 (1)」，此处不能只出裸数字
        label = number if number.startswith("(") else f"({number})"
        num_run = p.add_run(label)
        _set_run_font(num_run, size=SZ_BODY)
    return ok


def _emit_table(doc, rows: list[list[str]], *, borders: bool = True) -> None:
    """表格：表级 tblBorders 全 none + 单元格级 tcBorders single（模板的双层写法）。

    只写表级 none 会得到一张完全没线的表；模板视觉上是全框线表。
    """
    ncol = max(len(r) for r in rows)
    table = doc.add_table(rows=len(rows), cols=ncol)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    tblPr = table._tbl.tblPr
    for old in tblPr.findall(qn("w:tblBorders")):
        tblPr.remove(old)
    bd = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        _sub(bd, f"w:{side}", val="none", color="auto", sz="0", space="0")
    tblPr.append(bd)

    # 表宽留 40 twips 余量：正好等于版心时，边框描边与网格舍入会压出右边界
    tbl_total = TEXT_W_TWIPS - 40
    for old in tblPr.findall(qn("w:tblLayout")):
        tblPr.remove(old)
    _sub(tblPr, "w:tblLayout", type="fixed")
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.append(tblW)
    tblW.set(qn("w:w"), str(tbl_total))
    tblW.set(qn("w:type"), "dxa")
    # 单元格边距 左右 108
    for old in tblPr.findall(qn("w:tblCellMar")):
        tblPr.remove(old)
    mar = OxmlElement("w:tblCellMar")
    _sub(mar, "w:top", w="0", type="dxa")
    _sub(mar, "w:left", w="108", type="dxa")
    _sub(mar, "w:bottom", w="0", type="dxa")
    _sub(mar, "w:right", w="108", type="dxa")
    tblPr.append(mar)

    border_val = "single" if borders else "none"
    border_sz = "4" if borders else "0"

    for ri, row in enumerate(rows):
        for ci in range(ncol):
            cell = table.cell(ri, ci)
            par = cell.paragraphs[0]
            par.alignment = WD_ALIGN_PARAGRAPH.CENTER

            # 单元格级边框 + 垂直居中
            tcPr = cell._tc.get_or_add_tcPr()
            for old in tcPr.findall(qn("w:tcBorders")):
                tcPr.remove(old)
            tcb = OxmlElement("w:tcBorders")
            for side in ("top", "left", "bottom", "right"):
                _sub(tcb, f"w:{side}", val=border_val, color="000000",
                     sz=border_sz, space="0")
            tcPr.append(tcb)
            for old in tcPr.findall(qn("w:vAlign")):
                tcPr.remove(old)
            _sub(tcPr, "w:vAlign", val="center")

            for r in list(par.runs):
                r._element.getparent().remove(r._element)
            value = row[ci] if ci < len(row) else ""
            _emit_inline(par, value, size=SZ_TABLE, bold=(ri == 0))


def _zero_wrap_dist(paragraph) -> None:
    """把 wp:inline 的四周环绕距离显式写成 0。

    渲染器（LibreOffice）对缺省的 distL/distR 会套用 0.32cm 的默认环绕距离，
    实测把 14.66cm 的图整体右移 18px，右侧压出版心。显式写 0 才是真正贴版心。
    """
    for inline in paragraph._p.iter(
            "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}inline"):
        for k in ("distT", "distB", "distL", "distR"):
            if inline.get(k) is None:
                inline.set(k, "0")


def _emit_image(doc, path: Path, *, max_w_cm: float = MAX_IMG_W_CM,
                max_h_cm: float = MAX_IMG_H_CM) -> bool:
    if not path.is_file():
        return False
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _no_indent(p)
    run = p.add_run()
    try:
        shape = run.add_picture(str(path), width=Cm(max_w_cm))
    except Exception:
        return False
    # 高度超限则按高度重设
    try:
        h = shape.height.cm
        if h > max_h_cm:
            ratio = max_h_cm / h
            run._element.getparent().remove(run._element)
            run = p.add_run()
            run.add_picture(str(path), height=Cm(max_h_cm))
    except Exception:
        pass
    _zero_wrap_dist(p)
    return True


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------

def build(text: str, base_dir: Path, *, table_borders: bool = True,
          restart_page: bool = True) -> tuple[Document, dict]:
    blocks = parse_markdown(text)
    doc = Document()

    # Normal 样式
    style = doc.styles["Normal"]
    style.font.name = FONT_EN
    style.font.size = Pt(SZ_NORMAL / 2)
    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), "Calibri")
    rFonts.set(qn("w:hAnsi"), "Calibri")
    rFonts.set(qn("w:eastAsia"), FONT_ZH)
    rFonts.set(qn("w:cs"), "Times New Roman")

    pPr = style.element.get_or_add_pPr()
    for old in pPr.findall(qn("w:spacing")):
        pPr.remove(old)
    _sub(pPr, "w:spacing", line="360", lineRule="auto")
    for old in pPr.findall(qn("w:jc")):
        pPr.remove(old)
    _sub(pPr, "w:jc", val="both")
    for old in pPr.findall(qn("w:widowControl")):
        pPr.remove(old)
    widow = OxmlElement("w:widowControl")
    widow.set(qn("w:val"), "0")
    pPr.insert(0, widow)

    _setup_first_section(doc.sections[0])

    stats = {"sections": [], "math": 0, "math_fail": 0, "images": 0,
             "images_missing": 0, "tables": 0}

    current_section = None
    first_section_done = False

    for blk in blocks:
        t = blk["t"]

        if t == "section":
            name = blk["v"]
            stats["sections"].append(name)
            if not first_section_done:
                section = doc.sections[0]
                first_section_done = True
                _setup_first_section(section, restart_page=False)
            else:
                section = doc.add_section(WD_SECTION.NEW_PAGE)
                # 模板口径：第 2 节起每节页码重编为 1
                _setup_first_section(section, restart_page=restart_page)
            _style_header(section, SECTION_TITLES[name])
            _style_footer(section)
            current_section = name
            continue

        if not first_section_done:
            # 首个 ::: 之前的内容丢弃（防止误把说明写进正文）
            continue

        if t == "h1":
            _emit_title(doc, blk["v"])
        elif t == "h2":
            _emit_heading(doc, blk["v"])
        elif t == "h3":
            _emit_heading(doc, blk["v"])
        elif t == "p":
            v = blk["v"].strip()
            if not v:
                continue
            _emit_body(doc, v)
        elif t == "math":
            ok = _emit_math(doc, blk["v"], blk.get("num"))
            if ok:
                stats["math"] += 1
            else:
                stats["math_fail"] += 1
        elif t == "table":
            _emit_table(doc, blk["rows"], borders=table_borders)
            stats["tables"] += 1
        elif t == "img":
            src = blk["src"].strip()
            if src.startswith("http://") or src.startswith("https://"):
                stats["images_missing"] += 1
                continue
            path = (base_dir / src).resolve()
            if _emit_image(doc, path):
                stats["images"] += 1
            else:
                stats["images_missing"] += 1
            # 图题只按节规则生成：摘要附图无图题；说明书附图在图下方
            if current_section == "说明书附图":
                stats["figure_no"] = stats.get("figure_no", 0) + 1
                _emit_caption(doc, f"图{stats['figure_no']}")
            elif current_section != "摘要附图":
                alt = blk["alt"].strip()
                if alt:
                    _emit_caption(doc, alt)
        elif t == "pagebreak":
            doc.add_page_break()

    return doc, stats


def _count_omml(doc) -> int:
    m = "{http://schemas.openxmlformats.org/officeDocument/2006/math}oMath"
    return len(doc.element.body.findall(f".//{m}"))


def _emit_caption(doc, text: str) -> None:
    """图题：居中、宋体四号、不置粗。"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _first_line_indent(p)
    _mark_rpr(p)
    run = p.add_run(text)
    _set_run_font(run, size=SZ_BODY, bold=False)


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="专利 Markdown → 申请文件 Word")
    ap.add_argument("-i", "--input", required=True, type=Path)
    ap.add_argument("-o", "--output", type=Path)
    ap.add_argument("--check", action="store_true", help="只校验，不出 Word")
    ap.add_argument("--no-table-borders", dest="table_borders", action="store_false",
                    help="表格不加框线（默认加框线，与模板一致）")
    ap.add_argument("--continuous-page", dest="restart_page", action="store_false",
                    help="页码全文连续（默认每节从 1 重编，与模板一致）")
    args = ap.parse_args(argv)

    src = args.input.expanduser().resolve()
    if not src.is_file():
        print(f"DOCX: ok=0 reason=missing_input path={src}")
        return 1

    text = src.read_text(encoding="utf-8")
    base_dir = src.parent

    doc, stats = build(text, base_dir, table_borders=args.table_borders,
                       restart_page=args.restart_page)

    missing = [s for s in SECTION_ORDER if s not in stats["sections"]]
    ok = not missing

    print(f"PATENT_DOCX_SECTIONS: {'ok' if ok else 'missing'} "
          f"found={len(stats['sections'])}/{len(SECTION_ORDER)} "
          f"list={','.join(stats['sections']) or '-'}")
    if missing:
        print(f"PATENT_DOCX_WARN: missing_section={','.join(missing)}")
    print(f"PATENT_DOCX_MATH: omml={stats['math']} fallback={stats['math_fail']}")
    print(f"PATENT_DOCX_IMAGES: embedded={stats['images']} missing={stats['images_missing']}")
    print(f"PATENT_DOCX_TABLES: n={stats['tables']}")

    if args.check:
        print(f"PATENT_DOCX: ok={1 if ok else 0} check_only=1")
        return 0 if ok else 1

    if args.output is None:
        print("DOCX: ok=0 reason=missing_output（非 --check 时须给 -o）")
        return 1

    out = args.output.expanduser().resolve()
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out))
    print(f"已写入: {out}")
    print(f"PATENT_DOCX: ok={1 if ok else 0} written=1")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())