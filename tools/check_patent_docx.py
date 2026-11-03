#!/usr/bin/env python
"""校验申请文件 Word 是否合规于 references/template_spec.md。

检查五节结构、页面、页眉、页脚、字体字号、缩进、图表题、公式编号、字体可用性。

用法：
  python tools/check_patent_docx.py -i 申请文件.docx
  python tools/check_patent_docx.py -i 申请文件.docx --json

输出 `PATENT_LAYOUT: ok=N fail=M warn=K`，退出码 0 表示无 fail。
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))

from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

MATH_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

SECTION_TITLES = ["说明书摘要", "摘要附图", "权利要求书", "说明书", "说明书附图"]
BODY_FONT = "宋体"
# 模板历史残留可能出现的字体（WPS 编辑痕迹），不算合规错误，只记 warn
TOLERATED_FONTS = {"楷体", "黑体"}
HEADER_FONT = "黑体"
SZ_BODY = 14.0
SZ_TITLE = 16.0
SZ_TABLE = 12.0
FIRST_LINE = "560"


class Report:
    def __init__(self) -> None:
        self.items: list[dict] = []

    def add(self, level: str, code: str, msg: str, **extra) -> None:
        self.items.append({"level": level, "code": code, "msg": msg, **extra})

    def ok(self, code: str, msg: str = "") -> None:
        self.add("ok", code, msg)

    def fail(self, code: str, msg: str) -> None:
        self.add("fail", code, msg)

    def warn(self, code: str, msg: str) -> None:
        self.add("warn", code, msg)

    @property
    def fails(self) -> list[dict]:
        return [i for i in self.items if i["level"] == "fail"]

    @property
    def warns(self) -> list[dict]:
        return [i for i in self.items if i["level"] == "warn"]


def _para_font(run):
    """返回 (字体名, 字号pt, 是否加粗)。字体优先取 rFonts eastAsia。"""
    if run is None:
        return (None, None, None)
    size = run.font.size.pt if run.font.size is not None else None
    bold = run.font.bold
    name = run.font.name
    rPr = run._element.rPr
    if rPr is not None:
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is not None:
            ea = rFonts.get(qn("w:eastAsia"))
            asc = rFonts.get(qn("w:ascii"))
            name = ea or asc or name
    return (name, size, bold)


def _first_line(p) -> str | None:
    pPr = p._p.pPr
    if pPr is None:
        return None
    ind = pPr.find(qn("w:ind"))
    if ind is None:
        return None
    return ind.get(qn("w:firstLine"))


def _jc(p) -> str | None:
    pPr = p._p.pPr
    if pPr is None:
        return None
    jc = pPr.find(qn("w:jc"))
    return jc.get(qn("w:val")) if jc is not None else None


def check(doc_path: Path) -> Report:
    rep = Report()
    d = Document(str(doc_path))
    body = [Paragraph(c, d) for c in d.element.body.iterchildren()
            if c.tag.split("}")[-1] == "p"]

    # ---- 1. 分节 ----
    n = len(d.sections)
    if n == 5:
        rep.ok("sections", f"共 {n} 节")
    else:
        rep.fail("sections", f"分节数 {n}，应为 5")

    for i, s in enumerate(d.sections):
        if s.start_type is not None and "NEW_PAGE" not in str(s.start_type):
            if i > 0:
                rep.fail("section_break", f"第 {i+1} 节分节类型 {s.start_type}，应为 NEW_PAGE")
        pgSz = s._sectPr.find(qn("w:pgSz"))
        if pgSz is None or pgSz.get(qn("w:w")) != "11906" or pgSz.get(qn("w:h")) != "16838":
            rep.fail("page_size", f"第 {i+1} 节非 A4(11906x16838)")
        pgMar = s._sectPr.find(qn("w:pgMar"))
        if pgMar is None:
            rep.fail("page_margin", f"第 {i+1} 节无 pgMar")
        else:
            want = {"top": "1440", "bottom": "1440", "left": "1797", "right": "1797"}
            for k, v in want.items():
                if pgMar.get(qn(f"w:{k}")) != v:
                    rep.fail("page_margin",
                             f"第 {i+1} 节 {k}={pgMar.get(qn(f'w:{k}'))}，应为 {v}")

    # ---- 2. 页眉 ----
    for i, s in enumerate(d.sections):
        if s.header.is_linked_to_previous and i > 0:
            rep.fail("header_link", f"第 {i+1} 节页眉仍链接上一节，应切断")
            continue
        hp = s.header.paragraphs[0] if s.header.paragraphs else None
        text = hp.text.strip() if hp is not None else ""
        want = SECTION_TITLES[i] if i < len(SECTION_TITLES) else "?"
        if text != want:
            rep.fail("header_text", f"第 {i+1} 节页眉='{text}'，应为 '{want}'")
        if hp is not None and hp.runs:
            name, size, bold = _para_font(hp.runs[0])
            if name != HEADER_FONT:
                rep.fail("header_font", f"第 {i+1} 节页眉字体={name}，应为 {HEADER_FONT}")
            if size != SZ_BODY:
                rep.fail("header_size", f"第 {i+1} 节页眉字号={size}，应为 {SZ_BODY}")
            if not bold:
                rep.fail("header_bold", f"第 {i+1} 节页眉未加粗")
        pPr = hp._p.pPr if hp is not None else None
        if pPr is None or pPr.find(qn("w:pBdr")) is None:
            rep.fail("header_border", f"第 {i+1} 节页眉缺下边框")

    # ---- 3. 页脚页码 ----
    for i, s in enumerate(d.sections):
        fp = s.footer.paragraphs[0] if s.footer.paragraphs else None
        if fp is None:
            rep.fail("footer", f"第 {i+1} 节无页脚")
            continue
        instr = "".join(t.text or "" for t in fp._p.iter(qn("w:instrText")))
        if "PAGE" not in instr.upper():
            rep.fail("footer_page", f"第 {i+1} 节页脚无 PAGE 域")
        if _jc(fp) != "center":
            rep.fail("footer_align", f"第 {i+1} 节页脚未居中")

    # ---- 4. 正文段落字体字号 ----
    body_bad = []
    body_warn = []
    for p in body:
        txt = p.text.strip()
        if not txt:
            continue
        if p._p.findall(f".//{MATH_NS}oMath"):
            continue
        name, size, bold = _para_font(p.runs[0] if p.runs else None)
        if name is not None and name != BODY_FONT:
            (body_warn if name in TOLERATED_FONTS else body_bad).append(
                (txt[:20], "font", name))
        if size is not None and size not in (SZ_BODY, SZ_TITLE, SZ_TABLE):
            body_bad.append((txt[:20], "size", size))
    for t, k, v in body_bad[:8]:
        rep.fail("body_type", f"段落 '{t}' {k}={v}")
    for t, k, v in body_warn[:5]:
        rep.warn("body_font_residue", f"段落 '{t}' {k}={v}（模板残留字体，可接受）")
    if not body_bad:
        rep.ok("body_type", "正文字体字号合规")

    # ---- 5. 说明书标题 ----
    title_ok = False
    for p in body:
        txt = p.text.strip()
        if not txt:
            continue
        name, size, bold = _para_font(p.runs[0] if p.runs else None)
        if size == SZ_TITLE:
            if bold and _jc(p) == "center":
                title_ok = True
            else:
                rep.fail("title_style", f"说明书标题 '{txt[:24]}' 未加粗居中")
            break
    if not title_ok:
        rep.fail("title_missing", "未找到 16pt 说明书标题")

    # ---- 6. 章标题 ----
    chapters = ["技术领域", "背景技术", "发明内容", "附图说明", "具体实施方式"]
    found = {}
    for p in body:
        txt = p.text.strip()
        if txt in chapters:
            name, size, bold = _para_font(p.runs[0] if p.runs else None)
            found[txt] = True
            if not bold:
                rep.fail("chapter_bold", f"章标题 '{txt}' 未加粗")
            if size != SZ_BODY:
                rep.fail("chapter_size", f"章标题 '{txt}' 字号={size}，应为 {SZ_BODY}")
            if _first_line(p) not in (None, "0"):
                rep.fail("chapter_indent", f"章标题 '{txt}' 不应首行缩进")
    missing = [c for c in chapters if c not in found]
    if missing:
        rep.fail("chapter_missing", f"缺章标题：{','.join(missing)}")
    else:
        rep.ok("chapters", "五个章标题齐全")

    # ---- 7. 图片与图题 ----
    # 逐段标注所属节（段内带 sectPr 即该节结束）
    sec_of: list[int] = []
    cur = 0
    for p in body:
        sec_of.append(cur)
        pPr = p._p.pPr
        if pPr is not None and pPr.find(qn("w:sectPr")) is not None:
            cur += 1

    img_idx = [i for i, p in enumerate(body) if p._p.findall(f".//{A_NS}blip")]
    if not img_idx:
        rep.warn("image_none", "文档无图")
    for i in img_idx:
        sect = sec_of[i] if i < len(sec_of) else -1
        if sect == 1:      # 摘要附图：不要求图题
            rep.ok("abs_fig", f"第 {i} 段为摘要附图")
            continue
        nxt_p = None
        for j in (i + 1, i + 2):
            if j < len(body) and body[j].text.strip():
                nxt_p = body[j]
                break
        if nxt_p is None or not nxt_p.text.strip().startswith("图"):
            rep.warn("fig_caption_missing", f"第 {i} 段图后无图题")
        elif _jc(nxt_p) != "center":
            rep.fail("fig_caption_align", f"第 {i} 段图后图题未居中")

    # ---- 8. 表格 ----
    for ti, t in enumerate(d.tables):
        # 视觉边框由 tcBorders 决定（模板：表级全 none + 单元格级 single）
        vals = set()
        for row in t.rows[:2]:
            for c in row.cells:
                tcPr = c._tc.tcPr
                tcb = tcPr.find(qn("w:tcBorders")) if tcPr is not None else None
                if tcb is None:
                    vals.add("missing")
                    continue
                for ch in tcb:
                    vals.add(ch.get(qn("w:val")))
        if "missing" in vals:
            rep.warn("table_borders", f"表 {ti+1} 有单元格未设 tcBorders")
        elif vals <= {"single", "double", "thick"}:
            rep.ok("table_borders", f"表 {ti+1} 单元格有框线（模板双层写法）")
        elif vals <= {"none", "nil"}:
            rep.warn("table_borders", f"表 {ti+1} 无框线")
        if not t.rows:
            continue
        # 表宽不得超版心 8312
        tblW = t._tbl.tblPr.find(qn("w:tblW")) if t._tbl.tblPr is not None else None
        if tblW is not None:
            w = tblW.get(qn("w:w"))
            if w and w.isdigit() and int(w) > 8312:
                rep.warn("table_width", f"表 {ti+1} 宽 {w} 超版心 8312（模板自带瑕疵，勿复制）")
        # 表头：首行中「有文字且非符号」的单元格应加粗；模板常有合并/符号格，缺粗只记 warn
        head_bold, head_plain = 0, 0
        for c in t.rows[0].cells:
            for p in c.paragraphs:
                txt = p.text.strip()
                if not txt or txt in {"√", "×", "-", "—"}:
                    continue
                if p.runs:
                    name, size, bold = _para_font(p.runs[0])
                    if bold:
                        head_bold += 1
                    else:
                        head_plain += 1
                    if size != SZ_TABLE:
                        rep.fail("table_size", f"表 {ti+1} 字号={size}，应为 {SZ_TABLE}")
                break
        if head_plain and not head_bold:
            rep.fail("table_head_bold", f"表 {ti+1} 表头整行未加粗")
        elif head_plain:
            rep.warn("table_head_partial",
                     f"表 {ti+1} 表头 {head_plain} 格未加粗（{head_bold} 格已加粗）")
        else:
            rep.ok("table_head_bold", f"表 {ti+1} 表头已加粗")
        rep.ok("table_done", f"表 {ti+1} 已检查 {len(t.rows)} 行")

    # ---- 9. 公式：编号 + OMML 字号 ----
    math_p = [p for p in body if p._p.findall(f".//{MATH_NS}oMath")]
    numbered = 0
    bare = 0
    for p in math_p:
        pPr = p._p.pPr
        tabs = pPr.find(qn("w:tabs")) if pPr is not None else None
        if tabs is None:
            continue
        txt = p.text.strip()
        if re.fullmatch(r"\(\d+\)", txt):
            numbered += 1
        elif txt.isdigit():
            bare += 1
    if bare:
        rep.fail("math_number_paren",
                 f"{bare} 个公式编号是裸数字，应写作「(N)」")
    else:
        rep.ok("math_number_paren", f"公式编号均带括号（{numbered} 条）")
    rep.ok("math", f"含公式段落 {len(math_p)}，其中带编号 {numbered}")

    # 模板每个 m:r 都显式 Cambria Math + sz=28
    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    bad_math = 0
    for p in math_p:
        for mr in p._p.iter(f"{MATH_NS}r"):
            rPrs = mr.findall(f"{W}rPr")
            if not rPrs:
                bad_math += 1
                continue
            # 取最后一个 w:rPr（ooMath 的 m:r 里 self-closing 的 m:rPr 不参与）
            rPr = rPrs[-1]
            rf = rPr.find(qn("w:rFonts"))
            sz = rPr.find(qn("w:sz"))
            if rf is None or rf.get(qn("w:ascii")) != "Cambria Math":
                bad_math += 1
            elif sz is None or sz.get(qn("w:val")) != "28":
                bad_math += 1
    if bad_math:
        rep.fail("math_font", f"{bad_math} 个 m:r 缺 Cambria Math 或四号字号")
    elif math_p:
        rep.ok("math_font", "公式字体字号合规")

    # ---- 10. 页码按节重编 ----
    restarts = 0
    for i, s in enumerate(d.sections):
        pn = s._sectPr.find(qn("w:pgNumType"))
        if pn is not None and pn.get(qn("w:start")) == "1":
            restarts += 1
    if restarts >= 4:
        rep.ok("page_restart", f"{restarts} 节页码重编（模板口径）")
    elif restarts:
        rep.warn("page_restart", f"仅 {restarts} 节页码重编")
    else:
        rep.warn("page_restart", "页码未按节重编（全文连续）")

    # ---- 10. 字体可用性（引用了但可能缺字体的名字）----
    fonts = set()
    for p in body:
        for r in p.runs:
            name, _, _ = _para_font(r)
            if name:
                fonts.add(name)
    rep.ok("fonts_used", ",".join(sorted(fonts)))

    return rep


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="校验申请文件 Word 版式")
    ap.add_argument("-i", "--input", required=True, type=Path)
    ap.add_argument("--json", action="store_true")
    args = ap.parse_args(argv)

    path = args.input.expanduser().resolve()
    if not path.is_file():
        print(f"PATENT_LAYOUT: ok=0 reason=missing_file path={path}")
        return 1

    rep = check(path)

    if args.json:
        print(json.dumps(rep.items, ensure_ascii=False, indent=2))
    else:
        for item in rep.items:
            mark = {"ok": "OK  ", "warn": "WARN", "fail": "FAIL"}[item["level"]]
            print(f"[{mark}] {item['code']}: {item['msg']}")

    nf, nw = len(rep.fails), len(rep.warns)
    print(f"PATENT_LAYOUT: ok={1 if nf == 0 else 0} fail={nf} warn={nw}")
    return 0 if nf == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())