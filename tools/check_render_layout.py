# -*- coding: utf-8 -*-
"""渲染后版面体检：逐页检查墨迹是否越出版心、页眉页脚是否齐全、有无空页。

XML 级校验看不到「渲染器把内容画到哪儿」，本工具补这一层。
需要先用 LibreOffice Kit 把 docx 渲染成 PNG：

    node <cli> render --input 稿.docx --output-dir <新目录>

然后：

    python tools/check_render_layout.py --render-dir <新目录>

也可 --docx 直接给 docx，工具会尝试调用 DSH 自带的 LibreOffice Kit。
"""
from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))
try:
    from stdio_utf8 import ensure_utf8_stdio
    ensure_utf8_stdio()
except Exception:
    pass

# A4 @144dpi 下的版面常量（与 emit_patent_docx 的页面设置一致）
DPI = 144
TW = DPI / 1440.0          # twips -> px
PAGE_W_TWIPS = 11906
PAGE_H_TWIPS = 16838
MARGIN_TB = 1440
MARGIN_LR = 1797
PAGE_W = int(round(PAGE_W_TWIPS * TW))      # 1191
PAGE_H = int(round(PAGE_H_TWIPS * TW))      # 1684
M_L = int(MARGIN_LR * TW)                   # 179
M_R = PAGE_W - M_L                          # 1012
M_T = int(MARGIN_TB * TW)                   # 144
M_B = PAGE_H - M_T                          # 1540
TOL = 6                                     # 允许的像素容差

INK_THR = 200


def _px():
    try:
        from PIL import Image  # noqa: F401
    except ImportError:
        print("RENDER_LAYOUT: ok=0 reason=pillow_missing")
        raise SystemExit(1)


def _find_node() -> str | None:
    """找 node：PATH 优先，其次 DSH 运行时自带的那份。"""
    import os
    for env in ("DSH_NODE", "NODE_EXE"):
        v = os.environ.get(env)
        if v and Path(v).is_file():
            return v
    n = shutil.which("node")
    if n:
        return n
    for cand in (
        Path(r"D:\work\deepseekherness\resources\runtime\primary-runtime"
             r"\dependencies\node\bin\node.exe"),
    ):
        if cand.is_file():
            return str(cand)
    return None


def _find_cli() -> Path | None:
    import os
    v = os.environ.get("DSH_LIBREOFFICE_CLI")
    cands = [Path(v)] if v else []
    cands.append(Path(
        r"D:\work\deepseekherness\resources\app.asar.unpacked\dsh"
        r"\node_modules\@deepseek-ai\libreoffice-kit\lib\cli.js"))
    for c in cands:
        if c.is_file():
            return c
    return None


def render_dir_for(docx: Path) -> Path | None:
    """用 LibreOffice Kit 渲染；找不到渲染器就返回 None。"""
    node = _find_node()
    cli = _find_cli()
    if not node or cli is None:
        return None
    out = Path(tempfile.mkdtemp(prefix="renderchk-"))
    # Kit 要求 output-dir 尚不存在，用它返回的临时目录再套一层
    target = out / "pages"
    try:
        r = subprocess.run(
            [node, str(cli), "render", "--input", str(docx),
             "--output-dir", str(target)],
            capture_output=True, text=True, timeout=900,
            encoding="utf-8", errors="replace",
        )
    except Exception:
        return None
    if r.returncode != 0 or not any(target.glob("page-*.png")):
        return None
    return target


def analyze_page(path: Path) -> dict:
    from PIL import Image
    im = Image.open(path).convert("L")
    W, H = im.size
    px = im.load()
    res = {"page": path.name, "w": W, "h": H}

    # 正文区墨迹包围盒
    l = r = t = b = None
    for y in range(M_T, min(M_B, H)):
        xs = [x for x in range(W) if px[x, y] < INK_THR]
        if not xs:
            continue
        if t is None:
            t = y
        b = y
        x0, x1 = xs[0], xs[-1]
        l = x0 if l is None else min(l, x0)
        r = x1 if r is None else max(r, x1)
    res["body"] = (l, t, r, b)
    res["over_l"] = bool(l is not None and l < M_L - TOL)
    res["over_r"] = bool(r is not None and r > M_R + TOL)
    res["over_b"] = bool(b is not None and b > M_B + TOL)
    # 越界量（cm）
    res["over_r_cm"] = round((r - M_R) / (DPI / 2.54), 2) if r else 0.0
    res["over_l_cm"] = round((M_L - l) / (DPI / 2.54), 2) if l else 0.0

    # 页眉：页眉区有墨 + 底部有横线
    hdr_ink = False
    for y in range(0, min(M_T, H)):
        if any(px[x, y] < INK_THR for x in range(0, W, 2)):
            hdr_ink = True
            break
    res["header"] = hdr_ink
    rule = False
    for y in range(max(0, M_T - 26), min(H, M_T + 6)):
        dark = sum(1 for x in range(0, W, 4) if px[x, y] < 160)
        if dark > (W / 4) * 0.55:
            rule = True
            break
    res["header_rule"] = rule

    # 页脚
    ftr = False
    for y in range(min(M_B, H), H):
        if any(px[x, y] < INK_THR for x in range(0, W, 2)):
            ftr = True
            break
    res["footer"] = ftr

    # 空页
    res["ink_px"] = sum(
        1 for y in range(M_T, min(M_B, H))
        for x in range(0, W, 2) if px[x, y] < INK_THR
    )
    res["blank"] = res["ink_px"] < 200
    return res


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="渲染后版面越界体检")
    ap.add_argument("--render-dir", type=Path, help="已渲染好的 PNG 目录")
    ap.add_argument("--docx", type=Path, help="直接给 docx，工具自行渲染")
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--keep", action="store_true", help="保留临时渲染目录")
    ap.add_argument("--allow-skip", action="store_true",
                    help="渲染器不可用时返回通过（默认算失败，防止闸门空转）")
    args = ap.parse_args(argv)

    _px()

    d = args.render_dir
    tmp = None
    if d is None:
        if not args.docx:
            print("RENDER_LAYOUT: ok=0 reason=need_render_dir_or_docx")
            return 1
        tmp = render_dir_for(args.docx)
        if tmp is None:
            if args.allow_skip:
                print("RENDER_LAYOUT: ok=1 issues=0 "
                      "reason=renderer_unavailable (已 --allow-skip)")
                return 0
            print("RENDER_LAYOUT: ok=0 reason=renderer_unavailable "
                  "(设 DSH_NODE / DSH_LIBREOFFICE_CLI，或先手工渲染再传 "
                  "--render-dir；确实无渲染器才加 --allow-skip)")
            return 1
        d = tmp
    pages = sorted(Path(d).glob("page-*.png"))
    if not pages:
        print("RENDER_LAYOUT: ok=0 reason=no_pages")
        if tmp and not args.keep:
            shutil.rmtree(tmp, ignore_errors=True)
        return 1

    rows = [analyze_page(p) for p in pages]
    if tmp and not args.keep:
        shutil.rmtree(tmp, ignore_errors=True)

    fails = []
    for r in rows:
        if r["over_l"]:
            fails.append((r["page"], "over_left", f"{r['over_l_cm']}cm"))
        if r["over_r"]:
            fails.append((r["page"], "over_right", f"{r['over_r_cm']}cm"))
        if r["over_b"]:
            fails.append((r["page"], "over_bottom", ""))
        if not r["header"]:
            fails.append((r["page"], "no_header", ""))
        if not r["header_rule"]:
            fails.append((r["page"], "no_header_rule", ""))
        if not r["footer"]:
            fails.append((r["page"], "no_footer", ""))
        if r["blank"]:
            fails.append((r["page"], "blank_page", ""))

    if args.json:
        print(json.dumps({"pages": len(rows), "fails": fails,
                          "rows": rows}, ensure_ascii=False, indent=2))
    else:
        print(f"页数 {len(rows)}  版心 x[{M_L},{M_R}] y[{M_T},{M_B}] 像素")
        seen = {}
        for page, kind, extra in fails:
            seen.setdefault(kind, []).append((page, extra))
        names = {
            "over_left": "左越界", "over_right": "右越界",
            "over_bottom": "下越界", "no_header": "缺页眉",
            "no_header_rule": "缺页眉下横线", "no_footer": "缺页脚",
            "blank_page": "空页",
        }
        if not fails:
            print("[OK  ] 全部页面均在版心内，页眉/横线/页脚齐全")
        for kind, items in seen.items():
            detail = ", ".join(
                f"{p[-8:-4]}({e})" if e else p[-8:-4] for p, e in items
            )
            print(f"[FAIL] {names.get(kind, kind)}: {len(items)} 页 {detail}")

    print(f"RENDER_LAYOUT: ok={0 if fails else 1} pages={len(rows)} "
          f"issues={len(fails)}")
    return 1 if fails else 0


if __name__ == "__main__":
    raise SystemExit(main())