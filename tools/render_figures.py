#!/usr/bin/env python
"""附图渲染：mermaid 围栏 → 黑白 PNG（申请文件用）。

用法：
  # 从含 mermaid 围栏的 .md 批量出图
  python tools/render_figures.py -i figures_src.md --out-dir assets --blackwhite

  # 单块 mermaid 直出
  python tools/render_figures.py --block flow.mmd --out assets/fig1.png --blackwhite

  # 只把已有图片转黑白（材料图）
  python tools/render_figures.py --binarize assets/fig3.png --out assets/fig3_bw.png

依赖：playwright（本机已装）、tools/vendor/mermaid.min.js。

输出信号：
  FIG_RENDER: ok=N fail=M
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))

VENDOR_JS = HERE / "vendor" / "mermaid.min.js"

# 黑白主题：线条全黑、填充全白，符合申请附图要求
BW_INIT = (
    "%%{init: {'theme':'base', 'themeVariables': {"
    "'primaryColor':'#ffffff','primaryTextColor':'#000000',"
    "'primaryBorderColor':'#000000','lineColor':'#000000',"
    "'secondaryColor':'#ffffff','tertiaryColor':'#ffffff',"
    "'fontFamily':'SimSun, serif','fontSize':'16px'}}}%%"
)

HTML_TMPL = """<!doctype html>
<html><head><meta charset="utf-8">
<style>
  html, body {{ margin:0; padding:0; background:#fff; }}
  #c {{ display:inline-block; padding:10px; background:#fff; }}
  /* 不强制宽度：一强制就会按 viewBox 等比放大成巨图 */
</style>
<script src="file:///{vendor}"></script>
</head><body>
<div id="c"><pre class="mermaid">{code}</pre></div>
<script>
  mermaid.initialize({{ startOnLoad: true, securityLevel: 'loose' }});
</script>
</body></html>
"""

RE_FENCE = re.compile(r"^```mermaid\s*$(.*?)^```\s*$", re.MULTILINE | re.DOTALL)


def _ensure_init(code: str) -> str:
    """确保代码块首行有黑白 init 指令。"""
    stripped = code.lstrip()
    if stripped.startswith("%%{init"):
        return code
    return BW_INIT + "\n" + code


def render_blocks(md_text: str) -> list[str]:
    return [m.group(1) for m in RE_FENCE.finditer(md_text)]


def render_one(code: str, out_png: Path, *, scale: int = 2,
               width: int = 1400) -> bool:
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("FIG_RENDER: ok=0 reason=playwright_missing", file=sys.stderr)
        return False

    if not VENDOR_JS.is_file():
        print(f"FIG_RENDER: ok=0 reason=missing_vendor path={VENDOR_JS}", file=sys.stderr)
        return False

    import tempfile
    code = _ensure_init(code)
    html = HTML_TMPL.format(vendor=VENDOR_JS.as_posix(), code=code)

    out_png.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as td:
        html_path = Path(td) / "fig.html"
        html_path.write_text(html, encoding="utf-8")
        try:
            with sync_playwright() as pw:
                try:
                    browser = pw.chromium.launch(channel="chrome")
                except Exception:
                    browser = pw.chromium.launch()
                page = browser.new_page(
                    viewport={"width": width, "height": 900},
                    device_scale_factor=scale,
                )
                page.goto(html_path.as_uri())
                page.wait_for_timeout(900)
                el = page.query_selector("#c svg")
                if el is None:
                    browser.close()
                    print("FIG_RENDER: ok=0 reason=no_svg", file=sys.stderr)
                    return False
                el.screenshot(path=str(out_png))
                browser.close()
        except Exception as exc:
            print(f"FIG_RENDER: ok=0 reason={exc}", file=sys.stderr)
            return False
    return out_png.is_file()


def binarize(src: Path, dst: Path, *, threshold: int = 200,
             force: bool = False) -> tuple[bool, str]:
    """灰度 → 二值化，线条压黑、背景提白。

    自动保护：若原图本身带灰底（大面积浅灰），硬二值化会把背景压成黑块。
    此时如未 force，则退化为「提白背景 + 保留灰阶」，避免出脏图。
    返回 (是否成功, 说明)。
    """
    try:
        from PIL import Image
    except ImportError:
        print("FIG_RENDER: ok=0 reason=pillow_missing", file=sys.stderr)
        return False, "pillow_missing"

    im = Image.open(src).convert("RGBA")
    bg = Image.new("RGBA", im.size, (255, 255, 255, 255))
    flat = Image.alpha_composite(bg, im).convert("L")
    hist = flat.histogram()
    total = sum(hist) or 1
    dark_before = sum(hist[:32]) / total
    # 浅灰底：介于 190~250 之间且占比可观
    pale = sum(hist[190:251]) / total

    bw = flat.point(lambda p: 255 if p > threshold else 0, "L")
    bw_hist = bw.histogram()
    dark_after = sum(bw_hist[:32]) / total

    # 判定「脏二值化」：黑占比的绝对涨幅太大，说明浅色底被压成了黑块。
    # 注意只看比值会误判：线框细的图从 0.3% 涨到 4.5% 是正常的。
    dirty = pale > 0.05 and (dark_after - dark_before) > 0.10
    if dirty and not force:
        # 提白：把 250 以上的拉纯白，其余保留灰阶
        soft = flat.point(lambda p: 255 if p >= 248 else max(0, int(p * 0.85)), "L")
        dst.parent.mkdir(parents=True, exist_ok=True)
        soft.save(dst)
        return True, (f"gray_fallback pale={pale:.1%} "
                      f"dark {dark_before:.1%}->{dark_after:.1%}，"
                      f"改存灰阶并提白（如需硬二值化加 --force-bw）")

    dst.parent.mkdir(parents=True, exist_ok=True)
    bw.save(dst)
    return True, f"binary dark={dark_after:.1%}"


def trim(src: Path, dst: Path, *, pad: int = 12) -> bool:
    """按内容包围盒紧裁，去掉大片空白。"""
    try:
        from PIL import Image, ImageChops
    except ImportError:
        return False
    im = Image.open(src).convert("L")
    inv = ImageChops.invert(im)
    bbox = inv.getbbox()
    if not bbox:
        return False
    l, t, r, b = bbox
    l, t = max(0, l - pad), max(0, t - pad)
    r, b = min(im.width, r + pad), min(im.height, b + pad)
    im.crop((l, t, r, b)).save(dst)
    return True


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="附图渲染（mermaid → 黑白 PNG）")
    ap.add_argument("-i", "--input", type=Path, help="含 mermaid 围栏的 .md")
    ap.add_argument("--block", type=Path, help="单个 .mmd 文件")
    ap.add_argument("--binarize", type=Path, help="把该图片转黑白")
    ap.add_argument("--out", type=Path, help="输出 PNG（--block/--binarize 用）")
    ap.add_argument("--out-dir", type=Path, help="批量出图的目录")
    ap.add_argument("--prefix", default="fig", help="批量命名前缀（默认 fig）")
    ap.add_argument("--blackwhite", action="store_true", help="出图后二值化")
    ap.add_argument("--force-bw", action="store_true",
                    help="强制硬二值化（原图带灰底时会把背景压黑，慎用）")
    ap.add_argument("--no-trim", action="store_true", help="不紧裁")
    ap.add_argument("--scale", type=int, default=2)
    ap.add_argument("--width", type=int, default=1400)
    args = ap.parse_args(argv)

    ok = fail = 0

    if args.binarize:
        out = args.out or args.binarize.with_name(args.binarize.stem + "_bw.png")
        ok_bw, why = binarize(args.binarize, out, force=args.force_bw)
        if ok_bw:
            print(f"FIG_RENDER: ok=1 path={out} note={why}")
            return 0
        print("FIG_RENDER: ok=0")
        return 1

    if args.block:
        if not args.out:
            print("FIG_RENDER: ok=0 reason=missing_out")
            return 1
        code = args.block.read_text(encoding="utf-8")
        tmp = args.out.with_name(args.out.stem + "_raw.png")
        if render_one(code, tmp, scale=args.scale, width=args.width):
            final = args.out
            if args.blackwhite:
                bw = args.out.with_name(args.out.stem + "_bw.png")
                ok_bw, why = binarize(tmp, bw, force=args.force_bw)
                if not ok_bw:
                    print("FIG_RENDER: ok=0 reason=binarize_failed")
                    return 1
                if why.startswith("gray_fallback"):
                    print(f"FIG_RENDER_NOTE: {why}")
                tmp = bw
            if not args.no_trim:
                trim(tmp, final)
            else:
                trim(tmp, final) if tmp != final else None
            print(f"FIG_RENDER: ok=1 path={final}")
            return 0
        print("FIG_RENDER: ok=0")
        return 1

    if not args.input:
        print("FIG_RENDER: ok=0 reason=missing_input（或用 --block / --binarize）")
        return 1

    md = args.input.expanduser().resolve()
    if not md.is_file():
        print(f"FIG_RENDER: ok=0 reason=missing_file path={md}")
        return 1
    out_dir = args.out_dir or (md.parent / "assets")

    blocks = render_blocks(md.read_text(encoding="utf-8"))
    if not blocks:
        print("FIG_RENDER: ok=0 reason=no_mermaid_block")
        return 1

    for idx, code in enumerate(blocks, 1):
        raw = out_dir / f"{args.prefix}{idx}_raw.png"
        if not render_one(code, raw, scale=args.scale, width=args.width):
            fail += 1
            continue
        cur = raw
        if args.blackwhite:
            bw = out_dir / f"{args.prefix}{idx}_bw.png"
            ok_bw, why = binarize(cur, bw, force=args.force_bw)
            if ok_bw:
                if why.startswith("gray_fallback"):
                    print(f"FIG_RENDER_NOTE: fig{idx} {why}")
                cur = bw
        final = out_dir / f"{args.prefix}{idx}.png"
        if not args.no_trim:
            trim(cur, final)
        else:
            final.write_bytes(cur.read_bytes())
        print(f"FIG_RENDER: ok=1 path={final}")
        ok += 1
        # 清理中间件
        for p in (raw, out_dir / f"{args.prefix}{idx}_bw.png"):
            if p.exists() and p != final:
                try:
                    p.unlink()
                except OSError:
                    pass

    print(f"FIG_RENDER: ok={1 if fail == 0 else 0} ok_n={ok} fail={fail}")
    return 0 if fail == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())