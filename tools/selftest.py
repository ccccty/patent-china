#!/usr/bin/env python
"""一键回归：把 tests/smoke.md 走完整链路，四道闸全过才算通过。

用法：
  python tools/selftest.py
  python tools/selftest.py --keep     # 保留产物便于排查

输出 `SELFTEST: ok=N fail=M`，退出码 0 为通过。
"""
from __future__ import annotations

import argparse
import re
import subprocess
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))

try:  # 控制台可能是 GBK，统一转 UTF-8，否则中文日志会炸
    from stdio_utf8 import ensure_utf8_stdio
    ensure_utf8_stdio()
except Exception:
    for _s in (sys.stdout, sys.stderr):
        try:
            _s.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

SKILL = HERE.parent
TESTS = SKILL / "tests"
SMOKE = TESTS / "smoke.md"

EXPECT = {
    "PATENT_DOCX_SECTIONS": r"found=5/5",
    "PATENT_DOCX_MATH": r"fallback=0",
    "PATENT_DOCX_IMAGES": r"missing=0",
    "PATENT_DOCX": r"ok=1",
    "PATENT_LAYOUT": r"ok=1",
    "CLEAN_BODY": r"ok=1",
    "CONSISTENCY": r"ok=1",
}


def _run(args: list[str]) -> tuple[int, str]:
    r = subprocess.run([sys.executable, *args], capture_output=True, text=True,
                       encoding="utf-8", errors="replace")
    return r.returncode, (r.stdout or "") + (r.stderr or "")


def _render_and_check(docx: Path) -> tuple[int, str]:
    """渲染后查版面越界。渲染器不可用时显式 SKIP，不伪装成通过。"""
    rc, out = _run([str(HERE / "check_render_layout.py"),
                    "--docx", str(docx), "--allow-skip"])
    if "reason=renderer_unavailable" in out:
        return 2, out
    return rc, out


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="技能自测")
    ap.add_argument("--keep", action="store_true")
    args = ap.parse_args(argv)

    if not SMOKE.is_file():
        print(f"SELFTEST: ok=0 reason=missing_fixture path={SMOKE}")
        return 1

    ok = fail = 0
    tmpdir = Path(tempfile.mkdtemp(prefix="patent_china_"))
    out_docx = tmpdir / "smoke.docx"

    steps = [
        ("生成 Word", [str(HERE / "emit_patent_docx.py"),
                       "-i", str(SMOKE), "-o", str(out_docx)]),
        ("版式校验", [str(HERE / "check_patent_docx.py"), "-i", str(out_docx)]),
        ("正文体检", [str(HERE / "check_clean_body.py"), "-i", str(SMOKE)]),
        ("一致性", [str(HERE / "check_consistency.py"), "-i", str(SMOKE)]),
    ]

    all_out: dict[str, str] = {}
    for label, cmd in steps:
        code, out = _run(cmd)
        all_out[label] = out
        if code != 0:
            print(f"[FAIL] {label} exit={code}")
            print(out.rstrip())
            fail += 1
        else:
            print(f"[OK  ] {label}")
            ok += 1

    # 渲染级体检（XML 查不出的越界）
    code, out = _render_and_check(out_docx)
    all_out["渲染体检"] = out
    if code == 2:
        print("[SKIP] 渲染体检（本机无 LibreOffice Kit，越界未验证）")
    elif code != 0:
        print(f"[FAIL] 渲染体检 exit={code}")
        print(out.rstrip())
        fail += 1
    else:
        print("[OK  ] 渲染体检")
        ok += 1

    # 逐条核对关键信号
    print("\n---- 信号核对 ----")
    blob = "\n".join(all_out.values())
    for key, pat in EXPECT.items():
        line = next((ln for ln in blob.splitlines() if ln.startswith(key + ":")), "")
        hit = bool(re.search(pat, line))
        print(f"[{'OK  ' if hit else 'FAIL'}] {key}: {line.strip() or '(未出现)'}")
        ok, fail = (ok + 1, fail) if hit else (ok, fail + 1)

    print(f"\nSELFTEST: ok={1 if fail == 0 else 0} ok_n={ok} fail={fail}")
    if not args.keep:
        for p in tmpdir.glob("*"):
            try:
                p.unlink()
            except OSError:
                pass
        try:
            tmpdir.rmdir()
        except OSError:
            pass
    else:
        print(f"产物保留在: {tmpdir}")
    return 0 if fail == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())