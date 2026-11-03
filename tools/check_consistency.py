# -*- coding: utf-8 -*-
"""一致性闸：附图说明条目数 / 图数 / 如图N引用 / 公式编号引用 是否自洽。"""
from __future__ import annotations
import argparse
import re
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))
try:
    from stdio_utf8 import ensure_utf8_stdio
    ensure_utf8_stdio()
except Exception:
    pass


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="申请文件内在一致性校验")
    ap.add_argument("-i", "--input", required=True, type=Path)
    args = ap.parse_args(argv)

    f = args.input.expanduser().resolve()
    if not f.is_file():
        print(f"CONSISTENCY: ok=0 reason=missing_file")
        return 1
    text = f.read_text(encoding="utf-8").replace("\ufeff", "")

    ok = fail = 0

    def rep(passed: bool, name: str, msg: str):
        nonlocal ok, fail
        print(f"[{'OK  ' if passed else 'FAIL'}] {name}: {msg}")
        if passed:
            ok += 1
        else:
            fail += 1

    # 拆节
    parts = re.split(r"^:::\s*(摘要|摘要附图|权利要求书|说明书|说明书附图)\s*$",
                     text, flags=re.MULTILINE)
    sec: dict[str, str] = {}
    for i in range(1, len(parts) - 1, 2):
        sec[parts[i]] = parts[i + 1]

    spec = sec.get("说明书", "")
    figs_sec = sec.get("说明书附图", "")

    # 1) 附图说明条目
    cap_block = ""
    m = re.search(r"##\s*附图说明(.*?)(?=##\s|\Z)", spec, re.DOTALL)
    if m:
        cap_block = m.group(1)
    listed = sorted({int(x) for x in re.findall(r"图\s*(\d+)\s*[为是]", cap_block)})
    rep(bool(listed), "figure_list", f"附图说明列了 {len(listed)} 条 {listed}")

    # 2) 实际图片
    imgs = re.findall(r"!\[[^\]]*\]\(([^)]+)\)", figs_sec)
    rep(len(imgs) == len(listed),
        "figure_count_match",
        f"说明书附图 {len(imgs)} 张 vs 附图说明 {len(listed)} 条")

    # 3) 如图N引用必须存在
    cited = {int(x) for x in re.findall(r"如图\s*(\d+)\s*所示", spec)}
    missing = sorted(cited - set(listed))
    rep(not missing, "figure_ref_valid",
        f"引用 {sorted(cited)}，未定义 {missing}" if missing else f"引用 {sorted(cited)} 全部有定义")

    # 4) 图号连续
    cont = listed == list(range(1, len(listed) + 1))
    rep(cont, "figure_numbering", f"编号 {listed} {'连续' if cont else '不连续'}")

    # 5) 公式编号 vs 式 (N) 引用
    eqs = sorted({int(x) for x in re.findall(r"\$\$\s*\((\d+)\)", text)})
    if not eqs:
        eqs = sorted({int(x) for x in re.findall(r"^\s*\((\d+)\)\s*$", text, re.MULTILINE)})
    refs = sorted({int(x) for x in re.findall(r"[式如]\s*\((\d+)\)", spec)})
    bad_refs = sorted(set(refs) - set(eqs))
    if eqs:
        rep(not bad_refs, "equation_ref_valid",
            f"公式 {eqs}，引用 {refs}，悬空引用 {bad_refs}" if bad_refs
            else f"公式 {len(eqs)} 条，引用 {refs} 均有对应")
        rep(eqs == list(range(1, len(eqs) + 1)), "equation_numbering",
            f"编号 {eqs} {'连续' if eqs == list(range(1, len(eqs) + 1)) else '不连续'}")
    else:
        print("[----] equation: 无编号公式，跳过")

    # 6) 权项与说明书步骤数一致
    claims = sec.get("权利要求书", "")
    s_claim = sorted({int(x) for x in re.findall(r"^S(\d+)、", claims, re.MULTILINE)})
    steps = sorted({int(x) for x in re.findall(r"步骤\s*(\d+)\s*[：:]", spec)})
    if s_claim and steps:
        rep(s_claim == steps, "step_sync",
            f"权项 S{s_claim} vs 实施例 步骤{steps} {'一致' if s_claim == steps else '不一致'}")
    elif s_claim:
        print(f"[----] step_sync: 权项 S{s_claim}，说明书未用「步骤N：」，跳过")

    # 7) 摘要长度
    abs_txt = sec.get("摘要", "").strip()
    n = len(re.sub(r"\s", "", abs_txt))
    rep(0 < n <= 300, "abstract_len", f"摘要 {n} 字（上限 300）")

    print(f"CONSISTENCY: ok={1 if fail == 0 else 0} ok_n={ok} fail={fail}")
    return 0 if fail == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())