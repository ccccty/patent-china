#!/usr/bin/env python
"""把分段稿合并成单份申请文件 Markdown（带五节标记）。

输入目录需含（有则合并，无则跳过并告警）：
  摘要.md  摘要附图.md  权利要求书.md  说明书.md  说明书附图.md
       或
  说明书.md 内含全部节（此时直接补节标记）

输出：--out 指定的单份 Markdown，文件名形如 {案件名}_{YYYYMMDDHHmmss}.md

用法：
  python tools/assemble_application.py --dir <案件目录> --name "一种…方法"
  python tools/assemble_application.py --dir <案件目录> --name "…" --check
"""
from __future__ import annotations

import argparse
import re
import sys
from datetime import datetime
from pathlib import Path

HERE = Path(__file__).resolve().parent
if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))

SECTION_ORDER = ["摘要", "摘要附图", "权利要求书", "说明书", "说明书附图"]
SECTION_FILES = {name: f"{name}.md" for name in SECTION_ORDER}

RE_SECTION_LINE = re.compile(r"^:::\s*(摘要|摘要附图|权利要求书|说明书|说明书附图)\s*$")


def _strip_section_markers(text: str) -> str:
    # 有些编辑器写 UTF-8 with BOM，BOM 会在合并后跑到行首、破坏标题与图片解析
    text = text.replace("\ufeff", "")
    return "\n".join(
        ln for ln in text.splitlines() if not RE_SECTION_LINE.match(ln.strip())
    ).strip()


def _sanitize_name(name: str) -> str:
    name = re.sub(r'[\\/:*?"<>|\s]+', "", name).strip()
    return name or "专利申请"


def assemble(src_dir: Path, case_name: str, *, timestamp: str | None = None
             ) -> tuple[str, dict]:
    ts = timestamp or datetime.now().strftime("%Y%m%d%H%M%S")
    stats = {"found": [], "missing": [], "warn": []}
    parts: list[str] = []

    texts: dict[str, str] = {}
    for sec in SECTION_ORDER:
        f = src_dir / SECTION_FILES[sec]
        if f.is_file():
            raw = f.read_text(encoding="utf-8")
            body = _strip_section_markers(raw)
            texts[sec] = body
            stats["found"].append(sec)
        else:
            stats["missing"].append(sec)

    # 说明书.md 若是全集（自带全部五节），退回整体重排
    if not texts and (src_dir / "说明书.md").is_file():
        stats["warn"].append("说明书.md 自带全部节，按原文重排")
        return (src_dir / "说明书.md").read_text(encoding="utf-8"), stats

    for sec in SECTION_ORDER:
        body = texts.get(sec, "").strip()
        parts.append(f"::: {sec}")
        if body:
            parts.append(body)
        else:
            stats["warn"].append(f"{sec} 为空")
        parts.append("")

    # 摘要附图 / 说明书附图：若正文没写图，提示
    for sec in ("摘要附图", "说明书附图"):
        body = texts.get(sec, "")
        if body and not re.search(r"!\[[^\]]*\]\([^)]+\)", body):
            stats["warn"].append(f"{sec} 无图片引用")

    content = "\n".join(parts).rstrip() + "\n"
    for w in stats["warn"]:
        content  # 仅记录，不改内容
    return content, stats


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="合并申请文件分段稿")
    ap.add_argument("--dir", required=True, type=Path)
    ap.add_argument("--name", required=True, help="案件名（用于文件名）")
    ap.add_argument("--out", type=Path)
    ap.add_argument("--stamp", default=None, help="自定义时间戳，默认当前时间")
    ap.add_argument("--check", action="store_true", help="只检查不出文件")
    args = ap.parse_args(argv)

    d = args.dir.expanduser().resolve()
    if not d.is_dir():
        print(f"ASSEMBLE: ok=0 reason=missing_dir path={d}")
        return 1

    content, stats = assemble(d, args.name, timestamp=args.stamp)

    print(f"ASSEMBLE_FOUND: {','.join(stats['found']) or '-'}")
    if stats["missing"]:
        print(f"ASSEMBLE_MISSING: {','.join(stats['missing'])}")
    for wmsg in stats["warn"]:
        print(f"ASSEMBLE_WARN: {wmsg}")

    n_sec = sum(1 for s in SECTION_ORDER if f"::: {s}" in content)
    ok = n_sec == len(SECTION_ORDER)
    print(f"ASSEMBLE_SECTIONS: found={n_sec}/{len(SECTION_ORDER)}")

    if args.check:
        print(f"ASSEMBLE: ok={1 if ok else 0} check_only=1")
        return 0 if ok else 1

    ts = args.stamp or datetime.now().strftime("%Y%m%d%H%M%S")
    out = args.out or (d / f"{_sanitize_name(args.name)}_{ts}.md")
    if out.exists():
        print(f"ASSEMBLE: ok=0 reason=exists path={out}")
        return 1
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(content, encoding="utf-8", newline="\n")
    print(f"已写入: {out}")
    print(f"ASSEMBLE: ok={1 if ok else 0} written=1")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())