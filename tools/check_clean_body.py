#!/usr/bin/env python
"""正文体检：确认主交付 Markdown 里没有元信息泄漏。

申请文件的正文不许出现公开号、URL、检索腔、待办标记。查新证据一律去
《查新与区别定位》附件。本脚本是这条规矩的机器兜底。

用法：
  python tools/check_clean_body.py -i 申请文件.md
  python tools/check_clean_body.py -i 申请文件.md --json

输出 `CLEAN_BODY: ok=N fail=M warn=K`；退出码 0 表示无 fail。
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

# (规则名, 正则, 级别, 说明)
RULES: list[tuple[str, str, str, str]] = [
    ("publication_no", r"CN\s*\d{6,}[A-Z]?", "fail", "公开号只允许出现在查新附件"),
    ("other_pub_no", r"\b(?:US|EP|JP|KR|WO)\s*\d{6,}[A-Z]?\d?\b", "fail", "海外专利号只允许出现在查新附件"),
    ("url", r"https?://\S+", "fail", "正文不得出现 URL"),
    ("doi", r"\b(?:doi|DOI)\s*[:：]?\s*10\.\d{4,}", "fail", "正文不得出现 DOI"),
    ("search_talk", r"经检索|据检索|检索到|查新发现|通过检索", "fail", "正文不得出现检索腔"),
    ("todo", r"待核实|待确认|待补|占位|TODO|TBD|FIXME|XXX|【.*?】", "fail", "正文不得出现待办或占位标记"),
    ("draft", r"草稿|初稿|草拟|示例文本|此处应|如下填写", "fail", "正文不得出现草稿腔"),
    ("meta_skill", r"本技能|该技能|patent-china|prompt|skill", "fail", "正文不得出现工具元信息"),
    ("score", r"质量分|得分[:：]\s*\d|评分[:：]\s*\d", "fail", "正文不得出现自评分"),
    ("placeholder_x", r"某某|某公司|XXX公司|张三|李四", "fail", "正文不得出现占位人名/公司"),
    ("hype", r"填补空白|国际领先|革命性|颠覆性|史无前例|开创性", "warn", "避免空洞评价"),
    ("hype2", r"创新性强|显著优于|全面超越", "warn", "效果须对得上手段，慎用比较级"),
    ("citation_bracket", r"\[\d+\]", "warn", "方括号数字像文献引用，正文不需要"),
    ("chapter_num", r"^\s*(?:第[一二三四五六七八九十]+章|\d+\.\d+(?:\.\d+)?\s)", "fail",
     "章标题固定为五个中文名，不用章节号"),
    ("latex_bare_paren", r"\((?:[^()]*\\\\?(?:frac|sum|mathrm|alpha|beta|gamma|sigma|theta)"
                          r"[^()]*)\)", "fail", "行内公式须用 \\( \\)，不能用普通括号包 LaTeX"),
]


def _iter_lines(text: str):
    """跳过围栏块；公式 $$ … $$ 区间内的行也不是正文。"""
    in_fence = False
    in_math = False
    for no, line in enumerate(text.splitlines(), 1):
        stripped = line.strip()
        if stripped.startswith("```"):
            in_fence = not in_fence
            continue
        if in_fence:
            continue
        if stripped.startswith("$$"):
            # 单行 $$ … $$ 或独立起止围栏
            if stripped == "$$" or (stripped.startswith("$$") and stripped.endswith("$$")
                                    and len(stripped) > 4):
                in_math = not in_math if stripped == "$$" else in_math
                continue
            in_math = not in_math
            continue
        if in_math:
            continue
        yield no, line


# 模板固定套话：出现在正文时必须是逐字原文（见 references/template_spec.md）
BOILERPLATE = [
    ("附图说明引导",
     "为了更清楚地说明本发明实施例的技术方案，下面将对实施例描述所需要使用的附图作简单地介绍，"
     "显而易见地，下面描述中的附图仅仅是本发明的一些实施例，对于本领域普通技术人员来讲，"
     "在不付出创造性劳动的前提下，还可以根据这些附图获得其他的附图。"),
    ("具体实施方式引导",
     "下面将结合本发明实施例中的附图，对本发明实施例中的技术方案进行清楚、完整地描述，"
     "显然，所描述的实施例仅仅是本发明一部分实施例，而不是全部的实施例。"
     "基于本发明中的实施例，本领域普通技术人员在没有做出创造性劳动前提下所获得的所有其他实施例，"
     "都属于本发明保护的范围。"),
    ("有益效果收尾",
     "当然，实施本发明的任一产品并不一定需要同时达到以上描述的所有优点。"),
    ("实施例递进说明",
     "本说明书中的各个实施例均采用递进的方式描述，各个实施例之间相同相似的部分互相参见即可，"
     "每个实施例重点说明的都是与其他实施例的不同之处。"
     "尤其，对于装置、设备、非易失性计算机存储介质实施例而言，由于其基本相似于方法实施例，"
     "所以描述的比较简单，相关之处参见方法实施例的部分说明即可。"),
    ("保护范围收尾",
     "以上内容仅仅是对本发明的构思所作的举例和说明，所属本技术领域的技术人员对所描述的具体实施例"
     "做各种各样的修改或补充或采用类似的方式替代，只要不偏离发明的构思或者超越本申请所定义的范围，"
     "均应属于本发明的保护范围。"),
]

# 套话的标志性前缀：用来判断「这段是不是在写套话」
BOILERPLATE_PROBES = [
    ("附图说明引导", "为了更清楚地说明本发明实施例"),
    ("具体实施方式引导", "下面将结合本发明实施例中的附图"),
    ("有益效果收尾", "当然，实施本发明的任一产品"),
    ("实施例递进说明", "本说明书中的各个实施例均采用递进"),
    ("保护范围收尾", "以上内容仅仅是对本发明的构思"),
]


def check_boilerplate(text: str) -> list[dict]:
    """套话段必须逐字一致，改一个字都不行。

    只在「探针命中」时判定：稿件确实在写这条套话，就必须与模板一字不差。
    比对前统一去掉空白与 Markdown 强调符，避免假阳性。
    """
    findings: list[dict] = []
    flat = text.replace("*", "").replace("`", "")
    flat = re.sub(r"[ \t\u3000]+", "", flat)
    flat = re.sub(r"\n+", "", flat)  # 套话可能被折行，全抹平再比
    for name, probe in BOILERPLATE_PROBES:
        if probe not in flat:
            continue
        expected = dict(BOILERPLATE)[name]
        expected_flat = re.sub(r"[ \t\u3000]+", "", expected)
        if expected_flat not in flat:
            # 找出第一处分歧，便于用户一眼看到错在哪
            idx = flat.find(probe)
            seg = flat[idx: idx + len(expected_flat) + 10]
            k = 0
            while k < min(len(seg), len(expected_flat)) and seg[k] == expected_flat[k]:
                k += 1
            findings.append({
                "level": "fail", "rule": "boilerplate_verbatim", "line": 0,
                "match": name,
                "context": f"第 {k} 字起分歧：稿={seg[max(0,k-8):k+18]!r} 模板={expected_flat[max(0,k-8):k+18]!r}",
                "msg": f"套话「{name}」与模板不一致，须逐字照抄 references/template_spec.md",
            })
    return findings


def _strip_inline_math(line: str) -> str:
    """去掉规范的 \\( … \\) 行内公式，剩下的括号才可能是误用。"""
    return re.sub(r"\\\(.+?\\\)", "", line)


def check(text: str) -> list[dict]:
    findings: list[dict] = []
    for name, pattern, level, msg in RULES:
        rx = re.compile(pattern)
        hits = []
        for no, line in _iter_lines(text):
            if line.strip() in ("$$",):
                continue
            probe = _strip_inline_math(line) if name == "latex_bare_paren" else line
            for m in rx.finditer(probe):
                hits.append({"line": no, "match": m.group(0)[:60],
                             "context": line.strip()[:80]})
        for h in hits[:6]:
            findings.append({
                "level": level, "rule": name, "msg": msg,
                "line": h["line"], "match": h["match"], "context": h["context"],
            })
        if len(hits) > 6:
            findings.append({
                "level": level, "rule": name, "msg": f"{msg}（另有 {len(hits)-6} 处）",
                "line": hits[6]["line"], "match": "", "context": "",
            })
    return findings


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="正文元信息体检")
    ap.add_argument("-i", "--input", required=True, type=Path)
    ap.add_argument("--json", action="store_true")
    args = ap.parse_args(argv)

    path = args.input.expanduser().resolve()
    if not path.is_file():
        print(f"CLEAN_BODY: ok=0 reason=missing_file path={path}")
        return 1

    text = path.read_text(encoding="utf-8")
    findings = check(text) + check_boilerplate(text)

    if args.json:
        print(json.dumps(findings, ensure_ascii=False, indent=2))
    else:
        if not findings:
            print("[OK  ] 正文无元信息泄漏")
        for f in findings:
            mark = "FAIL" if f["level"] == "fail" else "WARN"
            ln = f"L{f['line']}" if f.get("line") else "-"
            print(f"[{mark}] {f['rule']} {ln}: {f['msg']}")
            if f.get("context"):
                print(f"        → {f['context']}")

    nf = sum(1 for f in findings if f["level"] == "fail")
    nw = sum(1 for f in findings if f["level"] == "warn")
    print(f"CLEAN_BODY: ok={1 if nf == 0 else 0} fail={nf} warn={nw}")
    return 0 if nf == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())