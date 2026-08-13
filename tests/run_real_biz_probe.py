# -*- coding: utf-8 -*-
"""本机真实业务文档抽检（复杂表/合并单元格/ZIP/长 PDF）。"""

from __future__ import annotations

import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from converter import convert_path, postcheck_result  # noqa: E402

OUT = Path(__file__).resolve().parent / "_real_biz_out.txt"

CASES = [
    {
        "path": Path(r"C:\Users\wluser\Desktop\周一至周五档案整理.xlsx"),
        "must_any": ["档案", "整理", "周一", "周"],
        "min_chars": 80,
        "note": "多表+合并排班",
    },
    {
        "path": Path(r"C:\Users\wluser\Desktop\王成-运维岗offer.docx"),
        "must_any": ["王成", "运维", "offer", "录用", "岗位"],
        "min_chars": 200,
        "note": "中文录用函",
    },
    {
        "path": Path(r"C:\Users\wluser\Desktop\档案整理活动.zip"),
        "must_any": ["档案", "整理", "签到", "活动"],
        "min_chars": 300,
        "note": "多 docx ZIP",
    },
    {
        "path": Path(r"C:\Users\wluser\Desktop\调度智能解耦创新材料正式版.zip"),
        "must_any": ["调度", "解耦", "html", "HTML", "创新"],
        "min_chars": 400,
        "note": "docx+html 混合包",
    },
    {
        "path": Path(
            r"C:\Users\wluser\Documents\WPSDrive\1359559257\WPS云盘\国奖分享会在线表格.xlsx"
        ),
        "must_any": ["国奖", "分享", "姓名", "学号", "表"],
        "min_chars": 60,
        "note": "在线表多 sheet",
    },
    {
        "path": Path(
            r"C:\Users\wluser\Documents\WPSDrive\1359559257\WPS云盘\2_返家乡人员信息汇总表(1).xlsx"
        ),
        "must_any": ["返家乡", "姓名", "学院", "人员"],
        "min_chars": 100,
        "note": "人员汇总宽表",
    },
    {
        "path": Path(
            r"C:\Users\wluser\Documents\WPSDrive\1359559257\WPS云盘\应用\输出为PDF\我的论文_20240908195642.pdf"
        ),
        "must_any": ["论文", "摘要", "章", "研究", "结论"],
        "min_chars": 500,
        "note": "长论文 PDF",
    },
    {
        "path": Path(
            r"C:\Users\wluser\Desktop\markdown\tests\office_exhaust\samples\P33_stress_pack.pptx"
        ),
        "must_any": ["P33", "stress", "PPT", "标题", "MARK"],
        "min_chars": 40,
        "note": "应力 PPTX 夹具",
    },
    {
        "path": Path(
            r"C:\Users\wluser\Desktop\markdown\tests\full_matrix\phase2\B3_notebook.ipynb"
        ),
        "must_any": ["#", "code", "print", "cell", "Notebook"],
        "min_chars": 20,
        "note": "notebook",
    },
]


def main() -> int:
    lines: list[str] = []
    ok = 0
    fail = 0
    skip = 0
    for case in CASES:
        p: Path = case["path"]
        title = f"{p.name} ({case['note']})"
        if not p.exists():
            lines.append(f"SKIP\t{title}\tmissing")
            skip += 1
            continue
        t0 = time.time()
        try:
            md = convert_path(str(p), local_only=False)
        except Exception as exc:  # noqa: BLE001
            lines.append(f"FAIL\t{title}\tEXC {exc}")
            fail += 1
            continue
        elapsed = time.time() - t0
        tip = postcheck_result(str(p), md) or ""
        must = case["must_any"]
        hit = [k for k in must if k.lower() in md.lower()]
        chars = len((md or "").strip())
        good = chars >= int(case["min_chars"]) and bool(hit)
        status = "PASS" if good else "WARN"
        if good:
            ok += 1
        else:
            fail += 1
        preview = (md or "").replace("\n", " ")[:160]
        lines.append(
            f"{status}\t{title}\tchars={chars}\thit={hit}\t"
            f"{elapsed:.2f}s\ttip={tip[:40]!r}\t{preview}"
        )

    summary = f"PASS={ok} FAIL/WARN={fail} SKIP={skip} total={len(CASES)}"
    report = summary + "\n" + "\n".join(lines) + "\n"
    OUT.write_text(report, encoding="utf-8")
    try:
        print(report)
    except UnicodeEncodeError:
        sys.stdout.buffer.write(report.encode("utf-8", errors="replace"))
    return 0 if fail == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
