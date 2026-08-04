"""Office 加深加压跑测 + 报告。"""

from __future__ import annotations

import json
import sys
import time
from dataclasses import dataclass, asdict
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from converter import convert_path  # noqa: E402
from test_helpers import IMAGE_TRACE

SAMPLES = Path(__file__).resolve().parent / "samples"
OUT = Path(__file__).resolve().parent / "output"
REPORT = Path(__file__).resolve().parent / "OFFICE_DEEP_REPORT.md"


@dataclass
class Case:
    file: str
    must: list[str]
    forbid: list[str]
    min_chars: int = 20
    expect_fail: bool = False
    note: str = ""


CASES: list[Case] = [
    # Word
    Case("WD01_text_only.docx", ["WD_TEXT", "WD_BODY_CN", "WD_BODY_EN"], ["NaN", "Unnamed"]),
    Case("WD02_text_table.docx", ["WD_TT", "华北", "1280"], ["NaN"]),
    Case("WD03_text_image.docx", ["WD_TI", "CAP_WD_N"], [], min_chars=10),
    Case("WD04_text_image_table.docx", ["WD_MIX", "96%", "CAP_WD_C"], ["NaN"]),
    Case("WD05_multi_table.docx", ["WD_MT1", "WD_MT2", "WD_MT3", "华北", "人力"], ["NaN"]),
    Case("WD06_multi_image.docx", ["WD_MI", "CAP_WD_N1", "CAP_WD_S1"], []),
    Case("WD07_merged_table.docx", ["WD_MERGE", "MERGE_HDR", "华北"], ["NaN"]),
    Case("WD08_multi_section.docx", ["WD_LONG", "CH1", "CH5", "SEC1"], []),
    Case("WD09_mixed_lang.docx", ["WD_MIXLANG", "WD_LANG_MARK", "North China"], []),
    Case("WD10_stress_pack.docx", ["WD_STRESS", "WD_STRESS_END", "CAP_ST_1"], ["NaN"]),
    Case("WD11_header_footer.docx", ["WD_HF", "BODY_WD_HF", "HEADER_WD", "FOOTER_WD"], []),
    Case("WD12_hyperlink.docx", ["WD_LINK", "WD_LINK_BODY", "https://example.com/wd"], []),
    Case("WD13_almost_empty.docx", [], [], min_chars=0, note="几乎空文档"),
    # PPT
    Case("PP01_text_only.pptx", ["PPT_T1", "PPT_TEXT_1", "PPT_TEXT_5"], []),
    Case("PP02_text_table.pptx", ["PPT_TT", "华北", "1280"], []),
    Case("PP03_text_image.pptx", ["PPT_TI", "CAPTION_PPT_N", IMAGE_TRACE], []),
    Case("PP04_text_image_table.pptx", ["PPT_MIX", "96%", IMAGE_TRACE], []),
    Case("PP05_multi_table.pptx", ["PPT_MT1", "PPT_MT2", "PPT_MT3"], []),
    Case("PP06_multi_image.pptx", ["PPT_MI", "CAP_PPT_N1", IMAGE_TRACE], []),
    Case("PP07_multi_section_notes.pptx", ["PPT_SEC1", "PPT_SECTION_1", "NOTE_PPT_1"], []),
    Case("PP08_styles_list.pptx", ["PPT_STYLE", "PPT_STYLE_MARK"], []),
    Case("PP09_hidden_slide.pptx", ["VISIBLE_MARK"], [], note="隐藏页可能仍输出"),
    Case("PP10_mixed_lang.pptx", ["PPT_LANG", "华北", "Overseas"], []),
    Case("PP11_stress_pack.pptx", ["PPT_STRESS_1", "PPT_STRESS_END", "NOTE_STRESS_1"], []),
    Case("PP12_blankish.pptx", [], [], min_chars=0, note="空白页"),
    # Excel
    Case("EX01_text_only.xlsx", ["EX_TEXT", "EX_BODY_CN"], ["NaN", "Unnamed", "| --- |"]),
    Case("EX02_text_table.xlsx", ["EX_TT", "华北", "1280"], ["NaN", "Unnamed"]),
    Case("EX03_text_image.xlsx", ["EX_TI", "CAP_EX_N", "embedded-image"], ["NaN"]),
    Case("EX04_text_image_table.xlsx", ["EX_MIX", "96%", "CAP_EX_C", "embedded-image"], ["NaN"]),
    Case("EX05_multi_table_side.xlsx", ["EX_MT_LEFT", "EX_MT_RIGHT", "华北", "人力"], ["NaN"]),
    Case("EX06_multi_sheet.xlsx", ["SHEET_N", "SHEET_S", "SHEET_O", "SHEET_SUM"], ["NaN"]),
    Case("EX07_merged_nested.xlsx", ["EX_MERGE_HDR", "华北组", "1280"], ["NaN", "Unnamed"]),
    Case("EX08_mixed_lang.xlsx", ["EX_LANG", "North China", "EX_LANG_MARK"], ["NaN"]),
    Case("EX09_wide_table.xlsx", ["EX_WIDE", "列1", "列20"], ["NaN"]),
    Case("EX10_stress_pack.xlsx", ["EX_STRESS_1", "EX_STRESS_END"], ["NaN", "Unnamed"]),
    Case("EX11_formula_format.xlsx", ["EX_FMT", "EX_FMT_MARK"], ["NaN"]),
    Case("EX12_emptyish.xlsx", [], [], min_chars=0, note="几乎空表"),
    Case("EX13_xls_basic.xls", ["EX_XLS", "华北", "1280"], ["NaN"]),
    # MSG
    Case("MSG01_official.msg", ["test.sender@example.com", "Test Email Message", "body of the test email"], [], min_chars=10),
    Case("MSG02_official_dl.msg", ["test.recipient@example.com", "Test Email Message"], [], min_chars=10),
    Case("MSG03_fake_ole.msg", [], [], expect_fail=True, note="伪 OLE 应失败或空"),
]


def check(md: str, case: Case) -> tuple[bool, list[str]]:
    from test_helpers import check_must

    issues: list[str] = []
    if case.expect_fail:
        return True, []  # 由调用方处理异常
    if len(md.strip()) < case.min_chars:
        issues.append(f"过短 len={len(md.strip())}")
    issues.extend(check_must(md, case.must))
    for f in case.forbid:
        if f in md:
            issues.append(f"含 forbid: {f}")
    return len(issues) == 0, issues


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    results = []
    pass_n = fail_n = skip_n = 0

    for case in CASES:
        path = SAMPLES / case.file
        if not path.exists():
            skip_n += 1
            results.append({"file": case.file, "status": "SKIP", "issues": ["文件不存在"], "note": case.note})
            print(f"SKIP {case.file}")
            continue

        t0 = time.time()
        try:
            md = convert_path(str(path))
            elapsed = time.time() - t0
            out_path = OUT / (path.stem + ".md")
            out_path.write_text(md, encoding="utf-8")

            if case.expect_fail:
                # 伪文件若仍产出内容，记 WARN；完全空算 PASS
                if len(md.strip()) > 50:
                    fail_n += 1
                    results.append(
                        {
                            "file": case.file,
                            "status": "FAIL",
                            "issues": ["期望失败但产出较长内容"],
                            "chars": len(md),
                            "sec": round(elapsed, 2),
                            "note": case.note,
                        }
                    )
                    print(f"FAIL {case.file} (expect fail but got content)")
                else:
                    pass_n += 1
                    results.append(
                        {
                            "file": case.file,
                            "status": "PASS",
                            "issues": [],
                            "chars": len(md),
                            "sec": round(elapsed, 2),
                            "note": case.note,
                        }
                    )
                    print(f"PASS {case.file} (empty-ish as expected)")
                continue

            ok, issues = check(md, case)
            if ok:
                pass_n += 1
                status = "PASS"
            else:
                fail_n += 1
                status = "FAIL"
            results.append(
                {
                    "file": case.file,
                    "status": status,
                    "issues": issues,
                    "chars": len(md),
                    "sec": round(elapsed, 2),
                    "note": case.note,
                }
            )
            print(f"{status} {case.file} ({elapsed:.2f}s) {issues or ''}")
        except Exception as e:
            elapsed = time.time() - t0
            if case.expect_fail:
                pass_n += 1
                results.append(
                    {
                        "file": case.file,
                        "status": "PASS",
                        "issues": [],
                        "error": str(e)[:200],
                        "sec": round(elapsed, 2),
                        "note": case.note + " | 异常符合预期",
                    }
                )
                print(f"PASS {case.file} (exception as expected: {e})")
            else:
                fail_n += 1
                results.append(
                    {
                        "file": case.file,
                        "status": "FAIL",
                        "issues": [f"异常: {e}"],
                        "sec": round(elapsed, 2),
                        "note": case.note,
                    }
                )
                print(f"FAIL {case.file}: {e}")

    # report
    lines = [
        "# Office 加深加压报告",
        "",
        f"- 总计: {len(results)} | PASS: {pass_n} | FAIL: {fail_n} | SKIP: {skip_n}",
        "",
        "| 文件 | 状态 | 字数 | 耗时 | 问题 |",
        "|---|---|---:|---:|---|",
    ]
    for r in results:
        issues = "; ".join(r.get("issues") or []) or (r.get("error") or "")
        lines.append(
            f"| `{r['file']}` | **{r['status']}** | {r.get('chars', '-')} | {r.get('sec', '-')} | {issues} |"
        )
    lines.append("")
    fails = [r for r in results if r["status"] == "FAIL"]
    if fails:
        lines.append("## 失败明细")
        for r in fails:
            lines.append(f"### {r['file']}")
            for i in r.get("issues") or []:
                lines.append(f"- {i}")
            if r.get("note"):
                lines.append(f"- note: {r['note']}")
            lines.append("")
    else:
        lines.append("## 全部通过")
        lines.append("")

    REPORT.write_text("\n".join(lines), encoding="utf-8")
    (Path(__file__).parent / "results.json").write_text(
        json.dumps({"pass": pass_n, "fail": fail_n, "skip": skip_n, "results": results}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(f"\nDONE pass={pass_n} fail={fail_n} skip={skip_n} -> {REPORT}")
    return 0 if fail_n == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
