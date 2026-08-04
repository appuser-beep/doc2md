"""Office 穷举跑测：Wave2 基础加强 + Wave3 深入。"""

from __future__ import annotations

import json
import sys
import time
from dataclasses import dataclass
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from converter import ConversionError, convert_path  # noqa: E402
from test_helpers import IMAGE_TRACE

SAMPLES = Path(__file__).resolve().parent / "samples"
OUT = Path(__file__).resolve().parent / "output"
REPORT = Path(__file__).resolve().parent / "OFFICE_EXHAUST_REPORT.md"


@dataclass
class Case:
    file: str
    must: list[str]
    forbid: list[str] = None
    min_chars: int = 15
    expect_fail: bool = False
    soft_fail_ok: bool = False  # 失败或极短均可接受
    wave: str = "W2"
    note: str = ""

    def __post_init__(self):
        if self.forbid is None:
            self.forbid = []


CASES: list[Case] = [
    # ---- Word Wave2 ----
    Case("W20_styles_basic.docx", ["W20_STYLES", "W20_BOLD", "W20_ITALIC", "W20_BODY"], ["NaN"], wave="W2"),
    Case("W21_multilevel_list.docx", ["W21_LIST", "L1_A", "L2_A", "N1", "W21_END"], wave="W2"),
    Case("W22_table_rich.docx", ["W22_TABLE", "华北", "1,280", "Overseas"], ["NaN"], wave="W2"),
    Case("W23_images_basic.docx", ["W23_IMG", "CAP_W23_A", "CAP_W23_B", "W23_END"], wave="W2"),
    Case("W24_mix_basic.docx", ["W24_MIX", "96.2%", "CAP_W24_C", "W24_END"], ["NaN"], wave="W2"),
    Case("W25_special_chars.docx", ["W25_CHARS", "W25_CN", "W25_EN", "WORD_TOKEN"], wave="W2"),
    # ---- Word Wave3 ----
    Case("W30_nested_table.docx", ["W30_NEST", "OUTER_L", "INNER_MARK", "SIDE_NOTE"], ["NaN"], wave="W3"),
    Case("W31_complex_merge.docx", ["W31_MERGE", "MERGE_TOP", "华北事业部", "W31_END"], ["NaN"], wave="W3"),
    Case("W32_multi_section_pagebreak.docx", ["W32_SEC", "CH01", "CH07", "W32_END"], wave="W3"),
    Case("W33_header_footer_mix.docx", ["W33_HF", "BODY_W33", "HDR_W33", "FTR_W33"], wave="W3"),
    Case("W34_hyperlink_rich.docx", ["W34_LINK", "LINK_W34", "example.com"], wave="W3"),
    Case("W35_stress_pack.docx", ["W35_STRESS", "W35_END", "CAP_W35_1"], ["NaN"], wave="W3"),
    Case("W36_almost_empty.docx", [], min_chars=0, wave="W3", note="几乎空"),
    Case("W37_corrupt.docx", [], expect_fail=True, soft_fail_ok=True, min_chars=0, wave="W3", note="损坏docx"),
    Case("W38_fake_old.doc", [], soft_fail_ok=True, min_chars=0, wave="W3", note="老doc负面：失败或乱码均可"),
    Case("W39 中文 空格 名.docx", ["W39_NAME", "W39_BODY"], wave="W3"),
    # ---- PPT Wave2 ----
    Case("P20_text_basic.pptx", ["P20_T1", "P20_TEXT_1", "NOTE_P20_1"], wave="W2"),
    Case("P21_table_basic.pptx", ["P21_TT", "华北", "1280", "P21_MARK"], wave="W2"),
    Case("P22_images_basic.pptx", ["P22_IMG", "P22_MARK", IMAGE_TRACE], wave="W2"),
    Case("P23_mix_basic.pptx", ["P23_MIX", "96%", IMAGE_TRACE], wave="W2"),
    Case("P24_styles_list.pptx", ["P24_STYLE", "P24_L1", "P24_L2"], wave="W2"),
    Case("P25_special_chars.pptx", ["P25_CHARS", "TOKEN_P25", "P25_MARK"], wave="W2"),
    # ---- PPT Wave3 ----
    Case("P30_multi_section.pptx", ["P30_SEC1", "P30_SEC8", "NOTE_P30_1", "P30_END"], wave="W3"),
    Case("P31_hidden_slide.pptx", ["VISIBLE_P31", "AFTER_P31"], wave="W3", note="隐藏页可能仍输出"),
    Case("P32_free_textbox.pptx", ["P32_FREEBOX", "P32_MARK"], wave="W3"),
    Case("P33_stress_pack.pptx", ["P33_S01", "P33_END", "NOTE_P33_1"], wave="W3"),
    Case("P34_blankish.pptx", [], min_chars=0, wave="W3"),
    Case("P35_corrupt.pptx", [], expect_fail=True, soft_fail_ok=True, min_chars=0, wave="W3"),
    Case("P36_fake_old.ppt", [], soft_fail_ok=True, min_chars=0, wave="W3", note="老ppt负面"),
    Case("P37 演示 中文名.pptx", ["P37_NAME", "P37_BODY"], wave="W3"),
    # ---- Excel Wave2 ----
    Case("E20_text_basic.xlsx", ["E20_TEXT", "E20_CN"], ["NaN", "Unnamed", "| --- |"], wave="W2"),
    Case("E21_table_basic.xlsx", ["E21_TT", "华北", "1280", "E21_MARK"], ["NaN"], wave="W2"),
    Case("E22_images_basic.xlsx", ["E22_IMG", "E22_MARK", "embedded-image"], ["NaN"], wave="W2"),
    Case("E23_mix_basic.xlsx", ["E23_MIX", "96%", "embedded-image"], ["NaN"], wave="W2"),
    Case("E24_special_chars.xlsx", ["E24_CHARS", "E24_MARK", "华北"], ["NaN"], wave="W2"),
    Case("E25_side_tables.xlsx", ["E25_LEFT", "E25_RIGHT", "华北", "人力"], ["NaN"], wave="W2"),
    # ---- Excel Wave3 ----
    Case("E30_complex_merge.xlsx", ["E30_MERGE_HDR", "华北组", "海外组", "E30_END"], ["NaN", "Unnamed"], wave="W3"),
    Case("E31_multi_sheet_hidden.xlsx", ["SHEET_华北", "SHEET_汇总", "E31_MULTI"], ["NaN"], wave="W3", note="隐藏表可能不输出"),
    Case("E32_types_formula.xlsx", ["E32_TYPES", "E32_MARK"], ["NaN"], wave="W3"),
    Case("E33_wide_tall.xlsx", ["E33_WIDE", "列1", "列30", "行24", "E33_MARK"], ["NaN"], wave="W3"),
    Case("E34_comment_chart_image.xlsx", ["E34_COMMENT", "E34_MARK", "embedded-image"], ["NaN"], wave="W3", note="批注/图表标题可能丢失"),
    Case("E35_stress_pack.xlsx", ["E35_S1", "E35_END"], ["NaN", "Unnamed"], wave="W3"),
    Case("E36_emptyish.xlsx", [], min_chars=0, wave="W3"),
    Case("E37_xls_multi.xls", ["XLS_N", "XLS_O", "E37_MARK", "1280"], ["NaN"], wave="W3"),
    Case("E38_corrupt.xlsx", [], expect_fail=True, soft_fail_ok=True, min_chars=0, wave="W3"),
    Case("E39 评分表 中文.xlsx", ["E39_NAME", "E39_BODY"], ["NaN"], wave="W3"),
    Case("E40_workbook_protect_meta.xlsx", ["E40_PROT", "E40_MARK"], ["NaN"], wave="W3"),
    # ---- MSG ----
    Case("M20_official.msg", ["test.sender@example.com", "Test Email Message"], wave="W2"),
    Case("M21 邮件 中文名.msg", ["Test Email Message", "body of the test email"], wave="W3"),
    Case("M30_truncated.msg", [], soft_fail_ok=True, min_chars=0, wave="W3", note="截断msg"),
    Case("M31_fake_ole.msg", [], expect_fail=True, soft_fail_ok=True, min_chars=0, wave="W3"),
    Case("M32_random.bin.msg", [], expect_fail=True, soft_fail_ok=True, min_chars=0, wave="W3"),
    Case(
        "M33_msg_as.docx",
        [],
        soft_fail_ok=True,
        min_chars=0,
        wave="W3",
        note="后缀伪装：失败或误读均可记录",
    ),
]


def evaluate(md: str, case: Case, exc: Exception | None) -> tuple[str, list[str]]:
    from test_helpers import check_must

    issues: list[str] = []
    if case.expect_fail or case.soft_fail_ok:
        if exc is not None:
            return "PASS", []
        if len(md.strip()) <= case.min_chars + 5 and case.expect_fail:
            return "PASS", []
        if case.soft_fail_ok:
            # 允许“转出乱码/短内容”，只要不崩溃
            return "PASS", []
        if len(md.strip()) > 80:
            return "FAIL", ["期望失败/空，但产出较长"]
        return "PASS", []

    if exc is not None:
        return "FAIL", [f"异常: {exc}"]

    if len(md.strip()) < case.min_chars:
        issues.append(f"过短 len={len(md.strip())}")
    issues.extend(check_must(md, case.must))
    for f in case.forbid:
        if f in md:
            issues.append(f"含 forbid: {f}")
    return ("PASS" if not issues else "FAIL"), issues


def main() -> int:
    OUT.mkdir(parents=True, exist_ok=True)
    results = []
    stats = {"PASS": 0, "FAIL": 0, "SKIP": 0}

    for case in CASES:
        path = SAMPLES / case.file
        if not path.exists():
            stats["SKIP"] += 1
            results.append({"file": case.file, "status": "SKIP", "wave": case.wave, "issues": ["不存在"], "note": case.note})
            print(f"SKIP {case.file}")
            continue

        t0 = time.time()
        md = ""
        exc = None
        try:
            md = convert_path(str(path))
            (OUT / (path.stem + ".md")).write_text(md, encoding="utf-8")
        except Exception as e:  # noqa: BLE001
            exc = e

        elapsed = round(time.time() - t0, 2)
        status, issues = evaluate(md, case, exc)
        stats[status] += 1
        results.append(
            {
                "file": case.file,
                "status": status,
                "wave": case.wave,
                "issues": issues,
                "chars": len(md) if md else 0,
                "sec": elapsed,
                "note": case.note,
                "error": (str(exc)[:180] if exc else ""),
            }
        )
        extra = issues or ([str(exc)[:80]] if exc and status == "PASS" else [])
        print(f"{status} [{case.wave}] {case.file} ({elapsed}s) {extra or ''}")

    # report
    lines = [
        "# Office 穷举报告（Wave2 基础加强 + Wave3 深入）",
        "",
        f"- 总计: {len(results)} | PASS: {stats['PASS']} | FAIL: {stats['FAIL']} | SKIP: {stats['SKIP']}",
        "",
        "| Wave | 文件 | 状态 | 字数 | 耗时 | 问题 |",
        "|---|---|---|---:|---:|---|",
    ]
    for r in results:
        iss = "; ".join(r["issues"]) or r.get("error") or ""
        lines.append(
            f"| {r['wave']} | `{r['file']}` | **{r['status']}** | {r.get('chars', '-')} | {r.get('sec', '-')} | {iss} |"
        )

    fails = [r for r in results if r["status"] == "FAIL"]
    lines.append("")
    if fails:
        lines.append("## 失败明细")
        for r in fails:
            lines.append(f"### {r['file']}")
            for i in r["issues"]:
                lines.append(f"- {i}")
            if r.get("note"):
                lines.append(f"- note: {r['note']}")
            if r.get("error"):
                lines.append(f"- error: {r['error']}")
            lines.append("")
    else:
        lines.append("## 全部通过")
        lines.append("")

    REPORT.write_text("\n".join(lines), encoding="utf-8")
    (Path(__file__).parent / "results.json").write_text(
        json.dumps({"stats": stats, "results": results}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(f"\nDONE pass={stats['PASS']} fail={stats['FAIL']} skip={stats['SKIP']} -> {REPORT}")
    return 0 if stats["FAIL"] == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
