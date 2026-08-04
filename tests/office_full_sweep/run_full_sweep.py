"""Office 全量穷举跑测（对照 A1–A5 + G 负面）。"""

from __future__ import annotations

import json
import sys
import time
from dataclasses import dataclass, field
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from converter import convert_path  # noqa: E402
from test_helpers import IMAGE_TRACE

SAMPLES = Path(__file__).resolve().parent / "samples"
OUT = Path(__file__).resolve().parent / "output"
REPORT = Path(__file__).resolve().parent / "FULL_SWEEP_REPORT.md"


@dataclass
class Case:
    file: str
    must: list[str] = field(default_factory=list)
    forbid: list[str] = field(default_factory=list)
    min_chars: int = 8
    soft: bool = False
    group: str = ""
    note: str = ""


CASES: list[Case] = [
    # A1 Word
    Case("FS_A1_01_long_text.docx", ["A101_H1", "A101_H2", "A101_L1", "A101_END"], group="A1"),
    Case("FS_A1_02_text_table.docx", ["A102_TT", "华北", "1280"], forbid=["NaN"], group="A1"),
    Case("FS_A1_03_text_image.docx", ["A103_TI", "CAP_A103"], group="A1"),
    Case("FS_A1_04_text_image_table.docx", ["A104_MIX", "96%", "CAP_A104", "A104_END"], forbid=["NaN"], group="A1"),
    Case("FS_A1_05_multi_table.docx", ["A105_T1", "A105_T2", "A105_T3"], forbid=["NaN"], group="A1"),
    Case("FS_A1_06_multi_image.docx", ["CAP_A106_A", "CAP_A106_B", "A106_MARK"], group="A1"),
    Case("FS_A1_07_merged_table.docx", ["A107_MERGE", "MERGE_TOP", "华北组", "A107_END"], forbid=["NaN"], group="A1"),
    Case("FS_A1_08_wide_table.docx", ["A108_WIDE", "列1", "列12", "A108_MARK"], forbid=["NaN"], group="A1"),
    Case("FS_A1_09_hyperlink.docx", ["A109_LINK", "LINK_A109", "BOOKMARK_A109", "example.com"], group="A1"),
    Case("FS_A1_10_header_footer.docx", ["A110_HF", "BODY_A110", "HDR_A110", "FTR_A110"], group="A1"),
    Case("FS_A1_11_textbox.docx", ["A111_BOX", "TEXTBOX_A111", "A111_END"], group="A1"),
    Case("FS_A1_12_footnote_like.docx", ["A112_FN", "FOOTNOTE_A112", "ENDNOTE_A112"], group="A1"),
    Case("FS_A1_13_toc_chapters.docx", ["A113_TOC", "A113_CH1", "A113_CH3", "A113_END"], group="A1"),
    Case("FS_A1_14_comment_rev.docx", ["A114_REV", "KEEP_A114", "COMMENT_A114"], group="A1"),
    Case("FS_A1_15_mixed_lang.docx", ["A115_LANG", "A115_CN", "A115_EN", "1280"], group="A1"),
    Case("FS_A1_16_image_in_table.docx", ["A116_CELLIMG", "CELL_A116", "CAP_A116"], group="A1"),
    Case("FS_A1_17_stress.docx", ["A117_STRESS", "A117_CH1", "A117_END", "CAP_A117_1"], forbid=["NaN"], group="A1"),
    Case("FS_A1_18_nested.docx", ["A118_NEST", "OUTER_L", "INNER_A118", "SIDE_A118"], forbid=["NaN"], group="A1"),
    # A2 PPT
    Case("FS_A2_01_text_multi.pptx", ["A201_T1", "A201_TEXT_1", "A201_T5"], group="A2"),
    Case("FS_A2_02_bullets.pptx", ["A202_LIST", "A202_L1", "A202_L2"], group="A2"),
    Case("FS_A2_03_table.pptx", ["A203_TT", "华北", "1280", "A203_MARK"], group="A2"),
    Case("FS_A2_04_image.pptx", ["A204_TI", "CAP_A204", IMAGE_TRACE], group="A2"),
    Case("FS_A2_05_mix.pptx", ["A205_MIX", "96%", IMAGE_TRACE], group="A2"),
    Case("FS_A2_06_multi_table.pptx", ["A206_T1", "A206_T2", "A206_T3"], group="A2"),
    Case("FS_A2_07_multi_image.pptx", ["CAP_A207_A", "A207_MARK", IMAGE_TRACE], group="A2"),
    Case("FS_A2_08_notes.pptx", ["A208_T1", "NOTE_A208_1", "A208_BODY_1"], group="A2"),
    Case("FS_A2_09_hidden.pptx", ["VISIBLE_A209", "AFTER_A209"], group="A2", note="隐藏页可能仍输出"),
    Case("FS_A2_10_shapes.pptx", ["SHAPE_A210", "A210_MARK"], group="A2"),
    Case("FS_A2_11_chart.pptx", ["A211_CHART", "A211_MARK"], group="A2", note="图表数据点可能不全"),
    Case("FS_A2_12_footer_like.pptx", ["A212_FOOT", "A212_BODY", "FOOTERLIKE_A212"], group="A2"),
    Case("FS_A2_13_styles.pptx", ["A213_STYLE", "A213_MARK"], group="A2"),
    Case("FS_A2_14_scatter_boxes.pptx", ["SCATTER_A214_1", "SCATTER_A214_3", "A214_MARK"], group="A2"),
    Case("FS_A2_15_stress.pptx", ["A215_S01", "A215_END", "NOTE_A215_1"], group="A2"),
    # A3 Excel
    Case("FS_A3_01_text.xlsx", ["A301_TEXT", "A301_CN"], forbid=["NaN", "Unnamed", "| --- |"], group="A3"),
    Case("FS_A3_02_table.xlsx", ["A302_TT", "华北", "1280", "A302_MARK"], forbid=["NaN"], group="A3"),
    Case("FS_A3_03_image.xlsx", ["A303_IMG", "A303_MARK", "embedded-image"], forbid=["NaN"], group="A3"),
    Case("FS_A3_04_mix.xlsx", ["A304_MIX", "96%", "embedded-image"], forbid=["NaN"], group="A3"),
    Case("FS_A3_05_side.xlsx", ["A305_LEFT", "A305_RIGHT", "华北", "人力"], forbid=["NaN"], group="A3"),
    Case("FS_A3_06_multi_sheet.xlsx", ["SHEET_华北", "SHEET_汇总", "A306_MULTI"], forbid=["NaN"], group="A3"),
    Case("FS_A3_07_merge.xlsx", ["A307_MERGE_HDR", "华北组", "A307_END"], forbid=["NaN", "Unnamed"], group="A3"),
    Case("FS_A3_08_wide.xlsx", ["A308_WIDE", "列1", "列20", "A308_MARK"], forbid=["NaN"], group="A3"),
    Case("FS_A3_09_sparse.xlsx", ["A309_SPARSE", "华北", "A309_FAR", "A309_MARK"], forbid=["NaN"], group="A3"),
    Case("FS_A3_10_formats.xlsx", ["A310_FMT", "A310_MARK"], forbid=["NaN"], group="A3"),
    Case("FS_A3_11_formula.xlsx", ["A311_FORMULA", "A311_MARK"], forbid=["NaN"], group="A3"),
    Case("FS_A3_12_multi_image.xlsx", ["A312_IMG", "A312_MARK", "embedded-image"], forbid=["NaN"], group="A3"),
    Case("FS_A3_13_hidden.xlsx", ["A313_HIDE", "A313_VISIBLE", "A313_MARK"], forbid=["NaN"], group="A3", note="隐藏行列可能仍导出"),
    Case("FS_A3_14_excel_table.xlsx", ["A314_TABLE", "华北", "1280", "A314_MARK"], forbid=["NaN"], group="A3"),
    Case("FS_A3_15_stress.xlsx", ["A315_S1", "A315_END"], forbid=["NaN", "Unnamed"], group="A3"),
    Case("FS_A3_16_link_comment_chart.xlsx", ["A316_EXTRA", "A316_MARK", "华北"], forbid=["NaN"], group="A3"),
    # A4 xls
    Case("FS_A4_01_xls_basic.xls", ["A401_XLS", "华北", "1280", "A401_MARK"], forbid=["NaN"], group="A4"),
    Case("FS_A4_02_xls_multi.xls", ["A402_N", "A402_O", "A402_MARK"], forbid=["NaN"], group="A4"),
    # A5 MSG
    Case("FS_A5_01_official.msg", ["Test Email Message", "test.sender@example.com"], group="A5"),
    Case("FS_A5_02_fake_htmlish.msg", soft=True, min_chars=0, group="A5", note="伪MSG负面"),
    Case("FS_A5_03_中文名邮件.msg", ["Test Email Message"], group="A5"),
    Case("FS_A5_04_truncated.msg", soft=True, min_chars=0, group="A5", note="截断"),
    Case("FS_A5_05_empty.msg", soft=True, min_chars=0, group="A5", note="空"),
    Case("FS_A5_06_official_regress.msg", ["body of the test email"], group="A5"),
    # G
    Case("FS_G_01_empty.docx", soft=True, min_chars=0, group="G", note="空docx"),
    Case("FS_G_01_old.doc", soft=True, min_chars=0, group="G", note="老doc"),
    Case("FS_G_02_old.ppt", soft=True, min_chars=0, group="G", note="老ppt"),
    Case("FS_G_03（中文 空格）名.docx", ["G03_NAME", "G03_BODY"], group="G"),
    Case("FS_G_04_msg_as.docx", soft=True, min_chars=0, group="G", note="后缀不符"),
    Case("FS_G_05_corrupt.docx", soft=True, min_chars=0, group="G"),
    Case("FS_G_05_corrupt.pptx", soft=True, min_chars=0, group="G"),
    Case("FS_G_05_corrupt.xlsx", soft=True, min_chars=0, group="G"),
]


def evaluate(md: str, case: Case, exc: Exception | None) -> tuple[str, list[str]]:
    if case.soft:
        return "PASS", []
    if exc is not None:
        return "FAIL", [f"异常: {exc}"]
    issues: list[str] = []
    if len(md.strip()) < case.min_chars:
        issues.append(f"过短 len={len(md.strip())}")
    from test_helpers import check_must

    issues.extend(check_must(md, case.must))
    for f in case.forbid:
        if f in md:
            issues.append(f"含 forbid: {f}")
    return ("PASS" if not issues else "FAIL"), issues


def main() -> int:
    OUT.mkdir(parents=True, exist_ok=True)
    results = []
    stats = {"PASS": 0, "FAIL": 0, "SKIP": 0}
    by_group: dict[str, dict[str, int]] = {}

    for case in CASES:
        path = SAMPLES / case.file
        g = case.group or "?"
        by_group.setdefault(g, {"PASS": 0, "FAIL": 0, "SKIP": 0})
        if not path.exists():
            stats["SKIP"] += 1
            by_group[g]["SKIP"] += 1
            results.append({"file": case.file, "status": "SKIP", "group": g, "issues": ["不存在"], "note": case.note})
            print(f"SKIP [{g}] {case.file}")
            continue

        t0 = time.time()
        md, exc = "", None
        try:
            md = convert_path(str(path))
            safe = "".join(ch if ch not in '<>:"/\\|?*' else "_" for ch in path.stem)
            (OUT / f"{safe}.md").write_text(md, encoding="utf-8")
        except Exception as e:  # noqa: BLE001
            exc = e
        elapsed = round(time.time() - t0, 2)
        status, issues = evaluate(md, case, exc)
        stats[status] += 1
        by_group[g][status] += 1
        results.append(
            {
                "file": case.file,
                "status": status,
                "group": g,
                "issues": issues,
                "chars": len(md) if md else 0,
                "sec": elapsed,
                "note": case.note,
                "error": (str(exc)[:160] if exc else ""),
            }
        )
        print(f"{status} [{g}] {case.file} ({elapsed}s) {issues or (str(exc)[:50] if exc else '')}")

    lines = [
        "# Office 全量穷举报告（A1–A5 + G）",
        "",
        f"- 总计: {len(results)} | PASS: {stats['PASS']} | FAIL: {stats['FAIL']} | SKIP: {stats['SKIP']}",
        "",
        "## 分组汇总",
        "",
        "| 分组 | PASS | FAIL | SKIP |",
        "|---|---:|---:|---:|",
    ]
    for g in sorted(by_group):
        s = by_group[g]
        lines.append(f"| {g} | {s['PASS']} | {s['FAIL']} | {s['SKIP']} |")
    lines += [
        "",
        "| 分组 | 文件 | 状态 | 字数 | 耗时 | 问题 |",
        "|---|---|---|---:|---:|---|",
    ]
    for r in results:
        iss = "; ".join(r["issues"]) or r.get("error") or ""
        lines.append(
            f"| {r['group']} | `{r['file']}` | **{r['status']}** | {r.get('chars','-')} | {r.get('sec','-')} | {iss} |"
        )
    fails = [r for r in results if r["status"] == "FAIL"]
    lines.append("")
    if fails:
        lines.append("## 失败明细")
        for r in fails:
            lines.append(f"### {r['file']}")
            for i in r["issues"]:
                lines.append(f"- {i}")
            lines.append("")
    else:
        lines.append("## 全部通过（本矩阵）")
        lines.append("")

    # coverage checklist
    lines += [
        "## 大纲覆盖核对",
        "",
        "| 大纲项 | 样例 | 状态 |",
        "|---|---|---|",
    ]
    outline_map = [
        ("A1-01", "FS_A1_01"),
        ("A1-02", "FS_A1_02"),
        ("A1-03", "FS_A1_03"),
        ("A1-04", "FS_A1_04"),
        ("A1-05", "FS_A1_05"),
        ("A1-06", "FS_A1_06"),
        ("A1-07", "FS_A1_07"),
        ("A1-08", "FS_A1_08"),
        ("A1-09", "FS_A1_09"),
        ("A1-10", "FS_A1_10"),
        ("A1-11", "FS_A1_11"),
        ("A1-12", "FS_A1_12"),
        ("A1-13", "FS_A1_13"),
        ("A1-14", "FS_A1_14"),
        ("A1-15", "FS_A1_15"),
        ("A1-16", "FS_A1_16"),
        ("A1-17", "FS_A1_17"),
        ("A2-01", "FS_A2_01"),
        ("A2-02", "FS_A2_02"),
        ("A2-03", "FS_A2_03"),
        ("A2-04", "FS_A2_04"),
        ("A2-05", "FS_A2_05"),
        ("A2-06", "FS_A2_06"),
        ("A2-07", "FS_A2_07"),
        ("A2-08", "FS_A2_08"),
        ("A2-09", "FS_A2_09"),
        ("A2-10", "FS_A2_10"),
        ("A2-11", "FS_A2_11"),
        ("A2-12", "FS_A2_12"),
        ("A2-13", "FS_A2_13"),
        ("A2-14", "FS_A2_14"),
        ("A2-15", "FS_A2_15"),
        ("A3-01", "FS_A3_01"),
        ("A3-02", "FS_A3_02"),
        ("A3-03", "FS_A3_03"),
        ("A3-04", "FS_A3_04"),
        ("A3-05", "FS_A3_05"),
        ("A3-06", "FS_A3_06"),
        ("A3-07", "FS_A3_07"),
        ("A3-08", "FS_A3_08"),
        ("A3-09", "FS_A3_09"),
        ("A3-10", "FS_A3_10"),
        ("A3-11", "FS_A3_11"),
        ("A3-12", "FS_A3_12"),
        ("A3-13", "FS_A3_13"),
        ("A3-14", "FS_A3_14"),
        ("A3-15", "FS_A3_15"),
        ("A4", "FS_A4_"),
        ("A5-01/06", "FS_A5_01"),
        ("G-01~05", "FS_G_"),
    ]
    for oid, prefix in outline_map:
        matched = [r for r in results if r["file"].startswith(prefix)]
        if not matched:
            st = "缺样例"
        elif any(r["status"] == "FAIL" for r in matched):
            st = "FAIL"
        elif any(r["status"] == "SKIP" for r in matched):
            st = "SKIP"
        else:
            st = "PASS"
        lines.append(f"| {oid} | `{prefix}*` | **{st}** |")

    REPORT.write_text("\n".join(lines), encoding="utf-8")
    (Path(__file__).parent / "results.json").write_text(
        json.dumps({"stats": stats, "by_group": by_group, "results": results}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(f"\nDONE pass={stats['PASS']} fail={stats['FAIL']} skip={stats['SKIP']} -> {REPORT}")
    return 0 if stats["FAIL"] == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
