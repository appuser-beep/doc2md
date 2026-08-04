"""Office Wave4 + PDF 穷举跑测。"""

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
REPORT = Path(__file__).resolve().parent / "WAVE4_PDF_REPORT.md"


@dataclass
class Case:
    file: str
    must: list[str] = field(default_factory=list)
    forbid: list[str] = field(default_factory=list)
    min_chars: int = 10
    soft: bool = False  # 失败/空/乱码均可
    expect_emptyish: bool = False  # 扫描件等：无关键 must，允许空或仅图片痕迹
    kind: str = "office"
    note: str = ""


CASES: list[Case] = [
    # Office Wave4 Word
    Case("W40_footnote_strike.docx", ["W40_NOTE", "W40_BODY", "FOOTNOTE_W40", "KEEP_W40"], kind="office-w4"),
    Case("W41_quote_block.docx", ["W41_QUOTE", "QUOTE_W41", "W41_END"], kind="office-w4"),
    Case("W42_multi_section_headers.docx", ["W42_SEC", "BODY_SEC1_W42", "BODY_SEC2_W42", "HDR_SEC1_W42"], kind="office-w4"),
    Case("W43_dual_table_image.docx", ["W43_DUAL", "W43_A", "W43_B", "华北", "CAP_W43"], forbid=["NaN"], kind="office-w4"),
    Case("W44_long_paragraphs.docx", ["W44_LONG", "W44_MARK", "W44_END"], kind="office-w4"),
    Case("W45_emoji_zwsp.docx", ["W45_EMOJI", "W45_MARK", "1280"], kind="office-w4"),
    Case("W46（综合）中文名.docx", ["W46_NAME", "W46_BODY", "CAP_W46"], kind="office-w4"),
    # PPT
    Case("P40_two_column_boxes.pptx", ["LEFT_P40", "RIGHT_P40", "P40_MARK", "NOTE_P40"], kind="office-w4"),
    Case("P41_multi_mix_stress.pptx", ["P41_S1", "P41_END", "NOTE_P41_1", IMAGE_TRACE], kind="office-w4"),
    Case("P42_long_notes.pptx", ["P42_BODY", "NOTE_P42_LONG", "NOTE_P42_END"], kind="office-w4"),
    Case("P43_emoji.pptx", ["P43_EMOJI", "P43_MARK"], kind="office-w4"),
    # Excel
    Case("E40_side_merge.xlsx", ["E40_LEFT", "E40_RIGHT", "华北", "人力", "E40_MARK"], forbid=["NaN", "Unnamed"], kind="office-w4"),
    Case("E41_many_sheets.xlsx", ["E41_SHEET_01", "E41_SHEET_12", "E41_END"], forbid=["NaN"], kind="office-w4"),
    Case("E42_sparse_wide.xlsx", ["E42_SPARSE", "H1", "H15", "E42_TAIL", "E42_MARK"], forbid=["NaN"], kind="office-w4"),
    Case("E43_styled_table.xlsx", ["E43_STYLE", "华北", "1280", "E43_MARK"], forbid=["NaN"], kind="office-w4"),
    Case("E44_multi_image.xlsx", ["E44_IMG", "E44_MARK", "embedded-image"], forbid=["NaN"], kind="office-w4"),
    Case("E45_xls_rich.xls", ["E45_XLS", "华北", "1280", "E45_MARK", "E45_NOTE_SHEET"], forbid=["NaN"], kind="office-w4"),
    # MSG
    Case("M40_official_wave4.msg", ["Test Email Message", "test.sender@example.com"], kind="office-w4"),
    Case("M41（括号）邮件.msg", ["Test Email Message"], kind="office-w4"),
    Case("M42_empty.msg", soft=True, min_chars=0, kind="office-w4", note="空msg"),
    # PDF
    Case("PDF01_text_only.pdf", ["PDF_TEXT", "PDF_BODY_CN", "PDF_BODY_EN"], kind="pdf"),
    Case("PDF02_text_table.pdf", ["PDF_TT", "华北", "1280"], kind="pdf"),
    Case("PDF03_text_image.pdf", ["PDF_TI", "CAP_PDF_A"], kind="pdf"),
    Case("PDF04_text_image_table.pdf", ["PDF_MIX", "96%", "CAP_PDF_C", "PDF_MIX_END"], kind="pdf"),
    Case("PDF05_multi_table.pdf", ["PDF_MT1", "PDF_MT2", "PDF_MT3"], kind="pdf"),
    Case("PDF06_multi_image.pdf", ["PDF_MI", "CAP_PDF_A1", "CAP_PDF_B1"], kind="pdf"),
    Case("PDF07_multipage_sections.pdf", ["PDF_SEC", "CH01", "CH05", "PDF_SEC_END"], kind="pdf"),
    Case("PDF08_mixed_lang.pdf", ["PDF_LANG", "North China", "PDF_LANG_MARK"], kind="pdf"),
    Case("PDF09_two_column.pdf", ["PDF_COL", "LEFT_PDF09", "RIGHT_PDF09", "PDF_COL_MARK"], kind="pdf"),
    Case("PDF10_stress_pack.pdf", ["PDF_STRESS", "PDF_STRESS_END", "CAP_PDF_ST_1"], kind="pdf"),
    Case("PDF11_wide_table.pdf", ["PDF_WIDE", "列1", "列8", "PDF_WIDE_MARK"], kind="pdf"),
    Case("PDF12_merged_table.pdf", ["PDF_MERGE", "MERGE_HDR", "华北", "PDF_MERGE_MARK"], kind="pdf"),
    Case("PDF13_almost_empty.pdf", min_chars=0, soft=True, kind="pdf", note="几乎空"),
    Case(
        "PDF14_scanned_like.pdf",
        must=[],
        expect_emptyish=True,
        min_chars=0,
        kind="pdf",
        note="扫描件：无文字层属预期；可有图片占位",
    ),
    Case("PDF15_corrupt.pdf", soft=True, min_chars=0, kind="pdf", note="损坏pdf"),
    Case("PDF16_fake.pdf", soft=True, min_chars=0, kind="pdf", note="伪pdf"),
    Case("PDF17 报告 中文名.pdf", ["PDF17_NAME", "PDF17_BODY"], kind="pdf"),
    Case("PDF18_encrypted.pdf", soft=True, min_chars=0, kind="pdf", note="加密pdf：失败或无法读密文均可"),
    Case("PDF19_many_pages.pdf", ["PDF19_MANY", "PDF19_P01", "PDF19_P12", "PDF19_END"], kind="pdf"),
    Case("PDF20_table_only.pdf", ["PDF20_MARK", "华北"], kind="pdf"),
]


def evaluate(md: str, case: Case, exc: Exception | None) -> tuple[str, list[str]]:
    if case.soft:
        return "PASS", []
    if case.expect_emptyish:
        if exc is not None:
            return "PASS", []
        # 扫描件：不应抽到 SCANNED_PDF_MARKER 文字（那是画在图上的）
        if "SCANNED_PDF_MARKER" in md:
            return "WARN", ["扫描件意外抽到图内文字"]
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
    stats = {"PASS": 0, "FAIL": 0, "SKIP": 0, "WARN": 0}

    for case in CASES:
        path = SAMPLES / case.file
        if not path.exists():
            stats["SKIP"] += 1
            results.append({"file": case.file, "status": "SKIP", "kind": case.kind, "issues": ["不存在"], "note": case.note})
            print(f"SKIP {case.file}")
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
        # WARN 计为通过但记录
        if status == "WARN":
            stats["WARN"] += 1
            stats["PASS"] += 1
            status_out = "PASS"
            issues = [f"WARN: {x}" for x in issues]
        else:
            stats[status] = stats.get(status, 0) + 1
            status_out = status

        results.append(
            {
                "file": case.file,
                "status": status_out,
                "kind": case.kind,
                "issues": issues,
                "chars": len(md) if md else 0,
                "sec": elapsed,
                "note": case.note,
                "error": (str(exc)[:160] if exc else ""),
            }
        )
        print(f"{status_out} [{case.kind}] {case.file} ({elapsed}s) {issues or (str(exc)[:60] if exc else '')}")

    lines = [
        "# Office Wave4 + PDF 穷举报告",
        "",
        f"- 总计: {len(results)} | PASS: {stats['PASS']} | FAIL: {stats['FAIL']} | SKIP: {stats['SKIP']} | WARN记入PASS: {stats['WARN']}",
        "",
        "| 类型 | 文件 | 状态 | 字数 | 耗时 | 问题 |",
        "|---|---|---|---:|---:|---|",
    ]
    for r in results:
        iss = "; ".join(r["issues"]) or r.get("error") or ""
        lines.append(
            f"| {r['kind']} | `{r['file']}` | **{r['status']}** | {r.get('chars','-')} | {r.get('sec','-')} | {iss} |"
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
