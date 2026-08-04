"""跑 Office 穷举测试并出报告。"""

from __future__ import annotations

import json
import re
import sys
from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from converter import convert_path  # noqa: E402
from test_helpers import has_image_trace, has_markdown_pipe_table, has_pdf_text_table

SAMPLES = Path(__file__).resolve().parent / "samples"
OUTPUT = Path(__file__).resolve().parent / "output"
REPORT = Path(__file__).resolve().parent / "OFFICE_REPORT.md"

# case_id -> 期望关键词（至少命中）
EXPECT = {
    "W01_text_only": ["WORD_TEXT_ONLY", "WORD_BODY_MARKER", "WORD_RED_MARKER"],
    "W02_text_table": ["WORD_TEXT_TABLE", "WORD_TT_MARKER", "华北", "1280"],
    "W03_text_image": ["WORD_TEXT_IMG", "WORD_TI_MARKER", "CAPTION_NORTH"],
    "W04_text_image_table": ["WORD_MIXED", "WORD_MIX_MARKER", "96%", "CHART_CAPTION"],
    "W05_multi_table": ["WORD_MULTI_TABLE", "WORD_MT1", "WORD_MT2", "人力"],
    "W06_multi_image": ["WORD_MULTI_IMG", "CAPTION_N1", "CAPTION_S1", "CAPTION_L1"],
    "W07_merged_table": ["WORD_MERGE_TABLE", "MERGE_HDR", "华北", "海外"],
    "W08_header_footer": ["WORD_HF", "BODY_WORD_HF"],
    "W09_hyperlink": ["WORD_LINK", "WORD_LINK_BODY"],
    "W10_long_headings": ["WORD_LONG", "CHAPTER_1", "CHAPTER_3", "SEC_2"],
    "E01_text_only": ["EXCEL_TEXT_ONLY", "EXCEL_BODY_MARKER"],
    "E02_text_table": ["EXCEL_TEXT_TABLE", "EXCEL_TT_MARKER", "华北", "1280"],
    "E03_text_image": ["EXCEL_TEXT_IMG", "EXCEL_TI_MARKER", "CAPTION_EXCEL_NORTH"],
    "E04_text_image_table": ["EXCEL_MIXED", "EXCEL_MIX_MARKER", "96%"],
    "E05_multi_table_side": ["EXCEL_MT_LEFT", "EXCEL_MT_RIGHT", "华北", "人力"],
    "E06_multi_sheet": ["SHEET_NORTH", "SHEET_OVERSEAS", "SHEET_SUMMARY"],
    "E07_merged": ["EXCEL_MERGE_HDR", "华北", "海外"],
    "E08_number_formats": ["EXCEL_FMT_MARKER"],
    "E09_formula": ["EXCEL_FORMULA_MARKER"],
    "E10_multi_image": ["EXCEL_MULTI_IMG", "EXCEL_MI_CAPTION"],
    "P01_text_only": ["PDF_TEXT_ONLY", "PDF_BODY_MARKER"],
    "P02_text_table": ["PDF_TEXT_TABLE", "PDF_TT_MARKER", "1280"],
    "P03_text_image": ["PDF_TEXT_IMG", "PDF_TI_MARKER", "CAPTION_PDF_NORTH"],
    "P04_text_image_table": ["PDF_MIXED", "PDF_MIX_MARKER", "96%", "CHART_PDF_CAPTION"],
    "P05_multi_table": ["PDF_MULTI_TABLE", "PDF_MT1", "PDF_MT2"],
    "P06_multi_image": ["PDF_MULTI_IMG", "CAPTION_PDF_N1", "CAPTION_PDF_S1"],
    "P07_multipage": ["PDF_MULTIPAGE", "PDF_PAGE_1", "PDF_PAGE_LAST"],
    "P08_scanned_like": [],  # 扫描件：允许无文字，单独判定
    "P09_two_column": ["PDF_TWO_COL", "PDF_LEFT_COL", "PDF_RIGHT_COL"],
}

# 含图用例：检查是否有图片痕迹（markdown 图 / data uri / 占位）
IMAGE_CASES = {
    "W03_text_image",
    "W04_text_image_table",
    "W06_multi_image",
    "E03_text_image",
    "E04_text_image_table",
    "E10_multi_image",
    "P03_text_image",
    "P04_text_image_table",
    "P06_multi_image",
}


@dataclass
class CaseResult:
    case_id: str
    fmt: str
    ok: bool
    chars: int = 0
    hit: int = 0
    need: int = 0
    missing: list[str] = field(default_factory=list)
    has_table: bool = False
    has_image_trace: bool = False
    notes: list[str] = field(default_factory=list)
    error: str = ""


def image_trace(md: str) -> bool:
    return has_image_trace(md)


# PDF 由 pdfminer 提取时常为纯文本表，不要求 | 管道
PDF_TABLE_CASES = {
    "P02_text_table",
    "P04_text_image_table",
    "P05_multi_table",
}


def _table_ok(case_id: str, fmt: str, md: str) -> bool:
    if has_markdown_pipe_table(md):
        return True
    if fmt == ".pdf" and case_id in PDF_TABLE_CASES:
        return has_pdf_text_table(md)
    return False


def eval_one(path: Path) -> CaseResult:
    case_id = path.stem
    fmt = path.suffix.lower()
    expect = EXPECT.get(case_id, [])
    res = CaseResult(case_id=case_id, fmt=fmt, ok=False, need=len(expect))
    try:
        md = convert_path(str(path))
        out = OUTPUT / f"{case_id}.md"
        out.write_text(md, encoding="utf-8")
        res.chars = len(md)
        res.has_table = _table_ok(case_id, fmt, md)
        res.has_image_trace = image_trace(md)
        from test_helpers import keyword_in_md

        missing = [k for k in expect if not keyword_in_md(md, k)]
        res.missing = missing
        res.hit = len(expect) - len(missing)
        # 扫描件特殊：无文字算“符合本地预期”，仍记 note
        if case_id == "P08_scanned_like":
            if "SCANNED_PDF_MARKER" in md:
                res.ok = True
                res.notes.append("扫描件意外抽到文字")
            else:
                res.ok = True
                res.notes.append("扫描件无文字层（本地预期）")
        else:
            res.ok = len(missing) == 0
        if case_id in IMAGE_CASES and not res.has_image_trace:
            res.notes.append("未检测到图片痕迹（可能只保留图注文字）")
        if ("table" in case_id or "merged" in case_id or case_id.endswith("_table")
                or "TABLE" in case_id.upper() or case_id in {
                    "W02_text_table", "W04_text_image_table", "W05_multi_table",
                    "W07_merged_table", "E02_text_table", "E04_text_image_table",
                    "E05_multi_table_side", "E07_merged", "P02_text_table",
                    "P04_text_image_table", "P05_multi_table"}):
            if not res.has_table:
                note = "期望有表格结构但未检测到"
                if fmt == ".pdf":
                    note += "（PDF 可能为纯文本布局，仅记说明）"
                    res.notes.append(note)
                else:
                    res.notes.append(note)
                    if res.ok:
                        res.ok = False
    except Exception as exc:  # noqa: BLE001
        res.error = f"{type(exc).__name__}: {exc}"
        res.ok = False
    return res


def main():
    OUTPUT.mkdir(parents=True, exist_ok=True)
    files = sorted(SAMPLES.glob("*.docx")) + sorted(SAMPLES.glob("*.xlsx")) + sorted(SAMPLES.glob("*.pdf"))
    results = []
    print(f"cases={len(files)}")
    for path in files:
        print(f"-> {path.name}")
        r = eval_one(path)
        status = "PASS" if r.ok else f"FAIL miss={r.missing or r.error}"
        print(f"   {status} chars={r.chars} notes={r.notes}")
        results.append(r)

    by = {"word": [], "excel": [], "pdf": []}
    for r in results:
        if r.fmt == ".docx":
            by["word"].append(r)
        elif r.fmt == ".xlsx":
            by["excel"].append(r)
        else:
            by["pdf"].append(r)

    lines = [
        "# Office 穷举测试报告",
        "",
        f"- 时间：{datetime.now(timezone.utc).astimezone().isoformat(timespec='seconds')}",
        f"- 合计：{len(results)} · 通过：{sum(1 for r in results if r.ok)} · 失败：{sum(1 for r in results if not r.ok)}",
        "",
    ]
    for name, items in by.items():
        ok = sum(1 for r in items if r.ok)
        lines += [f"## {name.upper()}（{ok}/{len(items)}）", ""]
        lines += ["| 用例 | 结果 | 命中 | 表 | 图痕迹 | 说明 |", "|------|------|------|----|--------|------|"]
        for r in items:
            st = "PASS" if r.ok else "FAIL"
            note = "; ".join(r.notes) if r.notes else (r.error[:60] if r.error else "")
            if r.missing:
                note = (note + "; " if note else "") + "缺:" + ",".join(r.missing[:5])
            lines.append(
                f"| `{r.case_id}` | {st} | {r.hit}/{r.need} | "
                f"{'Y' if r.has_table else 'N'} | {'Y' if r.has_image_trace else 'N'} | {note} |"
            )
        lines.append("")

    REPORT.write_text("\n".join(lines), encoding="utf-8")
    (Path(__file__).parent / "office_results.json").write_text(
        json.dumps([asdict(r) for r in results], ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(json.dumps({
        "total": len(results),
        "pass": sum(1 for r in results if r.ok),
        "fail": sum(1 for r in results if not r.ok),
        "report": str(REPORT),
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
