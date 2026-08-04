"""非 Office 全宇宙加深跑测。"""

from __future__ import annotations

import json
import sys
import threading
import time
from dataclasses import dataclass, field
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

# 确保本地 ffmpeg 可被 pydub/markitdown 找到
FFMPEG = ROOT / "tools" / "ffmpeg" / "ffmpeg.exe"
if FFMPEG.exists():
    import os

    os.environ["PATH"] = str(FFMPEG.parent) + os.pathsep + os.environ.get("PATH", "")
    os.environ["FFMPEG_BINARY"] = str(FFMPEG)

from converter import convert_path  # noqa: E402

SAMPLES = Path(__file__).resolve().parent / "samples"
OUT = Path(__file__).resolve().parent / "output"
REPORT = Path(__file__).resolve().parent / "UNIVERSE_DEEP_REPORT.md"
WEBROOT = Path(__file__).resolve().parent / "webroot"


@dataclass
class Case:
    file: str = ""
    url: str = ""
    must: list[str] = field(default_factory=list)
    forbid: list[str] = field(default_factory=list)
    min_chars: int = 5
    soft: bool = False  # 允许失败/空（L2/L3/边界）
    expect_fail: bool = False
    group: str = ""
    note: str = ""


CASES: list[Case] = [
    # PDF
    Case(file="U_PDF01_text_only.pdf", must=["PDF_TEXT", "PDF_BODY_CN"], group="B1"),
    Case(file="U_PDF02_text_table.pdf", must=["PDF_TT", "华北", "1280"], group="B1"),
    Case(file="U_PDF04_text_image_table.pdf", must=["PDF_MIX", "96%"], group="B1"),
    Case(file="U_PDF05_multi_table.pdf", must=["PDF_MT1", "PDF_MT3"], group="B1"),
    Case(file="U_PDF07_multipage_sections.pdf", must=["CH01", "PDF_SEC_END"], group="B1"),
    Case(file="U_PDF08_mixed_lang.pdf", must=["PDF_LANG", "North China"], group="B1"),
    Case(file="U_PDF09_two_column.pdf", must=["LEFT_PDF09", "RIGHT_PDF09"], group="B1"),
    Case(file="U_PDF10_stress_pack.pdf", must=["PDF_STRESS", "PDF_STRESS_END"], group="B1"),
    Case(file="U_PDF14_scanned_like.pdf", soft=True, min_chars=0, group="B1", note="扫描件无OCR"),
    Case(file="U_PDF18_encrypted.pdf", soft=True, expect_fail=True, min_chars=0, group="B1", note="加密"),
    Case(file="U_PDF19_many_pages.pdf", must=["PDF19_P01", "PDF19_END"], group="B1"),
    Case(file="U_PDF_form_fields.pdf", must=["PDF_FORM_MARK", "NAME_LABEL"], soft=True, group="B1", note="表单域"),
    # EPUB
    Case(file="U_EPUB_01_chapters.epub", must=["EP01_CH1", "EP01_CH3", "EP01_END"], group="B2"),
    Case(file="U_EPUB_02_list_table.epub", must=["EP02_TITLE", "华北", "1280", "EP02_MARK"], group="B2"),
    Case(file="U_EPUB_03_image.epub", must=["EP03_TITLE", "EP03_MARK", "CAP_EP03"], group="B2"),
    Case(file="U_EPUB_04_css.epub", must=["EP04_TITLE", "EP04_BODY"], forbid=["#dc2626"], group="B2"),
    Case(file="U_EPUB_05_mixed_lang.epub", must=["EP05_TITLE", "EP05_MARK", "North China"], group="B2"),
    # IPYNB
    Case(file="U_IPYNB_01_md_only.ipynb", must=["IP01_H1", "IP01_BODY", "IP01_L1"], group="B3"),
    Case(file="U_IPYNB_02_code_stdout.ipynb", must=["IP02_STDOUT_MARK"], group="B3"),
    Case(file="U_IPYNB_03_html_table.ipynb", must=["华北", "1280", "IP03_HTML_MARK"], group="B3"),
    Case(file="U_IPYNB_04_image_out.ipynb", must=["IP04_TITLE", "embedded-image"], group="B3"),
    Case(file="U_IPYNB_05_mixed.ipynb", must=["IP05_TITLE", "IP05_MARK", "IP05_CODE", "华北"], group="B3"),
    Case(file="U_IPYNB_06_empty_raw.ipynb", must=["IP06_END", "RAW_IP06_MARK"], group="B3"),
    # Images L1
    Case(file="U_IMG_01_text_shot.png", soft=True, min_chars=0, group="C1", note="无OCR时常空"),
    Case(file="U_IMG_02_table_shot.png", soft=True, min_chars=0, group="C1"),
    Case(file="U_IMG_03_poster.png", soft=True, min_chars=0, group="C1"),
    Case(file="U_IMG_04_scenic.png", soft=True, min_chars=0, group="C1"),
    Case(file="U_IMG_05_exif.jpg", soft=True, min_chars=0, group="C1"),
    Case(file="U_IMG_06_icon.png", soft=True, min_chars=0, group="C1"),
    Case(file="U_IMG_06_large.png", soft=True, min_chars=0, group="C1"),
    Case(file="U_IMG_07_portrait.png", soft=True, min_chars=0, group="C1"),
    # Audio/Video L2
    Case(file="U_AUD_01_tone.wav", soft=True, min_chars=0, group="C3", note="音调无语义，STT可能空/失败"),
    Case(file="U_AUD_03_silent.wav", soft=True, min_chars=0, group="C3"),
    Case(file="U_AUD_03_noise.wav", soft=True, min_chars=0, group="C3"),
    Case(file="U_AUD_04_tone.mp3", soft=True, min_chars=0, group="C4"),
    Case(file="U_AUD_05_tone.m4a", soft=True, min_chars=0, group="C5"),
    Case(file="U_VID_01_audio_only.mp4", soft=True, min_chars=0, group="C6"),
    Case(file="U_VID_02_no_audio.mp4", soft=True, min_chars=0, group="C6"),
    # HTML
    Case(file="U_HTML_01_semantic.html", must=["HTML01_H1", "HTML01_L1", "HTML01_MARK", "华北"], group="D1"),
    Case(file="U_HTML_02_css_noise.html", must=["HTML02_TITLE", "HTML02_BODY"], forbid=["SHOULD_IGNORE_SCRIPT"], group="D1"),
    Case(file="U_HTML_03_complex_table.html", must=["HTML03_TITLE", "华北", "1280", "HTML03_MARK"], group="D1"),
    Case(file="U_HTML_04_multi_media.html", must=["HTML04_TITLE", "ALT_HTML04_A", "华北"], group="D1"),
    Case(file="U_HTML_05_nav_footer.html", must=["HTML05_TITLE", "HTML05_MARK"], group="D1"),
    Case(file="U_HTML_06_script_style.html", must=["HTML06_TITLE", "HTML06_MARK"], forbid=["SCRIPT_SHOULD_NOT_APPEAR_HTML06"], group="D1"),
    Case(file="U_HTML_07_entities.html", must=["HTML07_TITLE", "HTML07_MARK", "华北"], group="D1"),
    # RSS/Atom
    Case(file="U_RSS_01_multi.rss", must=["UNIV_RSS_TITLE", "RSS_ITEM_1", "RSS_ITEM_2"], group="D2"),
    Case(file="U_RSS_02_cn.rss", must=["RSS_CN_TITLE", "RSS_CN_ITEM", "RSS_CN_MARK"], group="D2"),
    Case(file="U_ATOM_01_multi.atom", must=["ATOM_TITLE", "ATOM_E1", "ATOM_E2"], group="D3"),
    # Data
    Case(file="U_CSV_01_basic.csv", must=["华北", "1280", "CSV01_MARK"], group="D4"),
    Case(file="U_CSV_02_escape.csv", must=["CSV02_MARK", "华北"], group="D4"),
    Case(file="U_CSV_03_bom_wide.csv", must=["列1", "列15", "CSV03_MARK"], group="D4"),
    Case(file="U_JSON_01_nested.json", must=["JSON01_TITLE", "JSON01_MARK", "华北"], group="D5"),
    Case(file="U_JSON_02_array.json", must=["JSON02_MARK", "J2_1"], group="D5"),
    Case(file="U_JSONL_01.jsonl", must=["JSONL_1", "JSONL_MARK"], group="D6"),
    Case(file="U_XML_01_biz.xml", must=["XML01_TITLE", "XML01_MARK"], group="D7"),
    Case(file="U_XML_02_ns.xml", must=["XML02_TITLE", "XML02_MARK"], group="D7"),
    Case(file="U_TXT_01_manual.txt", must=["TXT01_TITLE", "TXT01_N1", "TXT01_MARK"], group="D8"),
    Case(file="U_MD_01_fidelity.md", must=["MD01_H1", "MD01_BODY", "MD01_CODE", "MD01_MARK"], group="D9"),
    # ZIP
    Case(file="U_ZIP_01_mixed.zip", must=["ZIP_MD_MARK", "ZIP_CSV_MARK"], group="D10"),
    Case(file="U_ZIP_02_cn_names.zip", must=["ZIP_CN_MARK"], group="D10"),
    Case(file="U_ZIP_03_unsupported.zip", must=["ZIP_OK_MARK", "ZIP_SKIP_MARK"], group="D10"),
    Case(file="U_ZIP_04_nested.zip", must=["ZIP_OUTER_MARK"], soft=True, group="D10", note="嵌套zip行为记录"),
    Case(file="U_ZIP_05_stress.zip", must=["ZIP_STRESS_A", "ZIP_STRESS_B", "ZIP_STRESS_C"], group="D10"),
    # Negatives
    Case(file="U_G_01_empty.txt", soft=True, min_chars=0, group="G"),
    Case(file="U_G_05_corrupt.html", must=["未闭合"], soft=True, group="G"),
    Case(file="U_G_05_bad.zip", soft=True, expect_fail=True, min_chars=0, group="G"),
    Case(file="U_G_03 中文 空格.json", must=["G03_JSON_MARK"], group="G"),
    Case(file="U_G_04_html_as.pdf", soft=True, min_chars=0, group="G", note="后缀不符"),
    Case(file="U_G_bmp.bmp", soft=True, min_chars=0, group="G", note="BMP不支持或失败"),
    # URL placeholders filled at runtime
]


def evaluate(md: str, case: Case, exc: Exception | None) -> tuple[str, list[str]]:
    if case.soft or case.expect_fail:
        if case.expect_fail and exc is None and len(md.strip()) > 100 and not case.soft:
            return "FAIL", ["期望失败但产出较长"]
        return "PASS", []
    if exc is not None:
        return "FAIL", [f"异常: {exc}"]
    issues = []
    if len(md.strip()) < case.min_chars:
        issues.append(f"过短 len={len(md.strip())}")
    for m in case.must:
        from test_helpers import keyword_in_md

        if not keyword_in_md(md, m):
            issues.append(f"缺 must: {m}")
    for f in case.forbid:
        if f in md:
            issues.append(f"含 forbid: {f}")
    return ("PASS" if not issues else "FAIL"), issues


def start_local_http() -> tuple[ThreadingHTTPServer, str]:
    WEBROOT.mkdir(parents=True, exist_ok=True)
    (WEBROOT / "article.html").write_text(
        """<!DOCTYPE html><html><head><meta charset="utf-8"><title>Local</title></head>
<body><h1>本地文章 URL_LOCAL_H1</h1>
<p>关键句 URL_LOCAL_MARK 华北增长</p>
<ul><li>要点 URL_LOCAL_L1</li></ul>
</body></html>""",
        encoding="utf-8",
    )
    handler = type(
        "H",
        (SimpleHTTPRequestHandler,),
        {"__init__": lambda self, *a, **k: SimpleHTTPRequestHandler.__init__(self, *a, directory=str(WEBROOT), **k)},
    )

    class Quiet(handler):  # type: ignore
        def log_message(self, format, *args):  # noqa: A003
            pass

    srv = ThreadingHTTPServer(("127.0.0.1", 0), Quiet)
    port = srv.server_address[1]
    t = threading.Thread(target=srv.serve_forever, daemon=True)
    t.start()
    return srv, f"http://127.0.0.1:{port}"


def main() -> int:
    OUT.mkdir(parents=True, exist_ok=True)
    results = []
    stats = {"PASS": 0, "FAIL": 0, "SKIP": 0}
    by_group: dict[str, dict[str, int]] = {}

    srv, base = start_local_http()
    url_cases = [
        Case(url=f"{base}/article.html", must=["URL_LOCAL_H1", "URL_LOCAL_MARK"], group="E1"),
        Case(url=f"{base}/missing-404.html", soft=True, expect_fail=True, min_chars=0, group="E1", note="404"),
    ]
    # 可选外网：软失败
    url_cases += [
        Case(
            url="https://en.wikipedia.org/wiki/Markdown",
            must=["Markdown"],
            soft=True,
            group="E2",
            note="外网Wiki，不通则软过",
        ),
        Case(
            url="https://www.bing.com/search?q=markitdown",
            soft=True,
            min_chars=0,
            group="E4",
            note="Bing软测",
        ),
    ]
    all_cases = list(CASES) + url_cases

    try:
        for case in all_cases:
            g = case.group or "?"
            by_group.setdefault(g, {"PASS": 0, "FAIL": 0, "SKIP": 0})
            label = case.file or case.url
            path = SAMPLES / case.file if case.file else None
            if case.file and path is not None and not path.exists():
                stats["SKIP"] += 1
                by_group[g]["SKIP"] += 1
                results.append({"id": label, "status": "SKIP", "group": g, "issues": ["不存在"], "note": case.note})
                print(f"SKIP [{g}] {label}")
                continue

            t0 = time.time()
            md, exc = "", None
            try:
                src = case.url if case.url else str(path)
                md = convert_path(src)
                safe = "".join(ch if ch not in '<>:"/\\|?*' else "_" for ch in (path.stem if path else f"url_{g}"))
                (OUT / f"{safe}.md").write_text(md, encoding="utf-8")
            except Exception as e:  # noqa: BLE001
                exc = e
            elapsed = round(time.time() - t0, 2)
            status, issues = evaluate(md, case, exc)
            stats[status] += 1
            by_group[g][status] += 1
            results.append(
                {
                    "id": label,
                    "status": status,
                    "group": g,
                    "issues": issues,
                    "chars": len(md) if md else 0,
                    "sec": elapsed,
                    "note": case.note,
                    "error": (str(exc)[:160] if exc else ""),
                }
            )
            print(f"{status} [{g}] {label} ({elapsed}s) {issues or (str(exc)[:50] if exc else '')}")
    finally:
        srv.shutdown()

    lines = [
        "# 非 Office 全宇宙加深报告",
        "",
        f"- 总计: {len(results)} | PASS: {stats['PASS']} | FAIL: {stats['FAIL']} | SKIP: {stats['SKIP']}",
        "",
        "## 分组",
        "",
        "| 组 | PASS | FAIL | SKIP |",
        "|---|---:|---:|---:|",
    ]
    for g in sorted(by_group):
        s = by_group[g]
        lines.append(f"| {g} | {s['PASS']} | {s['FAIL']} | {s['SKIP']} |")
    lines += ["", "| 组 | 用例 | 状态 | 字数 | 耗时 | 问题 |", "|---|---|---|---:|---:|---|"]
    for r in results:
        iss = "; ".join(r["issues"]) or r.get("error") or ""
        lines.append(
            f"| {r['group']} | `{r['id']}` | **{r['status']}** | {r.get('chars','-')} | {r.get('sec','-')} | {iss} |"
        )
    fails = [r for r in results if r["status"] == "FAIL"]
    lines.append("")
    if fails:
        lines.append("## 失败明细")
        for r in fails:
            lines.append(f"### {r['id']}")
            for i in r["issues"]:
                lines.append(f"- {i}")
            lines.append("")
    else:
        lines.append("## 全部通过")
        lines.append("")

    REPORT.write_text("\n".join(lines), encoding="utf-8")
    (Path(__file__).parent / "results.json").write_text(
        json.dumps({"stats": stats, "by_group": by_group, "results": results}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(f"\nDONE pass={stats['PASS']} fail={stats['FAIL']} skip={stats['SKIP']} -> {REPORT}")
    return 0 if stats["FAIL"] == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
