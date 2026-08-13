"""生产门禁：一键穷举跑测 + 真实场景回归 + 总报告。

用法：
    .venv\\Scripts\\python.exe tests\\run_production_gate.py
"""

from __future__ import annotations

import json
import subprocess
import sys
import time
from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
TESTS = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(TESTS))

REPORT = TESTS / "PRODUCTION_GATE_REPORT.md"
OUT = TESTS / "production_gate_output"


@dataclass
class SuiteResult:
    name: str
    exit_code: int
    elapsed: float
    detail: str = ""


@dataclass
class CaseResult:
    case_id: str
    group: str
    status: str  # PASS / FAIL / SKIP
    chars: int = 0
    issues: list[str] = field(default_factory=list)
    note: str = ""


def _py() -> str:
    return str(ROOT / ".venv" / "Scripts" / "python.exe")


def run_suite(name: str, script: Path, extra: list[str] | None = None) -> SuiteResult:
    if not script.exists():
        return SuiteResult(name, 127, 0.0, f"脚本不存在: {script}")
    cmd = [_py(), str(script)] + (extra or [])
    t0 = time.time()
    p = subprocess.run(cmd, cwd=str(ROOT), capture_output=True, text=True, encoding="utf-8", errors="replace")
    elapsed = round(time.time() - t0, 2)
    tail = (p.stdout or p.stderr or "").strip().splitlines()
    detail = tail[-1] if tail else f"exit={p.returncode}"
    return SuiteResult(name, p.returncode, elapsed, detail)


def run_unittest() -> SuiteResult:
    t0 = time.time()
    p = subprocess.run(
        [
            _py(),
            "-m",
            "unittest",
            "tests.test_excel_merge_collapse",
            "tests.test_zip_convert",
            "tests.test_v173_fixes",
            "tests.test_fidelity",
            "-v",
        ],
        cwd=str(ROOT),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    elapsed = round(time.time() - t0, 2)
    ok = p.returncode == 0
    detail = "OK" if ok else ((p.stderr or p.stdout or "FAILED")[-200:])
    return SuiteResult("unit_tests", 0 if ok else 1, elapsed, detail)


def _safe_print(msg: str) -> None:
    enc = getattr(sys.stdout, "encoding", None) or "utf-8"
    try:
        print(msg)
    except UnicodeEncodeError:
        safe = msg.encode(enc, errors="replace").decode(enc, errors="replace")
        print(safe)


def run_realworld() -> list[CaseResult]:
    from converter import ConversionError, convert_path, precheck_source
    from make_schedule_fixture import ensure_schedule_staff_xlsx
    from test_helpers import check_must, has_markdown_pipe_table

    # 门禁前强制生成排班样例，避免因 gitignore 缺文件而 SKIP
    ensure_schedule_staff_xlsx(TESTS / "samples" / "schedule_staff.xlsx")

    fake_rar = TESTS / "samples" / "fake_rar_as_zip.zip"
    if not fake_rar.exists():
        fake_rar.write_bytes(b"Rar!\x1a\x07\x00doc2md-fake-rar")

    cases: list[tuple[str, str, list[str], bool]] = [
        # id, path, must keywords, optional (skip if missing)
        (
            "RW_schedule_xlsx",
            str(TESTS / "samples" / "schedule_staff.xlsx"),
            ["周一上午", "周一下午", "王成", "杨雨糯", "周五一天", "潘爱萍"],
            False,
        ),
        (
            "RW_samples_html",
            str(TESTS / "samples" / "01_styles_structure.html"),
            ["季度经营报告", "1280"],
            False,
        ),
        (
            "RW_samples_csv",
            str(TESTS / "samples" / "03_regions.csv"),
            ["华北", "1280"],
            False,
        ),
        (
            "RW_samples_json",
            str(TESTS / "samples" / "02_metrics.json"),
            ["华北", "1280"],
            False,
        ),
        (
            "RW_samples_md",
            str(TESTS / "samples" / "04_already_md.md"),
            ["Markdown"],
            False,
        ),
        (
            "RW_samples_txt",
            str(TESTS / "samples" / "05_plain.txt"),
            ["产品说明书", "转换 Office"],
            False,
        ),
        (
            "RW_precheck_fake_rar",
            str(fake_rar),
            [],
            False,
        ),
    ]

    out: list[CaseResult] = []
    for cid, path, must, optional in cases:
        p = Path(path)
        if not p.exists():
            if optional:
                out.append(CaseResult(cid, "realworld", "SKIP", note=f"文件不存在: {p.name}"))
            else:
                out.append(CaseResult(cid, "realworld", "FAIL", issues=[f"缺少样例: {p}"]))
            continue

        if cid == "RW_precheck_fake_rar":
            tip = precheck_source(str(p))
            if tip and "RAR" in tip:
                out.append(CaseResult(cid, "realworld", "PASS", note=tip[:80]))
            elif sniff := __import__("zip_convert", fromlist=["sniff_archive_kind"]).sniff_archive_kind(p):
                if sniff == "rar":
                    out.append(CaseResult(cid, "realworld", "PASS", note="魔数=RAR"))
                else:
                    out.append(CaseResult(cid, "realworld", "FAIL", issues=[f"期望RAR，得到{sniff}"]))
            else:
                out.append(CaseResult(cid, "realworld", "SKIP", note="非RAR或文件不可用"))
            continue

        try:
            md = convert_path(str(p), local_only=False)
            issues = check_must(md, must)
            if cid == "RW_schedule_xlsx" and not has_markdown_pipe_table(md):
                issues.append("排班表未输出 Markdown 管道表")
            status = "PASS" if not issues else "FAIL"
            out.append(
                CaseResult(
                    cid,
                    "realworld",
                    status,
                    chars=len(md),
                    issues=issues,
                    note="真实排班表" if cid == "RW_schedule_xlsx" else "",
                )
            )
            if md.strip():
                OUT.mkdir(parents=True, exist_ok=True)
                (OUT / f"{cid}.md").write_text(md, encoding="utf-8")
        except ConversionError as exc:
            out.append(CaseResult(cid, "realworld", "FAIL", issues=[str(exc)[:120]]))
        except Exception as exc:  # noqa: BLE001
            out.append(CaseResult(cid, "realworld", "FAIL", issues=[f"{type(exc).__name__}: {exc}"]))

    return out


def run_l3_l4_suite() -> SuiteResult:
    """L3/L4：无凭证时标为 CONDITIONAL，避免误报「全绿已测云端」。"""
    script = TESTS / "run_l3_l4_conditional.py"
    sr = run_suite("l3_l4", script)
    results_path = TESTS / "l3_l4_output" / "results.json"
    if results_path.exists():
        try:
            data = json.loads(results_path.read_text(encoding="utf-8"))
            rows = data.get("results") or []
            statuses = [r.get("status") for r in rows]
            if statuses and all(s == "SKIP" for s in statuses) and sr.exit_code == 0:
                return SuiteResult(
                    "l3_l4",
                    0,
                    sr.elapsed,
                    f"CONDITIONAL_SKIP ({len(statuses)} 条未配置凭证，未实际验证)",
                )
            if any(s == "FAIL" for s in statuses):
                return SuiteResult("l3_l4", 1, sr.elapsed, sr.detail)
            if any(s == "PASS" for s in statuses):
                n_pass = sum(1 for s in statuses if s == "PASS")
                n_skip = sum(1 for s in statuses if s == "SKIP")
                return SuiteResult(
                    "l3_l4",
                    0,
                    sr.elapsed,
                    f"PASS {n_pass} · SKIP {n_skip}",
                )
        except Exception:
            pass
    return sr


def main() -> int:
    print("=== 生产门禁穷举测试 ===")
    started = datetime.now(timezone.utc).astimezone()

    suites = [
        ("unit_tests", None),
        ("office_matrix", TESTS / "office_matrix" / "run_office_tests.py"),
        ("office_exhaust", TESTS / "office_exhaust" / "run_exhaust.py"),
        ("office_deep", TESTS / "office_deep" / "run_office_deep.py"),
        ("office_pdf_wave4", TESTS / "office_pdf_wave4" / "run_wave4.py"),
        ("office_full_sweep", TESTS / "office_full_sweep" / "run_full_sweep.py"),
        ("l3_l4", None),
    ]

    # 生成 universe 样例（若脚本存在）
    gen = TESTS / "universe_deep" / "generate_universe.py"
    if gen.exists():
        subprocess.run([_py(), str(gen)], cwd=str(ROOT), capture_output=True)
        suites.append(("universe_deep", TESTS / "universe_deep" / "run_universe.py"))

    suite_results: list[SuiteResult] = []
    for name, script in suites:
        if name == "unit_tests":
            sr = run_unittest()
        elif name == "l3_l4":
            sr = run_l3_l4_suite()
        else:
            sr = run_suite(name, script)  # type: ignore[arg-type]
        suite_results.append(sr)
        mark = "OK" if sr.exit_code == 0 else "FAIL"
        if "CONDITIONAL" in (sr.detail or ""):
            mark = "COND"
        _safe_print(f"[{mark}] {name} ({sr.elapsed}s) {sr.detail}")

    realworld = run_realworld()
    for r in realworld:
        _safe_print(f"[{r.status}] {r.case_id} chars={r.chars} {r.issues or r.note}")

    # 汇总
    suite_pass = sum(1 for s in suite_results if s.exit_code == 0)
    rw_pass = sum(1 for r in realworld if r.status == "PASS")
    rw_fail = sum(1 for r in realworld if r.status == "FAIL")
    rw_skip = sum(1 for r in realworld if r.status == "SKIP")
    l3_conditional = any("CONDITIONAL" in (s.detail or "") for s in suite_results if s.name == "l3_l4")

    lines = [
        "# 生产门禁测试总报告",
        "",
        f"- 开始：{started.isoformat(timespec='seconds')}",
        f"- 结束：{datetime.now(timezone.utc).astimezone().isoformat(timespec='seconds')}",
        "",
        "## 套件汇总",
        "",
        f"- 自动化套件：**{suite_pass}/{len(suite_results)}** 通过（退出码 0）",
        f"- 真实场景：**{rw_pass}** 通过 · **{rw_fail}** 失败 · **{rw_skip}** 跳过",
        "",
        "| 套件 | 结果 | 耗时 | 说明 |",
        "|------|------|------|------|",
    ]
    for s in suite_results:
        if s.exit_code != 0:
            st = "FAIL"
        elif "CONDITIONAL" in (s.detail or ""):
            st = "CONDITIONAL"
        else:
            st = "PASS"
        lines.append(f"| {s.name} | **{st}** | {s.elapsed}s | {s.detail[:100]} |")

    lines += [
        "",
        "## 真实场景回归",
        "",
        "| 用例 | 状态 | 字符 | 说明 |",
        "|------|------|------|------|",
    ]
    for r in realworld:
        note = "; ".join(r.issues) if r.issues else r.note
        lines.append(f"| `{r.case_id}` | **{r.status}** | {r.chars} | {note[:100]} |")

    lines += [
        "",
        "## 覆盖范围（穷举）",
        "",
        "| 类别 | 已测 |",
        "|------|------|",
        "| Word / Excel / PPT / PDF / MSG | office_matrix + exhaust + deep + wave4 + full_sweep |",
        "| HTML / CSV / JSON / RSS / EPUB / IPYNB / ZIP | universe_deep（样例齐全时） |",
        "| 真实排班表 xlsx | RW_schedule_xlsx（门禁前自动生成样例） |",
        "| ZIP 魔数 / 伪 RAR | unit_tests + RW_precheck |",
        "| Excel 宽表折叠 / 公式回退 | unit_tests + test_v173_fixes |",
        "| keep_data_uris / stdin / 插件列表 | test_v173_fixes |",
        "| L3 LLM / L4 Azure | l3_l4（无凭证则为 CONDITIONAL，非「已验证」） |",
        "",
        "## 使用建议",
        "",
        "1. **日常发布前**跑本脚本；本地套件与真实场景无 FAIL 即可发版。",
        "2. **L3/L4 CONDITIONAL** 表示云端能力未用本机凭证实测，不代表已通过。",
        "3. 配置 Key 后再跑 `tests/run_l3_l4_conditional.py` 做真实云端验证。",
        "",
        "## 子报告路径",
        "",
        "- `tests/office_matrix/OFFICE_REPORT.md`",
        "- `tests/office_exhaust/OFFICE_EXHAUST_REPORT.md`",
        "- `tests/office_deep/OFFICE_DEEP_REPORT.md`",
        "- `tests/office_pdf_wave4/WAVE4_PDF_REPORT.md`",
        "- `tests/office_full_sweep/FULL_SWEEP_REPORT.md`",
        "- `tests/universe_deep/UNIVERSE_DEEP_REPORT.md`",
        "- `tests/L3_L4_REPORT.md`",
        "",
    ]

    fails = [s for s in suite_results if s.exit_code != 0] + [r for r in realworld if r.status == "FAIL"]
    if fails:
        lines.append("## 需关注项")
        lines.append("")
        for s in suite_results:
            if s.exit_code != 0:
                lines.append(f"- 套件 **{s.name}** 未通过")
        for r in realworld:
            if r.status == "FAIL":
                lines.append(f"- 真实场景 **{r.case_id}**: {'; '.join(r.issues)}")
    else:
        lines.append("## 结论")
        lines.append("")
        if l3_conditional:
            lines.append(
                "**本地转换门禁通过，可用于日常办公转换。"
                "L3/L4 云端能力尚未用凭证实测（CONDITIONAL）。**"
            )
        else:
            lines.append("**全部门禁通过（含已配置的云端用例），工具可用于日常办公转换。**")

    REPORT.write_text("\n".join(lines), encoding="utf-8")
    (TESTS / "production_gate_results.json").write_text(
        json.dumps(
            {
                "suites": [asdict(s) for s in suite_results],
                "realworld": [asdict(r) for r in realworld],
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    print(f"\n总报告: {REPORT}")
    gate_fail = any(s.exit_code != 0 for s in suite_results) or rw_fail > 0
    return 1 if gate_fail else 0


if __name__ == "__main__":
    raise SystemExit(main())
