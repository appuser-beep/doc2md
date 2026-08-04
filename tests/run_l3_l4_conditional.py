"""L3（LLM/OCR）与 L4（Azure）条件测试：有凭证则执行，否则 SKIP。"""

from __future__ import annotations

import json
import sys
import time
from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

OUT = Path(__file__).resolve().parent / "l3_l4_output"
REPORT = Path(__file__).resolve().parent / "L3_L4_REPORT.md"


@dataclass
class TierResult:
    case_id: str
    tier: str  # L3 / L4
    status: str  # PASS / FAIL / SKIP
    chars: int = 0
    elapsed: float = 0.0
    reason: str = ""
    notes: list[str] = field(default_factory=list)
    preview: str = ""


def _probe_l3() -> dict:
    from llm_settings import load_settings

    s = load_settings()
    return {
        "llm_ready": s.is_ready,
        "plugins": bool(s.enable_plugins),
        "model": s.model if s.enabled else "",
    }


def _probe_l4() -> dict:
    from azure_settings import load_settings

    s = load_settings()
    return {
        "docintel_ready": s.docintel_ready,
        "cu_ready": s.cu_ready,
        "cu_analyzer": s.cu_analyzer_id.strip() if s.cu_ready else "",
    }


def _make_text_png(path: Path) -> None:
    from PIL import Image, ImageDraw

    path.parent.mkdir(parents=True, exist_ok=True)
    im = Image.new("RGB", (480, 160), (255, 255, 255))
    dr = ImageDraw.Draw(im)
    dr.text((16, 24), "L3_OCR_MARKER", fill=(0, 0, 0))
    dr.text((16, 72), "North China revenue 1280", fill=(0, 0, 0))
    im.save(path)


def _run_l3_llm_image() -> TierResult:
    case_id = "L3-01_llm_image_description"
    probe = _probe_l3()
    if not probe["llm_ready"]:
        return TierResult(case_id, "L3", "SKIP", reason="未配置 LLM（需启用 + API Key）")

    from markitdown import MarkItDown
    from llm_settings import markitdown_llm_kwargs

    sample = OUT / "L3_text_screenshot.png"
    _make_text_png(sample)
    t0 = time.time()
    try:
        md = MarkItDown(**markitdown_llm_kwargs())
        text = (md.convert(str(sample)).text_content or "").strip()
    except Exception as exc:  # noqa: BLE001
        return TierResult(
            case_id, "L3", "FAIL", elapsed=time.time() - t0, reason=str(exc)[:200]
        )

    notes = []
    if len(text) < 20:
        notes.append("输出较短，可能 LLM 未返回描述")
    # 有 LLM 时应比纯元数据更长，或含 marker 语义
    ok = len(text) >= 20 or "1280" in text or "L3" in text.upper() or "revenue" in text.lower()
    return TierResult(
        case_id,
        "L3",
        "PASS" if ok else "FAIL",
        chars=len(text),
        elapsed=round(time.time() - t0, 2),
        reason="" if ok else "LLM 图片描述输出不足",
        notes=notes,
        preview=text[:400],
    )


def _run_l3_ocr_plugin() -> TierResult:
    case_id = "L3-02_markitdown_ocr_plugin"
    probe = _probe_l3()
    if not probe["llm_ready"]:
        return TierResult(case_id, "L3", "SKIP", reason="未配置 LLM")
    if not probe["plugins"]:
        return TierResult(case_id, "L3", "SKIP", reason="未启用 enable_plugins（markitdown-ocr）")

    # 用含嵌图 DOCX 测 OCR 插件是否加载（不强制抽到文字）
    sample = Path(__file__).resolve().parent / "office_matrix" / "samples" / "W03_text_image.docx"
    if not sample.exists():
        return TierResult(case_id, "L3", "SKIP", reason=f"样例不存在：{sample.name}")

    from converter import convert_path

    t0 = time.time()
    try:
        text = convert_path(sample).strip()
    except Exception as exc:  # noqa: BLE001
        return TierResult(case_id, "L3", "FAIL", elapsed=time.time() - t0, reason=str(exc)[:200])

    ok = "W03" in text or "WORD" in text or len(text) > 30
    return TierResult(
        case_id,
        "L3",
        "PASS" if ok else "FAIL",
        chars=len(text),
        elapsed=round(time.time() - t0, 2),
        reason="" if ok else "OCR 插件路径转换失败",
        notes=["插件已启用；嵌图 OCR 质量取决于 LLM"],
        preview=text[:400],
    )


def _run_l4_docintel() -> TierResult:
    case_id = "L4-01_docintel_scanned_pdf"
    probe = _probe_l4()
    if not probe["docintel_ready"]:
        return TierResult(case_id, "L4", "SKIP", reason="未配置 docintel_endpoint")

    candidates = [
        Path(__file__).resolve().parent / "office_matrix" / "samples" / "P08_scanned_like.pdf",
        Path(__file__).resolve().parent / "office_pdf_wave4" / "samples" / "PDF14_scanned_like.pdf",
    ]
    sample = next((p for p in candidates if p.exists()), None)
    if sample is None:
        return TierResult(case_id, "L4", "SKIP", reason="无扫描件 PDF 样例")

    from converter import convert_path

    t0 = time.time()
    try:
        text = convert_path(sample).strip()
    except Exception as exc:  # noqa: BLE001
        return TierResult(case_id, "L4", "FAIL", elapsed=time.time() - t0, reason=str(exc)[:200])

    # Doc Intel 对扫描件应比本地 pdfminer 有更多输出（本地通常 0 字）
    ok = len(text) > 10
    return TierResult(
        case_id,
        "L4",
        "PASS" if ok else "FAIL",
        chars=len(text),
        elapsed=round(time.time() - t0, 2),
        reason="" if ok else "Doc Intel 未抽出扫描件文字（检查 endpoint/配额）",
        notes=[f"样例：{sample.name}"],
        preview=text[:400],
    )


def _run_l4_cu_video_or_audio() -> TierResult:
    case_id = "L4-02_cu_multimodal"
    probe = _probe_l4()
    if not probe["cu_ready"]:
        return TierResult(case_id, "L4", "SKIP", reason="未配置 cu_endpoint")

    # 优先 wav（比视频轻）；若无则试 mp4
    media_dir = Path(__file__).resolve().parent / "full_matrix" / "phase4" / "media"
    wav = media_dir / "E3_tone.wav"
    if not wav.exists():
        try:
            import wave
            import math
            import struct

            media_dir.mkdir(parents=True, exist_ok=True)
            fr = 16000
            with wave.open(str(wav), "w") as w:
                w.setnchannels(1)
                w.setsampwidth(2)
                w.setframerate(fr)
                for i in range(fr):
                    val = int(32767 * 0.2 * math.sin(2 * math.pi * 440 * i / fr))
                    w.writeframes(struct.pack("<h", val))
        except Exception as exc:  # noqa: BLE001
            return TierResult(case_id, "L4", "SKIP", reason=f"无法生成 WAV：{exc}")

    from converter import convert_path

    t0 = time.time()
    try:
        text = convert_path(wav).strip()
    except Exception as exc:  # noqa: BLE001
        return TierResult(case_id, "L4", "FAIL", elapsed=time.time() - t0, reason=str(exc)[:200])

    ok = len(text) > 0  # CU 应返回结构化/YAML 或转写
    return TierResult(
        case_id,
        "L4",
        "PASS" if ok else "FAIL",
        chars=len(text),
        elapsed=round(time.time() - t0, 2),
        reason="" if ok else "CU 音频转换无输出",
        notes=[f"样例：{wav.name}", "完整视频理解需真实 MP4 + CU 配额"],
        preview=text[:400],
    )


def _write_report(results: list[TierResult], l3_probe: dict, l4_probe: dict) -> None:
    ok = sum(1 for r in results if r.status == "PASS")
    skip = sum(1 for r in results if r.status == "SKIP")
    fail = sum(1 for r in results if r.status == "FAIL")
    lines = [
        "# L3/L4 条件测试报告",
        "",
        f"- 时间：{datetime.now(timezone.utc).astimezone().isoformat(timespec='seconds')}",
        f"- 合计：{len(results)} · PASS：{ok} · FAIL：{fail} · SKIP：{skip}",
        "",
        "## 环境探测（不含密钥）",
        "",
        f"- L3 LLM 就绪：{l3_probe.get('llm_ready')} · plugins：{l3_probe.get('plugins')} · model：{l3_probe.get('model') or '(未启用)'}",
        f"- L4 DocIntel 就绪：{l4_probe.get('docintel_ready')} · CU 就绪：{l4_probe.get('cu_ready')}",
        "",
        "| 用例 | 层级 | 状态 | 字符 | 耗时 | 说明 |",
        "|------|------|------|------|------|------|",
    ]
    for r in results:
        note = r.reason or "; ".join(r.notes[:2])
        lines.append(
            f"| `{r.case_id}` | {r.tier} | **{r.status}** | {r.chars} | {r.elapsed}s | {note} |"
        )
    lines.append("")
    for r in results:
        lines.append(f"## {r.case_id}")
        lines.append(f"- 状态：**{r.status}**")
        if r.reason:
            lines.append(f"- 原因：{r.reason}")
        if r.preview:
            lines.append("")
            lines.append("```")
            lines.append(r.preview)
            lines.append("```")
        lines.append("")
    if skip == len(results) and len(results) > 0:
        lines.extend(
            [
                "## 如何启用 L3/L4",
                "",
                "**L3 大模型**：界面「大模型设置」→ 启用 + API Key + 可选 enable_plugins（OCR）。",
                "或设置环境变量 `OPENAI_API_KEY`。",
                "",
                "**L4 Azure**：界面「Azure 设置」→ docintel_endpoint / cu_endpoint + Key（或本机 `az login`）。",
                "",
            ]
        )
    REPORT.write_text("\n".join(lines), encoding="utf-8")
    (OUT / "results.json").write_text(
        json.dumps(
            {"l3_probe": l3_probe, "l4_probe": l4_probe, "results": [asdict(r) for r in results]},
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )


def main() -> int:
    OUT.mkdir(parents=True, exist_ok=True)
    l3_probe = _probe_l3()
    l4_probe = _probe_l4()
    print("=== L3/L4 条件测试 ===")
    print(f"L3 ready={l3_probe['llm_ready']} plugins={l3_probe['plugins']}")
    print(f"L4 docintel={l4_probe['docintel_ready']} cu={l4_probe['cu_ready']}")

    runners = [
        _run_l3_llm_image,
        _run_l3_ocr_plugin,
        _run_l4_docintel,
        _run_l4_cu_video_or_audio,
    ]
    results: list[TierResult] = []
    for fn in runners:
        r = fn()
        print(f"{r.case_id}: {r.status}" + (f" ({r.reason})" if r.reason else ""))
        results.append(r)

    _write_report(results, l3_probe, l4_probe)
    print(f"报告：{REPORT}")
    fail = sum(1 for r in results if r.status == "FAIL")
    return 1 if fail else 0


if __name__ == "__main__":
    raise SystemExit(main())
