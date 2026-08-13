"""Jupyter .ipynb 增强转换：保留 Markdown/代码，并抽取 stdout、HTML 表、图片占位。"""

from __future__ import annotations

import json
from pathlib import Path


def _join_source(source) -> str:
    if source is None:
        return ""
    if isinstance(source, list):
        return "".join(source)
    return str(source)


def _html_to_md(html: str) -> str:
    html = (html or "").strip()
    if not html:
        return ""
    try:
        from markitdown.converters._html_converter import HtmlConverter

        result = HtmlConverter().convert_string(html)
        return (getattr(result, "markdown", None) or "").strip()
    except Exception:
        import re

        text = re.sub(r"<br\s*/?>", "\n", html, flags=re.I)
        text = re.sub(r"</p>", "\n", text, flags=re.I)
        text = re.sub(r"<[^>]+>", "", text)
        return text.strip()


def _convert_output(output: dict) -> list[str]:
    bits: list[str] = []
    otype = output.get("output_type")
    if otype == "stream":
        text = _join_source(output.get("text"))
        if text.strip():
            bits.append(text.rstrip())
        return bits

    if otype in {"execute_result", "display_data"}:
        data = output.get("data") or {}
        html_md = ""
        if "text/html" in data:
            html_md = _html_to_md(_join_source(data.get("text/html")))
            if html_md:
                bits.append(html_md)
        # 无可用 HTML 时用 markdown；与图片可并存
        if not html_md and "text/markdown" in data:
            md = _join_source(data.get("text/markdown")).strip()
            if md:
                bits.append(md)
        if "image/png" in data or "image/jpeg" in data:
            bits.append("![Notebook图片](embedded-image)")
        if not bits and "text/plain" in data:
            plain = _join_source(data.get("text/plain")).strip()
            if plain:
                bits.append(plain)
        return bits

    if otype == "error":
        ename = output.get("ename", "Error")
        evalue = output.get("evalue", "")
        bits.append(f"**{ename}**: {evalue}")
    return bits


def convert_ipynb_to_markdown(path: str | Path) -> str:
    path = Path(path)
    nb = json.loads(path.read_text(encoding="utf-8"))
    parts: list[str] = []

    for cell in nb.get("cells", []):
        ctype = cell.get("cell_type", "")
        source = _join_source(cell.get("source")).rstrip()

        if ctype == "markdown":
            if source:
                parts.append(source)
            continue

        if ctype == "raw":
            if source:
                parts.append(f"```\n{source}\n```")
            continue

        if ctype == "code":
            if source:
                parts.append(f"```python\n{source}\n```")
            for output in cell.get("outputs") or []:
                for bit in _convert_output(output):
                    if bit.strip():
                        parts.append(bit)
            continue

    text = "\n\n".join(parts).strip()
    return text + ("\n" if text else "")
