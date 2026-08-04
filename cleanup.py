"""转换结果通用清理：去 NaN / Unnamed、修表头、还原转义下划线。"""

from __future__ import annotations

import re
from pathlib import Path


_RE_NAN = re.compile(r"\bNaN\b", re.IGNORECASE)
_RE_UNNAMED = re.compile(r"Unnamed:\s*\d+", re.IGNORECASE)
_RE_MULTI_NL = re.compile(r"\n{3,}")
# markdownify 常把 WORD_FOO 转成 WORD\_FOO
_RE_ESC_UNDERSCORE = re.compile(r"\\_")


def _split_table_row(line: str) -> list[str]:
    s = line.strip()
    if s.startswith("|"):
        s = s[1:]
    if s.endswith("|"):
        s = s[:-1]
    return [c.strip() for c in s.split("|")]


def _is_sep_row(cells: list[str]) -> bool:
    if not cells:
        return False
    return all(re.fullmatch(r":?-{3,}:?", c or "") for c in cells)


def _join_row(cells: list[str]) -> str:
    return "| " + " | ".join(cells) + " |"


def _clean_table_block(lines: list[str]) -> list[str]:
    if len(lines) < 2:
        return lines

    rows = [_split_table_row(ln) for ln in lines]
    width = max(len(r) for r in rows)
    rows = [r + [""] * (width - len(r)) for r in rows]

    for r in rows:
        for i, c in enumerate(r):
            c = _RE_NAN.sub("", c)
            c = _RE_UNNAMED.sub("", c)
            c = _RE_ESC_UNDERSCORE.sub("_", c)
            r[i] = c.strip()

    # Word/mammoth 常见：空表头行 + 分隔行 + 真实表头
    if (
        len(rows) >= 3
        and _is_sep_row(rows[1])
        and not any(c.strip() for c in rows[0])
        and any(c.strip() for c in rows[2])
    ):
        rows = rows[2:]  # 丢掉空表头与旧分隔，下面会重建分隔

    keep_cols = []
    for ci in range(len(rows[0]) if rows else 0):
        meaningful = []
        for ri, r in enumerate(rows):
            if _is_sep_row(r):
                continue
            if ci < len(r):
                meaningful.append(r[ci])
        if any(meaningful):
            keep_cols.append(ci)
    if not keep_cols:
        return []
    rows = [[r[i] for i in keep_cols] for r in rows]

    out_rows = []
    for ri, r in enumerate(rows):
        if ri == 0 or _is_sep_row(r):
            out_rows.append(r)
            continue
        if any(c.strip() for c in r):
            out_rows.append(r)

    if not out_rows:
        return []

    # 表头若仍全空，用列名占位
    if not any(c.strip() for c in out_rows[0]):
        out_rows[0] = [f"列{i+1}" for i in range(len(out_rows[0]))]

    if len(out_rows) >= 2 and _is_sep_row(out_rows[1]):
        out_rows[1] = ["---"] * len(out_rows[0])
    else:
        out_rows.insert(1, ["---"] * len(out_rows[0]))

    return [_join_row(r) for r in out_rows]


def clean_markdown_light(text: str) -> str:
    """轻量清理：尽量保持原样，仅做安全、可读的修正。"""
    if not text:
        return text

    text = text.replace("\r\n", "\n").replace("\r", "\n")
    # mammoth/markdownify 常把 WORD_FOO 转成 WORD\_FOO
    text = _RE_ESC_UNDERSCORE.sub("_", text)
    text = _RE_MULTI_NL.sub("\n\n", text)
    return text.strip() + "\n"


def clean_markdown(text: str) -> str:
    if not text:
        return text

    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = _RE_NAN.sub("", text)
    text = _RE_UNNAMED.sub("", text)
    text = _RE_ESC_UNDERSCORE.sub("_", text)

    # 缩短超长 data-uri 图片，保留可识别痕迹（避免预览爆炸）
    text = re.sub(
        r"!\[([^\]]*)\]\(data:image\/[a-zA-Z0-9+.-]+;base64,[A-Za-z0-9+/=\s]{200,}\)",
        r"![\1](data:image/...base64...)",
        text,
    )

    # PPT/Excel 等常留下不可解析的相对路径图链，统一改为占位（保留 alt）
    def _fix_dangling_img(m: re.Match[str]) -> str:
        alt, url = m.group(1), m.group(2).strip()
        low = url.lower()
        if low.startswith(("http://", "https://", "data:", "embedded-image")):
            return m.group(0)
        label = alt.strip() or Path(url).name or "embedded-image"
        return f"![{label}](embedded-image)"

    text = re.sub(r"!\[([^\]]*)\]\(([^)]+)\)", _fix_dangling_img, text)

    lines = text.split("\n")
    out: list[str] = []
    i = 0
    while i < len(lines):
        line = lines[i]
        if line.strip().startswith("|"):
            block = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                block.append(lines[i])
                i += 1
            cleaned = _clean_table_block(block)
            if cleaned:
                out.extend(cleaned)
                out.append("")
            continue
        out.append(line)
        i += 1

    text = "\n".join(out)
    text = _RE_MULTI_NL.sub("\n\n", text)
    return text.strip() + "\n"
