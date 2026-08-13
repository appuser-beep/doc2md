"""转换结果通用清理：去 NaN / Unnamed、修表头、还原转义下划线、修复表格。"""

from __future__ import annotations

import re
from pathlib import Path


_RE_NAN = re.compile(r"\bNaN\b", re.IGNORECASE)
_RE_UNNAMED = re.compile(r"Unnamed:\s*\d+", re.IGNORECASE)
_RE_MULTI_NL = re.compile(r"\n{3,}")
# markdownify 常把 WORD_FOO 转成 WORD\_FOO
_RE_ESC_UNDERSCORE = re.compile(r"\\_")
_RE_FENCE = re.compile(r"^(`{3,}|~{3,})")


def _split_table_row(line: str) -> list[str]:
    """按 | 拆分，支持已转义的 \\|。"""
    s = line.strip()
    if s.startswith("|"):
        s = s[1:]
    if s.endswith("|"):
        s = s[:-1]

    cells: list[str] = []
    cur: list[str] = []
    i = 0
    while i < len(s):
        ch = s[i]
        if ch == "\\" and i + 1 < len(s) and s[i + 1] == "|":
            cur.append("|")
            i += 2
            continue
        if ch == "|":
            cells.append("".join(cur).strip())
            cur = []
            i += 1
            continue
        cur.append(ch)
        i += 1
    cells.append("".join(cur).strip())
    return cells


def _fit_row_width(cells: list[str], width: int, *, pad: str = "") -> list[str]:
    """列数多于表宽时合并片段。优先合并更短的碎片（如「张|三」）；同分时靠后合并。

    pad：补齐短行时的填充；分隔行应传 ``---``，避免空串破坏分隔行判定。
    """
    if width <= 0:
        return cells
    if len(cells) < width:
        return cells + [pad] * (width - len(cells))
    if len(cells) == width:
        return cells

    cells = list(cells)
    while len(cells) > width:
        best_i = len(cells) - 2
        best_score = -10_000.0
        for i in range(len(cells) - 1):
            a = cells[i] or ""
            b = cells[i + 1] or ""
            score = 0.0
            la, lb = len(a), len(b)
            # 空片段优先与邻居合并
            if not a or not b:
                score += 40.0
            # 两段都短：更像单元格内未转义 |；总长越短越优先（张|三 > 三|研发）
            elif la <= 2 and lb <= 2:
                score += 30.0 - (la + lb)
                if re.search(r"[\u4e00-\u9fff]$", a) and re.search(
                    r"^[\u4e00-\u9fff]", b
                ):
                    score += 2.0
            elif la <= 3 and lb <= 3:
                score += 12.0 - 0.5 * (la + lb)
            # 末对若一侧像独立数字编码，略降分
            if i == len(cells) - 2 and (
                re.fullmatch(r"[\d.\-/%]+", a) or re.fullmatch(r"[\d.\-/%]+", b)
            ):
                score -= 3.0
            # 同分偏好靠后合并（多出的片段并入右侧）
            score += i * 0.01
            if score > best_score:
                best_score = score
                best_i = i
        cells = (
            cells[:best_i]
            + [f"{cells[best_i]}|{cells[best_i + 1]}"]
            + cells[best_i + 2 :]
        )
    return cells


def _is_sep_row(cells: list[str]) -> bool:
    if not cells:
        return False
    return all(re.fullmatch(r":?-{3,}:?", c or "") for c in cells)


def _escape_cell(text: str) -> str:
    """避免单元格内 | 破坏管道表；对已转义 \\| 保持幂等。"""
    s = text or ""
    out: list[str] = []
    i = 0
    while i < len(s):
        if s[i] == "\\" and i + 1 < len(s) and s[i + 1] == "|":
            out.append("\\|")
            i += 2
            continue
        if s[i] == "|":
            out.append("\\|")
            i += 1
            continue
        out.append(s[i])
        i += 1
    return "".join(out)


def _join_row(cells: list[str]) -> str:
    return "| " + " | ".join(_escape_cell(c) for c in cells) + " |"


def _sep_alignment_token(cell: str) -> str:
    """保留 GFM 对齐标记（:--- / ---: / :---:），否则用 ---。"""
    s = (cell or "").strip().replace(" ", "")
    if not s or not re.fullmatch(r":?-{1,}:?", s):
        return "---"
    left = s.startswith(":")
    right = s.endswith(":")
    if left and right:
        return ":---:"
    if left:
        return ":---"
    if right:
        return "---:"
    return "---"


def _infer_table_width(rows: list[list[str]]) -> int:
    """列宽以表头与分隔行为准，不用过宽数据行抬高，否则无法合并未转义 |。"""
    if not rows:
        return 0
    sep_w = 0
    for r in rows:
        if _is_sep_row(r):
            sep_w = len(r)
            break
    header_w = 0
    for r in rows:
        if not _is_sep_row(r):
            header_w = len(r)
            break
    return max(header_w, sep_w, 1)


def _table_needs_rewrite(rows: list[list[str]]) -> bool:
    """空表头、过宽行、或分隔行列数与表头不一致时需要重写。"""
    if len(rows) < 2:
        return False
    sep_w = 0
    header_w = 0
    for r in rows:
        if _is_sep_row(r):
            if sep_w <= 0:
                sep_w = len(r)
        elif header_w <= 0:
            header_w = len(r)
    width = _infer_table_width(rows)
    if any(len(r) > width for r in rows if not _is_sep_row(r)):
        return True
    if header_w and sep_w and header_w != sep_w:
        return True
    if (
        len(rows) >= 3
        and _is_sep_row(rows[1])
        and not any(c.strip() for c in rows[0])
        and any(c.strip() for c in rows[2])
    ):
        return True
    return False


def _clean_table_block(
    lines: list[str],
    *,
    drop_empty_cols: bool = True,
    strip_pandas_artifacts: bool = True,
    force: bool = False,
) -> list[str] | None:
    """返回清理后的表行；若 light 模式判定无需改写则返回 None（保留原文）。"""
    if len(lines) < 2:
        return lines

    rows = [_split_table_row(ln) for ln in lines]
    if not force and not drop_empty_cols and not strip_pandas_artifacts:
        if not _table_needs_rewrite(rows):
            return None

    # 保存原分隔行对齐（若存在）
    orig_sep: list[str] | None = None
    for r in rows:
        if _is_sep_row(r):
            orig_sep = r
            break

    width = _infer_table_width(rows)
    if width <= 0:
        width = max(len(r) for r in rows)
    fitted: list[list[str]] = []
    for r in rows:
        # 分隔行用 --- 补齐，避免短分隔被 pad 空串后变成「假数据行」
        if _is_sep_row(r):
            fitted.append(_fit_row_width(r, width, pad="---"))
        else:
            fitted.append(_fit_row_width(r, width, pad=""))
    rows = fitted

    for r in rows:
        for i, c in enumerate(r):
            if strip_pandas_artifacts:
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
        orig_sep = None

    if drop_empty_cols:
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
        if orig_sep is not None:
            orig_sep = [orig_sep[i] if i < len(orig_sep) else "---" for i in keep_cols]

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
        out_rows[0] = [f"列{i + 1}" for i in range(len(out_rows[0]))]

    ncol = len(out_rows[0])
    if len(out_rows) >= 2 and _is_sep_row(out_rows[1]):
        if orig_sep is not None and len(orig_sep) == ncol:
            out_rows[1] = [_sep_alignment_token(c) for c in orig_sep]
        else:
            out_rows[1] = ["---"] * ncol
    else:
        if orig_sep is not None and len(orig_sep) == ncol:
            out_rows.insert(1, [_sep_alignment_token(c) for c in orig_sep])
        else:
            out_rows.insert(1, ["---"] * ncol)

    return [_join_row(r) for r in out_rows]


def _rewrite_tables(
    text: str,
    *,
    drop_empty_cols: bool,
    strip_pandas_artifacts: bool,
    force: bool = False,
) -> str:
    """重写 Markdown 管道表；跳过 fenced code 内的伪表格。"""
    lines = text.split("\n")
    out: list[str] = []
    i = 0
    in_fence = False
    fence_char = ""
    fence_len = 0

    while i < len(lines):
        line = lines[i]
        fence_match = _RE_FENCE.match(line.strip())
        if fence_match:
            raw = fence_match.group(1)
            ch = raw[0]
            if not in_fence:
                in_fence = True
                fence_char = ch
                fence_len = len(raw)
                out.append(line)
                i += 1
                continue
            if ch == fence_char and len(raw) >= fence_len:
                in_fence = False
                fence_char = ""
                fence_len = 0
                out.append(line)
                i += 1
                continue

        if in_fence:
            out.append(line)
            i += 1
            continue

        if line.strip().startswith("|"):
            block = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                block.append(lines[i])
                i += 1
            cleaned = _clean_table_block(
                block,
                drop_empty_cols=drop_empty_cols,
                strip_pandas_artifacts=strip_pandas_artifacts,
                force=force,
            )
            if cleaned is None:
                out.extend(block)
            elif cleaned:
                out.extend(cleaned)
                out.append("")
            continue

        out.append(line)
        i += 1
    return "\n".join(out)


def _repair_md_images(text: str) -> str:
    """修复 alt 文本含 ] 时把 ![...]...](url) 截断的问题。"""
    if "![" not in text:
        return text
    markers = (
        "](data:",
        "](http://",
        "](https://",
        "](embedded-image",
        "](Picture",
    )
    out: list[str] = []
    i = 0
    n = len(text)
    while i < n:
        start = text.find("![", i)
        if start < 0:
            out.append(text[i:])
            break
        out.append(text[i:start])
        close = -1
        for marker in markers:
            pos = text.find(marker, start + 2)
            if pos >= 0 and (close < 0 or pos < close):
                close = pos
        if close < 0:
            out.append(text[start : start + 2])
            i = start + 2
            continue
        alt = text[start + 2 : close]
        paren = text.find(")", close + 2)
        if paren < 0:
            out.append(text[start:])
            break
        url = text[close + 2 : paren]
        alt_safe = (
            alt.replace("]", "_")
            .replace("[", "_")
            .replace("\n", " ")
            .strip()
        )
        # 过长路径 alt：保留文件名
        if len(alt_safe) > 120 and ("\\" in alt_safe or "/" in alt_safe):
            alt_safe = Path(alt_safe.replace("\\", "/")).name or "image"
        out.append(f"![{alt_safe}]({url})")
        i = paren + 1
    return "".join(out)


def clean_markdown_light(text: str) -> str:
    """轻量清理：保持正文，修复 Word/PPT 空表头与单元格内 |。"""
    if not text:
        return text

    text = text.replace("\r\n", "\n").replace("\r", "\n")
    # mammoth/markdownify 常把 WORD_FOO 转成 WORD\_FOO
    text = _RE_ESC_UNDERSCORE.sub("_", text)
    text = _repair_md_images(text)
    text = _rewrite_tables(
        text,
        drop_empty_cols=False,
        strip_pandas_artifacts=False,
    )
    text = _RE_MULTI_NL.sub("\n\n", text)
    return text.strip() + "\n"


def clean_markdown(text: str, *, keep_data_uris: bool = False) -> str:
    if not text:
        return text

    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = _RE_NAN.sub("", text)
    text = _RE_UNNAMED.sub("", text)
    text = _RE_ESC_UNDERSCORE.sub("_", text)
    text = _repair_md_images(text)

    # 默认缩短超长 data-uri；keep_data_uris=True 时保留完整内嵌图
    if not keep_data_uris:
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
        # 官方 PPT 输出的 Picture*.jpg freestanding 路径：保留文件名便于识别
        if re.search(r"picture\d+\.(jpg|jpeg|png|gif|webp)$", low):
            label = alt.strip() or Path(url).name
            return f"![{label}]({Path(url).name})"
        label = alt.strip() or Path(url).name or "embedded-image"
        return f"![{label}](embedded-image)"

    text = re.sub(r"!\[([^\]]*)\]\(([^)]+)\)", _fix_dangling_img, text)
    text = _rewrite_tables(
        text,
        drop_empty_cols=True,
        strip_pandas_artifacts=True,
        force=True,
    )
    text = _RE_MULTI_NL.sub("\n\n", text)
    return text.strip() + "\n"
