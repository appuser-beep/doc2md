"""测试公共辅助：图片痕迹、关键词判定。"""

from __future__ import annotations

import re

# 在 must 列表中替代硬编码 embedded-image（PPT 常保留 Picture*.jpg 等路径）
IMAGE_TRACE = "__IMAGE_TRACE__"


def has_image_trace(md: str) -> bool:
    """文档含图片痕迹：Markdown 图链 / embedded-image / Excel 嵌入图标注 / PPT Picture 路径。"""
    if re.search(r"!\[[^\]]*\]\([^)]+\)", md):
        return True
    if "embedded-image" in md:
        return True
    if "Excel图片" in md:
        return True
    if "PDF图片" in md:
        return True
    if re.search(r"(?i)picture\d*\.(jpg|jpeg|png)", md):
        return True
    return False


def keyword_in_md(md: str, keyword: str) -> bool:
    """关键词匹配；允许多词被 PDF 换行拆开（如 North\\nChina）。"""
    if keyword in md:
        return True
    parts = keyword.split()
    if len(parts) > 1 and all(p in md for p in parts):
        return True
    return False


def check_must(md: str, must: list[str]) -> list[str]:
    """检查 must 关键词；支持 IMAGE_TRACE 特殊项。"""
    issues: list[str] = []
    for token in must:
        if token == IMAGE_TRACE:
            if not has_image_trace(md):
                issues.append("缺图片痕迹（需 markdown 图链、Picture 路径或 embedded-image）")
        elif not keyword_in_md(md, token):
            issues.append(f"缺 must: {token}")
    return issues


def has_markdown_pipe_table(md: str) -> bool:
    """标准 Markdown 管道表。"""
    return bool(re.search(r"(?m)^\|.+\|", md))


def has_pdf_text_table(md: str) -> bool:
    """PDF 纯文本表格：无 | 管道，但有多行且含数字/表头词。"""
    lines = [ln.strip() for ln in md.splitlines() if ln.strip()]
    if len(lines) < 4:
        return False
    digit_lines = sum(1 for ln in lines if re.search(r"\d", ln))
    return digit_lines >= 2
