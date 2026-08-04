"""ZIP 本地增强：魔数校验、进度反馈、跳过二进制、源码按文本输出。"""

from __future__ import annotations

import io
import zipfile
from pathlib import Path
from typing import Callable

# 明确按「源码/文本」输出的扩展名
_TEXT_EXTS = {
    ".txt",
    ".text",
    ".md",
    ".markdown",
    ".java",
    ".py",
    ".pyw",
    ".c",
    ".h",
    ".cpp",
    ".hpp",
    ".cs",
    ".js",
    ".ts",
    ".jsx",
    ".tsx",
    ".json",
    ".jsonl",
    ".xml",
    ".html",
    ".htm",
    ".css",
    ".scss",
    ".less",
    ".yaml",
    ".yml",
    ".toml",
    ".ini",
    ".cfg",
    ".conf",
    ".properties",
    ".gradle",
    ".groovy",
    ".kt",
    ".kts",
    ".go",
    ".rs",
    ".sql",
    ".sh",
    ".bat",
    ".ps1",
    ".cmd",
    ".r",
    ".rb",
    ".php",
    ".swift",
    ".scala",
    ".lua",
    ".pl",
    ".csv",
    ".tsv",
    ".log",
    ".gitignore",
    ".dockerfile",
    ".makefile",
}

# 包内 Office / PDF 等：走标准转换引擎（与顶层文件一致）
_MARKITDOWN_EXTS = {
    ".pdf",
    ".docx",
    ".pptx",
    ".xlsx",
    ".xls",
    ".msg",
    ".epub",
    ".ipynb",
    ".csv",
    ".html",
    ".htm",
    ".xml",
    ".json",
    ".jsonl",
    ".rss",
    ".atom",
}

# 直接跳过的二进制 / 构建产物
_SKIP_EXTS = {
    ".class",
    ".jar",
    ".war",
    ".ear",
    ".exe",
    ".dll",
    ".so",
    ".dylib",
    ".o",
    ".obj",
    ".a",
    ".lib",
    ".pyc",
    ".pyo",
    ".pyd",
    ".wasm",
    ".bin",
    ".dat",
    ".db",
    ".sqlite",
    ".pak",
    ".pdb",
    ".ilk",
    ".exp",
    ".map",
    ".min.js",
    ".map.js",
    ".woff",
    ".woff2",
    ".ttf",
    ".otf",
    ".eot",
    ".ico",
    ".icns",
    ".mp3",
    ".wav",
    ".m4a",
    ".mp4",
    ".avi",
    ".mkv",
    ".mov",
    ".flv",
    ".zip",
    ".rar",
    ".7z",
    ".gz",
    ".tar",
    ".bz2",
    ".xz",
    ".iso",
    ".dmg",
}

_SKIP_DIR_PARTS = {
    "__macosx",
    ".git",
    ".svn",
    ".idea",
    ".vscode",
    "node_modules",
    "target",
    "build",
    "out",
    "dist",
    ".gradle",
    ".mvn",
    "__pycache__",
    ".pytest_cache",
}

_MAX_FILES = 80
_MAX_FILE_BYTES = 1_500_000
_MAX_TOTAL_BYTES = 25_000_000


def sniff_archive_kind(path: str | Path) -> str:
    """根据文件头判断真实压缩格式。返回 zip / rar / 7z / gzip / unknown。"""
    path = Path(path)
    try:
        head = path.read_bytes()[:16]
    except OSError:
        return "unknown"
    if head.startswith(b"PK\x03\x04") or head.startswith(b"PK\x05\x06") or head.startswith(
        b"PK\x07\x08"
    ):
        return "zip"
    if head.startswith(b"Rar!\x1a\x07"):
        return "rar"
    if head.startswith(b"7z\xbc\xaf\x27\x1c"):
        return "7z"
    if head.startswith(b"\x1f\x8b"):
        return "gzip"
    return "unknown"


def _lang_fence(ext: str) -> str:
    return {
        ".java": "java",
        ".py": "python",
        ".pyw": "python",
        ".js": "javascript",
        ".ts": "typescript",
        ".jsx": "jsx",
        ".tsx": "tsx",
        ".c": "c",
        ".h": "c",
        ".cpp": "cpp",
        ".hpp": "cpp",
        ".cs": "csharp",
        ".json": "json",
        ".xml": "xml",
        ".html": "html",
        ".htm": "html",
        ".css": "css",
        ".sql": "sql",
        ".sh": "bash",
        ".bat": "bat",
        ".ps1": "powershell",
        ".yaml": "yaml",
        ".yml": "yaml",
        ".md": "markdown",
        ".markdown": "markdown",
        ".gradle": "groovy",
        ".kt": "kotlin",
        ".go": "go",
        ".rs": "rust",
        ".php": "php",
        ".rb": "ruby",
    }.get(ext.lower(), "")


def _should_skip_name(name: str) -> bool:
    parts = name.replace("\\", "/").split("/")
    if any(p.lower() in _SKIP_DIR_PARTS for p in parts if p):
        return True
    base = parts[-1] if parts else name
    if not base or base.endswith("/"):
        return True
    if base.startswith("."):
        # 保留常见配置点文件
        if base.lower() not in {".gitignore", ".dockerignore", ".editorconfig"}:
            # .classpath / .project 等 Eclipse 配置仍可读
            if base.lower() in {".classpath", ".project", ".factorypath"}:
                return False
            if not any(base.lower().endswith(ext) for ext in _TEXT_EXTS):
                return True
    return False


def _decode_bytes(data: bytes) -> str:
    for enc in ("utf-8", "utf-8-sig", "gbk", "gb18030", "latin-1"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="replace")


def _looks_binary(data: bytes) -> bool:
    if not data:
        return False
    sample = data[:4096]
    if b"\x00" in sample:
        return True
    # 高比例不可打印控制字符
    ctrl = sum(1 for b in sample if b < 9 or (13 < b < 32))
    return ctrl / max(len(sample), 1) > 0.30


def _convert_member_bytes(
    name: str,
    data: bytes,
    *,
    progress: Callable[[str], None] | None = None,
) -> str | None:
    ext = Path(name).suffix.lower()
    if ext in _SKIP_EXTS:
        return None
    if len(data) > _MAX_FILE_BYTES:
        return f"_（跳过：单文件超过 {_MAX_FILE_BYTES // 1_000_000}MB）_"
    if _looks_binary(data) and ext not in _MARKITDOWN_EXTS:
        return None

    # 文档 / 结构化：标准引擎（Excel / Notebook 仍用本地增强）
    if ext in _MARKITDOWN_EXTS:
        import tempfile

        suffix = ext or ".bin"
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            tmp.write(data)
            tmp_path = Path(tmp.name)
        try:
            if ext in {".xlsx", ".xls"}:
                from excel_convert import convert_excel_to_markdown
                from cleanup import clean_markdown

                return clean_markdown(convert_excel_to_markdown(tmp_path).strip())
            if ext == ".ipynb":
                from ipynb_convert import convert_ipynb_to_markdown
                from cleanup import clean_markdown

                return clean_markdown(convert_ipynb_to_markdown(tmp_path).strip())
            from converter import _convert_options, _finalize_markdown, _get_markitdown

            if progress:
                progress(f"ZIP 内文档：{Path(name).name}")
            md = _get_markitdown()
            result = md.convert(str(tmp_path), **_convert_options())
            text = getattr(result, "markdown", None) or getattr(result, "text_content", "") or ""
            return _finalize_markdown(tmp_path, text.strip(), light=True)
        except Exception as exc:  # noqa: BLE001
            return f"_（转换失败：{exc}）_"
        finally:
            try:
                tmp_path.unlink(missing_ok=True)
            except OSError:
                pass

    # 文本 / 源码
    if ext in _TEXT_EXTS or ext == "":
        text = _decode_bytes(data).replace("\r\n", "\n").replace("\r", "\n")
        if not text.strip():
            return "_（空文件）_"
        # Markdown / 纯说明直接贴；其余包代码围栏
        if ext in {".md", ".markdown", ".txt", ".text", ".log"}:
            return text.strip()
        fence = _lang_fence(ext)
        return f"```{fence}\n{text.rstrip()}\n```"

    # 图片：仅占位，避免拖慢
    if ext in {".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp"}:
        return f"![ZIP内图片 · {Path(name).name}](embedded-image)"

    return None


def convert_zip_to_markdown(
    path: str | Path,
    *,
    progress: Callable[[str], None] | None = None,
) -> str:
    path = Path(path)
    kind = sniff_archive_kind(path)
    if kind == "rar":
        raise ValueError(
            "该文件扩展名是 .zip，但实际是 RAR 压缩包（文件头为 Rar!）。\n"
            "本工具只支持真正的 ZIP。\n"
            "请用 7-Zip / WinRAR 打开后「重新压缩为 ZIP」，或直接把包内文件拖出来转换。"
        )
    if kind == "7z":
        raise ValueError(
            "该文件实际是 7z 压缩包，不是 ZIP。\n"
            "请改用 ZIP 格式，或解压后选择其中的文档再转换。"
        )
    if kind == "gzip":
        raise ValueError("检测到 gzip（.gz）包，请先解压后再转换内容。")
    if kind != "zip":
        # 再让 zipfile 试一次，给出统一错误
        try:
            with zipfile.ZipFile(path, "r"):
                pass
        except zipfile.BadZipFile as exc:
            raise ValueError(
                "不是有效的 ZIP 文件（可能已损坏，或只是改了扩展名）。\n"
                "请确认压缩格式为 ZIP 后重试。"
            ) from exc

    if progress:
        progress("正在读取 ZIP 目录…")

    parts: list[str] = [f"# ZIP：{path.name}", ""]
    skipped: list[str] = []
    converted = 0
    total_bytes = 0

    try:
        zf = zipfile.ZipFile(path, "r")
    except zipfile.BadZipFile as exc:
        raise ValueError("ZIP 已损坏或不是有效压缩包。") from exc

    with zf:
        names = [n for n in zf.namelist() if not n.endswith("/")]
        candidates: list[str] = []
        for name in names:
            if _should_skip_name(name):
                skipped.append(name)
                continue
            ext = Path(name).suffix.lower()
            if ext in _SKIP_EXTS:
                skipped.append(name)
                continue
            candidates.append(name)

        if progress:
            progress(f"ZIP 内可处理条目约 {len(candidates)} 个…")

        for idx, name in enumerate(candidates, 1):
            if converted >= _MAX_FILES:
                skipped.extend(candidates[idx - 1 :])
                parts.append(
                    f"\n> 已达单次转换上限（{_MAX_FILES} 个文件），其余已跳过。"
                )
                break
            try:
                info = zf.getinfo(name)
            except KeyError:
                continue
            if info.file_size > _MAX_FILE_BYTES:
                skipped.append(name)
                continue
            if total_bytes + info.file_size > _MAX_TOTAL_BYTES:
                parts.append(
                    f"\n> 已达解压体积上限（{_MAX_TOTAL_BYTES // 1_000_000}MB），停止继续解析。"
                )
                skipped.extend(candidates[idx - 1 :])
                break

            if progress:
                progress(f"ZIP（{idx}/{len(candidates)}）：{name}")

            try:
                data = zf.read(name)
            except Exception as exc:  # noqa: BLE001
                parts.append(f"## 文件：`{name}`\n\n_（读取失败：{exc}）_\n")
                continue

            total_bytes += len(data)
            body = _convert_member_bytes(name, data, progress=progress)
            if body is None:
                skipped.append(name)
                continue

            parts.append(f"## 文件：`{name}`\n\n{body}\n")
            converted += 1

    if converted == 0:
        raise ValueError(
            "ZIP 内没有可转换的文本/文档内容"
            + ("（多为 class/jar 等二进制作业产物）。" if skipped else "。")
            + "\n请解压后选择 .java / .docx / .pdf 等源文件再转换。"
        )

    if skipped:
        show = skipped[:30]
        more = f"\n- …另有 {len(skipped) - 30} 个" if len(skipped) > 30 else ""
        parts.append(
            "## 已跳过\n\n"
            + "\n".join(f"- `{s}`" for s in show)
            + more
            + "\n"
        )

    parts.append(f"\n---\n转换了 **{converted}** 个条目，跳过 **{len(skipped)}** 个。\n")
    return "\n".join(parts).strip() + "\n"
