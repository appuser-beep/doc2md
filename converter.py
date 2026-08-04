"""文档转 Markdown 转换封装。"""

from __future__ import annotations

import io
from pathlib import Path
from typing import BinaryIO, Callable

from cleanup import clean_markdown, clean_markdown_light

APP_VERSION = "1.7.2"

_BUILTIN_EXTENSIONS = {
    ".pdf",
    ".docx",
    ".pptx",
    ".xlsx",
    ".xls",
    ".msg",
    ".epub",
    ".ipynb",
    ".jpg",
    ".jpeg",
    ".png",
    ".wav",
    ".mp3",
    ".m4a",
    ".mp4",
    ".html",
    ".htm",
    ".csv",
    ".json",
    ".jsonl",
    ".xml",
    ".rss",
    ".atom",
    ".txt",
    ".text",
    ".md",
    ".markdown",
    ".zip",
}

_CU_EXTENSIONS = {
    ".eml",
    ".rtf",
    ".heif",
    ".heic",
    ".bmp",
    ".tif",
    ".tiff",
    ".flac",
    ".ogg",
    ".aac",
    ".wma",
    ".m4v",
    ".mov",
    ".avi",
    ".mkv",
    ".webm",
    ".flv",
    ".wmv",
}

SUPPORTED_EXTENSIONS = _BUILTIN_EXTENSIONS | _CU_EXTENSIONS

UNSUPPORTED_BUT_RISKY = {
    ".doc": "老版 Word（.doc）不受支持，请先另存为 .docx。若强行转换，可能只得到乱码片段。",
    ".ppt": "老版 PowerPoint（.ppt）不受支持，请先另存为 .pptx。",
    ".xlsb": "不支持 .xlsb，请另存为 .xlsx。",
}

FILE_DIALOG_TYPES = [
    (
        "支持的文档",
        "*.pdf *.docx *.pptx *.xlsx *.xls *.msg *.eml *.epub *.ipynb "
        "*.jpg *.jpeg *.png *.bmp *.tif *.tiff *.heif *.heic "
        "*.wav *.mp3 *.m4a *.mp4 *.flac *.ogg *.aac *.wma "
        "*.mov *.avi *.mkv *.webm *.m4v *.flv *.wmv "
        "*.html *.htm *.rtf *.csv *.json *.jsonl *.xml *.rss *.atom *.txt *.md *.zip",
    ),
    ("PDF", "*.pdf"),
    ("Word", "*.docx"),
    ("PowerPoint", "*.pptx"),
    ("Excel", "*.xlsx *.xls"),
    ("Outlook / 邮件", "*.msg *.eml"),
    ("图片", "*.jpg *.jpeg *.png *.bmp *.tif *.tiff *.heif *.heic"),
    ("音视频", "*.wav *.mp3 *.m4a *.mp4 *.flac *.ogg *.aac *.wma *.mov *.avi *.mkv *.webm"),
    ("网页/文本", "*.html *.htm *.txt *.md *.rtf *.csv *.json *.xml"),
    ("压缩包", "*.zip"),
    ("所有文件", "*.*"),
]

SUPPORT_HELP = """格式与能力说明
版本 {version}

────────────────────────────────────────
1. 本地内置（安装依赖后可用）
────────────────────────────────────────

Office：Word (.docx)、Excel (.xlsx / .xls)、PowerPoint (.pptx)、Outlook (.msg)
文档：PDF、EPUB、Jupyter Notebook (.ipynb)
数据：HTML、RSS、Atom、CSV、JSON、JSONL、XML、TXT、Markdown
图片：JPG / PNG（含 EXIF）；音频 WAV / MP3 / M4A（语音转写，需联网）
视频：MP4 音轨转写（建议本机安装 ffmpeg）
压缩包：ZIP（遍历包内 Office、PDF、文本与源码）
网络：通用网页、Wikipedia、Bing 搜索、YouTube 字幕

────────────────────────────────────────
2. Azure 云端（「Azure 设置」中配置 Endpoint）
────────────────────────────────────────

Document Intelligence：扫描 PDF、复杂版面、BMP / TIFF 等图片
Content Understanding：视频 (MP4 / MOV / AVI / MKV / WebM 等)、EML、RTF、
  FLAC / OGG / AAC / WMA 等扩展音频

────────────────────────────────────────
3. 大模型（「大模型设置」中配置 API）
────────────────────────────────────────

图片描述、PPT 内嵌图说明
OCR 插件：PDF / DOCX / PPTX / XLSX 内嵌图文字识别（需在依赖中启用 OCR 插件）

────────────────────────────────────────
4. 高级选项（「高级设置」）
────────────────────────────────────────

Word 样式映射：自定义标题与段落转 Markdown 的规则
ExifTool 路径：读取图片 / 音频元数据（留空则自动查找）
窄接口模式：不启用 Excel / Notebook / ZIP 本地增强
保留内嵌图：Markdown 中保留 base64 图片（默认截断以减小体积）
自定义插件：加载含 register_converters 函数的 .py 脚本

────────────────────────────────────────
5. 命令行与 Docker
────────────────────────────────────────

doc2md-cli 文件.pdf -o 输出.md
type 文件.pdf | doc2md-cli -x pdf
doc2md-cli --local-only 文件.docx
doc2md-cli --keep-data-uris 文件.pptx
doc2md-cli --list-plugins

Docker：docker build -t doc2md .
  docker run --rm -v %cd%:/data doc2md 文件.pdf -o 文件.md

────────────────────────────────────────
6. 默认增强（未开启窄接口时）
────────────────────────────────────────

Excel：合并单元格、宽表折叠、去除无效 NaN
Notebook：保留代码单元输出
ZIP：魔数校验、跳过 .class 等二进制、优先源码与文档
Word：补充页眉、页脚文字

────────────────────────────────────────
7. 说明与限制
────────────────────────────────────────

视觉样式（颜色、字号）通常不保留；标题、列表、表格、链接结构会保留。
扫描件 PDF 需 Azure 或大模型 OCR 方可得到正文。
ZIP 仅支持标准 ZIP 格式；扩展名为 .zip 但实际为 RAR / 7z 时将提示错误。

不支持：老格式 .doc / .ppt / .xlsb；RAR / 7z 直接转换；Markdown 反向导出。
""".format(version=APP_VERSION)


class ConversionError(Exception):
    """转换失败。"""


def _result_text(result) -> str:
    text = getattr(result, "markdown", None) or getattr(result, "text_content", None)
    if text is None:
        raise ConversionError("转换完成，但未得到 Markdown 内容。")
    return text


def _get_markitdown():
    try:
        from markitdown import MarkItDown
    except ImportError as exc:
        raise ConversionError(
            "未安装转换引擎。请执行：pip install -r requirements.txt"
        ) from exc

    kwargs: dict = {"enable_plugins": False}
    try:
        from llm_settings import markitdown_llm_kwargs

        kwargs.update(markitdown_llm_kwargs())
    except Exception:
        pass
    try:
        from azure_settings import markitdown_azure_kwargs

        kwargs.update(markitdown_azure_kwargs())
    except Exception:
        pass
    try:
        from advanced_settings import markitdown_advanced_kwargs

        kwargs.update(markitdown_advanced_kwargs())
    except Exception:
        pass

    md = MarkItDown(**kwargs)
    try:
        from plugin_loader import apply_custom_converter_plugins

        apply_custom_converter_plugins(md, kwargs)
    except RuntimeError:
        raise
    except Exception:
        pass
    return md


def _convert_options() -> dict:
    try:
        from advanced_settings import convert_options

        return convert_options()
    except Exception:
        return {}


def precheck_source(source: str) -> str | None:
    source = (source or "").strip()
    if not source or source.startswith(("http://", "https://")):
        return None
    path = Path(source)
    if path.is_file() and path.stat().st_size == 0:
        return "文件为空（0 字节），无法转换。"
    ext = path.suffix.lower()
    if ext in UNSUPPORTED_BUT_RISKY:
        return UNSUPPORTED_BUT_RISKY[ext]
    if ext == ".rar":
        return "不支持 RAR。请用 7-Zip/WinRAR 重新压缩为 ZIP，或解压后选文件转换。"
    if ext == ".7z":
        return "不支持 7z。请改为 ZIP，或解压后选文件转换。"
    if ext == ".zip" and path.is_file():
        try:
            from zip_convert import sniff_archive_kind

            kind = sniff_archive_kind(path)
        except Exception:
            kind = "unknown"
        if kind == "rar":
            return (
                "扩展名是 .zip，但文件实际是 RAR（文件头 Rar!）。"
                "请重新压缩为真正的 ZIP 后再转。"
            )
        if kind == "7z":
            return "扩展名是 .zip，但文件实际是 7z。请改为真正的 ZIP。"
        if kind not in {"zip", "unknown"}:
            return f"扩展名是 .zip，但检测到格式为 {kind}，可能无法转换。"
    if ext and ext not in SUPPORTED_EXTENSIONS:
        return f"扩展名 {ext} 不在常规支持列表中，将尝试自动识别；结果可能不理想。"
    if path.is_file():
        try:
            from azure_settings import precheck_azure_extension

            azure_warn = precheck_azure_extension(path)
            if azure_warn:
                return azure_warn
        except Exception:
            pass
    return None


def _friendly_office_error(path: Path, exc: Exception) -> str | None:
    msg = str(exc)
    low = msg.lower()
    ext = path.suffix.lower()
    if "password" in low or "encrypt" in low:
        return f"文件可能已加密，无法直接转换（{ext}）。请先解除密码后重试。"
    if "badzipfile" in low or "not a zip" in low or "file is not a zip" in low:
        if ext == ".zip":
            return "ZIP 文件已损坏或不是有效压缩包。"
        return f"文件已损坏或不是有效的 Office/OpenXML 包（{ext}）。"
    if "notolefile" in low or "not an ole2" in low:
        return "不是有效的 Outlook .msg（OLE）文件。"
    if ext in {".doc", ".ppt"} and ("no converter" in low or "could not convert" in low):
        return f"不支持老格式 {ext}，请另存为新格式后重试。"
    if ext in {".wav", ".mp3", ".m4a", ".mp4"} and (
        "recognition" in low or "speech" in low or "transcri" in low or "connection" in low
    ):
        return "音频转写失败（可能需联网语音识别或本地 Whisper）。"
    if ext in {".bmp", ".tif", ".tiff", ".heif", ".heic"} and (
        "no converter" in low or "could not convert" in low
    ):
        return f"内置不支持 {ext}，请在「Azure 设置」启用 Document Intelligence 或 Content Understanding。"
    if ext in {".eml", ".rtf"} and ("no converter" in low or "could not convert" in low):
        return f"{ext} 需 Azure Content Understanding，请在「Azure 设置」配置 cu_endpoint。"
    if ext in {".avi", ".mkv", ".mov", ".webm", ".flv", ".wmv", ".m4v"} and (
        "no converter" in low or "could not convert" in low
    ):
        return f"{ext} 视频需 Azure Content Understanding（cu_endpoint）。"
    return None


def postcheck_result(source: str, text: str) -> str | None:
    stripped = (text or "").strip()
    lower = source.lower()
    if not stripped:
        if any(lower.endswith(x) for x in (".jpg", ".jpeg", ".png")):
            return (
                "转换完成，但几乎没有文本。\n"
                "图片默认不做 OCR；可在「大模型设置」启用视觉模型，或启用 OCR 插件。"
            )
        if lower.endswith(".pdf"):
            return (
                "转换完成，但几乎没有文本。\n"
                "若是扫描件/纯图片 PDF，需大模型 OCR 或 Azure Document Intelligence。"
            )
        return "转换完成，但结果为空。"
    if len(stripped) < 8 and any(lower.endswith(x) for x in (".jpg", ".jpeg", ".png")):
        return "图片输出很少，通常仅有元数据；正文 OCR 需额外能力。"
    return None


def _enrich_docx_chrome(path: Path, text: str) -> str:
    """补充 Word 页眉页脚文字。"""
    try:
        from docx import Document
    except Exception:
        return text
    try:
        doc = Document(str(path))
    except Exception:
        return text

    headers: list[str] = []
    footers: list[str] = []
    for sec in doc.sections:
        for p in sec.header.paragraphs:
            t = (p.text or "").strip()
            if t and t not in headers and t not in text:
                headers.append(t)
        for p in sec.footer.paragraphs:
            t = (p.text or "").strip()
            if t and t not in footers and t not in text:
                footers.append(t)
    if not headers and not footers:
        return text

    parts = [text.rstrip()]
    if headers:
        parts.append("\n## 页眉\n\n" + "\n\n".join(headers))
    if footers:
        parts.append("\n## 页脚\n\n" + "\n\n".join(footers))
    return "\n".join(parts).rstrip() + "\n"


def _finalize_markdown(path: Path | None, text: str, *, light: bool = True) -> str:
    cleaner = clean_markdown_light if light else clean_markdown
    text = cleaner(text)
    if light and path is not None and path.suffix.lower() == ".docx":
        text = _enrich_docx_chrome(path, text)
        text = clean_markdown_light(text)
    return text


def _should_use_local_only(local_only: bool | None) -> bool:
    if local_only is not None:
        return local_only
    try:
        from advanced_settings import load_settings

        return bool(load_settings().use_convert_local)
    except Exception:
        return False


def _run_markitdown_local(
    path: Path,
    *,
    progress: Callable[[str], None] | None = None,
    keep_data_uris: bool | None = None,
) -> str:
    if progress:
        progress("正在加载转换引擎…")
    md = _get_markitdown()
    opts = _convert_options()
    if keep_data_uris is True:
        opts = {**opts, "keep_data_uris": True}
    elif keep_data_uris is False and "keep_data_uris" in opts:
        opts = {k: v for k, v in opts.items() if k != "keep_data_uris"}
    try:
        if progress:
            progress(f"正在转换（窄接口）：{path.name}")
        result = md.convert_local(path, **opts)
    except Exception as exc:  # noqa: BLE001
        friendly = _friendly_office_error(path, exc)
        if friendly:
            raise ConversionError(friendly) from exc
        raise ConversionError(f"转换失败：{exc}") from exc

    if progress:
        progress("正在清理结果…")
    return _finalize_markdown(path, _result_text(result), light=True)


def convert_local_path(
    source: str | Path,
    *,
    progress: Callable[[str], None] | None = None,
    keep_data_uris: bool | None = None,
) -> str:
    """窄接口：仅 convert_local，不经过 Excel/Notebook/ZIP 本地增强。"""
    path = Path(source)
    if not path.exists():
        raise ConversionError(f"文件不存在：{source}")
    if not path.is_file():
        raise ConversionError(f"不是有效文件：{source}")
    if progress:
        progress(f"窄接口模式：{path.name}")
    return _run_markitdown_local(path, progress=progress, keep_data_uris=keep_data_uris)


def convert_stream(
    stream: BinaryIO,
    *,
    extension: str | None = None,
    mime_type: str | None = None,
    charset: str | None = None,
    progress: Callable[[str], None] | None = None,
    keep_data_uris: bool | None = None,
) -> str:
    """从二进制流转换（用于 stdin 管道）。"""
    try:
        from markitdown import StreamInfo
    except ImportError as exc:
        raise ConversionError("未安装转换引擎。") from exc

    ext = (extension or "").strip()
    if ext and not ext.startswith("."):
        ext = "." + ext

    stream_info = None
    if ext or mime_type or charset:
        stream_info = StreamInfo(
            extension=ext or None,
            mimetype=(mime_type or None),
            charset=(charset or None),
        )

    if progress:
        progress("正在从标准输入读取…")

    if not stream.seekable():
        buffer = io.BytesIO()
        while True:
            chunk = stream.read(4096)
            if not chunk:
                break
            buffer.write(chunk)
        buffer.seek(0)
        stream = buffer

    md = _get_markitdown()
    opts = _convert_options()
    if keep_data_uris is True:
        opts = {**opts, "keep_data_uris": True}
    elif keep_data_uris is False and "keep_data_uris" in opts:
        opts = {k: v for k, v in opts.items() if k != "keep_data_uris"}
    try:
        if progress:
            progress("正在转换标准输入流…")
        result = md.convert_stream(stream, stream_info=stream_info, **opts)
    except Exception as exc:  # noqa: BLE001
        raise ConversionError(f"转换失败：{exc}") from exc

    if progress:
        progress("正在清理结果…")
    return _finalize_markdown(None, _result_text(result), light=True)


def convert_path(
    source: str | Path,
    *,
    progress: Callable[[str], None] | None = None,
    local_only: bool | None = None,
    keep_data_uris: bool | None = None,
) -> str:
    """将本地文件或 URL 转为 Markdown。"""
    if isinstance(source, Path):
        source = str(source)
    source = (source or "").strip()
    if not source:
        raise ConversionError("请选择文件或输入 URL。")

    narrow = _should_use_local_only(local_only)

    if progress:
        progress("正在初始化转换引擎…")

    is_url = source.startswith(("http://", "https://"))
    path: Path | None = None
    if not is_url:
        path = Path(source)
        if not path.exists():
            raise ConversionError(f"文件不存在：{source}")
        if not path.is_file():
            raise ConversionError(f"不是有效文件：{source}")

        if narrow:
            return convert_local_path(path, progress=progress, keep_data_uris=keep_data_uris)

        if progress:
            progress(f"正在转换：{path.name}")

        if path.suffix.lower() in {".xlsx", ".xls"}:
            try:
                if progress:
                    progress("正在解析 Excel（合并单元格 / 分区域）…")
                from excel_convert import convert_excel_to_markdown

                text = convert_excel_to_markdown(path)
                text = clean_markdown(text)
                if progress:
                    progress("转换完成")
                return text
            except Exception as exc:  # noqa: BLE001
                if progress:
                    progress(f"Excel 增强转换失败，回退标准引擎…（{exc}）")

        if path.suffix.lower() == ".ipynb":
            try:
                if progress:
                    progress("正在解析 Jupyter Notebook（代码输出 / 表格）…")
                from ipynb_convert import convert_ipynb_to_markdown

                text = convert_ipynb_to_markdown(path)
                text = clean_markdown(text)
                if progress:
                    progress("转换完成")
                return text
            except Exception as exc:  # noqa: BLE001
                if progress:
                    progress(f"Notebook 增强转换失败，回退标准引擎…（{exc}）")

        if path.suffix.lower() == ".zip":
            try:
                if progress:
                    progress("正在解析 ZIP…")
                from zip_convert import convert_zip_to_markdown

                text = convert_zip_to_markdown(path, progress=progress)
                text = clean_markdown(text)
                if progress:
                    progress("转换完成")
                return text
            except ValueError as exc:
                raise ConversionError(str(exc)) from exc
            except Exception as exc:  # noqa: BLE001
                if progress:
                    progress(f"ZIP 增强转换失败，回退标准引擎…（{exc}）")
    else:
        if progress:
            progress(f"正在转换 URL：{source}")

    if progress:
        progress("正在加载转换引擎（首次可能较慢）…")
    md = _get_markitdown()
    opts = _convert_options()
    if keep_data_uris is True:
        opts = {**opts, "keep_data_uris": True}
    elif keep_data_uris is False and "keep_data_uris" in opts:
        opts = {k: v for k, v in opts.items() if k != "keep_data_uris"}
    try:
        if progress:
            progress(f"正在转换：{Path(source).name if path else source}")
        result = md.convert(source, **opts)
    except Exception as exc:  # noqa: BLE001
        msg = str(exc)
        if "Max retries" in msg or "ConnectTimeout" in msg or "Connection" in msg:
            raise ConversionError(
                f"网络请求失败，请检查网络或稍后重试。\n详情：{exc}"
            ) from exc
        if "404" in msg:
            raise ConversionError(f"资源不存在（404）。\n详情：{exc}") from exc
        if path is not None:
            friendly = _friendly_office_error(path, exc)
            if friendly:
                raise ConversionError(friendly) from exc
            if path.suffix.lower() == ".zip":
                raise ConversionError(
                    "ZIP 转换失败。若扩展名是 .zip 但实际是 RAR/7z，请先改为真正的 ZIP。\n"
                    f"详情：{exc}"
                ) from exc
        raise ConversionError(f"转换失败：{exc}") from exc

    text = _result_text(result)

    if progress:
        progress("正在清理结果…")
    text = _finalize_markdown(path, text, light=True)

    if progress:
        progress("转换完成")
    return text


def default_output_path(source: str) -> Path:
    if source.startswith(("http://", "https://")):
        return Path.cwd() / "converted.md"
    return Path(source).with_suffix(".md")


def is_supported_file(path: str | Path) -> bool:
    return Path(path).suffix.lower() in SUPPORTED_EXTENSIONS
