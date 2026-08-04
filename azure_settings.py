"""Azure 配置：Document Intelligence 与 Content Understanding。"""

from __future__ import annotations

import json
import os
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any


def settings_path() -> Path:
    return Path.home() / ".doc2md" / "azure_settings.json"


@dataclass
class AzureSettings:
    docintel_enabled: bool = False
    docintel_endpoint: str = ""
    docintel_api_key: str = ""
    docintel_api_version: str = ""
    docintel_file_types: str = ""  # 逗号分隔，如 pdf,docx；空=默认

    cu_enabled: bool = False
    cu_endpoint: str = ""
    cu_api_key: str = ""
    cu_analyzer_id: str = ""
    cu_file_types: str = ""  # 逗号分隔；空=默认

    @property
    def docintel_ready(self) -> bool:
        return bool(self.docintel_enabled and self.docintel_endpoint.strip())

    @property
    def cu_ready(self) -> bool:
        return bool(self.cu_enabled and self.cu_endpoint.strip())

    @property
    def is_ready(self) -> bool:
        return self.docintel_ready or self.cu_ready


def load_settings() -> AzureSettings:
    data: dict[str, Any] = {}
    path = settings_path()
    if path.exists():
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            data = {}

    s = AzureSettings(
        docintel_enabled=bool(data.get("docintel_enabled", False)),
        docintel_endpoint=str(data.get("docintel_endpoint", "") or ""),
        docintel_api_key=str(data.get("docintel_api_key", "") or ""),
        docintel_api_version=str(data.get("docintel_api_version", "") or ""),
        docintel_file_types=str(data.get("docintel_file_types", "") or ""),
        cu_enabled=bool(data.get("cu_enabled", False)),
        cu_endpoint=str(data.get("cu_endpoint", "") or ""),
        cu_api_key=str(data.get("cu_api_key", "") or ""),
        cu_analyzer_id=str(data.get("cu_analyzer_id", "") or ""),
        cu_file_types=str(data.get("cu_file_types", "") or ""),
    )
    if os.environ.get("AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT") and not s.docintel_endpoint:
        s.docintel_endpoint = os.environ["AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT"]
        s.docintel_enabled = True
    if os.environ.get("AZURE_DOCUMENT_INTELLIGENCE_KEY") and not s.docintel_api_key:
        s.docintel_api_key = os.environ["AZURE_DOCUMENT_INTELLIGENCE_KEY"]
    if os.environ.get("AZURE_CONTENT_UNDERSTANDING_ENDPOINT") and not s.cu_endpoint:
        s.cu_endpoint = os.environ["AZURE_CONTENT_UNDERSTANDING_ENDPOINT"]
        s.cu_enabled = True
    if os.environ.get("AZURE_CONTENT_UNDERSTANDING_KEY") and not s.cu_api_key:
        s.cu_api_key = os.environ["AZURE_CONTENT_UNDERSTANDING_KEY"]
    return s


def save_settings(settings: AzureSettings) -> Path:
    path = settings_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(asdict(settings), ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def _azure_key_credential(key: str):
    if not key.strip():
        return None
    try:
        from azure.core.credentials import AzureKeyCredential
    except ImportError as exc:
        raise RuntimeError(
            "未安装 Azure 依赖。请执行：pip install 'markitdown[az-doc-intel,az-content-understanding]'"
        ) from exc
    return AzureKeyCredential(key.strip())


def _parse_docintel_file_types(raw: str):
    from markitdown.converters._doc_intel_converter import DocumentIntelligenceFileType

    names = [t.strip().lower() for t in raw.split(",") if t.strip()]
    if not names:
        return None
    out = []
    for name in names:
        try:
            out.append(DocumentIntelligenceFileType(name))
        except ValueError as exc:
            raise ValueError(f"未知 Document Intelligence 类型：{name}") from exc
    return out


def _parse_cu_file_types(raw: str):
    from markitdown.converters._cu_converter import ContentUnderstandingFileType

    names = [t.strip().lower() for t in raw.split(",") if t.strip()]
    if not names:
        return None
    out = []
    for name in names:
        try:
            out.append(ContentUnderstandingFileType(name))
        except ValueError as exc:
            raise ValueError(f"未知 Content Understanding 类型：{name}") from exc
    return out


def markitdown_azure_kwargs(settings: AzureSettings | None = None) -> dict[str, Any]:
    s = settings or load_settings()
    kwargs: dict[str, Any] = {}

    if s.docintel_ready:
        kwargs["docintel_endpoint"] = s.docintel_endpoint.strip()
        cred = _azure_key_credential(s.docintel_api_key)
        if cred is not None:
            kwargs["docintel_credential"] = cred
        if s.docintel_api_version.strip():
            kwargs["docintel_api_version"] = s.docintel_api_version.strip()
        types = _parse_docintel_file_types(s.docintel_file_types)
        if types is not None:
            kwargs["docintel_file_types"] = types

    if s.cu_ready:
        kwargs["cu_endpoint"] = s.cu_endpoint.strip()
        cred = _azure_key_credential(s.cu_api_key)
        if cred is not None:
            kwargs["cu_credential"] = cred
        if s.cu_analyzer_id.strip():
            kwargs["cu_analyzer_id"] = s.cu_analyzer_id.strip()
        types = _parse_cu_file_types(s.cu_file_types)
        if types is not None:
            kwargs["cu_file_types"] = types

    return kwargs


CU_ONLY_EXTENSIONS = {
    ".eml",
    ".rtf",
    ".heif",
    ".heic",
    ".flac",
    ".ogg",
    ".aac",
    ".wma",
    ".m4v",
    ".webm",
    ".flv",
    ".wmv",
}

AZURE_EXTENDED_EXTENSIONS = {
    ".bmp",
    ".tif",
    ".tiff",
    ".avi",
    ".mkv",
    ".mov",
} | CU_ONLY_EXTENSIONS


def azure_status_text(settings: AzureSettings | None = None) -> str:
    s = settings or load_settings()
    parts: list[str] = []
    if s.docintel_ready:
        parts.append("DocIntel")
    if s.cu_ready:
        tag = "CU"
        if s.cu_analyzer_id.strip():
            tag += f"({s.cu_analyzer_id.strip()})"
        parts.append(tag)
    if not parts:
        return "Azure：未启用"
    return "Azure：" + " · ".join(parts)


def precheck_azure_extension(path: Path, settings: AzureSettings | None = None) -> str | None:
    s = settings or load_settings()
    ext = path.suffix.lower()
    if ext in CU_ONLY_EXTENSIONS and not s.cu_ready:
        return (
            f"{ext} 需 Azure Content Understanding。\n"
            "请在「Azure 设置」中启用并填写 cu_endpoint（及 API Key 或本机 az login）。"
        )
    if ext in AZURE_EXTENDED_EXTENSIONS:
        if ext in {".bmp", ".tif", ".tiff"} and not (s.docintel_ready or s.cu_ready):
            return (
                f"{ext} 内置转换不支持；需 Azure Document Intelligence 或 Content Understanding。\n"
                "请在「Azure 设置」中配置 endpoint。"
            )
        if ext in {".avi", ".mkv", ".mov"} and not s.cu_ready:
            return (
                f"{ext} 完整视频理解需 Azure Content Understanding。\n"
                "请在「Azure 设置」中启用 cu_endpoint。"
            )
    return None
