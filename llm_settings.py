"""大模型配置：图片描述与 OCR 插件。"""

from __future__ import annotations

import json
import os
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any


DEFAULT_MODEL = "gpt-4o"


def settings_path() -> Path:
    return Path.home() / ".doc2md" / "llm_settings.json"


@dataclass
class LlmSettings:
    enabled: bool = False
    api_key: str = ""
    base_url: str = ""  # 空 = 默认网关
    model: str = DEFAULT_MODEL
    prompt: str = ""  # 空 = 内置默认 prompt
    enable_plugins: bool = False  # 勾选 OCR 时再 True

    @property
    def resolved_api_key(self) -> str:
        return self.api_key.strip() or os.environ.get("OPENAI_API_KEY", "").strip()

    @property
    def is_ready(self) -> bool:
        return bool(self.enabled and self.resolved_api_key and self.model.strip())


def load_settings() -> LlmSettings:
    data: dict[str, Any] = {}
    path = settings_path()
    if path.exists():
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            data = {}

    s = LlmSettings(
        enabled=bool(data.get("enabled", False)),
        api_key=str(data.get("api_key", "") or ""),
        base_url=str(data.get("base_url", "") or ""),
        model=str(data.get("model", DEFAULT_MODEL) or DEFAULT_MODEL),
        prompt=str(data.get("prompt", "") or ""),
        enable_plugins=bool(data.get("enable_plugins", False)),
    )
    if os.environ.get("OPENAI_BASE_URL") and not s.base_url:
        s.base_url = os.environ["OPENAI_BASE_URL"]
    return s


def save_settings(settings: LlmSettings) -> Path:
    path = settings_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(asdict(settings), ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def build_openai_client(settings: LlmSettings | None = None):
    s = settings or load_settings()
    if not s.is_ready:
        return None
    try:
        from openai import OpenAI
    except ImportError as exc:
        raise RuntimeError("未安装 openai。请：pip install openai") from exc

    kwargs: dict[str, Any] = {}
    if s.resolved_api_key:
        kwargs["api_key"] = s.resolved_api_key
    base = (s.base_url or "").strip().rstrip("/")
    if base:
        kwargs["base_url"] = base
    return OpenAI(**kwargs)


def markitdown_llm_kwargs(settings: LlmSettings | None = None) -> dict[str, Any]:
    s = settings or load_settings()
    if not s.is_ready:
        return {"enable_plugins": False}

    client = build_openai_client(s)
    if client is None:
        return {"enable_plugins": False}

    kwargs: dict[str, Any] = {
        "enable_plugins": bool(s.enable_plugins),
        "llm_client": client,
        "llm_model": s.model.strip(),
    }
    if s.prompt.strip():
        kwargs["llm_prompt"] = s.prompt.strip()
    return kwargs


def llm_status_text(settings: LlmSettings | None = None) -> str:
    s = settings or load_settings()
    if not s.enabled:
        return "大模型：未启用"
    if not s.resolved_api_key:
        return "大模型：缺少 OPENAI_API_KEY"
    extra = "·plugins" if s.enable_plugins else ""
    return f"大模型：{s.model}{extra}"
