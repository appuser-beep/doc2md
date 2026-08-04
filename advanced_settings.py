"""高级转换参数：样式映射、ExifTool、窄接口、数据 URI、自定义插件。"""



from __future__ import annotations



import json

import os

import shutil

from dataclasses import asdict, dataclass

from pathlib import Path

from typing import Any





def settings_path() -> Path:

    return Path.home() / ".doc2md" / "advanced_settings.json"





@dataclass

class AdvancedSettings:

    style_map: str = ""

    exiftool_path: str = ""

    use_convert_local: bool = False

    keep_data_uris: bool = False

    custom_plugin_scripts: str = ""  # 每行一个 .py 路径



    @property

    def resolved_exiftool_path(self) -> str:

        raw = (self.exiftool_path or "").strip()

        if raw:

            return raw

        env = os.environ.get("EXIFTOOL_PATH", "").strip()

        if env:

            return env

        found = shutil.which("exiftool")

        return found or ""



    @property

    def has_style_map(self) -> bool:

        return bool(self.style_map.strip())



    @property

    def has_exiftool(self) -> bool:

        p = self.resolved_exiftool_path

        return bool(p and Path(p).is_file())



    @property

    def custom_plugin_paths(self) -> list[str]:

        out: list[str] = []

        for line in self.custom_plugin_scripts.splitlines():

            p = line.strip()

            if p and p not in out:

                out.append(p)

        return out



    @property

    def has_custom_plugins(self) -> bool:

        return bool(self.custom_plugin_paths)





def load_settings() -> AdvancedSettings:

    data: dict[str, Any] = {}

    path = settings_path()

    if path.exists():

        try:

            data = json.loads(path.read_text(encoding="utf-8"))

        except Exception:

            data = {}



    s = AdvancedSettings(

        style_map=str(data.get("style_map", "") or ""),

        exiftool_path=str(data.get("exiftool_path", "") or ""),

        use_convert_local=bool(data.get("use_convert_local", False)),

        keep_data_uris=bool(data.get("keep_data_uris", False)),

        custom_plugin_scripts=str(

            data.get("custom_plugin_scripts", data.get("custom_plugin_paths", "")) or ""

        ),

    )

    if os.environ.get("EXIFTOOL_PATH") and not s.exiftool_path:

        s.exiftool_path = os.environ["EXIFTOOL_PATH"]

    return s





def save_settings(settings: AdvancedSettings) -> Path:

    path = settings_path()

    path.parent.mkdir(parents=True, exist_ok=True)

    path.write_text(json.dumps(asdict(settings), ensure_ascii=False, indent=2), encoding="utf-8")

    return path





def markitdown_advanced_kwargs(settings: AdvancedSettings | None = None) -> dict[str, Any]:

    s = settings or load_settings()

    kwargs: dict[str, Any] = {}

    if s.style_map.strip():

        kwargs["style_map"] = s.style_map.strip()

    if s.exiftool_path.strip():

        kwargs["exiftool_path"] = s.exiftool_path.strip()

    elif s.resolved_exiftool_path and not s.exiftool_path.strip():

        kwargs["exiftool_path"] = s.resolved_exiftool_path

    return kwargs





def convert_options(settings: AdvancedSettings | None = None) -> dict[str, Any]:

    """传给 convert / convert_local / convert_stream 的选项。"""

    s = settings or load_settings()

    if s.keep_data_uris:

        return {"keep_data_uris": True}

    return {}





def advanced_status_text(settings: AdvancedSettings | None = None) -> str:

    s = settings or load_settings()

    parts: list[str] = []

    if s.has_style_map:

        parts.append("样式映射")

    if s.has_exiftool:

        parts.append("ExifTool")

    if s.use_convert_local:

        parts.append("窄接口")

    if s.keep_data_uris:

        parts.append("保留内嵌图")

    if s.has_custom_plugins:

        parts.append("自定义插件")

    if not parts:

        return "高级：默认"

    return "高级：" + " · ".join(parts)

