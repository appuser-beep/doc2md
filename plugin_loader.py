"""第三方与自定义转换插件加载。"""

from __future__ import annotations

import importlib.util
import sys
from importlib.metadata import entry_points
from pathlib import Path
from typing import Any


def list_plugin_entry_points() -> list[tuple[str, str]]:
    """返回 (名称, 包路径) 列表；兼容 PyInstaller 冻结包。"""
    out: list[tuple[str, str]] = []
    seen: set[str] = set()
    try:
        eps = entry_points()
        group = eps.select(group="markitdown.plugin") if hasattr(eps, "select") else eps.get(
            "markitdown.plugin", []
        )
        for ep in group:
            key = ep.name.lower()
            if key in seen:
                continue
            seen.add(key)
            out.append((ep.name, ep.value))
    except Exception:
        pass

    # 冻结 exe 时常丢 entry_points 元数据：探测已打包的 OCR 插件
    if "ocr" not in seen:
        try:
            import markitdown_ocr  # noqa: F401

            out.append(("ocr", "markitdown_ocr"))
            seen.add("ocr")
        except Exception:
            pass

    out.sort(key=lambda x: x[0].lower())
    return out


def format_plugin_list() -> str:
    lines = ["已安装的第三方转换插件：", ""]
    plugins = list_plugin_entry_points()
    if not plugins:
        lines.append("  （无）")
        lines.append("")
        lines.append("可在「大模型设置」中启用 OCR 插件。")
    else:
        for name, value in plugins:
            lines.append(f"  · {name:<16}  {value}")
        lines.append("")
        lines.append("启用 OCR：「大模型设置」→ 启用 OCR 插件。")
        lines.append("命令行查看：doc2md-cli --list-plugins")
    return "\n".join(lines)


def _load_module_from_path(path: Path):
    spec = importlib.util.spec_from_file_location(f"doc2md_plugin_{path.stem}", path)
    if spec is None or spec.loader is None:
        raise ImportError(f"无法加载插件模块：{path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def apply_custom_converter_plugins(markitdown: Any, engine_kwargs: dict[str, Any]) -> list[str]:
    """从 advanced_settings 中的路径加载 register_converters，返回已加载插件名。"""
    try:
        from advanced_settings import load_settings
    except Exception:
        return []

    paths = load_settings().custom_plugin_paths
    loaded: list[str] = []
    for raw in paths:
        path = Path(raw)
        if not path.is_file():
            continue
        try:
            module = _load_module_from_path(path)
        except Exception as exc:
            raise RuntimeError(f"加载自定义插件失败（{path}）：{exc}") from exc

        register = getattr(module, "register_converters", None)
        if register is None:
            raise RuntimeError(
                f"自定义插件 {path} 须定义 register_converters(markitdown, **kwargs)"
            )
        register(markitdown, **engine_kwargs)
        loaded.append(path.name)
    return loaded
