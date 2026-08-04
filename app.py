"""文档转 Markdown — 浅色办公风桌面工具。"""

from __future__ import annotations

import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

from converter import (
    APP_VERSION,
    FILE_DIALOG_TYPES,
    SUPPORT_HELP,
    ConversionError,
    convert_path,
    default_output_path,
    postcheck_result,
    precheck_source,
)
from llm_settings import LlmSettings, llm_status_text, load_settings, save_settings
from azure_settings import AzureSettings, azure_status_text, load_settings as load_azure_settings, save_settings as save_azure_settings
from advanced_settings import (
    AdvancedSettings,
    advanced_status_text,
    load_settings as load_advanced_settings,
    save_settings as save_advanced_settings,
)

# —— 浅色办公风配色 ——
COLORS = {
    "bg": "#F3F4F6",
    "surface": "#FFFFFF",
    "border": "#D8DEE6",
    "text": "#1F2937",
    "muted": "#6B7280",
    "accent": "#2563EB",
    "accent_hover": "#1D4ED8",
    "accent_soft": "#EFF6FF",
    "success": "#059669",
    "danger": "#DC2626",
    "warning": "#D97706",
    "preview_bg": "#FAFBFC",
}

APP_TITLE = "文档转 Markdown"
APP_SUBTITLE = "其他格式 → Markdown"


class App(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        self.title(f"{APP_TITLE} v{APP_VERSION}")
        self.geometry("1000x720")
        self.minsize(860, 600)
        self.configure(fg_color=COLORS["bg"])

        self._source_path = tk.StringVar(value="")
        self._url = tk.StringVar(value="")
        self._status = tk.StringVar(value="就绪 — 请选择文件或输入 URL")
        self._busy = False
        self._markdown = ""
        self._current_source = ""

        self._build_ui()

    def _build_ui(self) -> None:
        root = ctk.CTkFrame(self, fg_color=COLORS["bg"], corner_radius=0)
        root.pack(fill="both", expand=True, padx=20, pady=16)
        root.grid_columnconfigure(0, weight=1)
        root.grid_rowconfigure(3, weight=1)

        header = ctk.CTkFrame(root, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 14))
        header.grid_columnconfigure(0, weight=1)

        left = ctk.CTkFrame(header, fg_color="transparent")
        left.grid(row=0, column=0, sticky="w")
        ctk.CTkLabel(
            left,
            text=APP_TITLE,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=24, weight="bold"),
            text_color=COLORS["text"],
            anchor="w",
        ).pack(anchor="w")
        ctk.CTkLabel(
            left,
            text=APP_SUBTITLE,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=13),
            text_color=COLORS["muted"],
            anchor="w",
        ).pack(anchor="w", pady=(4, 0))

        btn_row = ctk.CTkFrame(header, fg_color="transparent")
        btn_row.grid(row=0, column=1, sticky="e")
        ctk.CTkButton(
            btn_row,
            text="高级设置",
            width=96,
            height=34,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=13),
            fg_color=COLORS["surface"],
            hover_color="#EEF2F7",
            text_color=COLORS["text"],
            border_width=1,
            border_color=COLORS["border"],
            command=self._show_advanced_settings,
        ).pack(side="left", padx=(0, 8))
        ctk.CTkButton(
            btn_row,
            text="Azure 设置",
            width=96,
            height=34,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=13),
            fg_color=COLORS["surface"],
            hover_color="#EEF2F7",
            text_color=COLORS["text"],
            border_width=1,
            border_color=COLORS["border"],
            command=self._show_azure_settings,
        ).pack(side="left", padx=(0, 8))
        ctk.CTkButton(
            btn_row,
            text="大模型设置",
            width=96,
            height=34,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=13),
            fg_color=COLORS["surface"],
            hover_color="#EEF2F7",
            text_color=COLORS["text"],
            border_width=1,
            border_color=COLORS["border"],
            command=self._show_llm_settings,
        ).pack(side="left", padx=(0, 8))
        ctk.CTkButton(
            btn_row,
            text="格式说明",
            width=96,
            height=34,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=13),
            fg_color=COLORS["surface"],
            hover_color="#EEF2F7",
            text_color=COLORS["text"],
            border_width=1,
            border_color=COLORS["border"],
            command=self._show_help,
        ).pack(side="left")

        input_card = ctk.CTkFrame(
            root,
            fg_color=COLORS["surface"],
            border_width=1,
            border_color=COLORS["border"],
            corner_radius=8,
        )
        input_card.grid(row=1, column=0, sticky="ew", pady=(0, 12))
        input_card.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            input_card,
            text="本地文件",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=13, weight="bold"),
            text_color=COLORS["text"],
            width=80,
            anchor="w",
        ).grid(row=0, column=0, sticky="w", padx=(16, 8), pady=(16, 6))

        self.path_entry = ctk.CTkEntry(
            input_card,
            textvariable=self._source_path,
            placeholder_text="PDF / Word(.docx) / Excel / PPT / HTML / ZIP / MSG…",
            height=36,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=13),
            fg_color=COLORS["preview_bg"],
            border_color=COLORS["border"],
            text_color=COLORS["text"],
        )
        self.path_entry.grid(row=0, column=1, sticky="ew", pady=(16, 6), padx=(0, 8))

        ctk.CTkButton(
            input_card,
            text="浏览…",
            width=88,
            height=36,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=13),
            fg_color=COLORS["accent_soft"],
            hover_color="#DBEAFE",
            text_color=COLORS["accent"],
            border_width=1,
            border_color="#BFDBFE",
            command=self._browse_file,
        ).grid(row=0, column=2, padx=(0, 16), pady=(16, 6))

        ctk.CTkLabel(
            input_card,
            text="或 URL",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=13, weight="bold"),
            text_color=COLORS["text"],
            width=80,
            anchor="w",
        ).grid(row=1, column=0, sticky="w", padx=(16, 8), pady=(6, 16))

        self.url_entry = ctk.CTkEntry(
            input_card,
            textvariable=self._url,
            placeholder_text="https://…（网页 / Wikipedia / YouTube 等，需可访问网络）",
            height=36,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=13),
            fg_color=COLORS["preview_bg"],
            border_color=COLORS["border"],
            text_color=COLORS["text"],
        )
        self.url_entry.grid(row=1, column=1, columnspan=2, sticky="ew", padx=(0, 16), pady=(6, 16))

        tip_row = ctk.CTkFrame(root, fg_color="transparent")
        tip_row.grid(row=2, column=0, sticky="ew", pady=(0, 8))
        tip_row.grid_columnconfigure(0, weight=1)
        tip_row.grid_columnconfigure(1, weight=0)
        tip_row.grid_columnconfigure(2, weight=0)
        tip_row.grid_columnconfigure(3, weight=0)
        ctk.CTkLabel(
            tip_row,
            text="提示：颜色/字号不会保留；标题、列表、表格会保留。扫描件/视频/EML 等需 Azure；图片描述需大模型。",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=12),
            text_color=COLORS["muted"],
            anchor="w",
        ).grid(row=0, column=0, sticky="w")
        self._advanced_status = tk.StringVar(value=advanced_status_text())
        ctk.CTkLabel(
            tip_row,
            textvariable=self._advanced_status,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=12),
            text_color=COLORS["muted"],
            anchor="e",
        ).grid(row=0, column=1, sticky="e", padx=(12, 8))
        self._azure_status = tk.StringVar(value=azure_status_text())
        ctk.CTkLabel(
            tip_row,
            textvariable=self._azure_status,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=12),
            text_color=COLORS["warning"],
            anchor="e",
        ).grid(row=0, column=2, sticky="e", padx=(0, 8))
        self._llm_status = tk.StringVar(value=llm_status_text())
        ctk.CTkLabel(
            tip_row,
            textvariable=self._llm_status,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=12),
            text_color=COLORS["accent"],
            anchor="e",
        ).grid(row=0, column=3, sticky="e")

        actions = ctk.CTkFrame(root, fg_color="transparent")
        actions.grid(row=3, column=0, sticky="ew", pady=(0, 12))

        self.convert_btn = ctk.CTkButton(
            actions,
            text="开始转换",
            width=120,
            height=38,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=14, weight="bold"),
            fg_color=COLORS["accent"],
            hover_color=COLORS["accent_hover"],
            text_color="#FFFFFF",
            command=self._start_convert,
        )
        self.convert_btn.pack(side="left")

        self.save_btn = ctk.CTkButton(
            actions,
            text="保存 Markdown",
            width=130,
            height=38,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=13),
            fg_color=COLORS["surface"],
            hover_color="#EEF2F7",
            text_color=COLORS["text"],
            border_width=1,
            border_color=COLORS["border"],
            state="disabled",
            command=self._save_markdown,
        )
        self.save_btn.pack(side="left", padx=(10, 0))

        self.copy_btn = ctk.CTkButton(
            actions,
            text="复制",
            width=72,
            height=38,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=13),
            fg_color=COLORS["surface"],
            hover_color="#EEF2F7",
            text_color=COLORS["text"],
            border_width=1,
            border_color=COLORS["border"],
            state="disabled",
            command=self._copy_markdown,
        )
        self.copy_btn.pack(side="left", padx=(10, 0))

        self.clear_btn = ctk.CTkButton(
            actions,
            text="清空",
            width=72,
            height=38,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=13),
            fg_color=COLORS["surface"],
            hover_color="#EEF2F7",
            text_color=COLORS["muted"],
            border_width=1,
            border_color=COLORS["border"],
            command=self._clear,
        )
        self.clear_btn.pack(side="left", padx=(10, 0))

        ctk.CTkLabel(
            actions,
            text="输出：.md",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=12),
            text_color=COLORS["muted"],
        ).pack(side="right")

        preview_wrap = ctk.CTkFrame(
            root,
            fg_color=COLORS["surface"],
            border_width=1,
            border_color=COLORS["border"],
            corner_radius=8,
        )
        preview_wrap.grid(row=4, column=0, sticky="nsew")
        root.grid_rowconfigure(4, weight=1)
        preview_wrap.grid_columnconfigure(0, weight=1)
        preview_wrap.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(
            preview_wrap,
            text="Markdown 预览",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=13, weight="bold"),
            text_color=COLORS["text"],
            anchor="w",
        ).grid(row=0, column=0, sticky="w", padx=16, pady=(12, 6))

        self.preview = ctk.CTkTextbox(
            preview_wrap,
            font=ctk.CTkFont(family="Consolas", size=13),
            fg_color=COLORS["preview_bg"],
            text_color=COLORS["text"],
            border_width=0,
            corner_radius=6,
            wrap="word",
        )
        self.preview.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 12))
        self.preview.insert("1.0", "转换结果将显示在这里。")
        self.preview.configure(state="disabled")

        status_bar = ctk.CTkFrame(root, fg_color="transparent", height=28)
        status_bar.grid(row=5, column=0, sticky="ew", pady=(10, 0))
        self.status_label = ctk.CTkLabel(
            status_bar,
            textvariable=self._status,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=12),
            text_color=COLORS["muted"],
            anchor="w",
        )
        self.status_label.pack(side="left", fill="x", expand=True)

        self.progress = ctk.CTkProgressBar(
            status_bar,
            width=140,
            height=8,
            progress_color=COLORS["accent"],
            fg_color="#E5E7EB",
        )
        self.progress.pack(side="right")
        self.progress.set(0)

    def _show_help(self) -> None:
        win = ctk.CTkToplevel(self)
        win.title(f"{APP_TITLE} — 格式说明")
        win.geometry("680x620")
        win.minsize(560, 480)
        win.transient(self)
        win.grab_set()
        win.configure(fg_color=COLORS["bg"])

        frame = ctk.CTkFrame(
            win,
            fg_color=COLORS["surface"],
            border_width=1,
            border_color=COLORS["border"],
            corner_radius=8,
        )
        frame.pack(fill="both", expand=True, padx=16, pady=16)
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)

        box = ctk.CTkTextbox(
            frame,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=12),
            fg_color=COLORS["preview_bg"],
            border_color=COLORS["border"],
            border_width=1,
            text_color=COLORS["text"],
            wrap="word",
        )
        box.grid(row=0, column=0, sticky="nsew", padx=12, pady=(12, 8))
        box.insert("1.0", SUPPORT_HELP)
        box.configure(state="disabled")

        ctk.CTkButton(
            frame,
            text="关闭",
            width=88,
            height=34,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=13),
            fg_color=COLORS["accent"],
            hover_color=COLORS["accent_hover"],
            command=win.destroy,
        ).grid(row=1, column=0, sticky="e", padx=12, pady=(0, 12))

    def _show_llm_settings(self) -> None:
        cur = load_settings()
        win = ctk.CTkToplevel(self)
        win.title("大模型设置")
        win.geometry("580x520")
        win.resizable(False, False)
        win.transient(self)
        win.grab_set()
        win.configure(fg_color=COLORS["bg"])

        frame = ctk.CTkFrame(
            win,
            fg_color=COLORS["surface"],
            border_width=1,
            border_color=COLORS["border"],
            corner_radius=8,
        )
        frame.pack(fill="both", expand=True, padx=16, pady=16)
        frame.grid_columnconfigure(1, weight=1)

        font = ctk.CTkFont(family="Microsoft YaHei UI", size=13)
        bold = ctk.CTkFont(family="Microsoft YaHei UI", size=13, weight="bold")

        enabled_var = tk.BooleanVar(value=cur.enabled)
        ctk.CTkCheckBox(
            frame,
            text="启用大模型（图片描述等）",
            variable=enabled_var,
            font=bold,
            text_color=COLORS["text"],
            fg_color=COLORS["accent"],
            hover_color=COLORS["accent_hover"],
        ).grid(row=1, column=0, columnspan=2, sticky="w", padx=16, pady=(0, 6))

        plugins_var = tk.BooleanVar(value=cur.enable_plugins)
        ctk.CTkCheckBox(
            frame,
            text="启用 OCR 插件（PDF / Office 内嵌图文字识别）",
            variable=plugins_var,
            font=font,
            text_color=COLORS["text"],
            fg_color=COLORS["accent"],
            hover_color=COLORS["accent_hover"],
        ).grid(row=2, column=0, columnspan=2, sticky="w", padx=16, pady=(0, 10))

        def add_row(r: int, label: str, value: str, show: str | None = None):
            ctk.CTkLabel(frame, text=label, font=bold, text_color=COLORS["text"], width=100, anchor="w").grid(
                row=r, column=0, sticky="w", padx=(16, 8), pady=6
            )
            entry = ctk.CTkEntry(
                frame,
                height=34,
                font=font,
                fg_color=COLORS["preview_bg"],
                border_color=COLORS["border"],
                text_color=COLORS["text"],
                show=show or "",
            )
            entry.insert(0, value)
            entry.grid(row=r, column=1, sticky="ew", padx=(0, 16), pady=6)
            return entry

        key_entry = add_row(3, "API Key", cur.api_key or "", show="*")
        url_entry = add_row(4, "Base URL", cur.base_url or "")
        model_entry = add_row(5, "llm_model", cur.model or "gpt-4o")

        ctk.CTkLabel(frame, text="llm_prompt", font=bold, text_color=COLORS["text"], width=100, anchor="nw").grid(
            row=6, column=0, sticky="nw", padx=(16, 8), pady=6
        )
        prompt_box = ctk.CTkTextbox(
            frame,
            height=90,
            font=font,
            fg_color=COLORS["preview_bg"],
            border_color=COLORS["border"],
            border_width=1,
            text_color=COLORS["text"],
        )
        prompt_box.grid(row=6, column=1, sticky="ew", padx=(0, 16), pady=6)
        prompt_box.insert("1.0", cur.prompt)

        ctk.CTkLabel(
            frame,
            text="Key 可留空并改用环境变量 OPENAI_API_KEY；Base URL 留空则使用默认网关；模型默认 gpt-4o。",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=12),
            text_color=COLORS["muted"],
            anchor="w",
            wraplength=480,
            justify="left",
        ).grid(row=7, column=0, columnspan=2, sticky="ew", padx=16, pady=(4, 8))

        btns = ctk.CTkFrame(frame, fg_color="transparent")
        btns.grid(row=8, column=0, columnspan=2, sticky="e", padx=16, pady=(8, 16))

        def on_save() -> None:
            from llm_settings import settings_path

            s = LlmSettings(
                enabled=bool(enabled_var.get()),
                api_key=key_entry.get().strip(),
                base_url=url_entry.get().strip(),
                model=model_entry.get().strip() or "gpt-4o",
                prompt=prompt_box.get("1.0", "end").strip(),
                enable_plugins=bool(plugins_var.get()),
            )
            save_settings(s)
            self._llm_status.set(llm_status_text(s))
            messagebox.showinfo(APP_TITLE, f"大模型设置已保存。\n\n{settings_path()}")
            win.destroy()

        ctk.CTkButton(
            btns,
            text="取消",
            width=88,
            height=34,
            font=font,
            fg_color=COLORS["surface"],
            hover_color="#EEF2F7",
            text_color=COLORS["text"],
            border_width=1,
            border_color=COLORS["border"],
            command=win.destroy,
        ).pack(side="left", padx=(0, 8))
        ctk.CTkButton(
            btns,
            text="保存",
            width=88,
            height=34,
            font=font,
            fg_color=COLORS["accent"],
            hover_color=COLORS["accent_hover"],
            command=on_save,
        ).pack(side="left")

    def _show_azure_settings(self) -> None:
        cur = load_azure_settings()
        win = ctk.CTkToplevel(self)
        win.title("Azure 设置")
        win.geometry("620x680")
        win.resizable(False, False)
        win.transient(self)
        win.grab_set()
        win.configure(fg_color=COLORS["bg"])

        frame = ctk.CTkFrame(
            win,
            fg_color=COLORS["surface"],
            border_width=1,
            border_color=COLORS["border"],
            corner_radius=8,
        )
        frame.pack(fill="both", expand=True, padx=16, pady=16)
        frame.grid_columnconfigure(1, weight=1)

        font = ctk.CTkFont(family="Microsoft YaHei UI", size=13)
        bold = ctk.CTkFont(family="Microsoft YaHei UI", size=13, weight="bold")

        docintel_var = tk.BooleanVar(value=cur.docintel_enabled)
        ctk.CTkCheckBox(
            frame,
            text="启用 Document Intelligence（扫描 PDF、BMP/TIFF 等）",
            variable=docintel_var,
            font=bold,
            text_color=COLORS["text"],
            fg_color=COLORS["accent"],
            hover_color=COLORS["accent_hover"],
        ).grid(row=1, column=0, columnspan=2, sticky="w", padx=16, pady=(0, 6))

        def add_row(r: int, label: str, value: str, show: str | None = None):
            ctk.CTkLabel(frame, text=label, font=bold, text_color=COLORS["text"], width=120, anchor="w").grid(
                row=r, column=0, sticky="w", padx=(16, 8), pady=6
            )
            entry = ctk.CTkEntry(
                frame,
                height=34,
                font=font,
                fg_color=COLORS["preview_bg"],
                border_color=COLORS["border"],
                text_color=COLORS["text"],
                show=show or "",
            )
            entry.insert(0, value)
            entry.grid(row=r, column=1, sticky="ew", padx=(0, 16), pady=6)
            return entry

        docintel_ep = add_row(2, "docintel_endpoint", cur.docintel_endpoint or "")
        docintel_key = add_row(3, "DocIntel Key", cur.docintel_api_key or "", show="*")
        docintel_ver = add_row(4, "docintel_api_version", cur.docintel_api_version or "")
        docintel_types = add_row(5, "docintel_file_types", cur.docintel_file_types or "")

        cu_var = tk.BooleanVar(value=cur.cu_enabled)
        ctk.CTkCheckBox(
            frame,
            text="启用 Content Understanding（视频、EML/RTF、更多音频）",
            variable=cu_var,
            font=bold,
            text_color=COLORS["text"],
            fg_color=COLORS["accent"],
            hover_color=COLORS["accent_hover"],
        ).grid(row=6, column=0, columnspan=2, sticky="w", padx=16, pady=(12, 6))

        cu_ep = add_row(7, "cu_endpoint", cur.cu_endpoint or "")
        cu_key = add_row(8, "CU Key", cur.cu_api_key or "", show="*")
        cu_analyzer = add_row(9, "cu_analyzer_id", cur.cu_analyzer_id or "")
        cu_types = add_row(10, "cu_file_types", cur.cu_file_types or "")

        ctk.CTkLabel(
            frame,
            text="file_types 逗号分隔（如 pdf,docx）；留空使用默认类型。Key 可留空并用 az login。",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=12),
            text_color=COLORS["muted"],
            anchor="w",
            wraplength=520,
            justify="left",
        ).grid(row=11, column=0, columnspan=2, sticky="ew", padx=16, pady=(4, 8))

        btns = ctk.CTkFrame(frame, fg_color="transparent")
        btns.grid(row=12, column=0, columnspan=2, sticky="e", padx=16, pady=(8, 16))

        def on_save() -> None:
            from azure_settings import settings_path

            s = AzureSettings(
                docintel_enabled=bool(docintel_var.get()),
                docintel_endpoint=docintel_ep.get().strip(),
                docintel_api_key=docintel_key.get().strip(),
                docintel_api_version=docintel_ver.get().strip(),
                docintel_file_types=docintel_types.get().strip(),
                cu_enabled=bool(cu_var.get()),
                cu_endpoint=cu_ep.get().strip(),
                cu_api_key=cu_key.get().strip(),
                cu_analyzer_id=cu_analyzer.get().strip(),
                cu_file_types=cu_types.get().strip(),
            )
            save_azure_settings(s)
            self._azure_status.set(azure_status_text(s))
            messagebox.showinfo(APP_TITLE, f"Azure 设置已保存。\n\n{settings_path()}")
            win.destroy()

        ctk.CTkButton(
            btns,
            text="取消",
            width=88,
            height=34,
            font=font,
            fg_color=COLORS["surface"],
            hover_color="#EEF2F7",
            text_color=COLORS["text"],
            border_width=1,
            border_color=COLORS["border"],
            command=win.destroy,
        ).pack(side="left", padx=(0, 8))
        ctk.CTkButton(
            btns,
            text="保存",
            width=88,
            height=34,
            font=font,
            fg_color=COLORS["accent"],
            hover_color=COLORS["accent_hover"],
            command=on_save,
        ).pack(side="left")

    def _show_advanced_settings(self) -> None:
        cur = load_advanced_settings()
        win = ctk.CTkToplevel(self)
        win.title("高级设置")
        win.geometry("640x680")
        win.resizable(False, False)
        win.transient(self)
        win.grab_set()
        win.configure(fg_color=COLORS["bg"])

        frame = ctk.CTkFrame(
            win,
            fg_color=COLORS["surface"],
            border_width=1,
            border_color=COLORS["border"],
            corner_radius=8,
        )
        frame.pack(fill="both", expand=True, padx=16, pady=16)
        frame.grid_columnconfigure(1, weight=1)

        font = ctk.CTkFont(family="Microsoft YaHei UI", size=13)
        bold = ctk.CTkFont(family="Microsoft YaHei UI", size=13, weight="bold")

        local_var = tk.BooleanVar(value=cur.use_convert_local)
        keep_uri_var = tk.BooleanVar(value=cur.keep_data_uris)
        ctk.CTkCheckBox(
            frame,
            text="窄接口模式（跳过 Excel / Notebook / ZIP 增强）",
            variable=local_var,
            font=bold,
            text_color=COLORS["text"],
            fg_color=COLORS["accent"],
            hover_color=COLORS["accent_hover"],
        ).grid(row=0, column=0, columnspan=2, sticky="w", padx=16, pady=(12, 6))
        ctk.CTkCheckBox(
            frame,
            text="保留内嵌图（输出中保留 base64 图片，文件会变大）",
            variable=keep_uri_var,
            font=font,
            text_color=COLORS["text"],
            fg_color=COLORS["accent"],
            hover_color=COLORS["accent_hover"],
        ).grid(row=1, column=0, columnspan=2, sticky="w", padx=16, pady=(0, 10))

        ctk.CTkLabel(
            frame,
            text="Word 样式映射",
            font=bold,
            text_color=COLORS["text"],
            width=100,
            anchor="nw",
        ).grid(row=2, column=0, sticky="nw", padx=(16, 8), pady=6)
        style_box = ctk.CTkTextbox(
            frame,
            height=100,
            font=font,
            fg_color=COLORS["preview_bg"],
            border_color=COLORS["border"],
            border_width=1,
            text_color=COLORS["text"],
        )
        style_box.grid(row=2, column=1, sticky="ew", padx=(0, 16), pady=6)
        style_box.insert("1.0", cur.style_map)

        ctk.CTkLabel(
            frame,
            text="ExifTool 路径",
            font=bold,
            text_color=COLORS["text"],
            width=100,
            anchor="w",
        ).grid(row=3, column=0, sticky="w", padx=(16, 8), pady=6)
        exif_row = ctk.CTkFrame(frame, fg_color="transparent")
        exif_row.grid(row=3, column=1, sticky="ew", padx=(0, 16), pady=6)
        exif_row.grid_columnconfigure(0, weight=1)
        exif_entry = ctk.CTkEntry(
            exif_row,
            height=34,
            font=font,
            fg_color=COLORS["preview_bg"],
            border_color=COLORS["border"],
            text_color=COLORS["text"],
        )
        exif_entry.insert(0, cur.exiftool_path)
        exif_entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))

        def browse_exiftool() -> None:
            path = filedialog.askopenfilename(
                title="选择 ExifTool 可执行文件",
                filetypes=[("可执行文件", "*.exe"), ("所有文件", "*.*")],
            )
            if path:
                exif_entry.delete(0, "end")
                exif_entry.insert(0, path)

        ctk.CTkButton(
            exif_row,
            text="浏览…",
            width=72,
            height=34,
            font=font,
            fg_color=COLORS["accent_soft"],
            hover_color="#DBEAFE",
            text_color=COLORS["accent"],
            border_width=1,
            border_color="#BFDBFE",
            command=browse_exiftool,
        ).grid(row=0, column=1)

        ctk.CTkLabel(
            frame,
            text="自定义插件",
            font=bold,
            text_color=COLORS["text"],
            width=100,
            anchor="nw",
        ).grid(row=4, column=0, sticky="nw", padx=(16, 8), pady=6)
        plugin_box = ctk.CTkTextbox(
            frame,
            height=72,
            font=font,
            fg_color=COLORS["preview_bg"],
            border_color=COLORS["border"],
            border_width=1,
            text_color=COLORS["text"],
        )
        plugin_box.grid(row=4, column=1, sticky="ew", padx=(0, 16), pady=6)
        plugin_box.insert("1.0", cur.custom_plugin_scripts)

        plugin_btn_row = ctk.CTkFrame(frame, fg_color="transparent")
        plugin_btn_row.grid(row=5, column=1, sticky="ew", padx=(0, 16), pady=(0, 4))

        def browse_plugin() -> None:
            path = filedialog.askopenfilename(
                title="选择自定义插件脚本",
                filetypes=[("Python", "*.py"), ("所有文件", "*.*")],
            )
            if path:
                existing = plugin_box.get("1.0", "end").strip()
                plugin_box.delete("1.0", "end")
                plugin_box.insert("1.0", (existing + "\n" + path).strip())

        def show_plugins() -> None:
            from plugin_loader import format_plugin_list

            messagebox.showinfo("已安装插件", format_plugin_list())

        ctk.CTkButton(
            plugin_btn_row,
            text="浏览…",
            width=72,
            height=30,
            font=font,
            fg_color=COLORS["accent_soft"],
            hover_color="#DBEAFE",
            text_color=COLORS["accent"],
            border_width=1,
            border_color="#BFDBFE",
            command=browse_plugin,
        ).pack(side="left", padx=(0, 8))
        ctk.CTkButton(
            plugin_btn_row,
            text="查看已安装插件",
            width=120,
            height=30,
            font=font,
            fg_color=COLORS["surface"],
            hover_color="#EEF2F7",
            text_color=COLORS["text"],
            border_width=1,
            border_color=COLORS["border"],
            command=show_plugins,
        ).pack(side="left")

        ctk.CTkLabel(
            frame,
            text="样式映射每行一条，如 p[style-name='Heading 1'] => h1。\n"
            "自定义插件 .py 须定义 register_converters(markitdown, **kwargs)。\n"
            "ExifTool 留空则自动查找；CLI：doc2md-cli --list-plugins",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=12),
            text_color=COLORS["muted"],
            anchor="w",
            wraplength=480,
            justify="left",
        ).grid(row=6, column=0, columnspan=2, sticky="ew", padx=16, pady=(4, 8))

        btns = ctk.CTkFrame(frame, fg_color="transparent")
        btns.grid(row=7, column=0, columnspan=2, sticky="e", padx=16, pady=(8, 16))

        def on_save() -> None:
            from advanced_settings import settings_path

            s = AdvancedSettings(
                style_map=style_box.get("1.0", "end").strip(),
                exiftool_path=exif_entry.get().strip(),
                use_convert_local=bool(local_var.get()),
                keep_data_uris=bool(keep_uri_var.get()),
                custom_plugin_scripts=plugin_box.get("1.0", "end").strip(),
            )
            save_advanced_settings(s)
            self._advanced_status.set(advanced_status_text(s))
            messagebox.showinfo(APP_TITLE, f"高级设置已保存。\n\n{settings_path()}")
            win.destroy()

        ctk.CTkButton(
            btns,
            text="取消",
            width=88,
            height=34,
            font=font,
            fg_color=COLORS["surface"],
            hover_color="#EEF2F7",
            text_color=COLORS["text"],
            border_width=1,
            border_color=COLORS["border"],
            command=win.destroy,
        ).pack(side="left", padx=(0, 8))
        ctk.CTkButton(
            btns,
            text="保存",
            width=88,
            height=34,
            font=font,
            fg_color=COLORS["accent"],
            hover_color=COLORS["accent_hover"],
            command=on_save,
        ).pack(side="left")

    def _browse_file(self) -> None:
        path = filedialog.askopenfilename(
            title="选择要转换的文件",
            filetypes=FILE_DIALOG_TYPES,
        )
        if path:
            self._source_path.set(path)
            self._url.set("")
            warn = precheck_source(path)
            if warn:
                self._set_status(warn, warning=True)
                messagebox.showwarning(APP_TITLE, warn)
            else:
                self._set_status(f"已选择：{Path(path).name}")

    def _resolve_source(self) -> str:
        url = self._url.get().strip()
        path = self._source_path.get().strip()
        if url:
            return url
        return path

    def _start_convert(self) -> None:
        if self._busy:
            return
        source = self._resolve_source()
        if not source:
            messagebox.showwarning(APP_TITLE, "请先选择本地文件，或输入 URL。")
            return

        warn = precheck_source(source)
        if warn and not source.startswith(("http://", "https://")):
            # 伪 ZIP / 明确不支持：直接阻断，避免一直转圈
            hard_block = any(
                k in warn
                for k in (
                    "实际是 RAR",
                    "实际是 7z",
                    "不支持 RAR",
                    "不支持 7z",
                    "文件为空",
                )
            )
            if hard_block:
                self._set_status(warn.splitlines()[0][:120], error=True)
                messagebox.showerror(APP_TITLE, warn)
                return
            if not messagebox.askyesno(APP_TITLE, f"{warn}\n\n是否仍要尝试转换？"):
                return

        self._busy = True
        self.convert_btn.configure(state="disabled")
        self.save_btn.configure(state="disabled")
        self.copy_btn.configure(state="disabled")
        self.progress.configure(mode="indeterminate")
        self.progress.start()
        self._set_status("正在转换，请稍候…")

        thread = threading.Thread(target=self._convert_worker, args=(source,), daemon=True)
        thread.start()

    def _convert_worker(self, source: str) -> None:
        try:

            def on_progress(msg: str) -> None:
                self.after(0, lambda m=msg: self._set_status(m))

            text = convert_path(source, progress=on_progress)
            tip = postcheck_result(source, text)
            self.after(0, lambda: self._on_convert_ok(source, text, tip))
        except ConversionError as exc:
            self.after(0, lambda: self._on_convert_fail(str(exc)))
        except Exception as exc:  # noqa: BLE001
            self.after(0, lambda: self._on_convert_fail(f"意外错误：{exc}"))

    def _on_convert_ok(self, source: str, text: str, tip: str | None = None) -> None:
        self._busy = False
        self.progress.stop()
        self.progress.configure(mode="determinate")
        self.progress.set(1)
        self.convert_btn.configure(state="normal")
        self._markdown = text
        self._current_source = source
        display = text if text.strip() else "（结果为空）"
        if tip and not text.strip():
            display = f"{display}\n\n——\n{tip}"
        self._set_preview(display)
        has_content = bool(text.strip())
        self.save_btn.configure(state="normal" if has_content else "disabled")
        self.copy_btn.configure(state="normal" if has_content else "disabled")
        chars = len(text)
        if tip:
            self._set_status(f"完成（{chars} 字符）— {tip.splitlines()[0]}", warning=True)
            messagebox.showwarning(APP_TITLE, tip)
        else:
            self._set_status(f"转换成功 — 共 {chars} 字符，可保存或复制", success=True)

    def _on_convert_fail(self, message: str) -> None:
        self._busy = False
        self.progress.stop()
        self.progress.configure(mode="determinate")
        self.progress.set(0)
        self.convert_btn.configure(state="normal")
        self._set_status(message.splitlines()[0][:120], error=True)
        messagebox.showerror(APP_TITLE, message)

    def _save_markdown(self) -> None:
        if not self._markdown.strip():
            messagebox.showinfo(APP_TITLE, "暂无转换结果可保存。")
            return
        initial = default_output_path(self._current_source or "converted.md")
        path = filedialog.asksaveasfilename(
            title="保存 Markdown",
            defaultextension=".md",
            initialfile=initial.name,
            initialdir=str(initial.parent),
            filetypes=[("Markdown", "*.md"), ("文本文件", "*.txt"), ("所有文件", "*.*")],
        )
        if not path:
            return
        try:
            Path(path).write_text(self._markdown, encoding="utf-8")
        except OSError as exc:
            messagebox.showerror(APP_TITLE, f"保存失败：{exc}")
            return
        self._set_status(f"已保存：{path}", success=True)
        messagebox.showinfo(APP_TITLE, f"已保存到：\n{path}")

    def _copy_markdown(self) -> None:
        if not self._markdown.strip():
            messagebox.showinfo(APP_TITLE, "暂无内容可复制。")
            return
        self.clipboard_clear()
        self.clipboard_append(self._markdown)
        self.update()
        self._set_status("已复制到剪贴板", success=True)

    def _clear(self) -> None:
        if self._busy:
            return
        self._source_path.set("")
        self._url.set("")
        self._markdown = ""
        self._current_source = ""
        self.save_btn.configure(state="disabled")
        self.copy_btn.configure(state="disabled")
        self.progress.set(0)
        self._set_preview("转换结果将显示在这里。")
        self._set_status("已清空 — 请选择文件或输入 URL")

    def _set_preview(self, text: str) -> None:
        self.preview.configure(state="normal")
        self.preview.delete("1.0", "end")
        self.preview.insert("1.0", text)
        self.preview.configure(state="disabled")

    def _set_status(
        self,
        message: str,
        *,
        success: bool = False,
        error: bool = False,
        warning: bool = False,
    ) -> None:
        self._status.set(message)
        color = COLORS["muted"]
        if success:
            color = COLORS["success"]
        elif error:
            color = COLORS["danger"]
        elif warning:
            color = COLORS["warning"]
        self.status_label.configure(text_color=color)


def main() -> None:
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
