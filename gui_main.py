#!/usr/bin/env python3
"""Yu_Works 2.0 基础排版桌面界面。"""

from __future__ import annotations

import os
import queue
import sys
import tempfile
import threading
import traceback
from datetime import datetime

import customtkinter as ctk
from tkinter import PhotoImage, filedialog

from core.config_manager import (
    DEFAULT_HEADING_SETTINGS,
    get_heading_settings,
    save_heading_settings,
)
from format_conversion import convert_markdown_to_docx, reformat_docx
from platform_utils import open_path


APP_TITLE = "Yu_Works 基础排版"
LOG_FILE = os.path.join(tempfile.gettempdir(), "yu_works_gui.log")

COLORS = {
    "window": "#F6F3ED",
    "card": "#FFFDF8",
    "line": "#DED8CC",
    "text": "#2E312E",
    "muted": "#777268",
    "accent": "#5F7B67",
    "accent_hover": "#4E6957",
    "soft": "#E9EEE9",
    "warning": "#8B633D",
}

FONT_OPTIONS = ["黑体", "宋体", "楷体", "仿宋", "微软雅黑"]


def _get_asset_path(*parts: str) -> str:
    root = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(root, *parts)


def _display_number(value: float) -> str:
    value = float(value)
    return str(int(value)) if value.is_integer() else str(value)


class ConsoleLog(ctk.CTkTextbox):
    def __init__(self, master, **kwargs):
        kwargs.setdefault("font", ("Consolas", 11))
        kwargs.setdefault("height", 105)
        kwargs.setdefault("wrap", "word")
        kwargs.setdefault("fg_color", COLORS["card"])
        kwargs.setdefault("text_color", COLORS["text"])
        kwargs.setdefault("border_width", 1)
        kwargs.setdefault("border_color", COLORS["line"])
        super().__init__(master, **kwargs)
        self.configure(state="disabled")

    def write(self, text: str) -> None:
        self.configure(state="normal")
        self.insert("end", text)
        self.see("end")
        self.configure(state="disabled")


class App(ctk.CTk):
    def __init__(self):
        try:
            import ctypes

            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
                "yuworks.basic.v2"
            )
        except Exception:
            pass

        super().__init__()
        self._ui_thread = threading.current_thread()
        self._ui_queue: queue.Queue = queue.Queue()
        self.heading_widgets: dict[str, dict] = {}

        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")
        self.title(APP_TITLE)
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        window_width = min(1060, max(820, screen_width - 60))
        window_height = min(820, max(640, screen_height - 80))
        self.geometry(f"{window_width}x{window_height}")
        self.minsize(820, 640)
        self.configure(fg_color=COLORS["window"])
        self._set_icon()

        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=0)
        self.grid_columnconfigure(0, weight=1)

        self._build_main_panel()
        self._build_action_panel()
        self._setup_dnd()
        self.after(25, self._drain_ui_queue)

    def _set_icon(self) -> None:
        png_path = _get_asset_path("assets", "app_icon.png")
        if os.path.exists(png_path):
            try:
                self._app_icon_photo = PhotoImage(file=png_path)
                self.iconphoto(True, self._app_icon_photo)
            except Exception:
                pass

    def _build_main_panel(self) -> None:
        scroll = ctk.CTkScrollableFrame(
            self,
            fg_color="transparent",
            scrollbar_button_color=COLORS["line"],
            scrollbar_button_hover_color=COLORS["muted"],
        )
        scroll.grid(row=0, column=0, sticky="nsew", padx=34, pady=(28, 12))

        ctk.CTkLabel(
            scroll,
            text="Yu_Works 基础排版",
            font=ctk.CTkFont(family="Microsoft YaHei", size=25, weight="bold"),
            text_color=COLORS["text"],
        ).pack(anchor="w")
        ctk.CTkLabel(
            scroll,
            text="只保留文档基础排版与三级标题设置",
            font=("Microsoft YaHei", 13),
            text_color=COLORS["muted"],
        ).pack(anchor="w", pady=(4, 20))

        self._build_file_card(scroll)
        self._build_heading_card(scroll)
        self._build_output_card(scroll)

    def _card(self, parent):
        card = ctk.CTkFrame(
            parent,
            fg_color=COLORS["card"],
            corner_radius=12,
            border_width=1,
            border_color=COLORS["line"],
        )
        card.pack(fill="x", pady=(0, 16))
        return card

    def _build_file_card(self, parent) -> None:
        card = self._card(parent)
        ctk.CTkLabel(
            card,
            text="1. 选择待排版文档",
            font=ctk.CTkFont(family="Microsoft YaHei", size=16, weight="bold"),
            text_color=COLORS["text"],
        ).pack(anchor="w", padx=20, pady=(16, 10))

        row = ctk.CTkFrame(card, fg_color="transparent")
        row.pack(fill="x", padx=20, pady=(0, 16))
        self.file_entry = ctk.CTkEntry(
            row,
            placeholder_text="选择或拖入 .md、.txt、.docx 文件",
            height=38,
            fg_color=COLORS["window"],
            text_color=COLORS["text"],
            border_color=COLORS["line"],
        )
        self.file_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.file_entry.configure(state="readonly")
        self.file_entry.bind("<Double-Button-1>", lambda _event: self._browse_input())

        ctk.CTkButton(
            row,
            text="浏览",
            width=90,
            height=38,
            fg_color=COLORS["accent"],
            hover_color=COLORS["accent_hover"],
            command=self._browse_input,
        ).pack(side="right")

    def _build_heading_card(self, parent) -> None:
        card = self._card(parent)
        ctk.CTkLabel(
            card,
            text="2. 标题格式",
            font=ctk.CTkFont(family="Microsoft YaHei", size=16, weight="bold"),
            text_color=COLORS["text"],
        ).pack(anchor="w", padx=20, pady=(16, 2))
        ctk.CTkLabel(
            card,
            text="平时直接使用默认值即可；需要时可分别调整一级、二级、三级标题。",
            font=("Microsoft YaHei", 12),
            text_color=COLORS["muted"],
        ).pack(anchor="w", padx=20, pady=(0, 12))

        grid = ctk.CTkFrame(card, fg_color="transparent")
        grid.pack(fill="x", padx=20)
        widths = [1, 2, 1, 1, 1, 1]
        for index, weight in enumerate(widths):
            grid.grid_columnconfigure(index, weight=weight)

        headers = ["级别", "中文字体", "字号（磅）", "段前（磅）", "段后（磅）", "居中"]
        for column, text in enumerate(headers):
            ctk.CTkLabel(
                grid,
                text=text,
                font=ctk.CTkFont(family="Microsoft YaHei", size=12, weight="bold"),
                text_color=COLORS["muted"],
            ).grid(row=0, column=column, sticky="w", padx=6, pady=(0, 7))

        settings = get_heading_settings()
        level_names = {"1": "一级标题", "2": "二级标题", "3": "三级标题"}
        for row_index, level in enumerate(("1", "2", "3"), start=1):
            values = settings[level]
            ctk.CTkLabel(
                grid,
                text=level_names[level],
                font=("Microsoft YaHei", 12),
                text_color=COLORS["text"],
            ).grid(row=row_index, column=0, sticky="w", padx=6, pady=6)

            font_var = ctk.StringVar(value=values["font_cn"])
            font_menu = ctk.CTkOptionMenu(
                grid,
                values=FONT_OPTIONS,
                variable=font_var,
                fg_color=COLORS["soft"],
                button_color=COLORS["accent"],
                button_hover_color=COLORS["accent_hover"],
                text_color=COLORS["text"],
                dropdown_fg_color=COLORS["card"],
            )
            font_menu.grid(row=row_index, column=1, sticky="ew", padx=6, pady=6)

            size_entry = self._number_entry(grid, values["size_pt"])
            size_entry.grid(row=row_index, column=2, sticky="ew", padx=6, pady=6)
            before_entry = self._number_entry(grid, values["space_before_pt"])
            before_entry.grid(row=row_index, column=3, sticky="ew", padx=6, pady=6)
            after_entry = self._number_entry(grid, values["space_after_pt"])
            after_entry.grid(row=row_index, column=4, sticky="ew", padx=6, pady=6)

            center_var = ctk.BooleanVar(value=values["center"])
            center_box = ctk.CTkCheckBox(
                grid,
                text="",
                variable=center_var,
                width=24,
                fg_color=COLORS["accent"],
                hover_color=COLORS["accent_hover"],
            )
            center_box.grid(row=row_index, column=5, padx=6, pady=6)

            self.heading_widgets[level] = {
                "font": font_var,
                "size": size_entry,
                "before": before_entry,
                "after": after_entry,
                "center": center_var,
            }

        ctk.CTkLabel(
            card,
            text="一级标题固定启用“段前分页”；默认：黑体、小三15磅、居中、段前40磅、段后20磅。",
            font=("Microsoft YaHei", 11),
            text_color=COLORS["warning"],
        ).pack(anchor="w", padx=20, pady=(10, 8))

        button_row = ctk.CTkFrame(card, fg_color="transparent")
        button_row.pack(fill="x", padx=20, pady=(0, 16))
        ctk.CTkButton(
            button_row,
            text="保存标题设置",
            width=130,
            fg_color=COLORS["accent"],
            hover_color=COLORS["accent_hover"],
            command=self._save_heading_settings,
        ).pack(side="left")
        ctk.CTkButton(
            button_row,
            text="恢复默认值",
            width=110,
            fg_color=COLORS["soft"],
            hover_color=COLORS["line"],
            text_color=COLORS["text"],
            command=self._reset_heading_settings,
        ).pack(side="left", padx=10)
        self.heading_status = ctk.CTkLabel(
            button_row,
            text="",
            font=("Microsoft YaHei", 11),
            text_color=COLORS["accent"],
        )
        self.heading_status.pack(side="left", padx=5)

    def _number_entry(self, parent, value: float):
        entry = ctk.CTkEntry(
            parent,
            height=34,
            fg_color=COLORS["window"],
            text_color=COLORS["text"],
            border_color=COLORS["line"],
        )
        entry.insert(0, _display_number(value))
        return entry

    def _build_output_card(self, parent) -> None:
        card = self._card(parent)
        ctk.CTkLabel(
            card,
            text="3. 输出设置",
            font=ctk.CTkFont(family="Microsoft YaHei", size=16, weight="bold"),
            text_color=COLORS["text"],
        ).pack(anchor="w", padx=20, pady=(16, 10))

        path_row = ctk.CTkFrame(card, fg_color="transparent")
        path_row.pack(fill="x", padx=20)
        self.same_dir_var = ctk.BooleanVar(value=True)
        self.out_entry = ctk.CTkEntry(
            path_row,
            placeholder_text="自动保存在原文件旁边",
            state="disabled",
            height=36,
            fg_color=COLORS["window"],
            text_color=COLORS["text"],
            border_color=COLORS["line"],
        )
        self.out_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.out_btn = ctk.CTkButton(
            path_row,
            text="选择位置",
            width=100,
            state="disabled",
            fg_color=COLORS["soft"],
            hover_color=COLORS["line"],
            text_color=COLORS["text"],
            command=self._browse_output,
        )
        self.out_btn.pack(side="right")

        option_row = ctk.CTkFrame(card, fg_color="transparent")
        option_row.pack(fill="x", padx=20, pady=(10, 14))
        ctk.CTkCheckBox(
            option_row,
            text="输出到原文件同目录",
            variable=self.same_dir_var,
            command=self._toggle_output,
            fg_color=COLORS["accent"],
            hover_color=COLORS["accent_hover"],
            text_color=COLORS["text"],
        ).pack(side="left")
        self.open_after_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(
            option_row,
            text="完成后打开文档",
            variable=self.open_after_var,
            fg_color=COLORS["accent"],
            hover_color=COLORS["accent_hover"],
            text_color=COLORS["text"],
        ).pack(side="left", padx=24)

        ctk.CTkLabel(
            card,
            text="目录处理规则：软件不生成、不重排目录；检测到旧目录时会跳过，完成后请在 Word/WPS 中自动生成。",
            font=("Microsoft YaHei", 11),
            text_color=COLORS["warning"],
            wraplength=850,
            justify="left",
        ).pack(anchor="w", padx=20, pady=(0, 16))

    def _build_action_panel(self) -> None:
        panel = ctk.CTkFrame(self, fg_color="transparent")
        panel.grid(row=1, column=0, sticky="ew", padx=34, pady=(0, 24))
        panel.grid_columnconfigure(0, weight=1)

        self.go_btn = ctk.CTkButton(
            panel,
            text="开始基础排版",
            height=42,
            font=ctk.CTkFont(family="Microsoft YaHei", size=14, weight="bold"),
            fg_color=COLORS["accent"],
            hover_color=COLORS["accent_hover"],
            command=self._start_conversion,
        )
        self.go_btn.grid(row=0, column=0, sticky="ew")

        self.progress = ctk.CTkProgressBar(
            panel,
            height=7,
            progress_color=COLORS["accent"],
            fg_color=COLORS["line"],
        )
        self.progress.grid(row=1, column=0, sticky="ew", pady=(10, 8))
        self.progress.set(0)

        self.log = ConsoleLog(panel)
        self.log.grid(row=2, column=0, sticky="ew")
        self._log("基础排版已就绪。请选择文件；标题设置不改时直接使用默认值。")

    def _collect_heading_settings(self) -> dict:
        settings = {}
        for level, widgets in self.heading_widgets.items():
            try:
                size = float(widgets["size"].get())
                before = float(widgets["before"].get())
                after = float(widgets["after"].get())
            except ValueError as exc:
                raise ValueError(f"{level}级标题的字号、段前、段后必须填写数字") from exc
            if not 8 <= size <= 72:
                raise ValueError(f"{level}级标题字号应在 8 到 72 磅之间")
            if not 0 <= before <= 200 or not 0 <= after <= 200:
                raise ValueError(f"{level}级标题段前段后应在 0 到 200 磅之间")
            settings[level] = {
                "font_cn": widgets["font"].get(),
                "size_pt": size,
                "space_before_pt": before,
                "space_after_pt": after,
                "center": widgets["center"].get(),
            }
        return settings

    def _save_heading_settings(self, *, quiet: bool = False) -> bool:
        try:
            save_heading_settings(self._collect_heading_settings())
        except (OSError, ValueError) as exc:
            self.heading_status.configure(text=f"保存失败：{exc}", text_color="#A33A32")
            self._log(f"❌ {exc}")
            return False
        if not quiet:
            self.heading_status.configure(text="已保存", text_color=COLORS["accent"])
            self._log("标题设置已保存")
        return True

    def _reset_heading_settings(self) -> None:
        for level, values in DEFAULT_HEADING_SETTINGS.items():
            widgets = self.heading_widgets[level]
            widgets["font"].set(values["font_cn"])
            for key, value_key in (
                ("size", "size_pt"),
                ("before", "space_before_pt"),
                ("after", "space_after_pt"),
            ):
                entry = widgets[key]
                entry.delete(0, "end")
                entry.insert(0, _display_number(values[value_key]))
            widgets["center"].set(values["center"])
        save_heading_settings(DEFAULT_HEADING_SETTINGS)
        self.heading_status.configure(text="已恢复默认值", text_color=COLORS["accent"])
        self._log("标题设置已恢复默认值")

    def _setup_dnd(self) -> None:
        try:
            from tkinterdnd2 import DND_FILES

            self.drop_target_register(DND_FILES)
            self.dnd_bind("<<Drop>>", self._on_drop)
        except Exception:
            pass

    def _on_drop(self, event) -> None:
        path = event.data.strip().strip("{}")
        if os.path.isfile(path):
            self._set_input_path(path)

    def _browse_input(self) -> None:
        path = filedialog.askopenfilename(
            parent=self,
            title="选择待排版文档",
            filetypes=[("支持的文档", "*.md *.txt *.docx"), ("所有文件", "*.*")],
        )
        if path:
            self._set_input_path(path)

    def _set_input_path(self, path: str) -> None:
        self.file_entry.configure(state="normal")
        self.file_entry.delete(0, "end")
        self.file_entry.insert(0, path)
        self.file_entry.configure(state="readonly")
        self._log(f"已选择：{os.path.basename(path)}")

    def _browse_output(self) -> None:
        path = filedialog.asksaveasfilename(
            parent=self,
            title="指定输出路径",
            defaultextension=".docx",
            filetypes=[("Word 文档", "*.docx")],
        )
        if path:
            self.out_entry.delete(0, "end")
            self.out_entry.insert(0, path)

    def _toggle_output(self) -> None:
        if self.same_dir_var.get():
            self.out_entry.configure(state="normal")
            self.out_entry.delete(0, "end")
            self.out_entry.configure(state="disabled", placeholder_text="自动保存在原文件旁边")
            self.out_btn.configure(state="disabled")
        else:
            self.out_entry.configure(state="normal", placeholder_text="请选择输出路径")
            self.out_btn.configure(state="normal")

    def _start_conversion(self) -> None:
        if not self._save_heading_settings(quiet=True):
            return
        input_path = self.file_entry.get().strip()
        if not input_path or not os.path.exists(input_path):
            self._log("❌ 请先选择有效文件")
            return
        ext = os.path.splitext(input_path)[1].lower()
        if ext not in (".md", ".txt", ".docx"):
            self._log("❌ 仅支持 .md、.txt、.docx")
            return

        if self.same_dir_var.get():
            output_path = os.path.splitext(input_path)[0] + "_已排版.docx"
        else:
            output_path = self.out_entry.get().strip()
            if not output_path:
                self._log("❌ 请选择输出路径")
                return
            if not output_path.lower().endswith(".docx"):
                output_path += ".docx"

        output_path = self._resolve_output_path(input_path, output_path)
        self.go_btn.configure(state="disabled", text="正在排版…")
        self.progress.set(0.2)
        threading.Thread(
            target=self._run_conversion,
            args=(input_path, output_path, ext, bool(self.open_after_var.get())),
            daemon=True,
        ).start()

    def _resolve_output_path(self, input_path: str, output_path: str) -> str:
        input_path = os.path.abspath(input_path)
        output_path = os.path.abspath(output_path)
        if os.path.normcase(input_path) == os.path.normcase(output_path):
            root, _ext = os.path.splitext(output_path)
            output_path = root + "_已排版.docx"
        os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
        if not os.path.exists(output_path):
            return output_path
        try:
            with open(output_path, "a+b"):
                return output_path
        except OSError:
            root, ext = os.path.splitext(output_path)
            return f"{root}_{datetime.now().strftime('%Y%m%d_%H%M%S')}{ext}"

    def _run_conversion(
        self, input_path: str, output_path: str, ext: str, open_after: bool
    ) -> None:
        try:
            self._log(f"开始处理：{os.path.basename(input_path)}")
            if ext == ".docx":
                reformat_docx(input_path, output_path)
            else:
                convert_markdown_to_docx(input_path, output_path)
            self._post_ui(lambda: self.progress.set(0.9))
            self._log(f"✅ 排版完成：{output_path}")
            if open_after:
                self._post_ui(lambda: open_path(output_path))
        except Exception:
            self._log("❌ 排版失败：")
            for line in traceback.format_exc().splitlines()[-8:]:
                self._log(f"   {line}")
            self._log(f"详细日志：{LOG_FILE}")
        finally:
            self._post_ui(self._conversion_done)

    def _conversion_done(self) -> None:
        self.progress.set(1.0)
        self.go_btn.configure(state="normal", text="开始基础排版")
        self.after(2500, lambda: self.progress.set(0))

    def _append_log_line(self, line: str) -> None:
        try:
            self.log.write(line + "\n")
        except Exception:
            pass

    def _post_ui(self, callback) -> None:
        self._ui_queue.put(callback)

    def _drain_ui_queue(self) -> None:
        try:
            for _ in range(100):
                try:
                    callback = self._ui_queue.get_nowait()
                except queue.Empty:
                    break
                try:
                    callback()
                except Exception:
                    with open(LOG_FILE, "a", encoding="utf-8") as stream:
                        stream.write(traceback.format_exc() + "\n")
        finally:
            if self.winfo_exists():
                self.after(25, self._drain_ui_queue)

    def _log(self, message: str) -> None:
        line = f"[{datetime.now().strftime('%H:%M:%S')}] {message}"
        try:
            with open(LOG_FILE, "a", encoding="utf-8") as stream:
                stream.write(line + "\n")
        except OSError:
            pass
        if threading.current_thread() is self._ui_thread:
            self._append_log_line(line)
        else:
            self._post_ui(lambda: self._append_log_line(line))


if __name__ == "__main__":
    app = App()
    app.mainloop()
