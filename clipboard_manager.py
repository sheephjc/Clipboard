# Author: HJC by codex
# Version: 2.0.0
from __future__ import annotations

import ctypes
import calendar
import contextlib
from dataclasses import dataclass
import datetime as dt
import hashlib
import html
import io
import json
import os
from pathlib import Path
import queue
import re
import shutil
import sqlite3
import struct
import subprocess
import sys
import threading
import time
import tkinter as tk
from html.parser import HTMLParser
from tkinter import messagebox
from typing import Any, Callable

from PIL import Image, ImageTk

IS_WINDOWS = sys.platform.startswith("win")
IS_MACOS = sys.platform == "darwin"
OCR_SUPPORTED_PLATFORM = IS_WINDOWS or IS_MACOS

if IS_WINDOWS:
    import winreg
    import win32api
    import win32clipboard
    import win32con
    import win32event
    import win32gui
else:
    winreg = None
    win32api = None
    win32clipboard = None
    win32con = None
    win32event = None
    win32gui = None


APP_VERSION = "2.0.0"
APP_NAME = "Clipboard"
APP_ID = "Clipboard.Desktop"
DISPLAY_NAME = "\u526a\u8d34\u677f"
APP_DIR_NAME = "Clipboard"
LEGACY_APP_NAME = "ClipboardTrayApp"
LEGACY_APP_ID = "ClipboardTrayApp.Desktop"
LEGACY_APP_DIR_NAME = "ClipboardTrayApp"
STARTUP_ARG = "--startup"
MAX_HISTORY = 200
POLL_INTERVAL_MS = 800
QUEUE_POLL_MS = 120
LIST_SUMMARY_LIMIT = 72
WINDOW_WIDTH = 1440
WINDOW_HEIGHT = 1000
WINDOW_MIN_WIDTH = 1140
WINDOW_MIN_HEIGHT = 900
ICON_VARIANT_SIZES = (16, 20, 24, 32, 40, 48, 64, 128, 256)
AUTO_DELETE_SETTING_KEY = "auto_delete_policy"
AUTO_DELETE_CHECK_INTERVAL_SECONDS = 300
OCR_BUTTON_TEXT = "识别文字"
OCR_BUTTON_BUSY_TEXT = "识别中…"
OCR_WORKING_STATUS_TEXT = "正在识别图片文字…"
SINGLE_INSTANCE_MUTEX_NAME = f"Local\\{APP_ID}.Singleton"
TRAY_WINDOW_CLASS_NAME = f"{APP_NAME}TrayWindow"
TRAY_WINDOW_NAME = f"{APP_NAME}TrayHost"
WAKE_INSTANCE_MESSAGE = win32api.RegisterWindowMessage(f"{APP_ID}.Wake") if IS_WINDOWS else 0
ERROR_ALREADY_EXISTS = 183
AUTO_DELETE_POLICIES = {
    "off": {"label": "不自动删除", "short_label": "关", "days": None},
    "day": {"label": "一天", "short_label": "天", "days": 1},
    "week": {"label": "一周", "short_label": "周", "days": 7},
    "month": {"label": "一个月", "short_label": "月", "days": 30},
}

COLORS = {
    "bg": "#0f1720",
    "panel": "#122033",
    "panel_alt": "#18283f",
    "card": "#1d314f",
    "card_active": "#245183",
    "border": "#294164",
    "accent": "#f08b54",
    "accent_soft": "#f3c9ac",
    "text": "#e9eef5",
    "text_dim": "#8fa6bf",
    "text_text": "#dff3ea",
    "text_image": "#f5d8c5",
    "text_other": "#d6d1f6",
    "success": "#4bc67a",
    "danger": "#f06767",
}

FONT_UI = ("Microsoft YaHei UI", 11)
FONT_TITLE = ("Microsoft YaHei UI", 18, "bold")
FONT_SMALL = ("Microsoft YaHei UI", 10)
FONT_TAG = ("Microsoft YaHei UI", 10, "bold")
FONT_TEXT = ("Microsoft YaHei UI", 12)
FONT_BUTTON = ("Microsoft YaHei UI", 11, "bold")

TYPE_LABELS = {
    "text": "文本",
    "image": "图片",
    "other": "其他",
}

TYPE_COLORS = {
    "text": COLORS["text_text"],
    "image": COLORS["text_image"],
    "other": COLORS["text_other"],
}

if IS_WINDOWS:
    STANDARD_FORMAT_NAMES = {
        win32con.CF_TEXT: "CF_TEXT",
        win32con.CF_BITMAP: "CF_BITMAP",
        win32con.CF_METAFILEPICT: "CF_METAFILEPICT",
        win32con.CF_SYLK: "CF_SYLK",
        win32con.CF_DIF: "CF_DIF",
        win32con.CF_TIFF: "CF_TIFF",
        win32con.CF_OEMTEXT: "CF_OEMTEXT",
        win32con.CF_DIB: "CF_DIB",
        win32con.CF_PALETTE: "CF_PALETTE",
        win32con.CF_PENDATA: "CF_PENDATA",
        win32con.CF_RIFF: "CF_RIFF",
        win32con.CF_WAVE: "CF_WAVE",
        win32con.CF_UNICODETEXT: "CF_UNICODETEXT",
        win32con.CF_ENHMETAFILE: "CF_ENHMETAFILE",
        win32con.CF_HDROP: "CF_HDROP",
        win32con.CF_LOCALE: "CF_LOCALE",
        win32con.CF_DIBV5: "CF_DIBV5",
    }
    HTML_CLIPBOARD_FORMAT = win32clipboard.RegisterClipboardFormat("HTML Format")
    RTF_CLIPBOARD_FORMAT = win32clipboard.RegisterClipboardFormat("Rich Text Format")
    PLAIN_TEXT_CLIPBOARD_FORMAT_NAME = "CF_UNICODETEXT"
    STANDARD_FORMAT_NAMES[HTML_CLIPBOARD_FORMAT] = "HTML Format"
    STANDARD_FORMAT_NAMES[RTF_CLIPBOARD_FORMAT] = "Rich Text Format"
else:
    STANDARD_FORMAT_NAMES = {}
    HTML_CLIPBOARD_FORMAT = "public.html"
    RTF_CLIPBOARD_FORMAT = "public.rtf"
    PLAIN_TEXT_CLIPBOARD_FORMAT_NAME = "public.utf8-plain-text"
HTML_FRAGMENT_START = "<!--StartFragment-->"
HTML_FRAGMENT_END = "<!--EndFragment-->"
FORMULA_COMPLEX_MARKERS = (
    r"\frac",
    r"\begin",
    r"\end",
    r"\sum",
    r"\int",
    r"\matrix",
    r"\sqrt",
    r"\left",
    r"\right",
)
FORMULA_INLINE_PATTERN = re.compile(
    r"(?<!\\)[A-Za-z0-9)\]]\s*[\^_]\s*(\{[^{}\r\n]{1,20}\}|[A-Za-z0-9+\-()=]{1,20})"
)


@dataclass
class ClipboardCapture:
    type: str
    content_hash: str
    summary: str
    plain_text: str | None = None
    html_content: str | None = None
    rtf_content: str | None = None
    source_formats_json: str | None = None
    has_rich_text: bool = False
    image: Image.Image | None = None
    other_kind: str | None = None
    other_payload_json: str | None = None
    is_favorite: bool = False

    @property
    def snapshot_key(self) -> tuple[str, str]:
        return (self.type, self.content_hash)

    @property
    def text_content(self) -> str | None:
        return self.plain_text


@dataclass
class ClipboardEntry:
    id: int
    type: str
    summary: str
    created_at: str
    content_hash: str
    plain_text: str | None = None
    html_content: str | None = None
    rtf_content: str | None = None
    source_formats_json: str | None = None
    has_rich_text: bool = False
    image_path: str | None = None
    other_kind: str | None = None
    other_payload_json: str | None = None
    is_favorite: bool = False

    @property
    def snapshot_key(self) -> tuple[str, str]:
        return (self.type, self.content_hash)

    @property
    def text_content(self) -> str | None:
        return self.plain_text


@dataclass
class RichTextPayload:
    plain_text: str
    html_content: str | None = None
    rtf_content: str | None = None


PREVIEW_STYLE_TO_HTML_TAG = {
    "preview_bold": "strong",
    "preview_italic": "em",
    "preview_underline": "u",
    "preview_superscript": "sup",
    "preview_subscript": "sub",
}
PREVIEW_STYLE_TAG_ORDER = tuple(PREVIEW_STYLE_TO_HTML_TAG.keys())
_RESOLVED_APP_DATA_DIR: Path | None = None


def text_source_formats_json(has_rich_text: bool) -> str:
    formats = [
        format_clipboard_name(win32con.CF_UNICODETEXT)
        if IS_WINDOWS
        else PLAIN_TEXT_CLIPBOARD_FORMAT_NAME
    ]
    if has_rich_text:
        formats.extend(["HTML Format", "Rich Text Format"])
    return json_dumps(formats)


def rich_payload_has_formatting(payload: RichTextPayload) -> bool:
    return bool(payload.html_content or payload.rtf_content)


def build_rich_payload_from_segments(segments: list[tuple[frozenset[str], str]]) -> RichTextPayload:
    plain_parts: list[str] = []
    html_parts: list[str] = []
    rtf_parts: list[str] = [r"{\rtf1\ansi\deff0 "]
    has_styles = False

    for active_tags, content in segments:
        if not content:
            continue
        plain_parts.append(content)
        normalized_tags = tuple(tag for tag in PREVIEW_STYLE_TAG_ORDER if tag in active_tags)
        if normalized_tags:
            has_styles = True

        html_segment = html_escape_preserving_newlines(content)
        for tag in reversed(normalized_tags):
            html_segment = f"<{PREVIEW_STYLE_TO_HTML_TAG[tag]}>{html_segment}</{PREVIEW_STYLE_TO_HTML_TAG[tag]}>"
        html_parts.append(html_segment)

        rtf_segment = rtf_escape(content)
        for tag in reversed(normalized_tags):
            if tag == "preview_bold":
                rtf_segment = r"{\b " + rtf_segment + "}"
            elif tag == "preview_italic":
                rtf_segment = r"{\i " + rtf_segment + "}"
            elif tag == "preview_underline":
                rtf_segment = r"{\ul " + rtf_segment + "}"
            elif tag == "preview_superscript":
                rtf_segment = r"{\super " + rtf_segment + "}"
            elif tag == "preview_subscript":
                rtf_segment = r"{\sub " + rtf_segment + "}"
        rtf_parts.append(rtf_segment)

    plain_text = "".join(plain_parts)
    if not has_styles:
        return RichTextPayload(plain_text=plain_text)

    rtf_parts.append("}")
    return RichTextPayload(
        plain_text=plain_text,
        html_content=build_clipboard_html("".join(html_parts)),
        rtf_content="".join(rtf_parts),
    )


def serialize_text_widget_rich_payload(
    text_widget: tk.Text,
    start_index: str = "1.0",
    end_index: str = "end-1c",
) -> RichTextPayload:
    start_index = text_widget.index(start_index)
    end_index = text_widget.index(end_index)
    if text_widget.compare(start_index, ">=", end_index):
        return RichTextPayload(plain_text="")

    segments: list[tuple[frozenset[str], str]] = []
    index = start_index
    while text_widget.compare(index, "<", end_index):
        next_index = text_widget.index(f"{index}+1c")
        char = text_widget.get(index, next_index)
        active_tags = frozenset(
            tag for tag in text_widget.tag_names(index) if tag in PREVIEW_STYLE_TO_HTML_TAG
        )
        if segments and segments[-1][0] == active_tags:
            segments[-1] = (active_tags, segments[-1][1] + char)
        else:
            segments.append((active_tags, char))
        index = next_index
    return build_rich_payload_from_segments(segments)


def rich_payload_matches_entry(entry: ClipboardEntry, payload: RichTextPayload) -> bool:
    return (
        (entry.plain_text or "") == payload.plain_text
        and (entry.html_content or "") == (payload.html_content or "")
        and (entry.rtf_content or "") == (payload.rtf_content or "")
    )


def entry_rich_payload(entry: ClipboardEntry) -> RichTextPayload:
    return RichTextPayload(
        plain_text=entry.plain_text or "",
        html_content=entry.html_content,
        rtf_content=entry.rtf_content,
    )


@dataclass
class TrayMenuItem:
    label: str
    callback: Callable[[], None]
    checked: bool = False
    enabled: bool = True


class AutoHideScrollbar(tk.Canvas):
    def __init__(
        self,
        master,
        *,
        bg: str,
        thumb_color: str,
        thumb_active_color: str,
        command: Callable[..., object],
        thickness: int = 10,
        hide_delay_ms: int = 900,
    ):
        super().__init__(
            master,
            width=0,
            highlightthickness=0,
            bd=0,
            relief="flat",
            bg=bg,
            cursor="hand2",
        )
        self.command = command
        self.thickness = thickness
        self.hide_delay_ms = hide_delay_ms
        self.thumb_color = thumb_color
        self.thumb_active_color = thumb_active_color
        self.first = 0.0
        self.last = 1.0
        self.visible = False
        self.needed = False
        self.hovered = False
        self.dragging = False
        self.hide_after_id: str | None = None
        self.bind("<Configure>", lambda _event: self._redraw())
        self.bind("<Button-1>", self._on_press)
        self.bind("<B1-Motion>", self._on_drag)
        self.bind("<ButtonRelease-1>", self._on_release)
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)

    def attach(self, widget: tk.Widget) -> None:
        widget.bind("<Enter>", self._on_widget_activity, add="+")
        widget.bind("<Leave>", lambda _event: self._schedule_hide(), add="+")
        for event_name in (
            "<Motion>",
            "<MouseWheel>",
            "<Button-1>",
            "<B1-Motion>",
            "<Up>",
            "<Down>",
            "<Prior>",
            "<Next>",
            "<Home>",
            "<End>",
            "<KeyPress>",
        ):
            widget.bind(event_name, self._on_widget_activity, add="+")

    def set(self, first: str | float, last: str | float) -> None:
        self.first = float(first)
        self.last = float(last)
        self.needed = (self.last - self.first) < 0.999
        if not self.needed:
            self._hide(immediate=True)
            return
        if self.visible:
            self._redraw()

    def pulse(self) -> None:
        if not self.needed:
            return
        self._show()
        self._schedule_hide()

    def _show(self) -> None:
        if self.visible:
            self._redraw()
            return
        self.visible = True
        self.configure(width=self.thickness + 6)
        self._redraw()

    def _hide(self, immediate: bool = False) -> None:
        if self.hide_after_id is not None:
            self.after_cancel(self.hide_after_id)
            self.hide_after_id = None
        if not immediate and (self.hovered or self.dragging):
            return
        self.visible = False
        self.configure(width=0)
        self.delete("all")

    def _schedule_hide(self) -> None:
        if self.hide_after_id is not None:
            self.after_cancel(self.hide_after_id)
        self.hide_after_id = self.after(self.hide_delay_ms, self._hide)

    def _thumb_geometry(self) -> tuple[float, float]:
        canvas_height = max(self.winfo_height(), 1)
        raw_top = self.first * canvas_height
        raw_bottom = self.last * canvas_height
        raw_height = max(raw_bottom - raw_top, 1.0)
        thumb_height = max(raw_height, 42.0)
        max_top = max(canvas_height - thumb_height, 0.0)
        if raw_height >= thumb_height:
            top = min(raw_top, max_top)
        else:
            center = (raw_top + raw_bottom) / 2
            top = min(max(center - thumb_height / 2, 0.0), max_top)
        bottom = min(canvas_height - 2, top + thumb_height)
        return top, bottom

    def _redraw(self) -> None:
        self.delete("all")
        if not self.visible or not self.needed:
            return
        top, bottom = self._thumb_geometry()
        color = self.thumb_active_color if (self.hovered or self.dragging) else self.thumb_color
        x0 = 3
        x1 = x0 + self.thickness
        self.create_rectangle(x0, top + 2, x1, bottom - 2, fill=color, outline="")

    def _fraction_from_y(self, y_pos: int) -> float:
        canvas_height = max(self.winfo_height(), 1)
        span_fraction = max(self.last - self.first, 0.0)
        raw_thumb_height = span_fraction * canvas_height
        thumb_height = max(raw_thumb_height, 42.0)
        track_height = max(canvas_height - thumb_height, 1.0)
        max_fraction = max(1.0 - span_fraction, 0.0)
        if max_fraction == 0.0:
            return 0.0
        top = max(0.0, min(track_height, y_pos - thumb_height / 2))
        return max(0.0, min(max_fraction, (top / track_height) * max_fraction))

    def _move_thumb(self, y_pos: int) -> None:
        if not self.needed:
            return
        self.command("moveto", self._fraction_from_y(y_pos))
        self.pulse()

    def _on_press(self, event) -> str:
        self.dragging = True
        self._move_thumb(event.y)
        return "break"

    def _on_drag(self, event) -> str:
        self._move_thumb(event.y)
        return "break"

    def _on_release(self, _event) -> str:
        self.dragging = False
        self._schedule_hide()
        return "break"

    def _on_enter(self, _event) -> None:
        self.hovered = True
        self._show()

    def _on_leave(self, _event) -> None:
        self.hovered = False
        self._schedule_hide()

    def _on_widget_activity(self, _event) -> None:
        self.pulse()


def resource_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(getattr(sys, "_MEIPASS", Path(sys.executable).resolve().parent))
    return Path(__file__).resolve().parent


def resource_path(*parts: str) -> Path:
    return resource_base_dir().joinpath(*parts)


def icon_ico_path() -> Path:
    return resource_path("assets", "clipboard_icon.ico")


def icon_png_path() -> Path:
    return resource_path("assets", "clipboard_icon.png")


def closest_icon_size(target_size: int) -> int:
    return min(ICON_VARIANT_SIZES, key=lambda size: abs(size - max(16, target_size)))


def icon_variant_path(target_size: int, extension: str) -> Path:
    return resource_path("assets", f"clipboard_icon_{closest_icon_size(target_size)}.{extension}")


def icon_variant_ico_path(target_size: int) -> Path:
    return icon_variant_path(target_size, "ico")


def icon_variant_png_path(target_size: int) -> Path:
    return icon_variant_path(target_size, "png")


def window_dpi(hwnd: int | None = None) -> int:
    if hwnd:
        with contextlib.suppress(Exception):
            dpi = int(ctypes.windll.user32.GetDpiForWindow(hwnd))
            if dpi > 0:
                return dpi
    return 96


def scaled_icon_size(base_size: int, hwnd: int | None = None) -> int:
    dpi = window_dpi(hwnd)
    return max(base_size, int(round(base_size * dpi / 96.0)))


def shell_icon_size(metric: int, base_size: int, hwnd: int | None = None) -> int:
    if not IS_WINDOWS:
        return scaled_icon_size(base_size, hwnd)

    dpi = window_dpi(hwnd)
    with contextlib.suppress(Exception):
        size = int(ctypes.windll.user32.GetSystemMetricsForDpi(metric, dpi))
        if size > 0:
            return size
    with contextlib.suppress(Exception):
        size = int(win32api.GetSystemMetrics(metric))
        if size > 0:
            return max(size, scaled_icon_size(base_size, hwnd))
    return scaled_icon_size(base_size, hwnd)


def small_shell_icon_size(hwnd: int | None = None) -> int:
    if not IS_WINDOWS:
        return scaled_icon_size(16, hwnd)
    return shell_icon_size(win32con.SM_CXSMICON, 16, hwnd)


def big_shell_icon_size(hwnd: int | None = None) -> int:
    if not IS_WINDOWS:
        return scaled_icon_size(32, hwnd)
    return shell_icon_size(win32con.SM_CXICON, 32, hwnd)


def load_image_safely(image_path: str | Path) -> Image.Image:
    with contextlib.redirect_stderr(io.StringIO()):
        with Image.open(image_path) as image:
            return image.convert("RGBA")


def is_startup_launch(argv: list[str] | None = None) -> bool:
    args = argv if argv is not None else sys.argv[1:]
    return STARTUP_ARG in args


def maybe_relaunch_with_pythonw() -> bool:
    if not IS_WINDOWS:
        return False
    if getattr(sys, "frozen", False):
        return False
    if os.environ.get("CLIPBOARD_TRAY_RELAUNCHED") == "1":
        return False

    python_executable = Path(sys.executable)
    if python_executable.name.lower() != "python.exe":
        return False

    pythonw = python_executable.with_name("pythonw.exe")
    if not pythonw.exists():
        return False

    env = os.environ.copy()
    env["CLIPBOARD_TRAY_RELAUNCHED"] = "1"
    creationflags = 0
    for flag_name in ("CREATE_NEW_PROCESS_GROUP", "DETACHED_PROCESS"):
        creationflags |= int(getattr(subprocess, flag_name, 0))

    subprocess.Popen(
        [str(pythonw), str(Path(__file__).resolve()), *sys.argv[1:]],
        env=env,
        close_fds=True,
        creationflags=creationflags,
    )
    return True


def enable_windows_ui_features() -> None:
    if not IS_WINDOWS:
        return

    dpi_context_applied = False
    with contextlib.suppress(Exception):
        ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))
        dpi_context_applied = True

    if not dpi_context_applied:
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(2)
        except Exception:
            with contextlib.suppress(Exception):
                ctypes.windll.user32.SetProcessDPIAware()

    with contextlib.suppress(Exception):
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_ID)


def app_data_root() -> Path:
    return Path(os.environ.get("APPDATA", Path.home()))


def portable_data_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent / "data"
    return Path(__file__).resolve().parent / "data"


def current_app_data_dir() -> Path:
    return app_data_root() / APP_DIR_NAME


def legacy_app_data_dir() -> Path:
    return app_data_root() / LEGACY_APP_DIR_NAME


def clipboard_data_exists(data_dir: Path) -> bool:
    image_dir = data_dir / "images"
    has_images = False
    if image_dir.exists():
        with contextlib.suppress(Exception):
            has_images = any(image_dir.iterdir())
    return (data_dir / "history.db").exists() or has_images


def migrate_clipboard_data(source_dir: Path, target_dir: Path) -> bool:
    if not source_dir.exists():
        return False
    with contextlib.suppress(Exception):
        if source_dir.resolve() == target_dir.resolve():
            return False

    try:
        shutil.copytree(str(source_dir), str(target_dir), dirs_exist_ok=True)
        return True
    except Exception as exc:
        print(
            f"[{APP_NAME}] Failed to migrate app data from {source_dir} to {target_dir}: {exc}",
            file=sys.stderr,
        )
        return False


def app_data_dir() -> Path:
    global _RESOLVED_APP_DATA_DIR
    if _RESOLVED_APP_DATA_DIR is not None:
        return _RESOLVED_APP_DATA_DIR

    preferred_dir = portable_data_dir()
    if clipboard_data_exists(preferred_dir):
        _RESOLVED_APP_DATA_DIR = preferred_dir
        return preferred_dir

    for source_dir in (current_app_data_dir(), legacy_app_data_dir()):
        if migrate_clipboard_data(source_dir, preferred_dir):
            _RESOLVED_APP_DATA_DIR = preferred_dir
            return preferred_dir

    _RESOLVED_APP_DATA_DIR = preferred_dir
    return preferred_dir


class SingleInstanceGuard:
    def __init__(self, name: str):
        self.name = name
        self.handle = None

    def acquire(self) -> bool:
        handle = win32event.CreateMutex(None, True, self.name)
        if handle is None:
            raise RuntimeError("???????????")
        self.handle = handle
        return win32api.GetLastError() != ERROR_ALREADY_EXISTS

    def release(self) -> None:
        if self.handle is None:
            return
        with contextlib.suppress(Exception):
            win32event.ReleaseMutex(self.handle)
        with contextlib.suppress(Exception):
            win32api.CloseHandle(int(self.handle))
        self.handle = None


def find_existing_instance_window() -> int | None:
    with contextlib.suppress(Exception):
        hwnd = win32gui.FindWindow(TRAY_WINDOW_CLASS_NAME, TRAY_WINDOW_NAME)
        if hwnd:
            return int(hwnd)
    return None


def signal_existing_instance(should_wake: bool = True, timeout_ms: int = 2500) -> bool:
    deadline = time.monotonic() + timeout_ms / 1000.0
    while time.monotonic() < deadline:
        hwnd = find_existing_instance_window()
        if hwnd is not None:
            if should_wake:
                with contextlib.suppress(Exception):
                    win32gui.PostMessage(hwnd, WAKE_INSTANCE_MESSAGE, 0, 0)
            return True
        time.sleep(0.05)
    return False


def json_dumps(payload: object) -> str:
    return json.dumps(payload, ensure_ascii=False, sort_keys=True)


def hash_text(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def hash_rich_text(plain_text: str, html_content: str | None = None, rtf_content: str | None = None) -> str:
    digest = hashlib.sha256()
    digest.update(plain_text.encode("utf-8"))
    digest.update(b"\0HTML\0")
    digest.update((html_content or "").encode("utf-8"))
    digest.update(b"\0RTF\0")
    digest.update((rtf_content or "").encode("latin-1", errors="replace"))
    return digest.hexdigest()


def hash_image(image: Image.Image) -> str:
    normalized = image.convert("RGBA")
    digest = hashlib.sha256()
    digest.update(f"{normalized.width}x{normalized.height}".encode("ascii"))
    digest.update(normalized.tobytes())
    return digest.hexdigest()


def hash_other(kind: str, payload: object) -> str:
    digest = hashlib.sha256()
    digest.update(kind.encode("utf-8"))
    digest.update(json_dumps(payload).encode("utf-8"))
    return digest.hexdigest()


def summarize_text(text: str) -> str:
    single_line = " ".join(text.replace("\r", "\n").split())
    if not single_line:
        return "空白文本"
    if len(single_line) > LIST_SUMMARY_LIMIT:
        return f"{single_line[:LIST_SUMMARY_LIMIT - 3]}..."
    return single_line


def summarize_files(paths: list[str]) -> str:
    names = [Path(path).name or path for path in paths]
    preview = "、".join(names[:2])
    if len(names) > 2:
        preview = f"{preview} 等"
    return f"文件/文件夹 {len(paths)} 项: {preview}"


def summarize_formats(format_names: list[str]) -> str:
    preview = "、".join(format_names[:3])
    if len(format_names) > 3:
        preview = f"{preview} 等"
    return f"其他格式: {preview}"


def parse_source_formats(source_formats_json: str | None) -> list[str]:
    if not source_formats_json:
        return []
    try:
        payload = json.loads(source_formats_json)
    except json.JSONDecodeError:
        return []
    if isinstance(payload, list):
        return [str(item) for item in payload]
    if isinstance(payload, dict):
        formats = payload.get("formats", [])
        if isinstance(formats, list):
            return [str(item) for item in formats]
    return []


def normalized_auto_delete_policy(policy_key: str | None) -> str:
    if policy_key in AUTO_DELETE_POLICIES:
        return str(policy_key)
    return "off"


def auto_delete_policy_days(policy_key: str | None) -> int | None:
    return AUTO_DELETE_POLICIES[normalized_auto_delete_policy(policy_key)]["days"]


def auto_delete_policy_label(policy_key: str | None) -> str:
    return str(AUTO_DELETE_POLICIES[normalized_auto_delete_policy(policy_key)]["label"])


def auto_delete_policy_short_label(policy_key: str | None) -> str:
    return str(AUTO_DELETE_POLICIES[normalized_auto_delete_policy(policy_key)]["short_label"])


def build_clipboard_html(fragment_html: str) -> str:
    prefix = f"<html><body>{HTML_FRAGMENT_START}"
    suffix = f"{HTML_FRAGMENT_END}</body></html>"
    html_document = prefix + fragment_html + suffix
    header_template = (
        "Version:1.0\r\n"
        "StartHTML:{start_html:010d}\r\n"
        "EndHTML:{end_html:010d}\r\n"
        "StartFragment:{start_fragment:010d}\r\n"
        "EndFragment:{end_fragment:010d}\r\n"
    )
    provisional_header = header_template.format(
        start_html=0,
        end_html=0,
        start_fragment=0,
        end_fragment=0,
    )
    start_html = len(provisional_header.encode("ascii"))
    fragment_prefix = prefix.encode("utf-8")
    fragment_bytes = fragment_html.encode("utf-8")
    suffix_bytes = suffix.encode("utf-8")
    start_fragment = start_html + len(fragment_prefix)
    end_fragment = start_fragment + len(fragment_bytes)
    end_html = start_html + len(fragment_prefix) + len(fragment_bytes) + len(suffix_bytes)
    header = header_template.format(
        start_html=start_html,
        end_html=end_html,
        start_fragment=start_fragment,
        end_fragment=end_fragment,
    )
    return header + html_document


def extract_html_fragment(raw_html: str | None) -> str | None:
    if not raw_html:
        return None
    start_marker = raw_html.find(HTML_FRAGMENT_START)
    end_marker = raw_html.find(HTML_FRAGMENT_END)
    if start_marker != -1 and end_marker != -1 and end_marker > start_marker:
        start_index = start_marker + len(HTML_FRAGMENT_START)
        return raw_html[start_index:end_marker]

    html_match = re.search(r"(?is)<html.*", raw_html)
    if html_match:
        return html_match.group(0)
    return raw_html


def plain_text_from_html(raw_html: str | None) -> str:
    fragment = extract_html_fragment(raw_html)
    if not fragment:
        return ""
    normalized = re.sub(r"(?is)<br\s*/?>", "\n", fragment)
    normalized = re.sub(r"(?is)</(p|div|li|tr|h[1-6])>", "\n", normalized)
    normalized = re.sub(r"(?is)<[^>]+>", "", normalized)
    normalized = html.unescape(normalized)
    normalized = normalized.replace("\xa0", " ")
    normalized = re.sub(r"\n{3,}", "\n\n", normalized)
    return normalized.strip()


def plain_text_from_rtf(raw_rtf: str | None) -> str:
    if not raw_rtf:
        return ""
    text = raw_rtf
    text = re.sub(r"\\par[d]? ?", "\n", text)
    text = re.sub(r"\\line ?", "\n", text)
    text = re.sub(r"\\'[0-9a-fA-F]{2}", "", text)
    text = re.sub(r"\\u-?\d+\??", "", text)
    text = re.sub(r"\\[A-Za-z]+-?\d* ?", "", text)
    text = text.replace(r"\{", "{").replace(r"\}", "}").replace(r"\\", "\\")
    text = text.replace("{", "").replace("}", "")
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def decode_clipboard_html(raw_data: bytes | str | None) -> str | None:
    if raw_data is None:
        return None
    if isinstance(raw_data, bytes):
        return raw_data.decode("utf-8", errors="replace")
    return str(raw_data)


def decode_clipboard_rtf(raw_data: bytes | str | None) -> str | None:
    if raw_data is None:
        return None
    if isinstance(raw_data, bytes):
        return raw_data.decode("latin-1", errors="replace")
    return str(raw_data)


def encode_clipboard_rtf(rtf_content: str) -> bytes:
    return rtf_content.encode("latin-1", errors="replace")


class OcrDependencyError(RuntimeError):
    pass


_rapid_ocr_engine_lock = threading.Lock()
_rapid_ocr_engine: Any | None = None
_rapid_ocr_engine_init_error: Exception | None = None


def _get_rapid_ocr_engine() -> Any:
    global _rapid_ocr_engine, _rapid_ocr_engine_init_error
    with _rapid_ocr_engine_lock:
        if _rapid_ocr_engine is not None:
            return _rapid_ocr_engine

        if _rapid_ocr_engine_init_error is not None:
            raise OcrDependencyError("OCR 依赖不可用，请先安装识别依赖。") from _rapid_ocr_engine_init_error

        try:
            from rapidocr_onnxruntime import RapidOCR
        except Exception as exc:
            _rapid_ocr_engine_init_error = exc
            raise OcrDependencyError("未检测到 rapidocr-onnxruntime / onnxruntime。") from exc

        try:
            _rapid_ocr_engine = RapidOCR()
        except Exception as exc:
            _rapid_ocr_engine_init_error = exc
            raise RuntimeError("OCR 引擎初始化失败。") from exc

        return _rapid_ocr_engine


def _extract_ocr_line_text(item: Any) -> str:
    if item is None:
        return ""

    if isinstance(item, str):
        return item.strip()

    if isinstance(item, dict):
        for key in ("text", "txt", "rec_txt", "ocr_text", "label"):
            value = item.get(key)
            if isinstance(value, str) and value.strip():
                return value.strip()
        for value in item.values():
            nested_text = _extract_ocr_line_text(value)
            if nested_text:
                return nested_text
        return ""

    if isinstance(item, (list, tuple, set)):
        sequence = list(item)
        if len(sequence) >= 2 and isinstance(sequence[1], str) and sequence[1].strip():
            return sequence[1].strip()
        if sequence and isinstance(sequence[0], str) and sequence[0].strip():
            return sequence[0].strip()
        for value in sequence:
            nested_text = _extract_ocr_line_text(value)
            if nested_text:
                return nested_text
        return ""

    return ""


def _normalize_ocr_output(result: Any) -> str:
    payload = result[0] if isinstance(result, tuple) and result else result
    if isinstance(payload, dict):
        payload = payload.get("rec_res", payload.get("result", payload))

    if isinstance(payload, str):
        return payload.strip()

    if isinstance(payload, (list, tuple, set)):
        lines: list[str] = []
        for item in payload:
            line_text = _extract_ocr_line_text(item)
            if line_text:
                lines.append(line_text)
        return "\n".join(lines).strip()

    return _extract_ocr_line_text(payload)


def recognize_image_text(image_path: Path) -> str:
    if not image_path.exists():
        raise FileNotFoundError("图片文件不存在。")

    ocr_engine = _get_rapid_ocr_engine()
    try:
        result = ocr_engine(str(image_path))
    except Exception as exc:
        raise RuntimeError("图片文字识别失败。") from exc

    return _normalize_ocr_output(result)


def normalize_formula_source_text(text: str) -> str:
    normalized = text.replace(r"\(", "").replace(r"\)", "")
    if FORMULA_INLINE_PATTERN.search(normalized):
        normalized = normalized.replace("$", "")
    return normalized


def parse_formula_script_target(text: str, start_index: int) -> tuple[str | None, int]:
    cursor = start_index
    while cursor < len(text) and text[cursor].isspace():
        cursor += 1
    if cursor >= len(text):
        return None, start_index

    if text[cursor] == "{":
        end_index = text.find("}", cursor + 1)
        if end_index == -1:
            return None, start_index
        token = text[cursor + 1:end_index].strip()
        if (
            not token
            or len(token) > 20
            or any(marker in token for marker in FORMULA_COMPLEX_MARKERS)
            or "{" in token
            or "}" in token
        ):
            return None, start_index
        return token, end_index + 1

    match = re.match(r"[A-Za-z0-9+\-()=]{1,20}", text[cursor:])
    current_char = text[cursor]
    if current_char.isdigit():
        digit_match = re.match(r"\d{1,20}", text[cursor:])
        if digit_match is None:
            return None, start_index
        token = digit_match.group(0)
    elif current_char.isalpha():
        token = current_char
    elif current_char in "+-=":
        token = current_char
    else:
        return None, start_index
    return token, cursor + len(token)


def rtf_escape(text: str) -> str:
    parts: list[str] = []
    for char in text:
        if char in {"\\", "{", "}"}:
            parts.append(f"\\{char}")
        elif char == "\n":
            parts.append(r"\line ")
        elif ord(char) > 127:
            parts.append(f"\\u{ord(char)}?")
        else:
            parts.append(char)
    return "".join(parts)


def html_escape_preserving_newlines(text: str) -> str:
    return html.escape(text).replace("\n", "<br>")


def merge_formula_segments(segments: list[tuple[str, str]]) -> list[tuple[str, str]]:
    merged: list[tuple[str, str]] = []
    for style, content in segments:
        if not content:
            continue
        if merged and merged[-1][0] == style:
            merged[-1] = (style, merged[-1][1] + content)
        else:
            merged.append((style, content))
    return merged


def build_formula_rich_payload(text: str) -> RichTextPayload | None:
    normalized = normalize_formula_source_text(text)
    segments: list[tuple[str, str]] = []
    plain_parts: list[str] = []
    has_conversion = False
    cursor = 0

    while cursor < len(normalized):
        char = normalized[cursor]
        if (
            char in {"^", "_"}
            and plain_parts
            and not plain_parts[-1].endswith((" ", "\n", "\t"))
        ):
            token, next_index = parse_formula_script_target(normalized, cursor + 1)
            if token is not None:
                style = "sup" if char == "^" else "sub"
                segments.append((style, token))
                plain_parts.append(char + token)
                has_conversion = True
                cursor = next_index
                continue

        segments.append(("normal", char))
        plain_parts.append(char)
        cursor += 1

    if not has_conversion:
        return None

    merged_segments = merge_formula_segments(segments)
    html_parts: list[str] = []
    rtf_parts: list[str] = [r"{\rtf1\ansi\deff0 "]
    for style, content in merged_segments:
        if style == "sup":
            html_parts.append(f"<sup>{html_escape_preserving_newlines(content)}</sup>")
            rtf_parts.append(r"{\super " + rtf_escape(content) + "}")
        elif style == "sub":
            html_parts.append(f"<sub>{html_escape_preserving_newlines(content)}</sub>")
            rtf_parts.append(r"{\sub " + rtf_escape(content) + "}")
        else:
            html_parts.append(html_escape_preserving_newlines(content))
            rtf_parts.append(rtf_escape(content))
    rtf_parts.append("}")

    plain_text = "".join(plain_parts)
    return RichTextPayload(
        plain_text=plain_text,
        html_content=build_clipboard_html("".join(html_parts)),
        rtf_content="".join(rtf_parts),
    )


def has_formula_candidate(text: str | None) -> bool:
    if not text:
        return False
    return build_formula_rich_payload(text) is not None


class HtmlPreviewRenderer(HTMLParser):
    STYLE_TAGS = {
        "b": "preview_bold",
        "strong": "preview_bold",
        "i": "preview_italic",
        "em": "preview_italic",
        "u": "preview_underline",
        "sub": "preview_subscript",
        "sup": "preview_superscript",
    }
    BLOCK_TAGS = {"p", "div"}

    def __init__(self, text_widget: tk.Text, insert_index: str = "end"):
        super().__init__(convert_charrefs=True)
        self.text_widget = text_widget
        self.tag_stack: list[str] = []
        self.cursor_index = self.text_widget.index(insert_index)

    def handle_starttag(self, tag: str, attrs) -> None:
        tag = tag.lower()
        mapped_tag = self.STYLE_TAGS.get(tag)
        if mapped_tag:
            self.tag_stack.append(mapped_tag)
        if tag == "br":
            self._insert("\n")
        elif tag in self.BLOCK_TAGS and self.cursor_index != "1.0":
            self._ensure_linebreak()

    def handle_endtag(self, tag: str) -> None:
        tag = tag.lower()
        mapped_tag = self.STYLE_TAGS.get(tag)
        if mapped_tag:
            for index in range(len(self.tag_stack) - 1, -1, -1):
                if self.tag_stack[index] == mapped_tag:
                    del self.tag_stack[index]
                    break
        if tag in self.BLOCK_TAGS:
            self._ensure_linebreak()

    def handle_data(self, data: str) -> None:
        if not data:
            return
        self._insert(data)

    def _insert(self, content: str) -> None:
        if not content:
            return
        tags = tuple(dict.fromkeys(self.tag_stack))
        self.text_widget.insert(self.cursor_index, content, tags)
        self.cursor_index = self.text_widget.index(f"{self.cursor_index}+{len(content)}c")

    def _ensure_linebreak(self) -> None:
        if self.cursor_index == "1.0":
            return
        current = self.text_widget.get(f"{self.cursor_index}-1c", self.cursor_index)
        if current != "\n":
            self._insert("\n")


class RtfPreviewRenderer:
    STYLE_WORD_TO_TAG = {
        "b": "preview_bold",
        "i": "preview_italic",
        "ul": "preview_underline",
    }
    DESTINATION_WORDS = {
        "colortbl",
        "datastore",
        "fonttbl",
        "generator",
        "info",
        "listoverridetable",
        "listtable",
        "pict",
        "stylesheet",
        "themedata",
        "xmlopen",
        "xmlclose",
    }

    def __init__(self, text_widget: tk.Text, insert_index: str = "end"):
        self.text_widget = text_widget
        self.cursor_index = self.text_widget.index(insert_index)
        self.tag_stack: list[set[str]] = [set()]
        self.skip_stack: list[bool] = [False]

    def render(self, rtf_content: str) -> None:
        cursor = 0
        text = rtf_content
        while cursor < len(text):
            char = text[cursor]
            if char == "{":
                self.tag_stack.append(set(self.tag_stack[-1]))
                self.skip_stack.append(self.skip_stack[-1])
                cursor += 1
                continue
            if char == "}":
                if len(self.tag_stack) > 1:
                    self.tag_stack.pop()
                    self.skip_stack.pop()
                cursor += 1
                continue
            if char == "\\":
                cursor = self._handle_control(text, cursor)
                continue
            if not self.skip_stack[-1]:
                self._insert(char)
            cursor += 1

    def _handle_control(self, text: str, cursor: int) -> int:
        cursor += 1
        if cursor >= len(text):
            return cursor

        char = text[cursor]
        if char in "\\{}":
            if not self.skip_stack[-1]:
                self._insert(char)
            return cursor + 1

        if char == "*":
            self.skip_stack[-1] = True
            return cursor + 1

        if char == "'":
            if cursor + 2 < len(text) and not self.skip_stack[-1]:
                hex_value = text[cursor + 1:cursor + 3]
                with contextlib.suppress(Exception):
                    decoded = bytes([int(hex_value, 16)]).decode("cp1252", errors="replace")
                    self._insert(decoded)
            return min(cursor + 3, len(text))

        word_start = cursor
        while cursor < len(text) and text[cursor].isalpha():
            cursor += 1
        control_word = text[word_start:cursor]

        sign = 1
        if cursor < len(text) and text[cursor] in "+-":
            sign = -1 if text[cursor] == "-" else 1
            cursor += 1

        number_start = cursor
        while cursor < len(text) and text[cursor].isdigit():
            cursor += 1
        has_number = cursor > number_start
        parameter = sign * int(text[number_start:cursor]) if has_number else None

        if cursor < len(text) and text[cursor] == " ":
            cursor += 1

        self._apply_control_word(control_word, parameter)

        if control_word == "u" and cursor < len(text) and text[cursor] == "?":
            cursor += 1
        return cursor

    def _apply_control_word(self, control_word: str, parameter: int | None) -> None:
        if not control_word:
            return
        if control_word in self.DESTINATION_WORDS:
            self.skip_stack[-1] = True
            return
        if self.skip_stack[-1]:
            return

        if control_word == "par" or control_word == "line":
            self._insert("\n")
            return
        if control_word == "tab":
            self._insert("\t")
            return
        if control_word == "plain":
            self.tag_stack[-1].clear()
            return
        if control_word == "ulnone":
            self.tag_stack[-1].discard("preview_underline")
            return
        if control_word == "nosupersub":
            self.tag_stack[-1].discard("preview_superscript")
            self.tag_stack[-1].discard("preview_subscript")
            return
        if control_word == "super":
            self.tag_stack[-1].discard("preview_subscript")
            self.tag_stack[-1].add("preview_superscript")
            return
        if control_word == "sub":
            self.tag_stack[-1].discard("preview_superscript")
            self.tag_stack[-1].add("preview_subscript")
            return
        if control_word == "u" and parameter is not None:
            codepoint = parameter if parameter >= 0 else 65536 + parameter
            with contextlib.suppress(Exception):
                self._insert(chr(codepoint))
            return

        mapped_tag = self.STYLE_WORD_TO_TAG.get(control_word)
        if mapped_tag:
            if parameter == 0:
                self.tag_stack[-1].discard(mapped_tag)
            else:
                self.tag_stack[-1].add(mapped_tag)

    def _insert(self, content: str) -> None:
        if not content:
            return
        tags = tuple(tag for tag in PREVIEW_STYLE_TAG_ORDER if tag in self.tag_stack[-1])
        self.text_widget.insert(self.cursor_index, content, tags)
        self.cursor_index = self.text_widget.index(f"{self.cursor_index}+{len(content)}c")


def format_clipboard_name(format_id: int) -> str:
    if format_id in STANDARD_FORMAT_NAMES:
        return STANDARD_FORMAT_NAMES[format_id]
    if not IS_WINDOWS:
        return str(format_id)
    try:
        return win32clipboard.GetClipboardFormatName(format_id)
    except Exception:
        return f"Format {format_id}"


@contextlib.contextmanager
def open_clipboard(retries: int = 10, delay: float = 0.05):
    last_error: Exception | None = None
    for _ in range(retries):
        try:
            win32clipboard.OpenClipboard()
            break
        except Exception as exc:
            last_error = exc
            time.sleep(delay)
    else:
        raise RuntimeError("无法访问系统剪贴板。") from last_error

    try:
        yield
    finally:
        with contextlib.suppress(Exception):
            win32clipboard.CloseClipboard()


def enum_clipboard_formats() -> list[int]:
    formats: list[int] = []
    current = 0
    while True:
        current = win32clipboard.EnumClipboardFormats(current)
        if current == 0:
            break
        formats.append(current)
    return formats


def dib_to_bmp_bytes(dib_data: bytes) -> bytes:
    if len(dib_data) < 40:
        raise ValueError("DIB 数据不完整。")

    header_size = struct.unpack("<I", dib_data[:4])[0]
    bits_per_pixel = struct.unpack("<H", dib_data[14:16])[0]
    colors_used = struct.unpack("<I", dib_data[32:36])[0] if len(dib_data) >= 36 else 0

    color_table_size = 0
    if bits_per_pixel <= 8:
        color_table_size = (colors_used or (1 << bits_per_pixel)) * 4

    pixel_offset = 14 + header_size + color_table_size
    file_size = 14 + len(dib_data)
    file_header = b"BM" + struct.pack("<IHHI", file_size, 0, 0, pixel_offset)
    return file_header + dib_data


def dib_to_image(dib_data: bytes) -> Image.Image:
    bmp_bytes = dib_to_bmp_bytes(dib_data)
    with Image.open(io.BytesIO(bmp_bytes)) as image:
        return image.convert("RGBA")


def set_clipboard_rich_text(
    plain_text: str,
    html_content: str | None = None,
    rtf_content: str | None = None,
) -> None:
    with open_clipboard():
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, plain_text)
        if html_content:
            win32clipboard.SetClipboardData(HTML_CLIPBOARD_FORMAT, html_content.encode("utf-8"))
        if rtf_content:
            win32clipboard.SetClipboardData(RTF_CLIPBOARD_FORMAT, encode_clipboard_rtf(rtf_content))


def set_clipboard_text(text: str) -> None:
    set_clipboard_rich_text(text)


def set_clipboard_image(image_path: str) -> None:
    with Image.open(image_path) as image:
        output = io.BytesIO()
        image.convert("RGB").save(output, format="BMP")
        dib_bytes = output.getvalue()[14:]

    with open_clipboard():
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32clipboard.CF_DIB, dib_bytes)


def read_clipboard_text_payload() -> RichTextPayload | None:
    with open_clipboard():
        format_ids = enum_clipboard_formats()
        if not format_ids:
            return None

        if not any(
            format_id in format_ids
            for format_id in (win32con.CF_UNICODETEXT, HTML_CLIPBOARD_FORMAT, RTF_CLIPBOARD_FORMAT)
        ):
            return None

        plain_text = ""
        if win32con.CF_UNICODETEXT in format_ids:
            text = win32clipboard.GetClipboardData(win32con.CF_UNICODETEXT)
            plain_text = (text or "").rstrip("\0")

        html_content = None
        if HTML_CLIPBOARD_FORMAT in format_ids:
            with contextlib.suppress(Exception):
                html_content = decode_clipboard_html(win32clipboard.GetClipboardData(HTML_CLIPBOARD_FORMAT))

        rtf_content = None
        if RTF_CLIPBOARD_FORMAT in format_ids:
            with contextlib.suppress(Exception):
                rtf_content = decode_clipboard_rtf(win32clipboard.GetClipboardData(RTF_CLIPBOARD_FORMAT))

        if not plain_text and html_content:
            plain_text = plain_text_from_html(html_content)
        if not plain_text and rtf_content:
            plain_text = plain_text_from_rtf(rtf_content)

        if plain_text == "" and not html_content and not rtf_content:
            return None

        return RichTextPayload(
            plain_text=plain_text,
            html_content=html_content,
            rtf_content=rtf_content,
        )


def read_clipboard_capture() -> ClipboardCapture | None:
    with open_clipboard():
        format_ids = enum_clipboard_formats()
        if not format_ids:
            return None

        if win32clipboard.CF_DIB in format_ids:
            dib_data = win32clipboard.GetClipboardData(win32clipboard.CF_DIB)
            image = dib_to_image(dib_data)
            return ClipboardCapture(
                type="image",
                content_hash=hash_image(image),
                summary=f"图片 {image.width}x{image.height}",
                image=image,
            )

        if win32clipboard.CF_HDROP in format_ids:
            paths = [str(Path(path)) for path in win32clipboard.GetClipboardData(win32clipboard.CF_HDROP)]
            if paths:
                payload = {"paths": paths}
                return ClipboardCapture(
                    type="other",
                    content_hash=hash_other("files", payload),
                    summary=summarize_files(paths),
                    other_kind="files",
                    other_payload_json=json_dumps(payload),
                )

        text_format_names = [
            format_clipboard_name(format_id)
            for format_id in format_ids
            if format_id in {win32con.CF_UNICODETEXT, HTML_CLIPBOARD_FORMAT, RTF_CLIPBOARD_FORMAT}
        ]
        if text_format_names:
            plain_text = ""
            if win32clipboard.CF_UNICODETEXT in format_ids:
                text = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
                plain_text = (text or "").rstrip("\0")

            html_content = None
            if HTML_CLIPBOARD_FORMAT in format_ids:
                with contextlib.suppress(Exception):
                    html_content = decode_clipboard_html(win32clipboard.GetClipboardData(HTML_CLIPBOARD_FORMAT))

            rtf_content = None
            if RTF_CLIPBOARD_FORMAT in format_ids:
                with contextlib.suppress(Exception):
                    rtf_content = decode_clipboard_rtf(win32clipboard.GetClipboardData(RTF_CLIPBOARD_FORMAT))

            if not plain_text and html_content:
                plain_text = plain_text_from_html(html_content)
            if not plain_text and rtf_content:
                plain_text = plain_text_from_rtf(rtf_content)

            if plain_text != "" or html_content or rtf_content:
                has_rich_text = bool(html_content or rtf_content)
                return ClipboardCapture(
                    type="text",
                    content_hash=hash_rich_text(plain_text, html_content, rtf_content),
                    summary=summarize_text(plain_text or plain_text_from_html(html_content)),
                    plain_text=plain_text,
                    html_content=html_content,
                    rtf_content=rtf_content,
                    source_formats_json=json_dumps(text_format_names),
                    has_rich_text=has_rich_text,
                )

        format_names = [format_clipboard_name(format_id) for format_id in format_ids]
        payload = {"formats": format_names}
        return ClipboardCapture(
            type="other",
            content_hash=hash_other("formats", payload),
            summary=summarize_formats(format_names),
            other_kind="formats",
            other_payload_json=json_dumps(payload),
        )

class ClipboardStore:
    def __init__(self, data_dir: Path, max_history: int = MAX_HISTORY):
        self.data_dir = data_dir
        self.image_dir = data_dir / "images"
        self.db_path = data_dir / "history.db"
        self.max_history = max_history

        self.data_dir.mkdir(parents=True, exist_ok=True)
        self.image_dir.mkdir(parents=True, exist_ok=True)
        self.conn = sqlite3.connect(self.db_path)
        self.conn.row_factory = sqlite3.Row
        self._init_db()

    def close(self) -> None:
        with contextlib.suppress(Exception):
            self.conn.close()

    def _init_db(self) -> None:
        self.conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS entries (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                type TEXT NOT NULL,
                summary TEXT NOT NULL,
                created_at TEXT NOT NULL,
                content_hash TEXT NOT NULL,
                plain_text TEXT,
                text_content TEXT,
                html_content TEXT,
                rtf_content TEXT,
                source_formats_json TEXT,
                has_rich_text INTEGER DEFAULT 0,
                image_path TEXT,
                other_kind TEXT,
                other_payload_json TEXT
            );

            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL
            );
            """
        )
        self._ensure_entries_columns()
        self.conn.commit()

    def _ensure_entries_columns(self) -> None:
        columns = {
            row["name"]
            for row in self.conn.execute("PRAGMA table_info(entries)").fetchall()
        }
        column_specs = {
            "plain_text": "TEXT",
            "html_content": "TEXT",
            "rtf_content": "TEXT",
            "source_formats_json": "TEXT",
            "has_rich_text": "INTEGER DEFAULT 0",
            "is_favorite": "INTEGER DEFAULT 0",
        }
        for column_name, column_type in column_specs.items():
            if column_name not in columns:
                self.conn.execute(f"ALTER TABLE entries ADD COLUMN {column_name} {column_type}")

        self.conn.execute(
            """
            UPDATE entries
            SET plain_text = text_content
            WHERE plain_text IS NULL AND text_content IS NOT NULL
            """
        )
        self.conn.execute(
            """
            UPDATE entries
            SET has_rich_text = CASE
                WHEN COALESCE(html_content, '') <> '' OR COALESCE(rtf_content, '') <> '' THEN 1
                ELSE COALESCE(has_rich_text, 0)
            END
            WHERE has_rich_text IS NULL OR has_rich_text = 0
            """
        )

    def _row_to_entry(self, row: sqlite3.Row) -> ClipboardEntry:
        return ClipboardEntry(
            id=row["id"],
            type=row["type"],
            summary=row["summary"],
            created_at=row["created_at"],
            content_hash=row["content_hash"],
            plain_text=row["plain_text"],
            html_content=row["html_content"],
            rtf_content=row["rtf_content"],
            source_formats_json=row["source_formats_json"],
            has_rich_text=bool(row["has_rich_text"]),
            image_path=row["image_path"],
            other_kind=row["other_kind"],
            other_payload_json=row["other_payload_json"],
            is_favorite=bool(row["is_favorite"]) if "is_favorite" in row.keys() else False,
        )

    def load_entries(self) -> list[ClipboardEntry]:
        rows = self.conn.execute(
            """
            SELECT id, type, summary, created_at, content_hash,
                   COALESCE(plain_text, text_content) AS plain_text,
                   html_content, rtf_content, source_formats_json, COALESCE(has_rich_text, 0) AS has_rich_text,
                   image_path, other_kind, other_payload_json, COALESCE(is_favorite, 0) AS is_favorite
            FROM entries
            ORDER BY id DESC
            LIMIT ?
            """,
            (self.max_history,),
        ).fetchall()
        return [self._row_to_entry(row) for row in rows]

    def get_latest_snapshot_key(self) -> tuple[str, str] | None:
        row = self.conn.execute(
            "SELECT type, content_hash FROM entries ORDER BY id DESC LIMIT 1"
        ).fetchone()
        if row is None:
            return None
        return (row["type"], row["content_hash"])

    def add_capture(self, capture: ClipboardCapture) -> ClipboardEntry:
        created_at = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        image_path = None
        if capture.type == "image" and capture.image is not None:
            image_path = self._save_image(capture.image, capture.content_hash)

        cursor = self.conn.execute(
            """
            INSERT INTO entries (
                type, summary, created_at, content_hash,
                plain_text, text_content, html_content, rtf_content, source_formats_json, has_rich_text,
                image_path, other_kind, other_payload_json, is_favorite
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                capture.type,
                capture.summary,
                created_at,
                capture.content_hash,
                capture.plain_text,
                capture.plain_text,
                capture.html_content,
                capture.rtf_content,
                capture.source_formats_json,
                1 if capture.has_rich_text else 0,
                image_path,
                capture.other_kind,
                capture.other_payload_json,
                1 if capture.is_favorite else 0,
            ),
        )
        self.conn.commit()
        self.prune_to_limit()
        return ClipboardEntry(
            id=int(cursor.lastrowid),
            type=capture.type,
            summary=capture.summary,
            created_at=created_at,
            content_hash=capture.content_hash,
            plain_text=capture.plain_text,
            html_content=capture.html_content,
            rtf_content=capture.rtf_content,
            source_formats_json=capture.source_formats_json,
            has_rich_text=capture.has_rich_text,
            image_path=image_path,
            other_kind=capture.other_kind,
            other_payload_json=capture.other_payload_json,
            is_favorite=capture.is_favorite,
        )

    def delete_entry(self, entry_id: int) -> None:
        self.conn.execute("DELETE FROM entries WHERE id = ?", (entry_id,))
        self.conn.commit()
        self.cleanup_unused_images()

    def update_entry_text(
        self,
        entry_id: int,
        plain_text: str,
        summary: str,
        content_hash: str,
        html_content: str | None = None,
        rtf_content: str | None = None,
        source_formats_json: str | None = None,
        has_rich_text: bool = False,
    ) -> None:
        self.conn.execute(
            """
            UPDATE entries
            SET plain_text = ?,
                text_content = ?,
                summary = ?,
                content_hash = ?,
                html_content = ?,
                rtf_content = ?,
                source_formats_json = ?,
                has_rich_text = ?
            WHERE id = ?
            """,
            (
                plain_text,
                plain_text,
                summary,
                content_hash,
                html_content,
                rtf_content,
                source_formats_json,
                1 if has_rich_text else 0,
                entry_id,
            ),
        )
        self.conn.commit()

    def update_favorite(self, entry_id: int, is_favorite: bool) -> None:
        self.conn.execute(
            "UPDATE entries SET is_favorite = ? WHERE id = ?",
            (1 if is_favorite else 0, entry_id)
        )
        self.conn.commit()

    def toggle_favorite(self, entry_id: int, is_favorite: bool) -> None:
        self.update_favorite(entry_id, is_favorite)

    def clear_entries(self) -> None:
        self.conn.execute("DELETE FROM entries")
        self.conn.commit()
        self.cleanup_unused_images()

    def delete_entries_older_than(self, cutoff_created_at: str) -> int:
        cursor = self.conn.execute(
            "DELETE FROM entries WHERE created_at < ? AND COALESCE(is_favorite, 0) = 0",
            (cutoff_created_at,),
        )
        deleted_count = cursor.rowcount if cursor.rowcount is not None else 0
        self.conn.commit()
        if deleted_count:
            self.cleanup_unused_images()
        return max(deleted_count, 0)

    def prune_to_limit(self) -> None:
        rows = self.conn.execute(
            "SELECT id FROM entries WHERE is_favorite = 0 ORDER BY id DESC LIMIT -1 OFFSET ?",
            (self.max_history,),
        ).fetchall()
        if not rows:
            return

        ids = [row["id"] for row in rows]
        placeholders = ",".join(["?"] * len(ids))
        self.conn.execute(f"DELETE FROM entries WHERE id IN ({placeholders})", ids)
        self.conn.commit()
        self.cleanup_unused_images()

    def cleanup_unused_images(self) -> None:
        used_paths = {
            Path(row["image_path"])
            for row in self.conn.execute(
                "SELECT image_path FROM entries WHERE image_path IS NOT NULL"
            ).fetchall()
        }
        for path in self.image_dir.glob("*.png"):
            if path not in used_paths:
                with contextlib.suppress(OSError):
                    path.unlink()

    def _save_image(self, image: Image.Image, content_hash: str) -> str:
        path = self.image_dir / f"{content_hash}.png"
        if not path.exists():
            normalized = image.convert("RGBA")
            clean_image = Image.new("RGBA", normalized.size)
            clean_image.alpha_composite(normalized)
            clean_image.save(path, format="PNG")
        return str(path)

    def get_setting(self, key: str, default: str | None = None) -> str | None:
        row = self.conn.execute("SELECT value FROM settings WHERE key = ?", (key,)).fetchone()
        if row is None:
            return default
        return row["value"]

    def set_setting(self, key: str, value: str) -> None:
        self.conn.execute(
            """
            INSERT INTO settings (key, value)
            VALUES (?, ?)
            ON CONFLICT(key) DO UPDATE SET value = excluded.value
            """,
            (key, value),
        )
        self.conn.commit()


class StartupManager:
    RUN_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"

    def __init__(self, value_name: str = APP_NAME, legacy_value_name: str = LEGACY_APP_NAME):
        self.value_name = value_name
        self.legacy_value_name = legacy_value_name
        self._migrate_legacy_value()

    def build_command(self) -> str:
        if getattr(sys, "frozen", False):
            return f'"{Path(sys.executable)}" {STARTUP_ARG}'

        python_executable = Path(sys.executable)
        if python_executable.name.lower() == "python.exe":
            pythonw = python_executable.with_name("pythonw.exe")
            if pythonw.exists():
                python_executable = pythonw

        script_path = Path(__file__).resolve()
        return f'"{python_executable}" "{script_path}" {STARTUP_ARG}'

    def _migrate_legacy_value(self) -> None:
        if self.legacy_value_name == self.value_name:
            return
        try:
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, self.RUN_KEY) as key:
                new_exists = True
                try:
                    winreg.QueryValueEx(key, self.value_name)
                except FileNotFoundError:
                    new_exists = False

                legacy_exists = True
                try:
                    winreg.QueryValueEx(key, self.legacy_value_name)
                except FileNotFoundError:
                    legacy_exists = False

                if not legacy_exists:
                    return

                if not new_exists:
                    winreg.SetValueEx(key, self.value_name, 0, winreg.REG_SZ, self.build_command())

                with contextlib.suppress(FileNotFoundError):
                    winreg.DeleteValue(key, self.legacy_value_name)
        except OSError:
            return

    def is_enabled(self) -> bool:
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, self.RUN_KEY, 0, winreg.KEY_READ) as key:
                winreg.QueryValueEx(key, self.value_name)
                return True
        except FileNotFoundError:
            return False

    def set_enabled(self, enabled: bool) -> None:
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, self.RUN_KEY) as key:
            if enabled:
                winreg.SetValueEx(key, self.value_name, 0, winreg.REG_SZ, self.build_command())
                if self.legacy_value_name != self.value_name:
                    with contextlib.suppress(FileNotFoundError):
                        winreg.DeleteValue(key, self.legacy_value_name)
            else:
                with contextlib.suppress(FileNotFoundError):
                    winreg.DeleteValue(key, self.value_name)
                if self.legacy_value_name != self.value_name:
                    with contextlib.suppress(FileNotFoundError):
                        winreg.DeleteValue(key, self.legacy_value_name)



class TrayIcon:
    WM_TRAYICON = (win32con.WM_USER if IS_WINDOWS else 0) + 20
    COMMAND_BASE = 2000

    def __init__(
        self,
        tooltip: str,
        menu_factory: Callable[[], list[TrayMenuItem | None]],
        default_action: Callable[[], None],
        wake_action: Callable[[], None] | None = None,
        icon_path: str | None = None,
    ):
        self.tooltip = tooltip
        self.menu_factory = menu_factory
        self.default_action = default_action
        self.wake_action = wake_action or default_action
        self.icon_path = icon_path

        self._ready = threading.Event()
        self._startup_error: Exception | None = None
        self._thread: threading.Thread | None = None
        self._commands: dict[int, Callable[[], None]] = {}
        self._hwnd: int | None = None
        self._class_name = TRAY_WINDOW_CLASS_NAME
        self._window_name = TRAY_WINDOW_NAME
        self._icon_handle = None

    def start(self) -> None:
        if self._thread and self._thread.is_alive():
            return
        self._thread = threading.Thread(target=self._run, name="TrayIconThread", daemon=True)
        self._thread.start()
        self._ready.wait(timeout=5)
        if self._startup_error is not None:
            raise RuntimeError("\u6258\u76d8\u521d\u59cb\u5316\u5931\u8d25\u3002") from self._startup_error

    def stop(self) -> None:
        if self._hwnd is not None:
            with contextlib.suppress(Exception):
                win32gui.PostMessage(self._hwnd, win32con.WM_CLOSE, 0, 0)
        if self._thread and self._thread.is_alive():
            self._thread.join(timeout=5)

    def _run(self) -> None:
        try:
            self._icon_handle = self._load_icon_handle()
            message_map = {
                win32con.WM_COMMAND: self._on_command,
                win32con.WM_DESTROY: self._on_destroy,
                win32con.WM_CLOSE: self._on_close,
                self.WM_TRAYICON: self._on_tray_notify,
                WAKE_INSTANCE_MESSAGE: self._on_wake_existing_instance,
            }

            window_class = win32gui.WNDCLASS()
            window_class.hInstance = win32api.GetModuleHandle(None)
            window_class.lpszClassName = self._class_name
            window_class.lpfnWndProc = message_map
            window_class.hCursor = win32gui.LoadCursor(0, win32con.IDC_ARROW)
            if self._icon_handle:
                window_class.hIcon = self._icon_handle
            class_atom = win32gui.RegisterClass(window_class)

            self._hwnd = win32gui.CreateWindow(
                class_atom,
                self._window_name,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                window_class.hInstance,
                None,
            )
            self._add_icon()
        except Exception as exc:
            self._startup_error = exc
        finally:
            self._ready.set()

        if self._startup_error is None:
            win32gui.PumpMessages()

    def _load_icon_handle(self):
        icon_source = None
        variant_source = icon_variant_ico_path(small_shell_icon_size())
        if variant_source.exists():
            icon_source = variant_source
        elif self.icon_path and Path(self.icon_path).exists():
            icon_source = Path(self.icon_path)

        if icon_source is not None:
            with contextlib.suppress(Exception):
                target_size = closest_icon_size(small_shell_icon_size())
                icon_handle = win32gui.LoadImage(
                    0,
                    str(icon_source),
                    win32con.IMAGE_ICON,
                    target_size,
                    target_size,
                    win32con.LR_LOADFROMFILE,
                )
                if icon_handle:
                    return icon_handle
        return win32gui.LoadIcon(0, win32con.IDI_APPLICATION)

    def _add_icon(self) -> None:
        assert self._hwnd is not None
        self._icon_handle = self._icon_handle or self._load_icon_handle()
        notify_id = (
            self._hwnd,
            0,
            win32gui.NIF_ICON | win32gui.NIF_MESSAGE | win32gui.NIF_TIP,
            self.WM_TRAYICON,
            self._icon_handle,
            self.tooltip,
        )
        win32gui.Shell_NotifyIcon(win32gui.NIM_ADD, notify_id)

    def _remove_icon(self) -> None:
        if self._hwnd is None:
            return
        notify_id = (self._hwnd, 0)
        with contextlib.suppress(Exception):
            win32gui.Shell_NotifyIcon(win32gui.NIM_DELETE, notify_id)

    def _on_close(self, hwnd, msg, wparam, lparam):
        win32gui.DestroyWindow(hwnd)
        return 0

    def _on_destroy(self, hwnd, msg, wparam, lparam):
        self._remove_icon()
        win32gui.PostQuitMessage(0)
        return 0

    def _on_command(self, hwnd, msg, wparam, lparam):
        command_id = wparam & 0xFFFF
        callback = self._commands.get(command_id)
        if callback is not None:
            callback()
        return 0

    def _on_tray_notify(self, hwnd, msg, wparam, lparam):
        if lparam == win32con.WM_LBUTTONUP:
            self.default_action()
        elif lparam == win32con.WM_RBUTTONUP:
            self._show_menu()
        return 0

    def _on_wake_existing_instance(self, hwnd, msg, wparam, lparam):
        self.wake_action()
        return 0

    def _show_menu(self) -> None:
        assert self._hwnd is not None
        menu = win32gui.CreatePopupMenu()
        self._commands.clear()

        command_id = self.COMMAND_BASE
        for item in self.menu_factory():
            if item is None:
                win32gui.AppendMenu(menu, win32con.MF_SEPARATOR, 0, "")
                continue

            flags = win32con.MF_STRING
            if item.checked:
                flags |= win32con.MF_CHECKED
            if not item.enabled:
                flags |= win32con.MF_GRAYED

            win32gui.AppendMenu(menu, flags, command_id, item.label)
            self._commands[command_id] = item.callback
            command_id += 1

        x_pos, y_pos = win32gui.GetCursorPos()
        win32gui.SetForegroundWindow(self._hwnd)
        win32gui.TrackPopupMenu(
            menu,
            win32con.TPM_LEFTALIGN | win32con.TPM_BOTTOMALIGN | win32con.TPM_RIGHTBUTTON,
            x_pos,
            y_pos,
            0,
            self._hwnd,
            None,
        )
        win32gui.PostMessage(self._hwnd, win32con.WM_NULL, 0, 0)
        win32gui.DestroyMenu(menu)


class NullStartupManager:
    def is_enabled(self) -> bool:
        return False

    def set_enabled(self, enabled: bool) -> None:
        if enabled:
            raise RuntimeError("This platform does not support startup registration yet.")


class NullSingleInstanceGuard:
    def acquire(self) -> bool:
        return True

    def release(self) -> None:
        return


class PlatformServices:
    platform_name = "generic"

    def app_data_dir(self) -> Path:
        return app_data_dir()

    def create_startup_manager(self):
        return NullStartupManager()

    def create_single_instance_guard(self):
        return NullSingleInstanceGuard()

    def signal_existing_instance(self, should_wake: bool = True) -> bool:
        return False

    def maybe_relaunch_with_pythonw(self) -> bool:
        return False

    def enable_ui_features(self) -> None:
        return

    def tray_icon_path(self) -> str | None:
        preferred_icon = icon_variant_ico_path(small_shell_icon_size())
        if preferred_icon.exists():
            return str(preferred_icon)
        if icon_ico_path().exists():
            return str(icon_ico_path())
        return None

    def create_tray_icon(
        self,
        tooltip: str,
        menu_factory: Callable[[], list[TrayMenuItem | None]],
        default_action: Callable[[], None],
        wake_action: Callable[[], None] | None = None,
        icon_path: str | None = None,
    ):
        return None

    def read_clipboard_capture(self) -> ClipboardCapture | None:
        raise NotImplementedError

    def read_clipboard_text_payload(self) -> RichTextPayload | None:
        raise NotImplementedError

    def set_clipboard_rich_text(
        self,
        plain_text: str,
        html_content: str | None = None,
        rtf_content: str | None = None,
    ) -> None:
        raise NotImplementedError

    def set_clipboard_text(self, text: str) -> None:
        self.set_clipboard_rich_text(text)

    def set_clipboard_image(self, image_path: str) -> None:
        raise NotImplementedError


class WindowsPlatformServices(PlatformServices):
    platform_name = "windows"

    def create_startup_manager(self):
        return StartupManager()

    def create_single_instance_guard(self):
        return SingleInstanceGuard(SINGLE_INSTANCE_MUTEX_NAME)

    def signal_existing_instance(self, should_wake: bool = True) -> bool:
        return signal_existing_instance(should_wake=should_wake)

    def maybe_relaunch_with_pythonw(self) -> bool:
        return maybe_relaunch_with_pythonw()

    def enable_ui_features(self) -> None:
        enable_windows_ui_features()

    def create_tray_icon(
        self,
        tooltip: str,
        menu_factory: Callable[[], list[TrayMenuItem | None]],
        default_action: Callable[[], None],
        wake_action: Callable[[], None] | None = None,
        icon_path: str | None = None,
    ):
        return TrayIcon(
            tooltip=tooltip,
            menu_factory=menu_factory,
            default_action=default_action,
            wake_action=wake_action,
            icon_path=icon_path,
        )

    def read_clipboard_capture(self) -> ClipboardCapture | None:
        return read_clipboard_capture()

    def read_clipboard_text_payload(self) -> RichTextPayload | None:
        return read_clipboard_text_payload()

    def set_clipboard_rich_text(
        self,
        plain_text: str,
        html_content: str | None = None,
        rtf_content: str | None = None,
    ) -> None:
        set_clipboard_rich_text(plain_text, html_content, rtf_content)

    def set_clipboard_text(self, text: str) -> None:
        set_clipboard_text(text)

    def set_clipboard_image(self, image_path: str) -> None:
        set_clipboard_image(image_path)


def create_platform_services() -> PlatformServices:
    if IS_MACOS:
        try:
            from platforms.macos.services import MacPlatformServices

            return MacPlatformServices(
                project_root=Path(__file__).resolve().parent,
                startup_arg=STARTUP_ARG,
            )
        except Exception as exc:
            print(f"[{APP_NAME}] Failed to initialize macOS services: {exc}", file=sys.stderr)
    return WindowsPlatformServices() if IS_WINDOWS else PlatformServices()


class ClipboardManagerApp(tk.Tk):
    def __init__(
        self,
        launched_from_startup: bool = False,
        platform_services: PlatformServices | None = None,
    ):
        super().__init__()
        self.platform_services = platform_services or create_platform_services()
        self.launched_from_startup = launched_from_startup
        self.icon_ico = icon_ico_path()
        self.icon_png = icon_png_path()
        self.window_has_been_presented = False
        self.is_pinned = False

        self.withdraw()
        self.title(DISPLAY_NAME)
        self.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}")
        self.minsize(WINDOW_MIN_WIDTH, WINDOW_MIN_HEIGHT)
        self.configure(bg=COLORS["bg"])
        self.option_add("*Font", FONT_UI)
        self.window_icon_photo: ImageTk.PhotoImage | None = None
        self.window_icon_photos: list[ImageTk.PhotoImage] = []
        self.window_shell_icon_handles: list[object] = []
        self.window_shell_icon_hwnd: int | None = None
        self.window_shell_icon_dpi: int | None = None
        self.header_icon_image: ImageTk.PhotoImage | None = None
        self.auto_delete_chip: tk.Frame | None = None
        self.auto_delete_chip_title: tk.Label | None = None
        self.auto_delete_chip_dot: tk.Label | None = None
        self.auto_delete_popup: tk.Toplevel | None = None
        self.date_popup: tk.Toplevel | None = None
        self.startup_chip: tk.Frame | None = None
        self.startup_chip_title: tk.Label | None = None
        self.startup_chip_dot: tk.Label | None = None
        self.help_button: tk.Label | None = None
        self.pin_button: tk.Label | None = None
        self._configure_window_appearance()

        self.store = ClipboardStore(self.platform_services.app_data_dir())
        self.startup_manager = self.platform_services.create_startup_manager()

        self.entries = self.store.load_entries()
        self.visible_entries: list[ClipboardEntry] = []
        self.selected_entry_id: int | None = None
        self.selected_date: str | None = None
        self.current_filter = "all"
        self.current_preview_mode = None
        self.preview_photo: ImageTk.PhotoImage | None = None
        self.zoom_photo: ImageTk.PhotoImage | None = None
        self.zoom_overlay: tk.Toplevel | None = None
        self.zoom_image_label: tk.Label | None = None
        self.zoom_hint_label: tk.Label | None = None
        self.zoom_source_image: Image.Image | None = None
        self.zoom_base_size: tuple[int, int] | None = None
        self.zoom_scale = 1.0
        self.preview_resize_after_id: str | None = None
        self.poll_after_id: str | None = None
        self.queue_after_id: str | None = None
        self.status_reset_after_id: str | None = None
        self.last_snapshot_key = self.store.get_latest_snapshot_key()
        self.suppressed_snapshot_key: tuple[str, str] | None = None
        self.last_auto_delete_check_monotonic = 0.0
        self.is_window_visible = False
        self.is_shutting_down = False
        self.ui_queue: queue.Queue[Callable[[], None]] = queue.Queue()

        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self._refresh_list())

        self.search_entry: tk.Entry | None = None
        self.filter_buttons: dict[str, tk.Label] = {}
        self.list_scrollbar: AutoHideScrollbar | None = None
        self.preview_text_scroll: AutoHideScrollbar | None = None
        self.formula_button: tk.Button | None = None
        self.ocr_button: tk.Button | None = None
        self.preview_header_row: tk.Frame | None = None
        self.preview_header_favorite_btn: tk.Label | None = None
        self.ocr_result_popup: tk.Toplevel | None = None
        self.ocr_running = False
        self.preview_text_is_programmatic_update = False
        self.draft_payload_by_entry_id: dict[int, RichTextPayload] = {}
        self.preview_history_by_entry_id: dict[int, list[RichTextPayload]] = {}
        self.preview_history_index_by_entry_id: dict[int, int] = {}
        self.preview_history_is_restoring = False
        self.header_icon_image = self._load_header_icon()

        self.protocol("WM_DELETE_WINDOW", self.hide_window)
        self.bind("<Unmap>", self._on_window_unmap, add="+")

        self._build_ui()
        self._ensure_default_startup()
        self._refresh_pin_button()
        self._refresh_auto_delete_chip()
        self._refresh_startup_chip()
        self._apply_auto_delete_policy(force=True)
        self._refresh_list()

        tray_icon_path = self.platform_services.tray_icon_path()
        self.tray = self.platform_services.create_tray_icon(
            tooltip=DISPLAY_NAME,
            menu_factory=self._build_tray_menu,
            default_action=lambda: self._queue_ui_action(self.toggle_window),
            wake_action=lambda: self._queue_ui_action(self._present_existing_instance),
            icon_path=tray_icon_path,
        )
        if self.tray is not None:
            try:
                self.tray.start()
            except Exception as exc:
                self.tray = None
                self.show_window(center=True)
                messagebox.showwarning(
                    "\u6258\u76d8\u521d\u59cb\u5316\u5931\u8d25",
                    f"\u6258\u76d8\u6ca1\u6709\u6210\u529f\u542f\u52a8\uff0c\u7a97\u53e3\u5c06\u76f4\u63a5\u663e\u793a\u3002\n\n{exc}",
                )
        else:
            self.show_window(center=True)

        self.queue_after_id = self.after(QUEUE_POLL_MS, self._process_ui_queue)
        self.poll_after_id = self.after(POLL_INTERVAL_MS, self._poll_clipboard)

        if not self.launched_from_startup and self.tray is not None:
            self.after(80, lambda: self.show_window(center=True))

    def _configure_window_appearance(self) -> None:
        self._apply_tk_scaling()
        self.window_icon_photos = []
        iconphoto_applied = False

        icon_sources = [
            icon_variant_png_path(16),
            icon_variant_png_path(24),
            icon_variant_png_path(32),
            icon_variant_png_path(48),
            icon_variant_png_path(64),
            icon_variant_png_path(128),
            self.icon_png,
        ]
        for icon_source in icon_sources:
            if not icon_source.exists():
                continue
            self.window_icon_photos.append(ImageTk.PhotoImage(load_image_safely(icon_source)))

        if self.window_icon_photos:
            self.window_icon_photo = self.window_icon_photos[-1]
            try:
                self.iconphoto(True, *self.window_icon_photos)
                iconphoto_applied = True
            except Exception:
                iconphoto_applied = False

        if not iconphoto_applied and self.icon_ico.exists():
            with contextlib.suppress(Exception):
                self.iconbitmap(str(self.icon_ico))

        self.after_idle(self._apply_window_shell_icons)

    def _load_window_shell_icon_handle(self, target_size: int):
        if not IS_WINDOWS:
            return None
        if not self.icon_ico.exists():
            return None
        icon_size = closest_icon_size(target_size)
        with contextlib.suppress(Exception):
            icon_handle = win32gui.LoadImage(
                0,
                str(self.icon_ico),
                win32con.IMAGE_ICON,
                icon_size,
                icon_size,
                win32con.LR_LOADFROMFILE,
            )
            if icon_handle:
                return icon_handle
        return None

    def _destroy_icon_handle(self, icon_handle) -> None:
        if IS_WINDOWS and icon_handle:
            with contextlib.suppress(Exception):
                win32gui.DestroyIcon(icon_handle)

    def _apply_window_shell_icons(self) -> None:
        if not IS_WINDOWS:
            return
        if not self.icon_ico.exists():
            return

        try:
            hwnd = int(self.winfo_id())
        except Exception:
            return
        if not hwnd:
            return

        dpi = window_dpi(hwnd)
        if (
            self.window_shell_icon_handles
            and self.window_shell_icon_hwnd == hwnd
            and self.window_shell_icon_dpi == dpi
        ):
            return

        small_icon = self._load_window_shell_icon_handle(small_shell_icon_size(hwnd))
        big_icon = self._load_window_shell_icon_handle(big_shell_icon_size(hwnd))
        if not small_icon and not big_icon:
            return

        previous_icons = self.window_shell_icon_handles
        try:
            if small_icon:
                win32gui.SendMessage(hwnd, win32con.WM_SETICON, win32con.ICON_SMALL, small_icon)
            if big_icon:
                win32gui.SendMessage(hwnd, win32con.WM_SETICON, win32con.ICON_BIG, big_icon)
        except Exception:
            self._destroy_icon_handle(small_icon)
            self._destroy_icon_handle(big_icon)
            return

        self.window_shell_icon_handles = [icon for icon in (small_icon, big_icon) if icon]
        self.window_shell_icon_hwnd = hwnd
        self.window_shell_icon_dpi = dpi
        for icon_handle in previous_icons:
            if icon_handle not in self.window_shell_icon_handles:
                self._destroy_icon_handle(icon_handle)

    def _release_window_shell_icons(self) -> None:
        if not IS_WINDOWS:
            return
        handles = getattr(self, "window_shell_icon_handles", [])
        with contextlib.suppress(Exception):
            hwnd = int(self.winfo_id())
            win32gui.SendMessage(hwnd, win32con.WM_SETICON, win32con.ICON_SMALL, 0)
            win32gui.SendMessage(hwnd, win32con.WM_SETICON, win32con.ICON_BIG, 0)
        for icon_handle in handles:
            self._destroy_icon_handle(icon_handle)
        self.window_shell_icon_handles = []
        self.window_shell_icon_hwnd = None
        self.window_shell_icon_dpi = None

    def _apply_tk_scaling(self) -> None:
        with contextlib.suppress(Exception):
            pixels_per_inch = float(self.winfo_fpixels("1i"))
            scaling = max(1.0, pixels_per_inch / 72.0)
            self.tk.call("tk", "scaling", scaling)

    def _load_header_icon(self) -> ImageTk.PhotoImage | None:
        icon_source = icon_variant_png_path(64) if icon_variant_png_path(64).exists() else self.icon_png
        if not icon_source.exists():
            return None
        icon_image = load_image_safely(icon_source)
        preview = icon_image.resize((60, 60), Image.LANCZOS)
        return ImageTk.PhotoImage(preview)

    def _build_ui(self) -> None:
        topbar = tk.Frame(self, bg=COLORS["panel"])
        topbar.pack(fill="x", side="top")

        header_row = tk.Frame(topbar, bg=COLORS["panel"])
        header_row.pack(fill="both", expand=True, padx=18, pady=(16, 14))

        brand = tk.Frame(header_row, bg=COLORS["panel"])
        brand.pack(side="left", fill="y")

        if self.header_icon_image is not None:
            tk.Label(brand, image=self.header_icon_image, bg=COLORS["panel"]).pack(side="left", padx=(0, 14))

        title_col = tk.Frame(brand, bg=COLORS["panel"])
        title_col.pack(side="left")
        tk.Label(
            title_col,
            text=DISPLAY_NAME,
            bg=COLORS["panel"],
            fg=COLORS["text"],
            font=FONT_TITLE,
        ).pack(anchor="w")
        tk.Label(
            title_col,
            text="\u6587\u672c\u53ef\u590d\u5236\uff0c\u56fe\u7247\u53ef\u56de\u5199\uff0c\u5176\u4ed6\u683c\u5f0f\u4e5f\u4f1a\u7559\u6863\u3002",
            bg=COLORS["panel"],
            fg=COLORS["text_dim"],
            font=FONT_SMALL,
        ).pack(anchor="w", pady=(4, 0))

        right_col = tk.Frame(header_row, bg=COLORS["panel"])
        right_col.pack(side="right", anchor="n")

        header_controls = tk.Frame(right_col, bg=COLORS["panel"])
        header_controls.pack(anchor="e")

        tk.Label(
            right_col,
            text="made by HJC",
            bg=COLORS["panel"],
            fg=COLORS["border"],
            font=("Microsoft YaHei UI", 9, "italic"),
        ).pack(anchor="e", pady=(6, 0))

        self.auto_delete_chip = tk.Frame(
            header_controls,
            bg=COLORS["panel_alt"],
            highlightthickness=1,
            highlightbackground=COLORS["border"],
            padx=14,
            pady=9,
            cursor="hand2",
        )
        self.auto_delete_chip.pack(side="left", anchor="n", padx=(0, 10))
        self.auto_delete_chip_dot = tk.Label(
            self.auto_delete_chip,
            text="\u25cf",
            bg=COLORS["panel_alt"],
            fg=COLORS["accent"],
            font=("Segoe UI Symbol", 12),
            cursor="hand2",
        )
        self.auto_delete_chip_dot.pack(side="left", padx=(0, 6))
        self.auto_delete_chip_title = tk.Label(
            self.auto_delete_chip,
            text="自动删除",
            anchor="center",
            bg=COLORS["panel_alt"],
            fg=COLORS["text"],
            font=FONT_SMALL,
            cursor="hand2",
        )
        self.auto_delete_chip_title.pack(side="left")

        for widget in (
            self.auto_delete_chip,
            self.auto_delete_chip_dot,
            self.auto_delete_chip_title,
        ):
            widget.bind("<Button-1>", self._toggle_auto_delete_popup)

        self.startup_chip = tk.Frame(
            header_controls,
            bg=COLORS["panel_alt"],
            highlightthickness=1,
            highlightbackground=COLORS["border"],
            padx=14,
            pady=9,
            cursor="hand2",
        )
        self.startup_chip.pack(side="left", anchor="n", padx=(0, 10))
        self.startup_chip_dot = tk.Label(
            self.startup_chip,
            text="\u25cf",
            bg=COLORS["panel_alt"],
            fg=COLORS["accent"],
            font=("Segoe UI Symbol", 12),
            cursor="hand2",
        )
        self.startup_chip_dot.pack(side="left", padx=(0, 6))
        self.startup_chip_title = tk.Label(
            self.startup_chip,
            text="\u5f00\u673a\u81ea\u542f",
            width=8,
            anchor="center",
            bg=COLORS["panel_alt"],
            fg=COLORS["text"],
            font=FONT_SMALL,
            cursor="hand2",
        )
        self.startup_chip_title.pack(side="left")

        for widget in (
            self.startup_chip,
            self.startup_chip_dot,
            self.startup_chip_title,
        ):
            widget.bind("<Button-1>", lambda _event: self._toggle_startup())

        status_chip = tk.Frame(
            header_controls,
            bg=COLORS["card"],
            highlightthickness=1,
            highlightbackground=COLORS["border"],
            padx=14,
            pady=9,
        )
        status_chip.pack(side="left", anchor="n")
        self.status_dot = tk.Label(
            status_chip,
            text="\u25cf",
            bg=COLORS["card"],
            fg=COLORS["success"],
            font=("Segoe UI Symbol", 12),
        )
        self.status_dot.pack(side="left", padx=(0, 6))
        self.status_label = tk.Label(
            status_chip,
            text="\u76d1\u542c\u4e2d",
            bg=COLORS["card"],
            fg=COLORS["text_dim"],
            font=FONT_SMALL,
        )
        self.status_label.pack(side="left")

        self.help_button = tk.Label(
            header_controls,
            text="?",
            width=2,
            anchor="center",
            font=("Microsoft YaHei UI", 11, "bold"),
            bg=COLORS["panel"],
            fg=COLORS["text_dim"],
            cursor="hand2",
        )
        self.help_button.pack(side="left", anchor="center", padx=(10, 0), pady=(0, 0))
        self.help_button.bind("<Button-1>", lambda _event: self._show_about_dialog())
        
        self.pin_button = tk.Label(
            header_controls,
            text="\uE718",
            font=("Segoe Fluent Icons, Segoe MDL2 Assets, Segoe UI Symbol", 13),
            bg=COLORS["panel"],
            fg=COLORS["text_dim"],
            cursor="hand2",
        )
        self.pin_button.pack(side="left", anchor="center", padx=(10, 0), pady=(5, 0))
        self.pin_button.bind("<Button-1>", lambda _event: self._toggle_pinned())

        toolbar = tk.Frame(self, bg=COLORS["bg"])
        toolbar.pack(fill="x", padx=18, pady=(14, 0))

        launch_chip = tk.Frame(
            toolbar,
            bg=COLORS["panel_alt"],
            highlightthickness=1,
            highlightbackground=COLORS["border"],
            padx=12,
            pady=12,
        )
        launch_chip.pack(side="right")
        tk.Label(
            launch_chip,
            text="\u624b\u52a8\u542f\u52a8\u663e\u793a\u7a97\u53e3\uff0c\u81ea\u542f\u52a8\u4ec5\u9a7b\u7559\u6258\u76d8",
            bg=COLORS["panel_alt"],
            fg=COLORS["accent_soft"],
            font=FONT_SMALL,
        ).pack()

        search_shell = tk.Frame(
            toolbar,
            bg=COLORS["card"],
            highlightthickness=1,
            highlightbackground=COLORS["border"],
        )
        search_shell.pack(side="left", fill="x", expand=True, padx=(0, 14))
        tk.Label(
            search_shell,
            text="\u641c\u7d22",
            bg=COLORS["card"],
            fg=COLORS["text_dim"],
            font=FONT_SMALL,
        ).pack(side="left", padx=(10, 6))
        self.search_entry = tk.Entry(
            search_shell,
            textvariable=self.search_var,
            bg=COLORS["card"],
            fg=COLORS["text"],
            insertbackground=COLORS["text"],
            relief="flat",
            bd=0,
            highlightthickness=0,
            font=FONT_UI,
        )
        self.search_entry.pack(side="left", fill="x", expand=True, padx=(0, 14), pady=12)

        filterbar = tk.Frame(self, bg=COLORS["bg"])
        filterbar.pack(fill="x", padx=18, pady=(14, 0))
        for key in ("all", "text", "image", "other", "favorite"):
            btn = tk.Label(
                filterbar,
                text="",
                bg=COLORS["card"],
                fg=COLORS["text"],
                font=FONT_TAG,
                padx=16,
                pady=8,
                cursor="hand2",
                highlightthickness=1,
                highlightbackground=COLORS["border"],
            )
            btn.pack(side="left", padx=(0, 8))
            btn.bind("<Button-1>", lambda _event, target=key: self._set_filter(target))
            self.filter_buttons[key] = btn

        self.date_filter_button = tk.Label(
            filterbar,
            text="\u6309\u65e5\u671f \u25bc",
            bg=COLORS["card"],
            fg=COLORS["text"],
            font=FONT_TAG,
            padx=16,
            pady=8,
            cursor="hand2",
            highlightthickness=1,
            highlightbackground=COLORS["border"],
        )
        self.date_filter_button.pack(side="left", padx=(16, 0))
        self.date_filter_button.bind("<Button-1>", self._show_date_picker)

        body = tk.PanedWindow(
            self,
            orient="horizontal",
            bg=COLORS["bg"],
            sashwidth=8,
            showhandle=False,
        )
        body.pack(fill="both", expand=True, padx=18, pady=18)

        left = tk.Frame(body, bg=COLORS["bg"])
        body.add(left, minsize=360, width=430)

        self.count_label = tk.Label(
            left,
            text="0 \u6761\u8bb0\u5f55",
            bg=COLORS["bg"],
            fg=COLORS["text_dim"],
            font=FONT_SMALL,
        )
        self.count_label.pack(anchor="w", padx=0, pady=(0, 8))

        tk.Button(
            left,
            text="\u6e05\u7a7a\u5386\u53f2",
            bg=COLORS["card"],
            fg=COLORS["accent"],
            relief="flat",
            cursor="hand2",
            font=FONT_BUTTON,
            padx=18,
            pady=11,
            command=self._clear_history,
        ).pack(side="bottom", fill="x", pady=(12, 0))

        list_shell = tk.Frame(
            left,
            bg=COLORS["panel_alt"],
            highlightthickness=1,
            highlightbackground=COLORS["border"],
        )
        list_shell.pack(side="top", fill="both", expand=True)

        self.listbox = tk.Listbox(
            list_shell,
            bg=COLORS["panel_alt"],
            fg=COLORS["text"],
            selectbackground=COLORS["card_active"],
            selectforeground=COLORS["text"],
            activestyle="none",
            relief="flat",
            bd=0,
            highlightthickness=0,
            font=FONT_UI,
            exportselection=False,
        )
        self.listbox.pack(side="left", fill="both", expand=True, padx=(12, 0), pady=12)
        self.listbox.bind("<<ListboxSelect>>", self._on_select)

        self.list_scrollbar = AutoHideScrollbar(
            list_shell,
            bg=COLORS["panel_alt"],
            thumb_color="#4a6281",
            thumb_active_color=COLORS["accent_soft"],
            command=self.listbox.yview,
        )
        self.list_scrollbar.pack(side="right", fill="y", pady=12)
        self.list_scrollbar.attach(self.listbox)
        self.listbox.configure(yscrollcommand=self.list_scrollbar.set)

        right = tk.Frame(body, bg=COLORS["bg"])
        body.add(right, minsize=430)

        actions = tk.Frame(right, bg=COLORS["bg"])
        actions.pack(side="bottom", fill="x", pady=(12, 0))
        actions.grid_columnconfigure(0, weight=1)
        actions.grid_columnconfigure(1, weight=0)
        actions.grid_columnconfigure(2, weight=0)
        actions.grid_columnconfigure(3, weight=0)

        self.copy_button = tk.Button(
            actions,
            text="\u590d\u5236\u5230\u526a\u8d34\u677f",
            bg=COLORS["accent"],
            fg="#111111",
            relief="flat",
            cursor="hand2",
            font=FONT_BUTTON,
            padx=18,
            pady=11,
            command=self._copy_selected,
            state="disabled",
        )
        self.copy_button.grid(row=0, column=0, sticky="ew", padx=(0, 8))

        self.formula_button = tk.Button(
            actions,
            text="\u590d\u5236\u7eaf\u6587\u672c",
            bg=COLORS["card"],
            fg=COLORS["accent_soft"],
            relief="flat",
            cursor="hand2",
            font=FONT_BUTTON,
            padx=16,
            pady=11,
            width=12,
            command=self._copy_formula_selected,
            state="disabled",
        )
        self.formula_button.grid(row=0, column=1, sticky="ew", padx=(0, 8))

        self.ocr_button = tk.Button(
            actions,
            text=OCR_BUTTON_TEXT,
            bg=COLORS["card"],
            fg=COLORS["accent_soft"],
            relief="flat",
            cursor="hand2",
            font=FONT_BUTTON,
            padx=16,
            pady=11,
            width=12,
            command=self._recognize_selected_image_text,
            state="disabled",
        )
        self.ocr_button.grid(row=0, column=2, sticky="ew", padx=(0, 8))

        self.delete_button = tk.Button(
            actions,
            text="\u5220\u9664\u8bb0\u5f55",
            bg=COLORS["card"],
            fg=COLORS["danger"],
            relief="flat",
            cursor="hand2",
            font=FONT_BUTTON,
            padx=18,
            pady=11,
            width=10,
            command=self._delete_selected,
            state="disabled",
        )
        self.delete_button.grid(row=0, column=3, sticky="ew", padx=(0, 8))

        self.preview_header_row = tk.Frame(right, bg=COLORS["bg"])
        self.preview_header_row.pack(side="top", fill="x", padx=0, pady=(0, 10))

        self.preview_header = tk.Label(
            self.preview_header_row,
            text="\u4ece\u5de6\u4fa7\u9009\u62e9\u4e00\u6761\u8bb0\u5f55\u67e5\u770b\u8be6\u60c5",
            bg=COLORS["bg"],
            fg=COLORS["text_dim"],
            font=FONT_UI,
            anchor="w",
        )
        self.preview_header.pack(side="left", fill="x", expand=True)
        self.preview_header_favorite_btn = tk.Label(
            self.preview_header_row,
            text="",
            bg=COLORS["bg"],
            fg=COLORS["text"],
            cursor="hand2",
            font=("Segoe UI", 17),
            padx=0,
            pady=0,
        )
        self.preview_header_favorite_btn.pack(side="right", padx=(16, 12))
        self.preview_header_favorite_btn.bind("<Button-1>", self._toggle_favorite)

        self.preview_frame = tk.Frame(
            right,
            bg=COLORS["panel"],
            highlightthickness=1,
            highlightbackground=COLORS["border"],
        )
        self.preview_frame.pack(side="top", fill="both", expand=True)
        self.preview_frame.bind("<Configure>", self._on_preview_resize)

        self.preview_empty = tk.Label(
            self.preview_frame,
            text="\u8fd8\u6ca1\u6709\u9009\u4e2d\u8bb0\u5f55\n\n\u4ece\u5de6\u4fa7\u70b9\u4e00\u6761\u5185\u5bb9\u5f00\u59cb\u67e5\u770b\u3002",
            bg=COLORS["panel"],
            fg=COLORS["text_dim"],
            font=FONT_UI,
            justify="center",
        )

        self.preview_text_frame = tk.Frame(self.preview_frame, bg=COLORS["panel"])
        self.preview_text = tk.Text(
            self.preview_text_frame,
            wrap="word",
            bg=COLORS["panel"],
            fg=COLORS["text"],
            insertbackground=COLORS["text"],
            relief="flat",
            bd=0,
            highlightthickness=0,
            font=FONT_TEXT,
            padx=18,
            pady=18,
            spacing1=2,
            spacing3=2,
            undo=True,
            autoseparators=True,
            maxundo=-1,
        )

        self.preview_text_actions = tk.Frame(self.preview_text_frame, bg=COLORS["panel"])
        self.preview_text_actions.pack(side="bottom", fill="x")

        self.preview_text.pack(side="left", fill="both", expand=True)

        self.preview_text_scroll = AutoHideScrollbar(
            self.preview_text_frame,
            bg=COLORS["panel"],
            thumb_color="#536b8a",
            thumb_active_color=COLORS["accent_soft"],
            command=self.preview_text.yview,
        )
        self.preview_text_scroll.pack(side="right", fill="y", pady=10)
        self.preview_text_scroll.attach(self.preview_text)
        self.preview_text.configure(yscrollcommand=self.preview_text_scroll.set)

        self.preview_text.bind("<Key>", self._block_preview_input)
        self.preview_text.bind("<Control-a>", self._select_all_preview)
        self.preview_text.bind("<Control-c>", self._copy_from_preview_selection)
        self.preview_text.bind("<<Copy>>", self._copy_from_preview_selection)
        self.preview_text.bind("<Control-z>", self._undo_preview_edit)
        self.preview_text.bind("<Control-y>", self._redo_preview_edit)
        self.preview_text.bind("<Control-Shift-Z>", self._redo_preview_edit)
        self.preview_text.bind("<Control-Shift-z>", self._redo_preview_edit)
        self.preview_text.bind("<Control-equal>", self._apply_subscript_shortcut)
        self.preview_text.bind("<Control-Shift-equal>", self._apply_superscript_shortcut)
        self.preview_text.bind("<Control-plus>", self._apply_superscript_shortcut)
        self.preview_text.bind("<Control-space>", self._clear_script_shortcut)
        self.preview_text.bind("<<Modified>>", self._on_preview_text_modified)
        self.preview_text.bind("<Control-v>", self._paste_into_preview)
        self.preview_text.bind("<Control-V>", self._paste_into_preview)
        self.preview_text.bind("<Shift-Insert>", self._paste_into_preview)
        self.preview_text.bind("<<Paste>>", self._paste_into_preview)
        self.preview_text.bind("<<Cut>>", self._on_preview_edit_command, add="+")
        self.preview_text.bind("<<Clear>>", self._on_preview_edit_command, add="+")
        self.preview_text.tag_configure("preview_bold", font=("Microsoft YaHei UI", 12, "bold"))
        self.preview_text.tag_configure("preview_italic", font=("Microsoft YaHei UI", 12, "italic"))
        self.preview_text.tag_configure("preview_underline", underline=True)
        self.preview_text.tag_configure("preview_superscript", offset=5, font=("Microsoft YaHei UI", 10))
        self.preview_text.tag_configure("preview_subscript", offset=-4, font=("Microsoft YaHei UI", 10))

        spacer = tk.Frame(self.preview_text_actions, bg=COLORS["panel"])
        spacer.pack(side="left", expand=True, fill="x")
        
        self.btn_reset_text = tk.Button(
            self.preview_text_actions,
            text="重置",
            bg=COLORS["card"],
            fg=COLORS["text"],
            relief="flat",
            cursor="hand2",
            padx=10,
            pady=4,
            command=self._reset_text,
        )
        self.btn_reset_text.pack(side="left", padx=(0, 8), pady=8)
        self.btn_save_text = tk.Button(
            self.preview_text_actions,
            text="保存",
            bg=COLORS["card"],
            fg=COLORS["text"],
            relief="flat",
            cursor="hand2",
            padx=10,
            pady=4,
            command=self._save_text,
        )
        self.btn_save_text.pack(side="left", padx=(0, 18), pady=8)


        self.preview_image_frame = tk.Frame(self.preview_frame, bg=COLORS["panel"])
        
        self.preview_image_label = tk.Label(
            self.preview_image_frame,
            bg=COLORS["panel"],
            fg=COLORS["text_dim"],
            cursor="arrow",
        )
        self.preview_image_label.pack(fill="both", expand=True)
        self.preview_image_label.bind("<Button-1>", self._open_zoomed_image)

        self._set_preview_mode("empty")

    def _set_filter(self, filter_key: str) -> None:
        self.current_filter = filter_key
        self._refresh_list()

    def _queue_ui_action(self, callback: Callable[[], None]) -> None:
        self.ui_queue.put(callback)

    def _process_ui_queue(self) -> None:
        if self.is_shutting_down:
            return
        while True:
            try:
                callback = self.ui_queue.get_nowait()
            except queue.Empty:
                break
            try:
                callback()
            except Exception:
                self._set_status("操作失败", healthy=False, temporary=True)
        self.queue_after_id = self.after(QUEUE_POLL_MS, self._process_ui_queue)

    def _set_status(self, text: str, healthy: bool = True, temporary: bool = False) -> None:
        self.status_label.config(text=text, fg=COLORS["text_dim"] if healthy else COLORS["danger"])
        self.status_dot.config(fg=COLORS["success"] if healthy else COLORS["danger"])

        if self.status_reset_after_id is not None:
            self.after_cancel(self.status_reset_after_id)
            self.status_reset_after_id = None

        if temporary:
            self.status_reset_after_id = self.after(1800, lambda: self._set_status("监听中", healthy=True))

    def _show_about_dialog(self) -> None:
        message = "\n".join(
            [
                f"Version: {APP_VERSION}",
                "Author\uff1aHJC by codex",
                "\u652f\u6301\u5f00\u673a\u81ea\u542f",
                "\u652f\u6301\u81ea\u52a8\u5220\u9664",
                "\u652f\u6301\u6536\u85cf",
                "\u652f\u6301\u65e5\u671f\u7b5b\u9009",
                "\u652f\u6301\u5728\u5185\u5bb9\u533a\u7f16\u8f91\u6587\u672c",
                "\u652f\u6301\u5bcc\u6587\u672c\u548c\u7eaf\u6587\u672c\u590d\u5236",
                "\u652f\u6301\u56fe\u7247\u5185\u6587\u5b57\u8bc6\u522b",
                "\u5173\u95ed\u9875\u9762\u6216\u5f00\u673a\u81ea\u542f\u9690\u85cf\u4e8e\u53f3\u4e0b\u89d2\u4efb\u52a1\u680f\u6258\u76d8\u5904",
            ]
        )
        messagebox.showinfo("\u5173\u4e8e\u526a\u8d34\u677f", message, parent=self)

    def _apply_pinned_state(self) -> None:
        with contextlib.suppress(Exception):
            self.attributes("-topmost", self.is_pinned)

    def _refresh_pin_button(self) -> None:
        if self.pin_button is None:
            return

        icon_color = COLORS["accent"] if self.is_pinned else COLORS["text_dim"]
        self.pin_button.config(fg=icon_color)

    def _toggle_pinned(self) -> None:
        self.is_pinned = not self.is_pinned
        self._apply_pinned_state()
        self._refresh_pin_button()

    def _get_auto_delete_policy(self) -> str:
        return normalized_auto_delete_policy(
            self.store.get_setting(AUTO_DELETE_SETTING_KEY, "off")
        )

    def _refresh_auto_delete_chip(self) -> None:
        if self.auto_delete_chip is None:
            return

        policy = self._get_auto_delete_policy()
        enabled = auto_delete_policy_days(policy) is not None
        dot_fg = COLORS["success"] if enabled else "#111111"
        title_text = f"自动删除：{auto_delete_policy_short_label(policy)}"
        self.auto_delete_chip.config(bg=COLORS["panel_alt"], highlightbackground=COLORS["border"])
        self.auto_delete_chip_dot.config(bg=COLORS["panel_alt"], fg=dot_fg)
        self.auto_delete_chip_title.config(
            text=title_text,
            bg=COLORS["panel_alt"],
            fg=COLORS["text"],
        )

    def _toggle_auto_delete_popup(self, event=None) -> None:
        if self.auto_delete_popup is not None and self.auto_delete_popup.winfo_exists():
            self._close_auto_delete_popup()
            return
        self._show_auto_delete_popup(event)

    def _show_auto_delete_popup(self, event=None) -> None:
        if self.auto_delete_chip is None:
            return

        popup = tk.Toplevel(self)
        popup.overrideredirect(True)
        popup.configure(bg=COLORS["panel"], highlightthickness=1, highlightbackground=COLORS["border"])
        self.auto_delete_popup = popup

        x_pos = self.auto_delete_chip.winfo_rootx()
        y_pos = self.auto_delete_chip.winfo_rooty() + self.auto_delete_chip.winfo_height() + 4
        popup.geometry(f"+{x_pos}+{y_pos}")

        current_policy = self._get_auto_delete_policy()
        for policy_key in ("off", "day", "week", "month"):
            is_selected = policy_key == current_policy
            if policy_key == "off":
                button_padding = (8, 0)
            elif policy_key == "month":
                button_padding = (6, 8)
            else:
                button_padding = (6, 0)
            button = tk.Button(
                popup,
                text=auto_delete_policy_label(policy_key),
                bg=COLORS["accent"] if is_selected else COLORS["card"],
                fg="#111111" if is_selected else COLORS["text"],
                font=FONT_UI,
                relief="flat",
                cursor="hand2",
                anchor="w",
                padx=14,
                pady=10,
                width=12,
                command=lambda selected=policy_key: self._set_auto_delete_policy(selected),
            )
            button.pack(fill="x", padx=8, pady=button_padding)

        popup.bind("<FocusOut>", lambda _event: self._close_auto_delete_popup())
        popup.bind("<Escape>", lambda _event: self._close_auto_delete_popup())
        popup.focus_force()

    def _close_auto_delete_popup(self) -> None:
        if self.auto_delete_popup is not None and self.auto_delete_popup.winfo_exists():
            self.auto_delete_popup.destroy()
        self.auto_delete_popup = None

    def _set_auto_delete_policy(self, policy_key: str) -> None:
        normalized_policy = normalized_auto_delete_policy(policy_key)
        self.store.set_setting(AUTO_DELETE_SETTING_KEY, normalized_policy)
        self._refresh_auto_delete_chip()
        self._close_auto_delete_popup()
        deleted_count = self._apply_auto_delete_policy(force=True)
        if deleted_count:
            self._set_status(f"已自动清理 {deleted_count} 条历史", healthy=True, temporary=True)
        else:
            self._set_status(
                f"自动删除已设为{auto_delete_policy_label(normalized_policy)}",
                healthy=True,
                temporary=True,
            )

    def _apply_auto_delete_policy(self, force: bool = False) -> int:
        now_monotonic = time.monotonic()
        if not force and (now_monotonic - self.last_auto_delete_check_monotonic) < AUTO_DELETE_CHECK_INTERVAL_SECONDS:
            return 0
        self.last_auto_delete_check_monotonic = now_monotonic

        policy = self._get_auto_delete_policy()
        retention_days = auto_delete_policy_days(policy)
        if retention_days is None:
            return 0

        cutoff = (dt.datetime.now() - dt.timedelta(days=retention_days)).strftime("%Y-%m-%d %H:%M:%S")
        deleted_count = self.store.delete_entries_older_than(cutoff)
        if deleted_count:
            self.entries = self.store.load_entries()
            if self.selected_entry_id is not None and self._find_entry_by_id(self.selected_entry_id) is None:
                self.selected_entry_id = None
            self._refresh_list()
        return deleted_count

    def _refresh_startup_chip(self) -> None:
        if self.startup_chip is None:
            return

        enabled = False
        with contextlib.suppress(Exception):
            enabled = self.startup_manager.is_enabled()

        chip_bg = COLORS["panel_alt"]
        border = COLORS["border"]
        title_fg = COLORS["text"]
        dot_fg = COLORS["success"] if enabled else "#111111"

        self.startup_chip.config(bg=chip_bg, highlightbackground=border)
        self.startup_chip_dot.config(bg=chip_bg, fg=dot_fg)
        self.startup_chip_title.config(bg=chip_bg, fg=title_fg)

    def _poll_clipboard(self) -> None:
        if self.is_shutting_down:
            return

        try:
            capture = self.platform_services.read_clipboard_capture()
            snapshot_key = capture.snapshot_key if capture is not None else None

            if snapshot_key != self.last_snapshot_key:
                suppressed = self.suppressed_snapshot_key is not None and snapshot_key == self.suppressed_snapshot_key
                self.suppressed_snapshot_key = None
                self.last_snapshot_key = snapshot_key

                if capture is not None and not suppressed:
                    self.store.add_capture(capture)
                    self.entries = self.store.load_entries()
                    self._refresh_list()

            self._apply_auto_delete_policy()
            self._set_status("监听中", healthy=True)
        except Exception:
            self._set_status("监听异常", healthy=False)
        finally:
            self.poll_after_id = self.after(POLL_INTERVAL_MS, self._poll_clipboard)

    def _filtered_entries(self) -> list[ClipboardEntry]:
        query = self.search_var.get().strip().lower()
        result: list[ClipboardEntry] = []
        for entry in self.entries:
            if self.current_filter == "favorite":
                if not entry.is_favorite:
                    continue
            elif self.current_filter != "all" and entry.type != self.current_filter:
                continue

            if self.selected_date is not None and not entry.created_at.startswith(self.selected_date):
                continue

            haystack = [entry.summary.lower()]
            if entry.plain_text:
                haystack.append(entry.plain_text.lower())
            if entry.other_payload_json:
                haystack.append(entry.other_payload_json.lower())
            if entry.source_formats_json:
                haystack.append(entry.source_formats_json.lower())

            if query and not any(query in value for value in haystack):
                continue
            result.append(entry)
        return result

    def _refresh_filter_buttons(self) -> None:
        counts = {
            "all": len(self.entries),
            "text": sum(1 for entry in self.entries if entry.type == "text"),
            "image": sum(1 for entry in self.entries if entry.type == "image"),
            "other": sum(1 for entry in self.entries if entry.type == "other"),
            "favorite": sum(1 for entry in self.entries if entry.is_favorite),
        }
        labels = {
            "all": f"全部 {counts['all']}",
            "text": f"文本 {counts['text']}",
            "image": f"图片 {counts['image']}",
            "other": f"其他 {counts['other']}",
            "favorite": f"收藏 {counts['favorite']}",
        }

        for key, button in self.filter_buttons.items():
            button.config(
                text=labels[key],
                bg=COLORS["accent"] if key == self.current_filter else COLORS["card"],
                fg="#111111" if key == self.current_filter else TYPE_COLORS.get(key, COLORS["text"]),
            )

        if self.selected_date:
            self.date_filter_button.config(
                text=f"{self.selected_date} \u25bc",
                bg=COLORS["accent"],
                fg="#111111",
            )
        else:
            self.date_filter_button.config(
                text="\u6309\u65e5\u671f \u25bc",
                bg=COLORS["card"],
                fg=COLORS["text"],
            )

    def _show_date_picker(self, event) -> None:
        if self.date_popup is not None and self.date_popup.winfo_exists():
            self._close_date_popup()
            return

        popup = tk.Toplevel(self)
        popup.withdraw()
        popup.overrideredirect(True)
        with contextlib.suppress(Exception):
            popup.transient(self)
        popup.configure(bg=COLORS["panel"], highlightthickness=1, highlightbackground=COLORS["border"])
        self.date_popup = popup

        now = dt.datetime.now()
        current_y = getattr(self, "_cal_year", now.year)
        current_m = getattr(self, "_cal_month", now.month)

        def render_cal(y, m):
            for widget in popup.winfo_children():
                widget.destroy()

            header = tk.Frame(popup, bg=COLORS["panel"])
            header.pack(fill="x", padx=8, pady=(8, 4))

            def move_month(dm):
                nm = m + dm
                ny = y
                if nm < 1: nm, ny = 12, ny - 1
                if nm > 12: nm, ny = 1, ny + 1
                self._cal_year, self._cal_month = ny, nm
                render_cal(ny, nm)

            tk.Button(header, text="\u25c0", command=lambda: move_month(-1), bg=COLORS["card"], fg=COLORS["text"], relief="flat", cursor="hand2", padx=6).pack(side="left")
            tk.Label(header, text=f"{y}年 {m}月", bg=COLORS["panel"], fg=COLORS["text"], font=FONT_TAG).pack(side="left", expand=True)
            tk.Button(header, text="\u25b6", command=lambda: move_month(1), bg=COLORS["card"], fg=COLORS["text"], relief="flat", cursor="hand2", padx=6).pack(side="right")

            grid = tk.Frame(popup, bg=COLORS["panel"])
            grid.pack(padx=8, pady=4)

            days = ["一", "二", "三", "四", "五", "六", "日"]
            for i, d in enumerate(days):
                tk.Label(grid, text=d, bg=COLORS["panel"], fg=COLORS["text_dim"], font=FONT_SMALL, width=4).grid(row=0, column=i, pady=(0,4))

            cal_days = calendar.monthcalendar(y, m)
            for r, week in enumerate(cal_days):
                for c, day in enumerate(week):
                    if day != 0:
                        date_str = f"{y:04d}-{m:02d}-{day:02d}"
                        is_selected = self.selected_date == date_str
                        btn_bg = COLORS["accent"] if is_selected else COLORS["card"]
                        btn_fg = "#111111" if is_selected else COLORS["text"]

                        btn = tk.Button(
                            grid, text=str(day), bg=btn_bg, fg=btn_fg, font=FONT_SMALL, relief="flat", cursor="hand2",
                            command=lambda d=date_str: select_date(d)
                        )
                        btn.grid(row=r+1, column=c, padx=2, pady=2, sticky="nsew")
                        grid.grid_columnconfigure(c, minsize=32)

            footer = tk.Frame(popup, bg=COLORS["panel"])
            footer.pack(fill="x", padx=8, pady=(4, 8))
            tk.Button(footer, text="清空日期", command=lambda: select_date(None), bg=COLORS["card"], fg=COLORS["danger"], font=FONT_SMALL, relief="flat", cursor="hand2", pady=4).pack(side="left", expand=True, fill="x", padx=(0, 4))
            tk.Button(footer, text="关闭日历", command=self._close_date_popup, bg=COLORS["card"], fg=COLORS["text"], font=FONT_SMALL, relief="flat", cursor="hand2", pady=4).pack(side="right", expand=True, fill="x", padx=(4, 0))

        def select_date(d_str):
            self.selected_date = d_str
            self._refresh_list()
            self._close_date_popup()

        render_cal(current_y, current_m)
        popup.update_idletasks()
        x = event.widget.winfo_rootx()
        y = event.widget.winfo_rooty() + event.widget.winfo_height() + 4
        popup.geometry(f"+{x}+{y}")
        popup.bind("<FocusOut>", self._on_date_popup_focus_out)
        popup.bind("<Escape>", lambda _event: self._close_date_popup())
        popup.bind("<Destroy>", self._on_date_popup_destroyed, add="+")
        popup.deiconify()
        popup.lift()
        popup.focus_force()

    def _widget_belongs_to_popup(self, widget: tk.Misc | None, popup: tk.Toplevel | None) -> bool:
        if widget is None or popup is None or not popup.winfo_exists():
            return False
        current = widget
        while current is not None:
            if current == popup:
                return True
            with contextlib.suppress(Exception):
                current = current.master
                continue
            break
        return False

    def _on_date_popup_focus_out(self, _event=None) -> None:
        popup = self.date_popup
        if popup is None or not popup.winfo_exists():
            return

        def close_if_focus_left() -> None:
            active_popup = self.date_popup
            if active_popup is None or not active_popup.winfo_exists():
                return
            focus_widget = self.focus_get()
            if not self._widget_belongs_to_popup(focus_widget, active_popup):
                self._close_date_popup()

        self.after_idle(close_if_focus_left)

    def _on_date_popup_destroyed(self, _event=None) -> None:
        self.date_popup = None

    def _close_date_popup(self) -> None:
        if self.date_popup is not None and self.date_popup.winfo_exists():
            self.date_popup.destroy()
        self.date_popup = None

    def _list_label_for_entry(self, entry: ClipboardEntry) -> str:
        short_time = entry.created_at[11:19]
        prefix = TYPE_LABELS.get(entry.type, "记录")
        badges: list[str] = []
        if entry.type == "text" and entry.has_rich_text:
            badges.append("富")
        if entry.type == "text" and has_formula_candidate(entry.plain_text):
            badges.append("公式")
        badge_text = f" [{' / '.join(badges)}]" if badges else ""
        return f"[{short_time}] {prefix}{badge_text}  {entry.summary}"

    def _purge_stale_drafts(self) -> None:
        valid_ids = {entry.id for entry in self.entries if entry.type == "text"}
        self.draft_payload_by_entry_id = {
            entry_id: payload
            for entry_id, payload in self.draft_payload_by_entry_id.items()
            if entry_id in valid_ids
        }
        self.preview_history_by_entry_id = {
            entry_id: history
            for entry_id, history in self.preview_history_by_entry_id.items()
            if entry_id in valid_ids
        }
        self.preview_history_index_by_entry_id = {
            entry_id: history_index
            for entry_id, history_index in self.preview_history_index_by_entry_id.items()
            if entry_id in valid_ids
        }

    def _refresh_list(self) -> None:
        self._purge_stale_drafts()
        self.visible_entries = self._filtered_entries()
        self.listbox.delete(0, "end")

        selected_index = None
        for index, entry in enumerate(self.visible_entries):
            self.listbox.insert("end", self._list_label_for_entry(entry))
            self.listbox.itemconfig(index, fg=TYPE_COLORS.get(entry.type, COLORS["text"]))
            if entry.id == self.selected_entry_id:
                selected_index = index

        self.count_label.config(text=f"{len(self.visible_entries)} 条记录（共 {len(self.entries)} 条）")
        self._refresh_filter_buttons()

        if selected_index is None:
            self.selected_entry_id = None
            self.listbox.selection_clear(0, "end")
            self._render_entry(None)
            return

        self.listbox.selection_set(selected_index)
        self.listbox.activate(selected_index)
        self.listbox.see(selected_index)
        self._render_entry(self.visible_entries[selected_index])

    def _find_entry_by_id(self, entry_id: int | None) -> ClipboardEntry | None:
        if entry_id is None:
            return None
        for entry in self.entries:
            if entry.id == entry_id:
                return entry
        return None

    def _selected_entry(self) -> ClipboardEntry | None:
        return self._find_entry_by_id(self.selected_entry_id)

    def _draft_payload_for_entry(self, entry: ClipboardEntry | None) -> RichTextPayload | None:
        if entry is None or entry.type != "text":
            return None
        return self.draft_payload_by_entry_id.get(entry.id)

    def _effective_text_for_entry(self, entry: ClipboardEntry | None) -> str | None:
        if entry is None or entry.type != "text":
            return None
        draft_payload = self._draft_payload_for_entry(entry)
        if draft_payload is not None:
            return draft_payload.plain_text
        return entry.plain_text or ""

    def _effective_text_payload(self, entry: ClipboardEntry | None) -> RichTextPayload | None:
        if entry is None or entry.type != "text":
            return None

        draft_payload = self._draft_payload_for_entry(entry)
        if draft_payload is not None:
            return draft_payload
        return entry_rich_payload(entry)

    def _capture_preview_text_payload(self) -> RichTextPayload:
        return serialize_text_widget_rich_payload(self.preview_text)

    def _store_draft_payload(self, entry: ClipboardEntry, payload: RichTextPayload) -> None:
        if rich_payload_matches_entry(entry, payload):
            self.draft_payload_by_entry_id.pop(entry.id, None)
        else:
            self.draft_payload_by_entry_id[entry.id] = payload

    def _ensure_preview_history(self, entry: ClipboardEntry) -> None:
        if entry.id in self.preview_history_by_entry_id and entry.id in self.preview_history_index_by_entry_id:
            return
        payload = self._effective_text_payload(entry) or entry_rich_payload(entry)
        self.preview_history_by_entry_id[entry.id] = [payload]
        self.preview_history_index_by_entry_id[entry.id] = 0

    def _reset_preview_history(self, entry: ClipboardEntry, payload: RichTextPayload) -> None:
        self.preview_history_by_entry_id[entry.id] = [payload]
        self.preview_history_index_by_entry_id[entry.id] = 0

    def _record_preview_history(self, entry: ClipboardEntry, payload: RichTextPayload) -> None:
        if self.preview_history_is_restoring:
            return
        self._ensure_preview_history(entry)
        history = self.preview_history_by_entry_id[entry.id]
        current_index = self.preview_history_index_by_entry_id.get(entry.id, len(history) - 1)
        current_payload = history[current_index]
        if (
            current_payload.plain_text == payload.plain_text
            and (current_payload.html_content or "") == (payload.html_content or "")
            and (current_payload.rtf_content or "") == (payload.rtf_content or "")
        ):
            return
        if current_index < len(history) - 1:
            history = history[: current_index + 1]
            self.preview_history_by_entry_id[entry.id] = history
        history.append(payload)
        self.preview_history_index_by_entry_id[entry.id] = len(history) - 1

    def _restore_preview_history(self, entry: ClipboardEntry, direction: int) -> str:
        self._ensure_preview_history(entry)
        history = self.preview_history_by_entry_id.get(entry.id, [])
        if not history:
            return "break"
        current_index = self.preview_history_index_by_entry_id.get(entry.id, len(history) - 1)
        target_index = current_index + direction
        if target_index < 0 or target_index >= len(history):
            return "break"
        self.preview_history_index_by_entry_id[entry.id] = target_index
        payload = history[target_index]
        self.preview_history_is_restoring = True
        try:
            self._render_payload_in_preview(payload)
            self._store_draft_payload(entry, payload)
        finally:
            self.preview_history_is_restoring = False
        self._sync_formula_button_for_entry(entry)
        return "break"

    def _current_full_text_payload(self, entry: ClipboardEntry) -> RichTextPayload:
        draft_payload = self._draft_payload_for_entry(entry)
        widget_payload = self._capture_preview_text_payload()
        if draft_payload is not None:
            return widget_payload

        entry_payload = entry_rich_payload(entry)
        if rich_payload_has_formatting(entry_payload):
            if entry_payload.html_content and entry_payload.rtf_content:
                return entry_payload
            if rich_payload_has_formatting(widget_payload):
                return RichTextPayload(
                    plain_text=entry_payload.plain_text,
                    html_content=entry_payload.html_content or widget_payload.html_content,
                    rtf_content=entry_payload.rtf_content or widget_payload.rtf_content,
                )
            return entry_payload

        return widget_payload

    def _current_visible_plain_text(self) -> str:
        if self.current_preview_mode != "text":
            return ""
        return self._capture_preview_text_payload().plain_text

    def _selected_text_payload(self) -> RichTextPayload | None:
        selected_range = self._selected_preview_range()
        if selected_range is None:
            return None
        return serialize_text_widget_rich_payload(self.preview_text, selected_range[0], selected_range[1])

    def _render_payload_in_preview(self, payload: RichTextPayload) -> None:
        if rich_payload_has_formatting(payload) and payload.html_content:
            self._set_preview_rich_text(payload.html_content, payload.plain_text)
        elif rich_payload_has_formatting(payload) and payload.rtf_content:
            self._set_preview_rtf_text(payload.rtf_content, payload.plain_text)
        else:
            self._set_preview_text(payload.plain_text)

    def _sync_formula_button_for_entry(self, entry: ClipboardEntry | None) -> None:
        if self.formula_button is None:
            return
        if entry is None or entry.type != "text":
            self.formula_button.config(state="disabled", bg=COLORS["card"], fg=COLORS["accent_soft"])
            return
        self.formula_button.config(state="normal", bg=COLORS["card"], fg=COLORS["accent"])

    def _sync_ocr_button_for_entry(self, entry: ClipboardEntry | None) -> None:
        if self.ocr_button is None:
            return

        if self.ocr_running:
            self.ocr_button.config(
                state="disabled",
                text=OCR_BUTTON_BUSY_TEXT,
                bg=COLORS["card"],
                fg=COLORS["accent_soft"],
            )
            return

        is_image_entry = (
            entry is not None
            and entry.type == "image"
            and bool(entry.image_path)
            and Path(entry.image_path).exists()
        )
        enabled = OCR_SUPPORTED_PLATFORM and is_image_entry
        self.ocr_button.config(
            state="normal" if enabled else "disabled",
            text=OCR_BUTTON_TEXT,
            bg=COLORS["card"],
            fg=COLORS["accent"] if enabled else COLORS["accent_soft"],
        )

    def _recognize_selected_image_text(self) -> None:
        if self.ocr_running:
            return

        entry = self._selected_entry()
        if entry is None or entry.type != "image" or not entry.image_path:
            self._sync_ocr_button_for_entry(entry)
            return

        if not OCR_SUPPORTED_PLATFORM:
            messagebox.showinfo("识别文字", "当前系统暂不支持 OCR。")
            return

        image_path = Path(entry.image_path)
        if not image_path.exists():
            self._sync_ocr_button_for_entry(entry)
            self._set_status("图片文件不存在，无法识别", healthy=False, temporary=True)
            return

        self.ocr_running = True
        self._sync_ocr_button_for_entry(entry)
        self._set_status(OCR_WORKING_STATUS_TEXT, healthy=True)
        threading.Thread(
            target=self._recognize_image_text_worker,
            args=(image_path,),
            name="OCRWorker",
            daemon=True,
        ).start()

    def _recognize_image_text_worker(self, image_path: Path) -> None:
        recognized_text = ""
        error: Exception | None = None
        try:
            recognized_text = recognize_image_text(image_path)
        except Exception as exc:
            error = exc

        def finish() -> None:
            self.ocr_running = False
            if self.is_shutting_down:
                return

            self._sync_ocr_button_for_entry(self._selected_entry())

            if error is not None:
                self._show_ocr_error(error)
                return

            text = recognized_text.strip()
            if not text:
                messagebox.showinfo("识别结果", "未识别到文字。")
                self._set_status("未识别到文字", healthy=False, temporary=True)
                return

            self._show_ocr_result_popup(text)
            self._set_status("识别完成", healthy=True, temporary=True)

        self._queue_ui_action(finish)

    def _show_ocr_error(self, error: Exception) -> None:
        if isinstance(error, OcrDependencyError):
            message = (
                "OCR 依赖未就绪。\n\n"
                "请先安装：\n"
                "pip install rapidocr-onnxruntime onnxruntime\n\n"
                f"详情：{error}"
            )
        else:
            message = f"没有成功识别图片文字。\n\n{error}"

        messagebox.showerror("识别失败", message)
        self._set_status("识别失败", healthy=False, temporary=True)

    def _show_ocr_result_popup(self, recognized_text: str) -> None:
        self._close_ocr_result_popup()

        popup_width = 760
        popup_height = 520
        popup = tk.Toplevel(self)
        popup.title("识别结果")
        popup.configure(bg=COLORS["panel"])
        popup.geometry(f"{popup_width}x{popup_height}")
        popup.minsize(520, 360)
        with contextlib.suppress(Exception):
            popup.transient(self)
        popup.protocol("WM_DELETE_WINDOW", self._close_ocr_result_popup)

        screen_width = popup.winfo_screenwidth()
        screen_height = popup.winfo_screenheight()
        x_pos = max(0, (screen_width - popup_width) // 2)
        y_pos = max(0, (screen_height - popup_height) // 2)
        popup.geometry(f"{popup_width}x{popup_height}+{x_pos}+{y_pos}")

        container = tk.Frame(popup, bg=COLORS["panel"])
        container.pack(fill="both", expand=True, padx=12, pady=12)

        text_box = tk.Text(
            container,
            bg=COLORS["panel_alt"],
            fg=COLORS["text"],
            insertbackground=COLORS["text"],
            relief="flat",
            bd=0,
            wrap="word",
            font=FONT_TEXT,
            padx=12,
            pady=12,
        )
        text_box.pack(side="left", fill="both", expand=True)

        scroll = tk.Scrollbar(container, command=text_box.yview)
        scroll.pack(side="right", fill="y")
        text_box.configure(yscrollcommand=scroll.set)

        text_box.insert("1.0", recognized_text)
        text_box.configure(state="disabled")

        actions = tk.Frame(popup, bg=COLORS["panel"])
        actions.pack(fill="x", padx=12, pady=(0, 12))

        copy_btn = tk.Button(
            actions,
            text="复制识别文本",
            bg=COLORS["accent"],
            fg="#111111",
            relief="flat",
            cursor="hand2",
            font=FONT_BUTTON,
            padx=14,
            pady=9,
            command=lambda: self._copy_ocr_text(recognized_text),
        )
        copy_btn.pack(side="left")

        close_btn = tk.Button(
            actions,
            text="关闭",
            bg=COLORS["card"],
            fg=COLORS["text"],
            relief="flat",
            cursor="hand2",
            font=FONT_BUTTON,
            padx=14,
            pady=9,
            command=self._close_ocr_result_popup,
        )
        close_btn.pack(side="right")

        self.ocr_result_popup = popup
        popup.lift()
        popup.focus_force()

    def _close_ocr_result_popup(self) -> None:
        if self.ocr_result_popup is not None and self.ocr_result_popup.winfo_exists():
            self.ocr_result_popup.destroy()
        self.ocr_result_popup = None

    def _copy_ocr_text(self, text: str) -> None:
        content = text.strip()
        if not content:
            return

        try:
            self.platform_services.set_clipboard_text(content)
        except Exception as exc:
            messagebox.showerror("复制失败", f"没有成功复制识别文本。\n\n{exc}")
            return

        self.suppressed_snapshot_key = (
            "text",
            hash_rich_text(content),
        )
        self._set_status("已复制识别文本", healthy=True, temporary=True)

    def _set_preview_text_actions(self, *, show_reset: bool, show_save: bool) -> None:
        if show_reset:
            if not self.btn_reset_text.winfo_manager():
                self.btn_reset_text.pack(side="left", padx=(0, 8), pady=8)
        elif self.btn_reset_text.winfo_manager():
            self.btn_reset_text.pack_forget()

        if show_save:
            if not self.btn_save_text.winfo_manager():
                self.btn_save_text.pack(side="left", padx=(0, 18), pady=8)
        elif self.btn_save_text.winfo_manager():
            self.btn_save_text.pack_forget()

    def _on_select(self, _event) -> None:
        selection = self.listbox.curselection()
        if not selection:
            self.selected_entry_id = None
            self._render_entry(None)
            return

        index = selection[0]
        if index >= len(self.visible_entries):
            return

        entry = self.visible_entries[index]
        self.selected_entry_id = entry.id
        self._render_entry(entry)

    def _set_preview_mode(self, mode: str) -> None:
        if self.current_preview_mode == mode:
            return
        for widget in (self.preview_empty, self.preview_text_frame, self.preview_image_frame):
            widget.pack_forget()

        if mode == "empty":
            self.preview_empty.pack(expand=True)
        elif mode == "text":
            self.preview_text_frame.pack(fill="both", expand=True)
        elif mode == "image":
            self.preview_image_frame.pack(fill="both", expand=True)

        self.current_preview_mode = mode

    def _set_preview_text(self, content: str, retain_edit: bool = False) -> None:
        self._set_preview_mode("text")
        if retain_edit:
            return
        self.preview_text_is_programmatic_update = True
        try:
            self.preview_text.delete("1.0", "end")
            self.preview_text.insert("1.0", content)
            self.preview_text.mark_set("insert", "1.0")
            self.preview_text.edit_modified(False)
        finally:
            self.preview_text_is_programmatic_update = False

    def _reset_text(self) -> None:
        entry = self._selected_entry()
        if entry is not None and entry.text_content is not None:
            self.draft_payload_by_entry_id.pop(entry.id, None)
            payload = entry_rich_payload(entry)
            self._render_payload_in_preview(payload)
            self._reset_preview_history(entry, payload)
            self._sync_formula_button_for_entry(entry)

    def _save_text(self) -> None:
        entry = self._selected_entry()
        if entry is None or entry.type != "text":
            return
        payload = self._capture_preview_text_payload()
        has_rich_text = rich_payload_has_formatting(payload)
        new_hash = hash_rich_text(payload.plain_text, payload.html_content, payload.rtf_content)
        if not rich_payload_matches_entry(entry, payload):
            entry.plain_text = payload.plain_text
            entry.summary = summarize_text(payload.plain_text)
            entry.content_hash = new_hash
            entry.html_content = payload.html_content
            entry.rtf_content = payload.rtf_content
            entry.source_formats_json = text_source_formats_json(has_rich_text)
            entry.has_rich_text = has_rich_text
            self.draft_payload_by_entry_id.pop(entry.id, None)
            self.store.update_entry_text(
                entry.id,
                plain_text=payload.plain_text,
                summary=entry.summary,
                content_hash=new_hash,
                html_content=payload.html_content,
                rtf_content=payload.rtf_content,
                source_formats_json=entry.source_formats_json,
                has_rich_text=has_rich_text,
            )
            self._reset_preview_history(entry, payload)
            self._refresh_list()
            self._set_status("保存成功", healthy=True, temporary=True)
            return

        self.draft_payload_by_entry_id.pop(entry.id, None)
        self._reset_preview_history(entry, entry_rich_payload(entry))
        self._sync_formula_button_for_entry(entry)

    def _toggle_favorite_status(self) -> None:
        pass

    def _set_preview_rich_text(self, html_content: str, fallback_text: str) -> None:
        fragment = extract_html_fragment(html_content)
        if not fragment:
            self._set_preview_text(fallback_text)
            return

        self._set_preview_mode("text")
        self.preview_text_is_programmatic_update = True
        try:
            self.preview_text.delete("1.0", "end")
            try:
                renderer = HtmlPreviewRenderer(self.preview_text, "1.0")
                renderer.feed(fragment)
                renderer.close()
            except Exception:
                self._set_preview_text(fallback_text)
                return

            if not self.preview_text.get("1.0", "end-1c").strip():
                self._set_preview_text(fallback_text)
                return
            self.preview_text.mark_set("insert", "1.0")
            self.preview_text.edit_modified(False)
        finally:
            self.preview_text_is_programmatic_update = False

    def _set_preview_rtf_text(self, rtf_content: str, fallback_text: str) -> None:
        self._set_preview_mode("text")
        self.preview_text_is_programmatic_update = True
        try:
            self.preview_text.delete("1.0", "end")
            try:
                renderer = RtfPreviewRenderer(self.preview_text, "1.0")
                renderer.render(rtf_content)
            except Exception:
                self._set_preview_text(fallback_text)
                return

            if not self.preview_text.get("1.0", "end-1c").strip():
                self._set_preview_text(fallback_text)
                return
            self.preview_text.mark_set("insert", "1.0")
            self.preview_text.edit_modified(False)
        finally:
            self.preview_text_is_programmatic_update = False

    def _insert_payload_into_preview(self, payload: RichTextPayload) -> str:
        selected_range = self._selected_preview_range()
        if selected_range is not None:
            start_index, end_index = selected_range
        else:
            start_index = self.preview_text.index("insert")
            end_index = start_index

        self.preview_text_is_programmatic_update = True
        try:
            self.preview_text.delete(start_index, end_index)
            insert_start = self.preview_text.index(start_index)
            if payload.html_content:
                fragment = extract_html_fragment(payload.html_content)
                if fragment:
                    renderer = HtmlPreviewRenderer(self.preview_text, insert_start)
                    renderer.feed(fragment)
                    renderer.close()
                    final_index = renderer.cursor_index
                else:
                    self.preview_text.insert(insert_start, payload.plain_text)
                    final_index = self.preview_text.index(f"{insert_start}+{len(payload.plain_text)}c")
            elif payload.rtf_content:
                renderer = RtfPreviewRenderer(self.preview_text, insert_start)
                renderer.render(payload.rtf_content)
                final_index = renderer.cursor_index
            else:
                self.preview_text.insert(insert_start, payload.plain_text)
                final_index = self.preview_text.index(f"{insert_start}+{len(payload.plain_text)}c")
            self.preview_text.mark_set("insert", final_index)
            self.preview_text.tag_remove("sel", "1.0", "end")
            self.preview_text.edit_modified(False)
            return final_index
        finally:
            self.preview_text_is_programmatic_update = False

    def _selected_preview_range(self) -> tuple[str, str] | None:
        try:
            return (self.preview_text.index("sel.first"), self.preview_text.index("sel.last"))
        except tk.TclError:
            return None

    def _apply_preview_script_tag(self, tag_name: str) -> None:
        entry = self._selected_entry()
        if entry is None or entry.type != "text":
            return
        selected_range = self._selected_preview_range()
        if selected_range is None:
            self._set_status("先选中文本，再设置上标或下标", healthy=False, temporary=True)
            return
        start_index, end_index = selected_range
        opposite_tag = "preview_subscript" if tag_name == "preview_superscript" else "preview_superscript"
        self.preview_text.tag_remove(opposite_tag, start_index, end_index)
        self.preview_text.tag_add(tag_name, start_index, end_index)
        payload = self._capture_preview_text_payload()
        self._store_draft_payload(entry, payload)
        self._record_preview_history(entry, payload)
        self._sync_formula_button_for_entry(entry)

    def _clear_preview_script_tags(self) -> None:
        entry = self._selected_entry()
        if entry is None or entry.type != "text":
            return
        selected_range = self._selected_preview_range()
        if selected_range is None:
            self._set_status("先选中文本，再恢复为正文", healthy=False, temporary=True)
            return
        start_index, end_index = selected_range
        self.preview_text.tag_remove("preview_superscript", start_index, end_index)
        self.preview_text.tag_remove("preview_subscript", start_index, end_index)
        payload = self._capture_preview_text_payload()
        self._store_draft_payload(entry, payload)
        self._record_preview_history(entry, payload)
        self._sync_formula_button_for_entry(entry)

    def _set_preview_image(self, image_path: str | None) -> None:
        if not image_path or not Path(image_path).exists():
            self.preview_image_label.config(cursor="arrow")
            self._set_preview_text("图片文件不存在，可能已经被清理。")
            return

        self._set_preview_mode("image")
        max_width = max(self.preview_frame.winfo_width() - 24, 200)
        max_height = max(self.preview_frame.winfo_height() - 24, 200)

        preview = load_image_safely(image_path)
        preview.thumbnail((max_width, max_height), Image.LANCZOS)
        self.preview_photo = ImageTk.PhotoImage(preview)
        self.preview_image_label.config(image=self.preview_photo, text="", cursor="hand2")

    def _open_zoomed_image(self, _event=None) -> None:
        entry = self._selected_entry()
        if entry is None or entry.type != "image" or not entry.image_path or not Path(entry.image_path).exists():
            return

        self._close_zoomed_image()

        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        max_width = max(int(screen_width * 0.86), 400)
        max_height = max(int(screen_height * 0.82), 300)

        self.zoom_source_image = load_image_safely(entry.image_path)
        source_width, source_height = self.zoom_source_image.size
        fit_scale = min(max_width / source_width, max_height / source_height, 1.0)
        self.zoom_base_size = (
            max(1, int(round(source_width * fit_scale))),
            max(1, int(round(source_height * fit_scale))),
        )
        self.zoom_scale = 1.0

        overlay = tk.Toplevel(self)
        overlay.withdraw()
        overlay.overrideredirect(True)
        overlay.configure(bg="#070c14")
        overlay.geometry(f"{screen_width}x{screen_height}+0+0")
        overlay.attributes("-topmost", True)
        with contextlib.suppress(Exception):
            overlay.attributes("-alpha", 0.96)

        dimmer = tk.Frame(overlay, bg="#070c14")
        dimmer.pack(fill="both", expand=True)

        image_shell = tk.Frame(
            dimmer,
            bg=COLORS["panel"],
            highlightthickness=1,
            highlightbackground=COLORS["border"],
            padx=18,
            pady=18,
        )
        image_shell.place(relx=0.5, rely=0.5, anchor="center")

        self.zoom_image_label = tk.Label(image_shell, bg=COLORS["panel"], bd=0)
        self.zoom_image_label.pack()
        self.zoom_hint_label = tk.Label(
            image_shell,
            text="滚轮缩放，点击任意位置或按 Esc 关闭",
            bg=COLORS["panel"],
            fg=COLORS["text_dim"],
            font=FONT_SMALL,
        )
        self.zoom_hint_label.pack(pady=(10, 0))

        for widget in (overlay, dimmer, image_shell, self.zoom_image_label, self.zoom_hint_label):
            widget.bind("<Button-1>", self._close_zoomed_image)
            widget.bind("<MouseWheel>", self._on_zoom_mousewheel)
            widget.bind("<Button-4>", self._on_zoom_mousewheel)
            widget.bind("<Button-5>", self._on_zoom_mousewheel)
        overlay.bind("<Escape>", self._close_zoomed_image)

        overlay.deiconify()
        overlay.focus_force()
        self.zoom_overlay = overlay
        self._render_zoomed_image()

    def _close_zoomed_image(self, _event=None) -> None:
        if self.zoom_overlay is not None and self.zoom_overlay.winfo_exists():
            self.zoom_overlay.destroy()
        self.zoom_overlay = None
        self.zoom_photo = None
        self.zoom_image_label = None
        self.zoom_hint_label = None
        self.zoom_source_image = None
        self.zoom_base_size = None
        self.zoom_scale = 1.0

    def _render_zoomed_image(self) -> None:
        if self.zoom_source_image is None or self.zoom_base_size is None or self.zoom_image_label is None:
            return

        width = max(1, int(round(self.zoom_base_size[0] * self.zoom_scale)))
        height = max(1, int(round(self.zoom_base_size[1] * self.zoom_scale)))
        preview = self.zoom_source_image.resize((width, height), Image.LANCZOS)
        self.zoom_photo = ImageTk.PhotoImage(preview)
        self.zoom_image_label.config(image=self.zoom_photo)
        if self.zoom_hint_label is not None:
            self.zoom_hint_label.config(
                text=f"滚轮缩放 {int(round(self.zoom_scale * 100))}% ，点击任意位置或按 Esc 关闭"
            )

    def _on_zoom_mousewheel(self, event):
        if self.zoom_overlay is None or not self.zoom_overlay.winfo_exists():
            return "break"

        delta = 0
        if hasattr(event, "delta") and event.delta:
            delta = 1 if event.delta > 0 else -1
        elif getattr(event, "num", None) == 4:
            delta = 1
        elif getattr(event, "num", None) == 5:
            delta = -1

        if delta == 0:
            return "break"

        factor = 1.12 if delta > 0 else (1 / 1.12)
        self.zoom_scale = min(5.0, max(0.55, self.zoom_scale * factor))
        self._render_zoomed_image()
        return "break"

    def _other_preview_text(self, entry: ClipboardEntry) -> str:
        if not entry.other_payload_json:
            return "没有可预览的数据。"

        try:
            payload = json.loads(entry.other_payload_json)
        except json.JSONDecodeError:
            return entry.other_payload_json

        if entry.other_kind == "files":
            paths = payload.get("paths", [])
            return "文件/文件夹列表\n\n" + "\n".join(paths)

        if entry.other_kind == "formats":
            formats = payload.get("formats", [])
            return "剪贴板格式摘要\n\n" + "\n".join(formats)

        return json.dumps(payload, ensure_ascii=False, indent=2)

    def _render_entry(self, entry: ClipboardEntry | None) -> None:
        if entry is None:
            self.preview_header.config(text="从左侧选择一条记录查看详情", fg=COLORS["text_dim"])
            if self.preview_header_favorite_btn is not None:
                self.preview_header_favorite_btn.config(text="", fg=COLORS["text_dim"])
            self.copy_button.config(state="disabled", text="复制到剪贴板", bg=COLORS["accent"])
            self._sync_formula_button_for_entry(None)
            self._sync_ocr_button_for_entry(None)
            self.delete_button.config(state="disabled")
            self.preview_photo = None
            self.preview_image_label.config(image="", text="", cursor="arrow")
            self._set_preview_text_actions(show_reset=False, show_save=False)
            self._set_preview_mode("empty")
            return

        badges: list[str] = []
        if entry.type == "text" and entry.has_rich_text:
            badges.append("富文本")
        if entry.type == "text" and has_formula_candidate(entry.plain_text):
            badges.append("公式文本")
        badge_text = f"    状态：{' / '.join(badges)}" if badges else ""
        self.preview_header.config(
            text=f"类型：{TYPE_LABELS.get(entry.type, entry.type)}    时间：{entry.created_at}{badge_text}",
            fg=COLORS["text"],
        )
        self.delete_button.config(state="normal")
        self._update_favorite_icon(entry)

        if entry.type == "text":
            self.copy_button.config(state="normal", text="复制到剪贴板", bg=COLORS["accent"])
            self._sync_formula_button_for_entry(entry)
            self._sync_ocr_button_for_entry(None)
            self._set_preview_text_actions(show_reset=True, show_save=True)
            payload = self._effective_text_payload(entry)
            if payload is not None:
                self._render_payload_in_preview(payload)
                self._ensure_preview_history(entry)
            return

        if entry.type == "image":
            can_copy = "normal" if entry.image_path and Path(entry.image_path).exists() else "disabled"
            self.copy_button.config(state=can_copy, text="复制到剪贴板", bg=COLORS["accent"])
            self._sync_formula_button_for_entry(None)
            self._sync_ocr_button_for_entry(entry)
            self._set_preview_text_actions(show_reset=False, show_save=False)
            self._set_preview_image(entry.image_path)
            return

        self.copy_button.config(state="disabled", text="复制到剪贴板", bg=COLORS["accent"])
        self._sync_formula_button_for_entry(None)
        self._sync_ocr_button_for_entry(None)
        self._set_preview_text_actions(show_reset=False, show_save=False)
        self._set_preview_text(self._other_preview_text(entry))

    def _block_preview_input(self, event):
        entry = self._selected_entry()
        if entry is None or entry.type != "text":
            return "break"
        return None

    def _on_preview_edit_command(self, _event=None):
        entry = self._selected_entry()
        if entry is None or entry.type != "text":
            return "break"
        return None

    def _paste_into_preview(self, _event=None):
        entry = self._selected_entry()
        if entry is None or entry.type != "text":
            return "break"

        try:
            payload = self.platform_services.read_clipboard_text_payload()
        except Exception as exc:
            messagebox.showerror("粘贴失败", f"没有成功读取系统剪贴板。\n\n{exc}")
            return "break"

        if payload is None:
            return "break"

        self._insert_payload_into_preview(payload)
        payload_after_paste = self._capture_preview_text_payload()
        self._store_draft_payload(entry, payload_after_paste)
        self._record_preview_history(entry, payload_after_paste)
        self._sync_formula_button_for_entry(entry)
        return "break"

    def _copy_from_preview_selection(self, _event=None):
        entry = self._selected_entry()
        if entry is None or self.current_preview_mode != "text":
            return "break"

        try:
            if entry.type == "text":
                payload = self._selected_text_payload()
                if payload is None or payload.plain_text == "":
                    payload = self._current_full_text_payload(entry)
                self.platform_services.set_clipboard_rich_text(
                    payload.plain_text,
                    payload.html_content,
                    payload.rtf_content,
                )
                self.suppressed_snapshot_key = (
                    "text",
                    hash_rich_text(payload.plain_text, payload.html_content, payload.rtf_content),
                )
            else:
                selected_text = self.preview_text.selection_get()
                self.platform_services.set_clipboard_text(selected_text)
        except Exception as exc:
            messagebox.showerror("复制失败", f"没有成功写回系统剪贴板。\n\n{exc}")
            return "break"

        self._flash_copy_button("已复制")
        self._set_status("已复制到系统剪贴板", healthy=True, temporary=True)
        return "break"

    def _undo_preview_edit(self, _event=None):
        entry = self._selected_entry()
        if entry is None or entry.type != "text":
            return "break"
        return self._restore_preview_history(entry, -1)

    def _redo_preview_edit(self, _event=None):
        entry = self._selected_entry()
        if entry is None or entry.type != "text":
            return "break"
        return self._restore_preview_history(entry, 1)

    def _apply_superscript_shortcut(self, _event=None):
        self._apply_preview_script_tag("preview_superscript")
        return "break"

    def _apply_subscript_shortcut(self, _event=None):
        self._apply_preview_script_tag("preview_subscript")
        return "break"

    def _clear_script_shortcut(self, _event=None):
        self._clear_preview_script_tags()
        return "break"

    def _on_preview_text_modified(self, _event=None):
        if self.preview_text_is_programmatic_update:
            return
        if not self.preview_text.edit_modified():
            return
        self.preview_text.edit_modified(False)

        entry = self._selected_entry()
        if entry is None or entry.type != "text":
            return

        payload = self._capture_preview_text_payload()
        self._store_draft_payload(entry, payload)
        self._record_preview_history(entry, payload)
        self._sync_formula_button_for_entry(entry)

    def _select_all_preview(self, _event):
        self.preview_text.tag_add("sel", "1.0", "end-1c")
        return "break"

    def _on_preview_resize(self, _event) -> None:
        if self.current_preview_mode != "image":
            return
        if self.preview_resize_after_id is not None:
            self.after_cancel(self.preview_resize_after_id)
        self.preview_resize_after_id = self.after(120, self._rerender_selected_image)

    def _rerender_selected_image(self) -> None:
        self.preview_resize_after_id = None
        entry = self._selected_entry()
        if entry and entry.type == "image":
            self._set_preview_image(entry.image_path)

    def _flash_copy_button(self, text: str) -> None:
        self.copy_button.config(text=text, bg=COLORS["success"])
        self.after(1400, lambda: self.copy_button.config(text="复制到剪贴板", bg=COLORS["accent"]))

    def _flash_formula_button(self, text: str) -> None:
        if self.formula_button is None:
            return
        self.formula_button.config(text=text, bg=COLORS["success"], fg="#111111")

        def restore_formula_button() -> None:
            if self.formula_button is None:
                return
            self.formula_button.config(text="复制纯文本", bg=COLORS["card"], fg=COLORS["accent"])

        self.after(1400, restore_formula_button)

    def _copy_selected(self) -> None:
        entry = self._selected_entry()
        if entry is None:
            return

        try:
            if entry.type == "text":
                payload = self._current_full_text_payload(entry)
                self.platform_services.set_clipboard_rich_text(
                    payload.plain_text,
                    payload.html_content,
                    payload.rtf_content,
                )
                self.suppressed_snapshot_key = (
                    "text",
                    hash_rich_text(payload.plain_text, payload.html_content, payload.rtf_content),
                )
            elif entry.type == "image" and entry.image_path:
                self.platform_services.set_clipboard_image(entry.image_path)
                self.suppressed_snapshot_key = entry.snapshot_key
            else:
                return
        except Exception as exc:
            messagebox.showerror("复制失败", f"没有成功写回系统剪贴板。\n\n{exc}")
            return

        self._flash_copy_button("已复制")
        self._set_status("已复制到系统剪贴板", healthy=True, temporary=True)

    def _copy_formula_selected(self) -> None:
        entry = self._selected_entry()
        source_text = self._current_visible_plain_text()
        if entry is None or entry.type != "text" or not source_text:
            return

        try:
            self.platform_services.set_clipboard_text(source_text)
        except Exception as exc:
            messagebox.showerror("复制失败", f"没有成功复制纯文本。\n\n{exc}")
            return

        self.suppressed_snapshot_key = (
            "text",
            hash_rich_text(source_text),
        )
        self._flash_formula_button("纯文本已复制")
        self._set_status("已复制为纯文本", healthy=True, temporary=True)

    def _toggle_favorite(self, event=None):
        entry = self._selected_entry()
        if entry is None:
            return
        
        entry.is_favorite = not getattr(entry, 'is_favorite', False)
        self.store.update_favorite(entry.id, entry.is_favorite)
        self._update_favorite_icon(entry)
        self._refresh_list()
        self._set_status("已收藏" if entry.is_favorite else "已取消收藏", healthy=True, temporary=True)
        
    def _update_favorite_icon(self, entry):
        star_text = "★" if getattr(entry, "is_favorite", False) else "☆"
        star_fg = "#FFD700" if getattr(entry, "is_favorite", False) else COLORS["text"]
        if self.preview_header_favorite_btn is not None:
            self.preview_header_favorite_btn.config(text=star_text, fg=star_fg)

    def _delete_selected(self) -> None:
        entry = self._selected_entry()
        if entry is None:
            return

        current_index = next(
            (index for index, visible_entry in enumerate(self.visible_entries) if visible_entry.id == entry.id),
            None,
        )
        self.store.delete_entry(entry.id)
        self.draft_payload_by_entry_id.pop(entry.id, None)
        self.preview_history_by_entry_id.pop(entry.id, None)
        self.preview_history_index_by_entry_id.pop(entry.id, None)
        self.entries = self.store.load_entries()
        self.selected_entry_id = None
        self._refresh_list()

        if current_index is not None and self.visible_entries:
            next_index = min(current_index, len(self.visible_entries) - 1)
            self.listbox.selection_set(next_index)
            self.listbox.event_generate("<<ListboxSelect>>")

    def _clear_history(self) -> None:
        if not messagebox.askyesno("清空历史", "确认清空所有已保存的剪贴板记录吗？"):
            return

        self.store.clear_entries()
        self.entries = []
        self.draft_payload_by_entry_id.clear()
        self.preview_history_by_entry_id.clear()
        self.preview_history_index_by_entry_id.clear()
        self.selected_entry_id = None
        self._refresh_list()
        self._set_status("历史已清空", healthy=True, temporary=True)

    def _place_window_centered(self) -> None:
        self.update_idletasks()
        width = self.winfo_width() or WINDOW_WIDTH
        height = self.winfo_height() or WINDOW_HEIGHT
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x_pos = max(0, (screen_width - width) // 2)
        y_pos = max(0, (screen_height - height) // 2 - 24)
        self.geometry(f"{width}x{height}+{x_pos}+{y_pos}")

    def show_window(self, center: bool = False) -> None:
        if self.is_shutting_down:
            return
        if center or not self.window_has_been_presented:
            self._place_window_centered()
        self.deiconify()
        self._apply_window_shell_icons()
        self.lift()
        self.attributes("-topmost", True)
        if self.is_pinned:
            self.after(120, self._apply_pinned_state)
        else:
            self.after(200, lambda: self.attributes("-topmost", False))
        self.focus_force()
        self.is_window_visible = True
        self.window_has_been_presented = True
        if self.search_entry is not None and self.search_entry.winfo_exists():
            self.after(120, lambda: self.search_entry.focus_set())

    def _on_window_unmap(self, _event=None) -> None:
        with contextlib.suppress(Exception):
            if self.state() == "iconic":
                self._close_auto_delete_popup()
                self._close_date_popup()

    def hide_window(self) -> None:
        self._close_auto_delete_popup()
        self._close_date_popup()
        self._close_zoomed_image()
        self._close_ocr_result_popup()
        self.withdraw()
        self.is_window_visible = False

    def toggle_window(self) -> None:
        if self.is_window_visible:
            self.hide_window()
        else:
            self.show_window(center=not self.window_has_been_presented)

    def _present_existing_instance(self) -> None:
        self.show_window(center=not self.window_has_been_presented)

    def _set_startup_enabled(self, enabled: bool) -> None:
        self.startup_manager.set_enabled(enabled)
        self.store.set_setting("startup_opt_out", "0" if enabled else "1")
        self._refresh_startup_chip()

    def _toggle_startup(self) -> None:
        enabled = self.startup_manager.is_enabled()
        try:
            self._set_startup_enabled(not enabled)
        except Exception as exc:
            messagebox.showerror("设置失败", f"开机自启设置没有成功更新。\n\n{exc}")

    def _ensure_default_startup(self) -> None:
        opt_out = self.store.get_setting("startup_opt_out", "0") == "1"
        if getattr(sys, "frozen", False) and not opt_out and not self.startup_manager.is_enabled():
            with contextlib.suppress(Exception):
                self.startup_manager.set_enabled(True)

    def _build_tray_menu(self) -> list[TrayMenuItem | None]:
        startup_enabled = False
        with contextlib.suppress(Exception):
            startup_enabled = self.startup_manager.is_enabled()

        return [
            TrayMenuItem(
                label="隐藏窗口" if self.is_window_visible else "显示窗口",
                callback=lambda: self._queue_ui_action(self.toggle_window),
            ),
            TrayMenuItem(
                label="窗口置顶",
                callback=lambda: self._queue_ui_action(self._toggle_pinned),
                checked=self.is_pinned,
            ),
            TrayMenuItem(
                label="开机自启",
                callback=lambda: self._queue_ui_action(self._toggle_startup),
                checked=startup_enabled,
            ),
            None,
            TrayMenuItem(
                label="清空历史",
                callback=lambda: self._queue_ui_action(self._clear_history),
            ),
            TrayMenuItem(
                label="退出",
                callback=lambda: self._queue_ui_action(self.request_exit),
            ),
        ]

    def request_exit(self) -> None:
        if self.is_shutting_down:
            return
        self.is_shutting_down = True
        self._close_auto_delete_popup()
        self._close_date_popup()
        self._close_zoomed_image()
        self._close_ocr_result_popup()

        for after_id in (self.poll_after_id, self.queue_after_id, self.status_reset_after_id, self.preview_resize_after_id):
            if after_id is not None:
                with contextlib.suppress(Exception):
                    self.after_cancel(after_id)

        if self.tray is not None:
            self.tray.stop()
            self.tray = None

        self.store.close()
        self.destroy()

    def destroy(self) -> None:
        self._release_window_shell_icons()
        super().destroy()


def main() -> None:
    platform_services = create_platform_services()
    launched_from_startup = is_startup_launch()
    instance_guard = None

    try:
        if launched_from_startup:
            sys.argv = [arg for arg in sys.argv if arg != STARTUP_ARG]
        elif platform_services.maybe_relaunch_with_pythonw():
            return

        instance_guard = platform_services.create_single_instance_guard()
        if not instance_guard.acquire():
            if not launched_from_startup:
                platform_services.signal_existing_instance(should_wake=True)
            return

        platform_services.enable_ui_features()
        app = ClipboardManagerApp(
            launched_from_startup=launched_from_startup,
            platform_services=platform_services,
        )
        app.mainloop()
    finally:
        if instance_guard is not None:
            instance_guard.release()


if __name__ == "__main__":
    main()
