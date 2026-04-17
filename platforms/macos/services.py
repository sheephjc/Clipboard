from __future__ import annotations

import contextlib
import io
from pathlib import Path
import sys
import threading
from typing import Callable

from PIL import Image

from shared.data_paths import portable_data_dir
from shared.models import ClipboardCapture, RichTextPayload
from shared.rich_text import (
    hash_image,
    hash_other,
    hash_rich_text,
    json_dumps,
    plain_text_from_html,
    plain_text_from_rtf,
    summarize_files,
    summarize_formats,
    summarize_text,
)


class MacStartupManager:
    def is_enabled(self) -> bool:
        return False

    def set_enabled(self, enabled: bool) -> None:
        if enabled:
            raise RuntimeError("macOS startup registration is not implemented in the Windows-prepared source yet.")


class MacSingleInstanceGuard:
    def acquire(self) -> bool:
        return True

    def release(self) -> None:
        return


class MacMenuBarTray:
    """Best-effort rumps menu bar bridge for later macOS validation."""

    def __init__(
        self,
        tooltip: str,
        menu_factory: Callable[[], list[object | None]],
        default_action: Callable[[], None],
        wake_action: Callable[[], None] | None = None,
        icon_path: str | None = None,
    ) -> None:
        self.tooltip = tooltip
        self.menu_factory = menu_factory
        self.default_action = default_action
        self.wake_action = wake_action or default_action
        self.icon_path = icon_path
        self._thread: threading.Thread | None = None
        self._app = None

    def start(self) -> None:
        try:
            import rumps
        except Exception:
            return

        def run() -> None:
            app = rumps.App(self.tooltip, icon=self.icon_path, quit_button=None)
            self._app = app

            @rumps.clicked("Show / Hide")
            def _toggle(_sender):
                self.default_action()

            @rumps.clicked("Quit")
            def _quit(_sender):
                for item in self.menu_factory():
                    if item is not None and getattr(item, "label", "") in {"退出", "Quit"}:
                        item.callback()
                        return

            app.run()

        self._thread = threading.Thread(target=run, name="MacMenuBarThread", daemon=True)
        self._thread.start()

    def stop(self) -> None:
        if self._app is not None:
            with contextlib.suppress(Exception):
                self._app.quit_application()


class MacPlatformServices:
    platform_name = "macos"

    def __init__(self, project_root: Path, startup_arg: str = "--startup") -> None:
        self.project_root = project_root
        self.startup_arg = startup_arg

    def app_data_dir(self) -> Path:
        return portable_data_dir(self.project_root)

    def create_startup_manager(self) -> MacStartupManager:
        return MacStartupManager()

    def create_single_instance_guard(self) -> MacSingleInstanceGuard:
        return MacSingleInstanceGuard()

    def signal_existing_instance(self, should_wake: bool = True) -> bool:
        return False

    def maybe_relaunch_with_pythonw(self) -> bool:
        return False

    def enable_ui_features(self) -> None:
        return

    def tray_icon_path(self) -> str | None:
        icon = self.project_root / "assets" / "clipboard_icon_32.png"
        return str(icon) if icon.exists() else None

    def create_tray_icon(
        self,
        tooltip: str,
        menu_factory,
        default_action,
        wake_action=None,
        icon_path: str | None = None,
    ):
        return MacMenuBarTray(tooltip, menu_factory, default_action, wake_action, icon_path)

    def _pasteboard(self):
        from AppKit import NSPasteboard

        return NSPasteboard.generalPasteboard()

    def _types(self) -> list[str]:
        types = self._pasteboard().types() or []
        return [str(item) for item in types]

    def read_clipboard_text_payload(self) -> RichTextPayload | None:
        from AppKit import NSPasteboardTypeHTML, NSPasteboardTypeRTF, NSPasteboardTypeString

        pasteboard = self._pasteboard()
        types = self._types()
        if NSPasteboardTypeString not in types and NSPasteboardTypeHTML not in types and NSPasteboardTypeRTF not in types:
            return None

        plain_text = pasteboard.stringForType_(NSPasteboardTypeString) or ""
        html_content = pasteboard.stringForType_(NSPasteboardTypeHTML) if NSPasteboardTypeHTML in types else None
        rtf_content = None
        if NSPasteboardTypeRTF in types:
            data = pasteboard.dataForType_(NSPasteboardTypeRTF)
            if data is not None:
                rtf_content = bytes(data).decode("utf-8", errors="ignore")

        if not plain_text and html_content:
            plain_text = plain_text_from_html(html_content)
        if not plain_text and rtf_content:
            plain_text = plain_text_from_rtf(rtf_content)
        if plain_text == "" and not html_content and not rtf_content:
            return None
        return RichTextPayload(plain_text=plain_text, html_content=html_content, rtf_content=rtf_content)

    def read_clipboard_capture(self) -> ClipboardCapture | None:
        from AppKit import NSPasteboardTypeHTML, NSPasteboardTypePNG, NSPasteboardTypeRTF, NSPasteboardTypeString, NSPasteboardTypeTIFF

        pasteboard = self._pasteboard()
        types = self._types()

        file_paths = self._read_file_paths()
        if file_paths:
            payload = {"paths": file_paths}
            return ClipboardCapture(
                type="other",
                content_hash=hash_other("files", payload),
                summary=summarize_files(file_paths),
                other_kind="files",
                other_payload_json=json_dumps(payload),
            )

        for image_type in (NSPasteboardTypePNG, NSPasteboardTypeTIFF):
            if image_type in types:
                data = pasteboard.dataForType_(image_type)
                if data is None:
                    continue
                with Image.open(io.BytesIO(bytes(data))) as image:
                    captured = image.convert("RGBA")
                return ClipboardCapture(
                    type="image",
                    content_hash=hash_image(captured),
                    summary=f"图片 {captured.width}x{captured.height}",
                    image=captured,
                )

        text_payload = self.read_clipboard_text_payload()
        if text_payload is not None:
            names = [
                name
                for name in (NSPasteboardTypeString, NSPasteboardTypeHTML, NSPasteboardTypeRTF)
                if name in types
            ]
            return ClipboardCapture(
                type="text",
                content_hash=hash_rich_text(
                    text_payload.plain_text,
                    text_payload.html_content,
                    text_payload.rtf_content,
                ),
                summary=summarize_text(text_payload.plain_text or plain_text_from_html(text_payload.html_content)),
                plain_text=text_payload.plain_text,
                html_content=text_payload.html_content,
                rtf_content=text_payload.rtf_content,
                source_formats_json=json_dumps(names),
                has_rich_text=bool(text_payload.html_content or text_payload.rtf_content),
            )

        if types:
            payload = {"formats": types}
            return ClipboardCapture(
                type="other",
                content_hash=hash_other("formats", payload),
                summary=summarize_formats(types),
                other_kind="formats",
                other_payload_json=json_dumps(payload),
            )
        return None

    def _read_file_paths(self) -> list[str]:
        pasteboard = self._pasteboard()
        paths: list[str] = []
        with contextlib.suppress(Exception):
            items = pasteboard.propertyListForType_("NSFilenamesPboardType") or []
            paths.extend(str(item) for item in items)
        if paths:
            return paths

        with contextlib.suppress(Exception):
            file_url = pasteboard.stringForType_("public.file-url")
            if file_url:
                from Foundation import NSURL

                url = NSURL.URLWithString_(file_url)
                if url is not None and url.isFileURL():
                    paths.append(str(url.path()))
        return paths

    def set_clipboard_rich_text(
        self,
        plain_text: str,
        html_content: str | None = None,
        rtf_content: str | None = None,
    ) -> None:
        from AppKit import NSPasteboardTypeHTML, NSPasteboardTypeRTF, NSPasteboardTypeString
        from Foundation import NSData

        pasteboard = self._pasteboard()
        pasteboard.clearContents()
        pasteboard.setString_forType_(plain_text, NSPasteboardTypeString)
        if html_content:
            pasteboard.setString_forType_(html_content, NSPasteboardTypeHTML)
        if rtf_content:
            encoded = rtf_content.encode("utf-8")
            data = NSData.dataWithBytes_length_(encoded, len(encoded))
            pasteboard.setData_forType_(data, NSPasteboardTypeRTF)

    def set_clipboard_text(self, text: str) -> None:
        self.set_clipboard_rich_text(text)

    def set_clipboard_image(self, image_path: str) -> None:
        from AppKit import NSPasteboardTypePNG
        from Foundation import NSData

        with Image.open(image_path) as image:
            output = io.BytesIO()
            image.convert("RGBA").save(output, format="PNG")
            png_bytes = output.getvalue()

        pasteboard = self._pasteboard()
        pasteboard.clearContents()
        data = NSData.dataWithBytes_length_(png_bytes, len(png_bytes))
        pasteboard.setData_forType_(data, NSPasteboardTypePNG)

