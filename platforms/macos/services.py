from __future__ import annotations

import argparse
import contextlib
import io
import os
from pathlib import Path
import plistlib
import socket
import subprocess
import sys
import threading
from typing import Callable, Iterable
from urllib.parse import unquote, urlparse

from PIL import Image

from shared.data_paths import portable_data_dir
from shared.models import ClipboardCapture, RichTextPayload
from shared.rich_text import (
    IMAGE_FILE_EXTENSIONS,
    decode_clipboard_rtf,
    hash_images,
    hash_mixed_content,
    hash_other,
    hash_rich_text,
    html_with_local_image_paths,
    images_from_html,
    json_dumps,
    load_image_file_list,
    plain_text_from_html,
    plain_text_from_rtf,
    rich_payload_has_visible_text,
    summarize_files,
    summarize_formats,
    summarize_text,
    unique_images,
)

TEXT_TYPES = (
    "public.utf8-plain-text",
    "NSStringPboardType",
    "public.utf16-plain-text",
    "public.plain-text",
)
HTML_TYPES = ("public.html", "Apple HTML pasteboard type")
RTF_TYPES = ("public.rtf", "NeXT Rich Text Format v1.0 pasteboard type")
IMAGE_TYPES = ("public.png", "public.tiff", "public.jpeg")
FILE_URL_TYPE = "public.file-url"
FILENAMES_TYPE = "NSFilenamesPboardType"
LAUNCH_AGENT_LABEL = "com.hjc.clipboard"


def _dedupe_strings(values: Iterable[str]) -> list[str]:
    result: list[str] = []
    seen: set[str] = set()
    for value in values:
        text = str(value)
        if not text or text in seen:
            continue
        seen.add(text)
        result.append(text)
    return result


def _nsdata_to_bytes(data) -> bytes:
    if data is None:
        return b""
    if isinstance(data, bytes):
        return data
    if isinstance(data, bytearray):
        return bytes(data)
    return bytes(data)


def _nsdata_from_bytes(payload: bytes):
    try:
        from Foundation import NSData

        return NSData.dataWithBytes_length_(payload, len(payload))
    except Exception:
        return payload


def _file_url_to_path(value: str | None) -> str | None:
    if not value:
        return None
    text = str(value)
    try:
        from Foundation import NSURL

        url = NSURL.URLWithString_(text)
        if url is not None and url.isFileURL():
            return str(url.path())
    except Exception:
        pass

    parsed = urlparse(text)
    if parsed.scheme.lower() != "file":
        return None
    path = unquote(parsed.path or "")
    if parsed.netloc and parsed.netloc.lower() != "localhost":
        path = f"//{parsed.netloc}{path}"
    if sys.platform.startswith("win") and len(path) >= 4 and path[0] == "/" and path[2] == ":":
        path = path[1:]
    return path or None


def _is_image_path(path_text: str) -> bool:
    path = Path(path_text)
    return path.is_file() and path.suffix.lower() in IMAGE_FILE_EXTENSIONS


def _encode_rtf_for_pasteboard(rtf_content: str) -> bytes:
    return rtf_content.encode("latin-1", errors="replace")


def _decode_text_data(payload: bytes | None, encoding: str = "utf-8") -> str | None:
    if not payload:
        return None
    return payload.decode(encoding, errors="replace")


class MacStartupManager:
    def __init__(self, project_root: Path, startup_arg: str = "--startup") -> None:
        self.project_root = project_root
        self.startup_arg = startup_arg
        self.launch_agents_dir = Path.home() / "Library" / "LaunchAgents"
        self.plist_path = self.launch_agents_dir / f"{LAUNCH_AGENT_LABEL}.plist"

    def build_command(self) -> list[str]:
        if getattr(sys, "frozen", False):
            return [str(Path(sys.executable).resolve()), self.startup_arg]
        return [str(Path(sys.executable).resolve()), str(self.project_root / "clipboard_manager.py"), self.startup_arg]

    def _read_plist(self) -> dict | None:
        try:
            with self.plist_path.open("rb") as handle:
                payload = plistlib.load(handle)
        except FileNotFoundError:
            return None
        except Exception:
            return None
        return payload if isinstance(payload, dict) else None

    def is_enabled(self) -> bool:
        payload = self._read_plist()
        if not payload:
            return False
        return payload.get("Label") == LAUNCH_AGENT_LABEL

    def ensure_current_command(self) -> None:
        if not self.is_enabled():
            return
        payload = self._read_plist() or {}
        if payload.get("ProgramArguments") != self.build_command():
            self.set_enabled(True)

    def set_enabled(self, enabled: bool) -> None:
        if enabled:
            self.launch_agents_dir.mkdir(parents=True, exist_ok=True)
            payload = {
                "Label": LAUNCH_AGENT_LABEL,
                "ProgramArguments": self.build_command(),
                "RunAtLoad": True,
                "KeepAlive": False,
                "ProcessType": "Interactive",
            }
            with self.plist_path.open("wb") as handle:
                plistlib.dump(payload, handle, sort_keys=True)
            self._launchctl("bootstrap")
            return

        self._launchctl("bootout")
        with contextlib.suppress(FileNotFoundError):
            self.plist_path.unlink()

    def _launchctl(self, action: str) -> None:
        if sys.platform != "darwin":
            return
        target = f"gui/{os.getuid()}"
        command = ["launchctl", action, target, str(self.plist_path)]
        with contextlib.suppress(Exception):
            subprocess.run(command, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=5)


class MacSingleInstanceGuard:
    def __init__(self, data_dir: Path) -> None:
        self.data_dir = data_dir
        self.lock_path = data_dir / "clipboard.lock"
        self._handle = None

    def acquire(self) -> bool:
        if sys.platform != "darwin":
            return True
        import fcntl

        self.data_dir.mkdir(parents=True, exist_ok=True)
        self._handle = self.lock_path.open("a+", encoding="utf-8")
        try:
            fcntl.flock(self._handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
        except BlockingIOError:
            with contextlib.suppress(Exception):
                self._handle.close()
            self._handle = None
            return False
        self._handle.seek(0)
        self._handle.truncate()
        self._handle.write(str(os.getpid()))
        self._handle.flush()
        return True

    def release(self) -> None:
        if self._handle is None:
            return
        if sys.platform == "darwin":
            with contextlib.suppress(Exception):
                import fcntl

                fcntl.flock(self._handle.fileno(), fcntl.LOCK_UN)
        with contextlib.suppress(Exception):
            self._handle.close()
        self._handle = None


def send_macos_command(socket_path: Path, command_id: str, timeout: float = 1.2) -> bool:
    try:
        client = socket.socket(socket.AF_UNIX, socket.SOCK_STREAM)
        client.settimeout(timeout)
        with client:
            client.connect(str(socket_path))
            client.sendall((command_id + "\n").encode("utf-8"))
        return True
    except Exception:
        return False


class MacMenuBarTray:
    COMMAND_SHOW_HIDE = "show_hide"
    COMMAND_PIN = "toggle_pinned"
    COMMAND_STARTUP = "toggle_startup"
    COMMAND_CLEAR = "clear_history"
    COMMAND_QUIT = "quit"
    COMMAND_WAKE = "wake"
    COMMAND_STOP = "__stop__"

    def __init__(
        self,
        tooltip: str,
        menu_factory: Callable[[], list[object | None]],
        default_action: Callable[[], None],
        wake_action: Callable[[], None] | None = None,
        icon_path: str | None = None,
        socket_path: Path | None = None,
        project_root: Path | None = None,
        helper_arg: str = "--menubar-helper",
    ) -> None:
        self.tooltip = tooltip
        self.menu_factory = menu_factory
        self.default_action = default_action
        self.wake_action = wake_action or default_action
        self.icon_path = icon_path
        self.socket_path = socket_path or (Path.home() / ".clipboard-menubar.sock")
        self.project_root = project_root or Path.cwd()
        self.helper_arg = helper_arg
        self._server_thread: threading.Thread | None = None
        self._helper_process: subprocess.Popen | None = None
        self._stopping = threading.Event()

    def start(self) -> None:
        if self._server_thread and self._server_thread.is_alive():
            return
        self.socket_path.parent.mkdir(parents=True, exist_ok=True)
        with contextlib.suppress(FileNotFoundError):
            self.socket_path.unlink()
        self._server_thread = threading.Thread(target=self._serve_commands, name="MacMenuCommandServer", daemon=True)
        self._server_thread.start()
        self._spawn_helper()

    def stop(self) -> None:
        self._stopping.set()
        send_macos_command(self.socket_path, self.COMMAND_STOP, timeout=0.2)
        if self._helper_process is not None:
            with contextlib.suppress(Exception):
                self._helper_process.terminate()
            with contextlib.suppress(Exception):
                self._helper_process.wait(timeout=2)
            self._helper_process = None
        if self._server_thread and self._server_thread.is_alive():
            self._server_thread.join(timeout=2)
        with contextlib.suppress(FileNotFoundError):
            self.socket_path.unlink()

    def _spawn_helper(self) -> None:
        if sys.platform != "darwin":
            return
        if getattr(sys, "frozen", False):
            command = [str(Path(sys.executable).resolve())]
        else:
            command = [str(Path(sys.executable).resolve()), str(self.project_root / "clipboard_manager.py")]
        command.extend(
            [
                self.helper_arg,
                "--socket",
                str(self.socket_path),
                "--tooltip",
                self.tooltip,
            ]
        )
        if self.icon_path:
            command.extend(["--icon", self.icon_path])
        with contextlib.suppress(Exception):
            self._helper_process = subprocess.Popen(
                command,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                close_fds=True,
            )

    def _serve_commands(self) -> None:
        server = socket.socket(socket.AF_UNIX, socket.SOCK_STREAM)
        with server:
            server.bind(str(self.socket_path))
            server.listen(8)
            server.settimeout(0.4)
            while not self._stopping.is_set():
                try:
                    conn, _addr = server.accept()
                except socket.timeout:
                    continue
                except OSError:
                    break
                with conn:
                    command_id = conn.recv(4096).decode("utf-8", errors="replace").strip()
                if command_id == self.COMMAND_STOP:
                    break
                self._dispatch(command_id)

    def _dispatch(self, command_id: str) -> None:
        if command_id == self.COMMAND_WAKE:
            self.wake_action()
            return
        if command_id == self.COMMAND_SHOW_HIDE:
            self.default_action()
            return
        for item in self.menu_factory():
            if item is None:
                continue
            item_id = getattr(item, "command_id", None)
            if item_id == command_id:
                item.callback()
                return


class MacPlatformServices:
    platform_name = "macos"

    def __init__(
        self,
        project_root: Path,
        startup_arg: str = "--startup",
        menubar_helper_arg: str = "--menubar-helper",
        pasteboard_factory: Callable[[], object] | None = None,
    ) -> None:
        self.project_root = project_root
        self.startup_arg = startup_arg
        self.menubar_helper_arg = menubar_helper_arg
        self._pasteboard_factory = pasteboard_factory
        self._data_dir = portable_data_dir(project_root)
        self.command_socket_path = self._data_dir / "clipboard.sock"

    def app_data_dir(self) -> Path:
        return self._data_dir

    def create_startup_manager(self) -> MacStartupManager:
        return MacStartupManager(self.project_root, self.startup_arg)

    def create_single_instance_guard(self) -> MacSingleInstanceGuard:
        return MacSingleInstanceGuard(self.app_data_dir())

    def signal_existing_instance(self, should_wake: bool = True) -> bool:
        return send_macos_command(
            self.command_socket_path,
            MacMenuBarTray.COMMAND_WAKE if should_wake else MacMenuBarTray.COMMAND_SHOW_HIDE,
        )

    def maybe_relaunch_with_pythonw(self) -> bool:
        return False

    def enable_ui_features(self) -> None:
        return

    def clipboard_change_count(self) -> int | None:
        with contextlib.suppress(Exception):
            return int(self._pasteboard().changeCount())
        return None

    def activate_app(self) -> None:
        try:
            from AppKit import NSApplication

            NSApplication.sharedApplication().activateIgnoringOtherApps_(True)
        except Exception:
            return

    def open_path(self, path: Path) -> None:
        subprocess.Popen(["open", str(path)])

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
        return MacMenuBarTray(
            tooltip,
            menu_factory,
            default_action,
            wake_action,
            icon_path,
            socket_path=self.command_socket_path,
            project_root=self.project_root,
            helper_arg=self.menubar_helper_arg,
        )

    def _pasteboard(self):
        if self._pasteboard_factory is not None:
            return self._pasteboard_factory()
        from AppKit import NSPasteboard

        return NSPasteboard.generalPasteboard()

    def _types(self, source=None) -> list[str]:
        source = source or self._pasteboard()
        with contextlib.suppress(Exception):
            return [str(item) for item in (source.types() or [])]
        return []

    def _items(self) -> list[object]:
        pasteboard = self._pasteboard()
        with contextlib.suppress(Exception):
            items = pasteboard.pasteboardItems() or []
            return list(items)
        return []

    def _string_for(self, source, type_name: str) -> str | None:
        with contextlib.suppress(Exception):
            value = source.stringForType_(type_name)
            if value is not None:
                return str(value)
        return None

    def _data_for(self, source, type_name: str) -> bytes | None:
        with contextlib.suppress(Exception):
            data = source.dataForType_(type_name)
            if data is not None:
                return _nsdata_to_bytes(data)
        return None

    def _property_list_for(self, source, type_name: str):
        with contextlib.suppress(Exception):
            return source.propertyListForType_(type_name)
        return None

    def _read_text_payload_from_source(self, source) -> RichTextPayload | None:
        types = self._types(source)
        plain_text = ""
        html_content = None
        rtf_content = None

        for type_name in TEXT_TYPES:
            if type_name in types:
                plain_text = self._string_for(source, type_name) or ""
                if plain_text:
                    break

        for type_name in HTML_TYPES:
            if type_name not in types:
                continue
            html_content = self._string_for(source, type_name)
            if html_content is None:
                html_content = _decode_text_data(self._data_for(source, type_name), "utf-8")
            if html_content:
                break

        for type_name in RTF_TYPES:
            if type_name not in types:
                continue
            rtf_content = self._string_for(source, type_name)
            if rtf_content is None:
                rtf_content = decode_clipboard_rtf(self._data_for(source, type_name))
            if rtf_content:
                break

        if not plain_text and html_content:
            plain_text = plain_text_from_html(html_content)
        if not plain_text and rtf_content:
            plain_text = plain_text_from_rtf(rtf_content)
        if plain_text == "" and not html_content and not rtf_content:
            return None
        return RichTextPayload(plain_text=plain_text, html_content=html_content, rtf_content=rtf_content)

    def read_clipboard_text_payload(self) -> RichTextPayload | None:
        pasteboard = self._pasteboard()
        payload = self._read_text_payload_from_source(pasteboard)
        if payload is not None:
            return payload
        for item in self._items():
            payload = self._read_text_payload_from_source(item)
            if payload is not None:
                return payload
        return None

    def _read_file_paths(self) -> list[str]:
        pasteboard = self._pasteboard()
        paths: list[str] = []

        payload = self._property_list_for(pasteboard, FILENAMES_TYPE)
        if isinstance(payload, (list, tuple)):
            paths.extend(str(item) for item in payload if item)

        for source in [pasteboard, *self._items()]:
            payload = self._property_list_for(source, FILENAMES_TYPE)
            if isinstance(payload, (list, tuple)):
                paths.extend(str(item) for item in payload if item)

            file_url = self._string_for(source, FILE_URL_TYPE)
            file_path = _file_url_to_path(file_url)
            if file_path:
                paths.append(file_path)

        return _dedupe_strings(paths)

    def _read_images_from_pasteboard(self) -> list[Image.Image]:
        images: list[Image.Image] = []
        sources = [*self._items(), self._pasteboard()]
        for source in sources:
            types = self._types(source)
            for image_type in IMAGE_TYPES:
                if image_type not in types:
                    continue
                data = self._data_for(source, image_type)
                if not data:
                    continue
                with contextlib.suppress(Exception):
                    with Image.open(io.BytesIO(data)) as image:
                        images.append(image.convert("RGBA"))
                    break
        return unique_images(images)

    def read_clipboard_capture(self) -> ClipboardCapture | None:
        pasteboard = self._pasteboard()
        types = self._types(pasteboard)
        item_types = [type_name for item in self._items() for type_name in self._types(item)]
        all_types = _dedupe_strings(types + item_types)

        file_paths = self._read_file_paths()
        if file_paths and not all(_is_image_path(path) for path in file_paths):
            payload = {"paths": file_paths}
            return ClipboardCapture(
                type="other",
                content_hash=hash_other("files", payload),
                summary=summarize_files(file_paths),
                other_kind="files",
                other_payload_json=json_dumps(payload),
            )

        text_payload = self.read_clipboard_text_payload()

        images: list[Image.Image] = []
        image_format_names: list[str] = []
        has_file_images = False
        if file_paths:
            file_images = load_image_file_list(file_paths) or []
            if file_images:
                images.extend(file_images)
                image_format_names.append(FILE_URL_TYPE)
                has_file_images = True

        pasteboard_images = [] if has_file_images else self._read_images_from_pasteboard()
        if pasteboard_images:
            images.extend(pasteboard_images)
            image_format_names.extend([type_name for type_name in all_types if type_name in IMAGE_TYPES])

        if text_payload is not None and text_payload.html_content:
            html_images = images_from_html(text_payload.html_content)
            if html_images:
                images.extend(html_images)
                image_format_names.append("HTML image")

        images = unique_images(images)
        text_format_names = [
            type_name
            for type_name in all_types
            if type_name in (*TEXT_TYPES, *HTML_TYPES, *RTF_TYPES)
        ]

        if rich_payload_has_visible_text(text_payload) and images:
            text_summary = summarize_text(text_payload.plain_text or plain_text_from_html(text_payload.html_content))
            image_summary = f"图片 {len(images)} 张" if len(images) > 1 else f"图片 {images[0].width}x{images[0].height}"
            summary = f"{text_summary} + {image_summary}" if text_summary else image_summary
            return ClipboardCapture(
                type="mixed",
                content_hash=hash_mixed_content(
                    text_payload.plain_text,
                    text_payload.html_content,
                    text_payload.rtf_content,
                    images,
                ),
                summary=summary,
                plain_text=text_payload.plain_text,
                html_content=text_payload.html_content,
                rtf_content=text_payload.rtf_content,
                source_formats_json=json_dumps(text_format_names + image_format_names),
                has_rich_text=bool(text_payload.html_content or text_payload.rtf_content),
                image=images[0],
                images=images,
            )

        if images:
            image = images[0]
            summary = f"图片 {len(images)} 张" if len(images) > 1 else f"图片 {image.width}x{image.height}"
            return ClipboardCapture(
                type="image",
                content_hash=hash_images(images),
                summary=summary,
                source_formats_json=json_dumps(image_format_names) if image_format_names else None,
                image=image,
                images=images,
            )

        if text_payload is not None:
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
                source_formats_json=json_dumps(text_format_names),
                has_rich_text=bool(text_payload.html_content or text_payload.rtf_content),
            )

        if all_types:
            payload = {"formats": all_types}
            return ClipboardCapture(
                type="other",
                content_hash=hash_other("formats", payload),
                summary=summarize_formats(all_types),
                other_kind="formats",
                other_payload_json=json_dumps(payload),
            )
        return None

    def _set_string(self, pasteboard, value: str, type_name: str) -> None:
        with contextlib.suppress(Exception):
            pasteboard.setString_forType_(value, type_name)

    def _set_data(self, pasteboard, payload: bytes, type_name: str) -> None:
        with contextlib.suppress(Exception):
            pasteboard.setData_forType_(_nsdata_from_bytes(payload), type_name)

    def _set_file_urls(self, pasteboard, image_paths: list[str]) -> None:
        existing_paths = [str(Path(path)) for path in image_paths if Path(path).exists()]
        if not existing_paths:
            return
        with contextlib.suppress(Exception):
            pasteboard.setPropertyList_forType_(existing_paths, FILENAMES_TYPE)
        if existing_paths:
            with contextlib.suppress(Exception):
                pasteboard.setString_forType_(Path(existing_paths[0]).resolve().as_uri(), FILE_URL_TYPE)

    def set_clipboard_rich_text(
        self,
        plain_text: str,
        html_content: str | None = None,
        rtf_content: str | None = None,
    ) -> None:
        pasteboard = self._pasteboard()
        pasteboard.clearContents()
        self._set_string(pasteboard, plain_text, TEXT_TYPES[0])
        if html_content:
            self._set_string(pasteboard, html_content, HTML_TYPES[0])
        if rtf_content:
            self._set_data(pasteboard, _encode_rtf_for_pasteboard(rtf_content), RTF_TYPES[0])

    def set_clipboard_text(self, text: str) -> None:
        self.set_clipboard_rich_text(text)

    def set_clipboard_image(self, image_path: str) -> None:
        with Image.open(image_path) as image:
            normalized = image.convert("RGBA")
            png_output = io.BytesIO()
            normalized.save(png_output, format="PNG")
            tiff_output = io.BytesIO()
            normalized.save(tiff_output, format="TIFF")

        pasteboard = self._pasteboard()
        pasteboard.clearContents()
        self._set_data(pasteboard, png_output.getvalue(), "public.png")
        self._set_data(pasteboard, tiff_output.getvalue(), "public.tiff")
        self._set_file_urls(pasteboard, [image_path])

    def set_clipboard_rich_text_and_image(
        self,
        plain_text: str,
        html_content: str | None,
        rtf_content: str | None,
        image_path: str,
    ) -> None:
        self.set_clipboard_rich_text_and_images(plain_text, html_content, rtf_content, [image_path])

    def set_clipboard_rich_text_and_images(
        self,
        plain_text: str,
        html_content: str | None,
        rtf_content: str | None,
        image_paths: list[str],
    ) -> None:
        local_html_content = html_with_local_image_paths(plain_text, html_content, image_paths)
        pasteboard = self._pasteboard()
        pasteboard.clearContents()

        if plain_text:
            self._set_string(pasteboard, plain_text, TEXT_TYPES[0])
        if local_html_content:
            self._set_string(pasteboard, local_html_content, HTML_TYPES[0])
        if rtf_content:
            self._set_data(pasteboard, _encode_rtf_for_pasteboard(rtf_content), RTF_TYPES[0])
        self._set_file_urls(pasteboard, image_paths)

        if len(image_paths) == 1 and not plain_text and not local_html_content and not rtf_content:
            self.set_clipboard_image(image_paths[0])


def run_menubar_helper(socket_path: Path, tooltip: str, icon_path: str | None = None) -> None:
    try:
        import rumps
    except Exception:
        return

    app = rumps.App(tooltip, icon=icon_path, quit_button=None)

    def send(command_id: str) -> None:
        send_macos_command(socket_path, command_id, timeout=0.5)

    app.menu = [
        rumps.MenuItem("显示/隐藏", callback=lambda _sender: send(MacMenuBarTray.COMMAND_SHOW_HIDE)),
        rumps.MenuItem("窗口置顶", callback=lambda _sender: send(MacMenuBarTray.COMMAND_PIN)),
        rumps.MenuItem("开机自启", callback=lambda _sender: send(MacMenuBarTray.COMMAND_STARTUP)),
        None,
        rumps.MenuItem("清空历史", callback=lambda _sender: send(MacMenuBarTray.COMMAND_CLEAR)),
        rumps.MenuItem("退出", callback=lambda _sender: send(MacMenuBarTray.COMMAND_QUIT)),
    ]
    app.run()


def run_menubar_helper_from_argv(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--menubar-helper", action="store_true")
    parser.add_argument("--socket", required=True)
    parser.add_argument("--tooltip", default="Clipboard")
    parser.add_argument("--icon")
    args, _extra = parser.parse_known_args(argv)
    run_menubar_helper(Path(args.socket), args.tooltip, args.icon)
