from __future__ import annotations

from dataclasses import dataclass, field
import contextlib
import json

from PIL import Image


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
    images: list[Image.Image] = field(default_factory=list)
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
    image_paths_json: str | None = None
    other_kind: str | None = None
    other_payload_json: str | None = None
    is_favorite: bool = False

    @property
    def snapshot_key(self) -> tuple[str, str]:
        return (self.type, self.content_hash)

    @property
    def text_content(self) -> str | None:
        return self.plain_text


TEXT_ENTRY_TYPES = {"text", "mixed"}
IMAGE_ENTRY_TYPES = {"image", "mixed"}


def entry_has_text(entry: ClipboardEntry | None) -> bool:
    return entry is not None and entry.type in TEXT_ENTRY_TYPES


def entry_image_paths(entry: ClipboardEntry | None) -> list[str]:
    if entry is None:
        return []
    paths: list[str] = []
    if entry.image_paths_json:
        with contextlib.suppress(json.JSONDecodeError, TypeError):
            decoded = json.loads(entry.image_paths_json)
            if isinstance(decoded, list):
                paths.extend(str(path) for path in decoded if path)
    if not paths and entry.image_path:
        paths.append(entry.image_path)
    elif entry.image_path and entry.image_path not in paths:
        paths.insert(0, entry.image_path)
    return paths


def entry_has_image(entry: ClipboardEntry | None) -> bool:
    return entry is not None and entry.type in IMAGE_ENTRY_TYPES and bool(entry_image_paths(entry))


@dataclass
class RichTextPayload:
    plain_text: str
    html_content: str | None = None
    rtf_content: str | None = None

