from __future__ import annotations

from dataclasses import dataclass

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

