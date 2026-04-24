from __future__ import annotations

import base64
import contextlib
import hashlib
import html
import io
import json
import re
import zipfile
from html.parser import HTMLParser
from pathlib import Path
from urllib.parse import unquote, urlparse

from PIL import Image

from shared.models import ClipboardCapture, ClipboardEntry, RichTextPayload

HTML_FRAGMENT_START = "<!--StartFragment-->"
HTML_FRAGMENT_END = "<!--EndFragment-->"
LIST_SUMMARY_LIMIT = 72
IMAGE_FILE_EXTENSIONS = {
    ".bmp",
    ".dib",
    ".gif",
    ".jfif",
    ".jpeg",
    ".jpg",
    ".png",
    ".tif",
    ".tiff",
    ".webp",
}
GVML_MEDIA_PREFIX = "clipboard/media/"


def json_dumps(payload: object) -> str:
    return json.dumps(payload, ensure_ascii=False, sort_keys=True)


def hash_text(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def hash_rich_text(
    plain_text: str,
    html_content: str | None = None,
    rtf_content: str | None = None,
) -> str:
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


def hash_images(images: list[Image.Image]) -> str:
    image_list = unique_images(images)
    if len(image_list) == 1:
        return hash_image(image_list[0])
    digest = hashlib.sha256()
    digest.update(b"IMAGES\0")
    for image in image_list:
        digest.update(hash_image(image).encode("ascii"))
        digest.update(b"\0")
    return digest.hexdigest()


def unique_images(images: list[Image.Image]) -> list[Image.Image]:
    unique: list[Image.Image] = []
    seen_hashes: set[str] = set()
    for image in images:
        image_hash = hash_image(image)
        if image_hash in seen_hashes:
            continue
        seen_hashes.add(image_hash)
        unique.append(image)
    return unique


def capture_image_list(capture: ClipboardCapture) -> list[Image.Image]:
    images = list(capture.images)
    if not images and capture.image is not None:
        images.append(capture.image)
    return unique_images(images)


def hash_mixed_content(
    plain_text: str,
    html_content: str | None,
    rtf_content: str | None,
    images: Image.Image | list[Image.Image],
) -> str:
    image_list = [images] if isinstance(images, Image.Image) else images
    digest = hashlib.sha256()
    digest.update(b"MIXED\0")
    digest.update(plain_text.encode("utf-8"))
    digest.update(b"\0HTML\0")
    digest.update((html_content or "").encode("utf-8"))
    digest.update(b"\0RTF\0")
    digest.update((rtf_content or "").encode("latin-1", errors="replace"))
    for image in unique_images(image_list):
        digest.update(b"\0IMAGE\0")
        digest.update(hash_image(image).encode("ascii"))
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


def rich_payload_has_formatting(payload: RichTextPayload) -> bool:
    return bool(payload.html_content or payload.rtf_content)


def rich_payload_has_visible_text(payload: RichTextPayload | None) -> bool:
    return bool(payload and payload.plain_text.strip())


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


URL_PATTERN = re.compile(r"(?i)(?:https?://|www\.)[^\s<>'\"]+")
URL_TRAILING_PUNCTUATION = ".,;:!?)]}，。；：！？）】、"


def normalize_web_url(value: str | None) -> str | None:
    if not value:
        return None
    url = html.unescape(value).strip().strip("\"'")
    url = url.rstrip(URL_TRAILING_PUNCTUATION)
    if not url:
        return None
    parsed = urlparse(url)
    if parsed.scheme.lower() in {"http", "https"} and parsed.netloc:
        return url
    if not parsed.scheme and url.lower().startswith("www."):
        return f"https://{url}"
    return None


def unique_urls(urls: list[str]) -> list[str]:
    unique: list[str] = []
    seen: set[str] = set()
    for url in urls:
        normalized = normalize_web_url(url)
        if normalized is None:
            continue
        key = normalized.lower()
        if key in seen:
            continue
        seen.add(key)
        unique.append(normalized)
    return unique


def web_urls_from_text(text: str | None) -> list[str]:
    if not text:
        return []
    return unique_urls([match.group(0) for match in URL_PATTERN.finditer(text)])


class HtmlLinkUrlParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.urls: list[str] = []

    def handle_starttag(self, tag: str, attrs) -> None:
        if tag.lower() != "a":
            return
        attr_map = {str(name).lower(): str(value) for name, value in attrs if value is not None}
        href = attr_map.get("href")
        if href:
            self.urls.append(href)


def html_link_urls(raw_html: str | None) -> list[str]:
    fragment = extract_html_fragment(raw_html)
    if not fragment:
        return []
    parser = HtmlLinkUrlParser()
    with contextlib.suppress(Exception):
        parser.feed(fragment)
        parser.close()
    return unique_urls(parser.urls)


def rich_payload_link_hint_urls(payload: RichTextPayload | None) -> list[str]:
    if payload is None:
        return []
    plain_text = payload.plain_text or ""
    visible_text = plain_text_from_html(payload.html_content) if payload.html_content else plain_text
    visible_urls = {url.lower() for url in web_urls_from_text(visible_text)}
    urls = unique_urls(html_link_urls(payload.html_content) + web_urls_from_text(plain_text))
    return [url for url in urls if url.lower() not in visible_urls]


class HtmlImageSourceParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.sources: list[str] = []

    def handle_starttag(self, tag: str, attrs) -> None:
        if tag.lower() != "img":
            return
        attr_map = {str(name).lower(): str(value) for name, value in attrs if value is not None}
        src = attr_map.get("src")
        if src:
            self.sources.append(html.unescape(src).strip())


def html_image_sources(raw_html: str | None) -> list[str]:
    fragment = extract_html_fragment(raw_html)
    if not fragment:
        return []
    parser = HtmlImageSourceParser()
    with contextlib.suppress(Exception):
        parser.feed(fragment)
        parser.close()
    return parser.sources


def image_from_html_source(src: str) -> Image.Image | None:
    source = src.strip().strip("\"'")
    if not source:
        return None

    if source.lower().startswith("data:image/"):
        match = re.match(r"(?is)data:image/[^;,\s]+;base64,(.+)", source)
        if not match:
            return None
        try:
            image_bytes = base64.b64decode(match.group(1), validate=False)
            with Image.open(io.BytesIO(image_bytes)) as image:
                return image.convert("RGBA")
        except Exception:
            return None

    path_text = ""
    parsed = urlparse(source)
    if parsed.scheme.lower() == "file":
        path_text = unquote(parsed.path or "")
        if parsed.netloc and parsed.netloc.lower() != "localhost":
            path_text = f"//{parsed.netloc}{path_text}"
        if re.match(r"^/[A-Za-z]:[/\\]", path_text):
            path_text = path_text[1:]
    elif not parsed.scheme:
        path_text = unquote(source)

    if not path_text:
        return None

    path = Path(path_text)
    if not path.exists() or not path.is_file():
        return None

    try:
        with Image.open(path) as image:
            return image.convert("RGBA")
    except Exception:
        return None


def images_from_html(raw_html: str | None) -> list[Image.Image]:
    images: list[Image.Image] = []
    for src in html_image_sources(raw_html):
        image = image_from_html_source(src)
        if image is not None:
            images.append(image)
    return unique_images(images)


def load_image_file_list(paths: list[str]) -> list[Image.Image] | None:
    if not paths:
        return None

    images: list[Image.Image] = []
    for path_text in paths:
        path = Path(path_text)
        if not path.is_file() or path.suffix.lower() not in IMAGE_FILE_EXTENSIONS:
            return None
        try:
            with Image.open(path) as image:
                images.append(image.convert("RGBA"))
        except Exception:
            return None
    return unique_images(images)


def images_from_gvml_data(raw_data: bytes | bytearray | memoryview | None) -> list[Image.Image]:
    if raw_data is None:
        return []

    data = bytes(raw_data)
    try:
        archive = zipfile.ZipFile(io.BytesIO(data))
    except zipfile.BadZipFile:
        return []

    images: list[Image.Image] = []
    with archive:
        for info in archive.infolist():
            normalized_name = info.filename.replace("\\", "/")
            lower_name = normalized_name.lower()
            suffix = Path(lower_name).suffix
            if info.is_dir() or not lower_name.startswith(GVML_MEDIA_PREFIX) or suffix not in IMAGE_FILE_EXTENSIONS:
                continue
            try:
                with archive.open(info) as image_file:
                    with Image.open(image_file) as image:
                        images.append(image.convert("RGBA"))
            except Exception:
                continue
    return unique_images(images)


def first_image_from_html(raw_html: str | None) -> Image.Image | None:
    images = images_from_html(raw_html)
    if images:
        return images[0]
    return None


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


def clipboard_file_uri(image_path: str) -> str:
    return Path(image_path).resolve().as_uri()


def html_image_tags_for_paths(image_paths: list[str]) -> str:
    return "".join(
        f'<br><img src="{html.escape(clipboard_file_uri(image_path), quote=True)}">'
        for image_path in image_paths
    )


def replace_html_image_sources(raw_html: str, image_paths: list[str]) -> tuple[str, int]:
    image_iter = iter(image_paths)
    replaced_count = 0

    def replace_match(match: re.Match[str]) -> str:
        nonlocal replaced_count
        try:
            image_path = next(image_iter)
        except StopIteration:
            return match.group(0)
        replaced_count += 1
        return (
            f"{match.group(1)}{match.group(2)}"
            f"{html.escape(clipboard_file_uri(image_path), quote=True)}"
            f"{match.group(4)}"
        )

    updated = re.sub(r"(?is)(<img\b[^>]*?\bsrc\s*=\s*)([\"'])(.*?)(\2)", replace_match, raw_html)
    return updated, replaced_count


def html_with_local_image_paths(
    plain_text: str,
    html_content: str | None,
    image_paths: list[str],
) -> str | None:
    valid_paths = [path for path in image_paths if Path(path).exists()]
    if not valid_paths:
        return html_content

    if html_content:
        updated, replaced_count = replace_html_image_sources(html_content, valid_paths)
        remaining_paths = valid_paths[replaced_count:]
        if remaining_paths:
            addition = html_image_tags_for_paths(remaining_paths)
            if HTML_FRAGMENT_END in updated:
                updated = updated.replace(HTML_FRAGMENT_END, addition + HTML_FRAGMENT_END, 1)
            elif re.search(r"(?is)</body>", updated):
                updated = re.sub(r"(?is)</body>", addition + "</body>", updated, count=1)
            else:
                updated += addition
        return updated

    fragment = html_escape_preserving_newlines(plain_text)
    if fragment:
        fragment += "<br>"
    fragment += html_image_tags_for_paths(valid_paths).lstrip("<br>")
    return build_clipboard_html(fragment)
