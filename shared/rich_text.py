from __future__ import annotations

import hashlib
import html
import json
import re
from html.parser import HTMLParser

from PIL import Image

HTML_FRAGMENT_START = "<!--StartFragment-->"
HTML_FRAGMENT_END = "<!--EndFragment-->"


def json_dumps(payload: object) -> str:
    return json.dumps(payload, ensure_ascii=False, separators=(",", ":"))


def hash_rich_text(
    plain_text: str,
    html_content: str | None = None,
    rtf_content: str | None = None,
) -> str:
    payload = json_dumps(
        {
            "plain": plain_text,
            "html": html_content or "",
            "rtf": rtf_content or "",
        }
    )
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def hash_image(image: Image.Image) -> str:
    normalized = image.convert("RGBA")
    return hashlib.sha256(normalized.tobytes() + repr(normalized.size).encode("utf-8")).hexdigest()


def hash_other(kind: str, payload: object) -> str:
    return hashlib.sha256(f"{kind}:{json_dumps(payload)}".encode("utf-8")).hexdigest()


def summarize_text(text: str) -> str:
    single_line = re.sub(r"\s+", " ", text).strip()
    return single_line[:72] + ("..." if len(single_line) > 72 else "")


def summarize_files(paths: list[str]) -> str:
    if len(paths) == 1:
        return paths[0]
    return f"{len(paths)} files: {paths[0]}"


def summarize_formats(format_names: list[str]) -> str:
    if not format_names:
        return "Unknown clipboard content"
    return ", ".join(format_names[:4]) + ("..." if len(format_names) > 4 else "")


def build_clipboard_html(fragment_html: str) -> str:
    html_doc = f"<html><body>{HTML_FRAGMENT_START}{fragment_html}{HTML_FRAGMENT_END}</body></html>"
    prefix_template = (
        "Version:0.9\r\n"
        "StartHTML:{start_html:010d}\r\n"
        "EndHTML:{end_html:010d}\r\n"
        "StartFragment:{start_fragment:010d}\r\n"
        "EndFragment:{end_fragment:010d}\r\n"
    )
    empty_prefix = prefix_template.format(start_html=0, end_html=0, start_fragment=0, end_fragment=0)
    start_html = len(empty_prefix.encode("utf-8"))
    start_fragment = start_html + html_doc.index(HTML_FRAGMENT_START) + len(HTML_FRAGMENT_START)
    end_fragment = start_html + html_doc.index(HTML_FRAGMENT_END)
    end_html = start_html + len(html_doc.encode("utf-8"))
    prefix = prefix_template.format(
        start_html=start_html,
        end_html=end_html,
        start_fragment=start_fragment,
        end_fragment=end_fragment,
    )
    return prefix + html_doc


def extract_html_fragment(raw_html: str | None) -> str | None:
    if not raw_html:
        return None
    start_marker = raw_html.find(HTML_FRAGMENT_START)
    end_marker = raw_html.find(HTML_FRAGMENT_END)
    if start_marker != -1 and end_marker != -1 and end_marker > start_marker:
        start = start_marker + len(HTML_FRAGMENT_START)
        return raw_html[start:end_marker]
    return raw_html


class _PlainTextHtmlParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.parts: list[str] = []

    def handle_data(self, data: str) -> None:
        self.parts.append(data)

    def handle_starttag(self, tag: str, attrs) -> None:
        if tag in {"br", "p", "div", "li"}:
            self.parts.append("\n")


def plain_text_from_html(raw_html: str | None) -> str:
    fragment = extract_html_fragment(raw_html)
    if not fragment:
        return ""
    parser = _PlainTextHtmlParser()
    parser.feed(fragment)
    return html.unescape("".join(parser.parts)).strip()


def plain_text_from_rtf(raw_rtf: str | None) -> str:
    if not raw_rtf:
        return ""
    text = re.sub(r"\\'[0-9a-fA-F]{2}", "", raw_rtf)
    text = re.sub(r"\\[a-zA-Z]+\d* ?", "", text)
    text = text.replace("{", "").replace("}", "")
    return text.strip()

