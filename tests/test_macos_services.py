from __future__ import annotations

import io
from pathlib import Path
import tempfile
import unittest

from PIL import Image

from platforms.macos.services import HTML_TYPES, MacPlatformServices, TEXT_TYPES


class FakePasteboardSource:
    def __init__(
        self,
        *,
        strings: dict[str, str] | None = None,
        data: dict[str, bytes] | None = None,
        plist: dict[str, object] | None = None,
        types: list[str] | None = None,
    ) -> None:
        self.strings = strings or {}
        self.data = data or {}
        self.plist = plist or {}
        self._types = types or list(self.strings) + list(self.data) + list(self.plist)

    def types(self):
        return self._types

    def stringForType_(self, type_name: str):
        return self.strings.get(type_name)

    def dataForType_(self, type_name: str):
        return self.data.get(type_name)

    def propertyListForType_(self, type_name: str):
        return self.plist.get(type_name)


class FakePasteboard(FakePasteboardSource):
    def __init__(self, **kwargs) -> None:
        super().__init__(**kwargs)
        self.items: list[FakePasteboardSource] = []
        self.cleared = False
        self.written_strings: dict[str, str] = {}
        self.written_data: dict[str, bytes] = {}
        self.written_plist: dict[str, object] = {}

    def pasteboardItems(self):
        return self.items

    def clearContents(self):
        self.cleared = True
        self.written_strings.clear()
        self.written_data.clear()
        self.written_plist.clear()

    def setString_forType_(self, value: str, type_name: str):
        self.written_strings[type_name] = value
        return True

    def setData_forType_(self, value, type_name: str):
        self.written_data[type_name] = bytes(value)
        return True

    def setPropertyList_forType_(self, value, type_name: str):
        self.written_plist[type_name] = value
        return True


def png_bytes(color: str = "red") -> bytes:
    output = io.BytesIO()
    Image.new("RGBA", (6, 7), color).save(output, format="PNG")
    return output.getvalue()


class MacServicesTests(unittest.TestCase):
    def service_for(self, pasteboard: FakePasteboard) -> MacPlatformServices:
        return MacPlatformServices(Path.cwd(), pasteboard_factory=lambda: pasteboard)

    def test_reads_rich_text_payload(self) -> None:
        pasteboard = FakePasteboard(
            strings={
                TEXT_TYPES[0]: "Hello",
                HTML_TYPES[0]: "<b>Hello</b>",
            }
        )
        capture = self.service_for(pasteboard).read_clipboard_capture()

        self.assertIsNotNone(capture)
        self.assertEqual(capture.type, "text")
        self.assertEqual(capture.plain_text, "Hello")
        self.assertTrue(capture.has_rich_text)

    def test_reads_png_image_item(self) -> None:
        pasteboard = FakePasteboard()
        pasteboard.items = [FakePasteboardSource(data={"public.png": png_bytes()})]

        capture = self.service_for(pasteboard).read_clipboard_capture()

        self.assertIsNotNone(capture)
        self.assertEqual(capture.type, "image")
        self.assertEqual(len(capture.images), 1)
        self.assertEqual(capture.images[0].size, (6, 7))

    def test_finder_image_files_become_image_collection(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            first = Path(temp_dir) / "a.png"
            second = Path(temp_dir) / "b.png"
            Image.new("RGBA", (3, 3), "red").save(first)
            Image.new("RGBA", (4, 4), "blue").save(second)
            pasteboard = FakePasteboard(plist={"NSFilenamesPboardType": [str(first), str(second)]})
            capture = self.service_for(pasteboard).read_clipboard_capture()

        self.assertIsNotNone(capture)
        self.assertEqual(capture.type, "image")
        self.assertEqual(len(capture.images), 2)

    def test_non_image_files_stay_other(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            document = Path(temp_dir) / "a.txt"
            document.write_text("hello", encoding="utf-8")
            pasteboard = FakePasteboard(plist={"NSFilenamesPboardType": [str(document)]})
            capture = self.service_for(pasteboard).read_clipboard_capture()

        self.assertIsNotNone(capture)
        self.assertEqual(capture.type, "other")
        self.assertEqual(capture.other_kind, "files")

    def test_public_file_url_is_read_as_file(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            document = Path(temp_dir) / "a.txt"
            document.write_text("hello", encoding="utf-8")
            pasteboard = FakePasteboard(strings={"public.file-url": document.resolve().as_uri()})
            capture = self.service_for(pasteboard).read_clipboard_capture()

        self.assertIsNotNone(capture)
        self.assertEqual(capture.type, "other")
        self.assertIn("a.txt", capture.summary)

    def test_text_plus_image_becomes_mixed(self) -> None:
        pasteboard = FakePasteboard(strings={TEXT_TYPES[0]: "caption"})
        pasteboard.items = [FakePasteboardSource(data={"public.png": png_bytes("blue")})]
        capture = self.service_for(pasteboard).read_clipboard_capture()

        self.assertIsNotNone(capture)
        self.assertEqual(capture.type, "mixed")
        self.assertEqual(capture.plain_text, "caption")
        self.assertEqual(len(capture.images), 1)

    def test_writes_rich_text_and_image_paths(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            image_path = Path(temp_dir) / "a.png"
            Image.new("RGBA", (3, 3), "red").save(image_path)
            pasteboard = FakePasteboard()
            self.service_for(pasteboard).set_clipboard_rich_text_and_images(
                "caption",
                None,
                None,
                [str(image_path)],
            )

        self.assertTrue(pasteboard.cleared)
        self.assertEqual(pasteboard.written_strings[TEXT_TYPES[0]], "caption")
        self.assertIn(HTML_TYPES[0], pasteboard.written_strings)
        self.assertIn("NSFilenamesPboardType", pasteboard.written_plist)


if __name__ == "__main__":
    unittest.main()
