from __future__ import annotations

from pathlib import Path
import tempfile
import unittest

from PIL import Image

from shared.models import ClipboardCapture
from shared.rich_text import (
    hash_images,
    hash_mixed_content,
    hash_rich_text,
    html_with_local_image_paths,
)
from shared.store import ClipboardStore


class SharedCoreTests(unittest.TestCase):
    def test_store_keeps_favorites_when_pruning(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            store = ClipboardStore(Path(temp_dir), max_history=1)
            first = store.add_capture(
                ClipboardCapture(
                    type="text",
                    content_hash=hash_rich_text("first"),
                    summary="first",
                    plain_text="first",
                    is_favorite=True,
                )
            )
            store.add_capture(
                ClipboardCapture(
                    type="text",
                    content_hash=hash_rich_text("second"),
                    summary="second",
                    plain_text="second",
                )
            )
            rows = store.conn.execute("SELECT id, is_favorite FROM entries ORDER BY id").fetchall()
            entries = [(row["id"], bool(row["is_favorite"])) for row in rows]
            store.close()

        self.assertEqual({entry_id for entry_id, _favorite in entries}, {first.id, first.id + 1})
        self.assertTrue(next(favorite for entry_id, favorite in entries if entry_id == first.id))

    def test_store_saves_and_cleans_multiple_images(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            store = ClipboardStore(Path(temp_dir), max_history=5)
            images = [
                Image.new("RGBA", (8, 8), "red"),
                Image.new("RGBA", (9, 9), "blue"),
            ]
            entry = store.add_capture(
                ClipboardCapture(
                    type="image",
                    content_hash=hash_images(images),
                    summary="图片 2 张",
                    image=images[0],
                    images=images,
                )
            )
            self.assertIsNotNone(entry.image_paths_json)
            for image_path in store.image_dir.glob("*.png"):
                self.assertTrue(image_path.exists())

            store.delete_entry(entry.id)
            remaining = list(store.image_dir.glob("*.png"))
            store.close()

        self.assertEqual(remaining, [])

    def test_rich_text_and_mixed_hashes_include_images(self) -> None:
        image_a = Image.new("RGBA", (4, 4), "red")
        image_b = Image.new("RGBA", (4, 4), "blue")
        base = hash_mixed_content("hello", None, None, [image_a])
        changed = hash_mixed_content("hello", None, None, [image_b])
        self.assertNotEqual(base, changed)

    def test_html_with_local_image_paths_adds_file_uri(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            image_path = Path(temp_dir) / "sample.png"
            Image.new("RGBA", (4, 4), "green").save(image_path)
            html = html_with_local_image_paths("hello", None, [str(image_path)])

        self.assertIsNotNone(html)
        self.assertIn("file://", html or "")


if __name__ == "__main__":
    unittest.main()
