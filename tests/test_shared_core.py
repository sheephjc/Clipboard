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
    def _add_text(
        self,
        store: ClipboardStore,
        text: str,
        is_favorite: bool = False,
    ):
        return store.add_capture(
            ClipboardCapture(
                type="text",
                content_hash=hash_rich_text(text),
                summary=text,
                plain_text=text,
                is_favorite=is_favorite,
            )
        )

    def test_store_keeps_favorites_when_pruning(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            store = ClipboardStore(Path(temp_dir), max_history=1)
            first = self._add_text(store, "first", is_favorite=True)
            self._add_text(store, "second")
            rows = store.conn.execute("SELECT id, is_favorite FROM entries ORDER BY id").fetchall()
            entries = [(row["id"], bool(row["is_favorite"])) for row in rows]
            store.close()

        self.assertEqual({entry_id for entry_id, _favorite in entries}, {first.id, first.id + 1})
        self.assertTrue(next(favorite for entry_id, favorite in entries if entry_id == first.id))

    def test_load_entries_includes_all_favorites_plus_limited_nonfavorites(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            store = ClipboardStore(Path(temp_dir), max_history=1)
            favorite_a = self._add_text(store, "favorite-a", is_favorite=True)
            self._add_text(store, "plain-a")
            favorite_b = self._add_text(store, "favorite-b", is_favorite=True)
            plain_b = self._add_text(store, "plain-b")
            loaded = store.load_entries()
            store.close()

        self.assertEqual([entry.id for entry in loaded], [plain_b.id, favorite_b.id, favorite_a.id])

    def test_unlimited_history_does_not_prune_or_limit_loads(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            store = ClipboardStore(Path(temp_dir), max_history=None)
            first = self._add_text(store, "first")
            second = self._add_text(store, "second")
            third = self._add_text(store, "third")
            loaded = store.load_entries()
            rows = store.conn.execute("SELECT id FROM entries ORDER BY id").fetchall()
            store.close()

        self.assertEqual([row["id"] for row in rows], [first.id, second.id, third.id])
        self.assertEqual([entry.id for entry in loaded], [third.id, second.id, first.id])

    def test_unfavoriting_reapplies_history_limit(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            store = ClipboardStore(Path(temp_dir), max_history=1)
            favorite = self._add_text(store, "favorite", is_favorite=True)
            plain = self._add_text(store, "plain")
            store.update_favorite(favorite.id, False)
            store.prune_to_limit()
            rows = store.conn.execute("SELECT id FROM entries ORDER BY id").fetchall()
            store.close()

        self.assertEqual([row["id"] for row in rows], [plain.id])

    def test_clear_history_entries_keeps_favorites(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            store = ClipboardStore(Path(temp_dir), max_history=10)
            favorite = self._add_text(store, "favorite", is_favorite=True)
            plain_a = self._add_text(store, "plain-a")
            plain_b = self._add_text(store, "plain-b")
            deleted_count = store.clear_history_entries()
            rows = store.conn.execute("SELECT id FROM entries ORDER BY id").fetchall()
            store.close()

        self.assertEqual(deleted_count, 2)
        self.assertEqual([row["id"] for row in rows], [favorite.id])
        self.assertNotIn(plain_a.id, [row["id"] for row in rows])
        self.assertNotIn(plain_b.id, [row["id"] for row in rows])

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
