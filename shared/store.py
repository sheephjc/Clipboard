from __future__ import annotations

import contextlib
import datetime as dt
import json
from pathlib import Path
import sqlite3

from PIL import Image

from shared.models import ClipboardCapture, ClipboardEntry, IMAGE_ENTRY_TYPES
from shared.rich_text import hash_image, json_dumps


class ClipboardStore:
    def __init__(self, data_dir: Path, max_history: int | None = 200):
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
                image_paths_json TEXT,
                other_kind TEXT,
                other_payload_json TEXT,
                is_favorite INTEGER DEFAULT 0
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
            "image_paths_json": "TEXT",
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
        rows = self.conn.execute(
            """
            SELECT id, image_path
            FROM entries
            WHERE image_path IS NOT NULL
              AND COALESCE(image_paths_json, '') = ''
            """
        ).fetchall()
        for row in rows:
            self.conn.execute(
                "UPDATE entries SET image_paths_json = ? WHERE id = ?",
                (json_dumps([row["image_path"]]), row["id"]),
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
            image_paths_json=row["image_paths_json"],
            other_kind=row["other_kind"],
            other_payload_json=row["other_payload_json"],
            is_favorite=bool(row["is_favorite"]) if "is_favorite" in row.keys() else False,
        )

    def load_entries(self) -> list[ClipboardEntry]:
        if self.max_history is None:
            rows = self.conn.execute(
                """
                SELECT id, type, summary, created_at, content_hash,
                       COALESCE(plain_text, text_content) AS plain_text,
                       html_content, rtf_content, source_formats_json, COALESCE(has_rich_text, 0) AS has_rich_text,
                       image_path, image_paths_json, other_kind, other_payload_json, COALESCE(is_favorite, 0) AS is_favorite
                FROM entries
                ORDER BY id DESC
                """
            ).fetchall()
            return [self._row_to_entry(row) for row in rows]

        rows = self.conn.execute(
            """
            SELECT id, type, summary, created_at, content_hash,
                   COALESCE(plain_text, text_content) AS plain_text,
                   html_content, rtf_content, source_formats_json, COALESCE(has_rich_text, 0) AS has_rich_text,
                   image_path, image_paths_json, other_kind, other_payload_json, COALESCE(is_favorite, 0) AS is_favorite
            FROM entries
            WHERE COALESCE(is_favorite, 0) = 1
               OR id IN (
                   SELECT id
                   FROM entries
                   WHERE COALESCE(is_favorite, 0) = 0
                   ORDER BY id DESC
                   LIMIT ?
               )
            ORDER BY id DESC
            """,
            (max(self.max_history, 0),),
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
        image_paths: list[str] = []
        if capture.type in IMAGE_ENTRY_TYPES and (capture.image is not None or capture.images):
            images = list(capture.images) or [capture.image]
            for image in images:
                image_paths.append(self._save_image(image, hash_image(image)))
            if image_paths:
                image_path = image_paths[0]
        image_paths_json = json_dumps(image_paths) if image_paths else None

        cursor = self.conn.execute(
            """
            INSERT INTO entries (
                type, summary, created_at, content_hash,
                plain_text, text_content, html_content, rtf_content, source_formats_json, has_rich_text,
                image_path, image_paths_json, other_kind, other_payload_json, is_favorite
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
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
                image_paths_json,
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
            image_paths_json=image_paths_json,
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
            (1 if is_favorite else 0, entry_id),
        )
        self.conn.commit()

    def clear_entries(self) -> None:
        self.conn.execute("DELETE FROM entries")
        self.conn.commit()
        self.cleanup_unused_images()

    def clear_history_entries(self) -> int:
        cursor = self.conn.execute(
            "DELETE FROM entries WHERE COALESCE(is_favorite, 0) = 0"
        )
        deleted_count = cursor.rowcount if cursor.rowcount is not None else 0
        self.conn.commit()
        if deleted_count:
            self.cleanup_unused_images()
        return max(deleted_count, 0)

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
        if self.max_history is None:
            return

        rows = self.conn.execute(
            "SELECT id FROM entries WHERE COALESCE(is_favorite, 0) = 0 ORDER BY id DESC LIMIT -1 OFFSET ?",
            (max(self.max_history, 0),),
        ).fetchall()
        if not rows:
            return

        ids = [row["id"] for row in rows]
        placeholders = ",".join(["?"] * len(ids))
        self.conn.execute(f"DELETE FROM entries WHERE id IN ({placeholders})", ids)
        self.conn.commit()
        self.cleanup_unused_images()

    def cleanup_unused_images(self) -> None:
        used_paths: set[Path] = set()
        for row in self.conn.execute(
            "SELECT image_path, image_paths_json FROM entries WHERE image_path IS NOT NULL OR image_paths_json IS NOT NULL"
        ).fetchall():
            if row["image_path"]:
                used_paths.add(Path(row["image_path"]))
            if row["image_paths_json"]:
                with contextlib.suppress(json.JSONDecodeError, TypeError):
                    decoded = json.loads(row["image_paths_json"])
                    if isinstance(decoded, list):
                        used_paths.update(Path(path) for path in decoded if path)
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

