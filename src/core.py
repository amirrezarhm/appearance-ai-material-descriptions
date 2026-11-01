"""Deterministic, human-supervised vocabulary classification for 90 images."""
from __future__ import annotations

from collections import Counter
from contextlib import contextmanager
from datetime import datetime, timezone
from hashlib import sha256
from pathlib import Path
from typing import Any
import json
import re
import sqlite3

from openpyxl import load_workbook

CATEGORIES = ("Color", "Transparency", "Glossiness", "Texture", "Shape", "Material", "Other")
APPEARANCE = ("Color", "Transparency", "Glossiness", "Texture")
MODELS = ("gemini", "qwen", "llama", "mistral")
SHEET_NAME = "Word_Level_1800"
REQUIRED_COLS = ("row_id", "image_name", "model_key", "position", "word")
EXPECTED_IMAGES = 90
EXPECTED_WORDS = 1800


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def normalized_key(text: str) -> str:
    """Normalize casing/spacing only; never silently collapse synonyms."""
    return " ".join(str(text or "").strip().lower().split())


def safe_canonical(text: str) -> str:
    word = normalized_key(text)
    if not word or len(word) > 120 or "\n" in word:
        raise ValueError("Canonical descriptor must be 1 to 120 characters.")
    return word


def occurrence_id(row_id: int, model: str, position: int) -> str:
    return f"{row_id:03d}:{model}:{position}"


def _find_header(ws) -> tuple[int, dict[str, int]]:
    for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 20), values_only=False):
        values = {str(c.value).strip().lower(): idx for idx, c in enumerate(row) if c.value is not None}
        if all(col in values for col in REQUIRED_COLS):
            return row[0].row, values
    raise ValueError(f"Cannot locate required columns in sheet {ws.title}: {REQUIRED_COLS}")


def load_word_rows(excel_path: Path) -> tuple[list[dict[str, Any]], str]:
    """Read authoritative word-level rows with strict 90x4x5 validation."""
    excel_path = Path(excel_path)
    if not excel_path.is_file():
        raise FileNotFoundError(f"Excel input not found: {excel_path}")
    excel_hash = sha256(excel_path.read_bytes()).hexdigest()
    wb = load_workbook(excel_path, read_only=True, data_only=True)
    try:
        if SHEET_NAME not in wb.sheetnames:
            raise ValueError(f"Missing sheet '{SHEET_NAME}'. Found: {wb.sheetnames}")
        ws = wb[SHEET_NAME]
        header_row, headers = _find_header(ws)
        result: list[dict[str, Any]] = []
        seen: set[str] = set()
        for cells in ws.iter_rows(min_row=header_row + 1, values_only=True):
            if all(v is None for v in cells):
                continue
            data = {col: cells[idx] if idx < len(cells) else None for col, idx in headers.items()}
            if not any(data.get(k) is not None for k in REQUIRED_COLS):
                continue
            try:
                image_id = int(data["row_id"])
                position = int(data["position"])
            except (ValueError, TypeError) as exc:
                raise ValueError(f"Invalid image ID / position at Excel row {len(result) + header_row + 1}") from exc
            model = normalized_key(data["model_key"])
            word = str(data["word"] or "").strip()
            image_name = str(data["image_name"] or "").strip()
            if image_id not in range(1, 91) or model not in MODELS or position not in range(1, 6):
                raise ValueError(f"Invalid key for image {image_id}, {model}, position {position}")
            if not word or not image_name:
                raise ValueError(f"Missing word/image_name for image {image_id}, {model}, position {position}")
            image_stem = Path(image_name).stem
            if not image_stem.isdecimal() or int(image_stem) != image_id:
                raise ValueError(f"Image name and ID mismatch: {image_name} vs {image_id}")
            key = occurrence_id(image_id, model, position)
            if key in seen:
                raise ValueError(f"Duplicate word position: {key}")
            seen.add(key)
            result.append({
                "occurrence_id": key,
                "image_id": image_id,
                "image_name": image_name,
                "model": model,
                "position": position,
                "raw_word": word,
                "word_key": normalized_key(word),
                "five_words": str(data.get("five_words") or ""),
                "image_md5": str(data.get("image_md5") or ""),
            })
        if len(result) != EXPECTED_WORDS:
            raise ValueError(f"Expected 1800 descriptors; found {len(result)}. Original data not modified.")
        totals = Counter((r["image_id"], r["model"]) for r in result)
        if len(totals) != 360 or any(v != 5 for v in totals.values()):
            raise ValueError("Each of 90 images must have exactly 5 descriptors from each of 4 models.")
        result.sort(key=lambda r: (r["image_id"], MODELS.index(r["model"]), r["position"]))
        return result, excel_hash
    finally:
        wb.close()


def connect(db_path: Path) -> sqlite3.Connection:
    conn = sqlite3.connect(str(db_path), timeout=30)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys=ON")
    conn.execute("PRAGMA busy_timeout=30000")
    return conn


@contextmanager
def db_session(db_path: Path):
    """Transaction scope that also closes its SQLite connection promptly."""
    conn = connect(db_path)
    try:
        with conn:
            yield conn
    finally:
        conn.close()


def initialize(excel_path: Path, db_path: Path) -> dict[str, Any]:
    """Idempotent first-time setup. Refuse silently changed source datasets."""
    rows, fingerprint = load_word_rows(excel_path)
    db_path = Path(db_path)
    db_path.parent.mkdir(parents=True, exist_ok=True)
    with db_session(db_path) as conn:
        conn.executescript("""
        CREATE TABLE IF NOT EXISTS metadata (key TEXT PRIMARY KEY, value TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS occurrences (
            occurrence_id TEXT PRIMARY KEY, image_id INTEGER NOT NULL,
            image_name TEXT NOT NULL, model TEXT NOT NULL,
            position INTEGER NOT NULL, raw_word TEXT NOT NULL,
            word_key TEXT NOT NULL, five_words TEXT NOT NULL,
            image_md5 TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS dictionary (
            word_key TEXT PRIMARY KEY,
            canonical_word TEXT NOT NULL,
            category TEXT,
            updated_at TEXT);
        CREATE TABLE IF NOT EXISTS overrides (
            occurrence_id TEXT PRIMARY KEY REFERENCES occurrences(occurrence_id),
            category TEXT NOT NULL, updated_at TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS decisions (
            decision_id INTEGER PRIMARY KEY AUTOINCREMENT,
            recorded_at TEXT NOT NULL, scope TEXT NOT NULL,
            occurrence_id TEXT, word_key TEXT NOT NULL,
            previous_value TEXT, new_value TEXT NOT NULL,
            note TEXT NOT NULL DEFAULT '');
        CREATE INDEX IF NOT EXISTS idx_occ_word ON occurrences(word_key);
        CREATE INDEX IF NOT EXISTS idx_occ_image_model ON occurrences(image_id, model);
        """)
        past = conn.execute("SELECT value FROM metadata WHERE key='input_sha256'").fetchone()
        if past and past["value"] != fingerprint:
            raise RuntimeError(
                "The input Excel has changed since initialization. "
                "Existing decisions were preserved. Back up the database and "
                "review changes before migrating to a new input file."
            )
        if not past:
            conn.execute("INSERT INTO metadata(key,value) VALUES (?,?)", ("input_sha256", fingerprint))
            conn.execute("INSERT INTO metadata(key,value) VALUES (?,?)", ("created_at", utc_now()))
            conn.executemany("""
                INSERT INTO occurrences
                (occurrence_id,image_id,image_name,model,position,raw_word,word_key,five_words,image_md5)
                VALUES (:occurrence_id,:image_id,:image_name,:model,:position,:raw_word,:word_key,:five_words,:image_md5)
            """, rows)
            conn.executemany(
                "INSERT OR IGNORE INTO dictionary(word_key,canonical_word) VALUES (?,?)",
                [(key, key) for key in sorted(set(r["word_key"] for r in rows))]
            )
        else:
            db_count = conn.execute("SELECT COUNT(*) FROM occurrences").fetchone()[0]
            if db_count != len(rows):
                raise RuntimeError("Database is incomplete; database and original Excel were not altered.")
        return {
            "words": conn.execute("SELECT COUNT(*) FROM occurrences").fetchone()[0],
            "unique_words": conn.execute("SELECT COUNT(*) FROM dictionary").fetchone()[0],
            "images": conn.execute("SELECT COUNT(DISTINCT image_id) FROM occurrences").fetchone()[0],
            "source_sha256": fingerprint,
            "database": str(db_path),
        }


def is_initialized(db_path: Path) -> bool:
    return Path(db_path).is_file()


def _validate_category(category: str):
    if category not in CATEGORIES:
        raise ValueError(f"Unknown category: {category}. Expected one of {CATEGORIES}")


def save_decision(db_path: Path, *, occurrence_key: str, category: str,
                  scope: str, canonical_word: str | None = None, note: str = ""):
    """Scope: all occurrences of a word, or this occurrence only."""
    _validate_category(category)
    if scope not in ("all", "occurrence"):
        raise ValueError("scope must be 'all' or 'occurrence'")
    note = str(note or "")[:500]
    with db_session(db_path) as conn:
        row = conn.execute(
            "SELECT word_key FROM occurrences WHERE occurrence_id=?", (occurrence_key,)
        ).fetchone()
        if row is None:
            raise ValueError(f"Unknown occurrence: {occurrence_key}")
        word_key = row["word_key"]
        old = None
        if scope == "all":
            old_row = conn.execute("SELECT category, canonical_word FROM dictionary WHERE word_key=?", (word_key,)).fetchone()
            old = old_row["category"] if old_row else None
            canonical = safe_canonical(canonical_word) if canonical_word is not None else old_row["canonical_word"]
            conn.execute("UPDATE dictionary SET category=?, canonical_word=?, updated_at=? WHERE word_key=?",
                         (category, canonical, utc_now(), word_key))
        else:
            old_row = conn.execute("SELECT category FROM overrides WHERE occurrence_id=?", (occurrence_key,)).fetchone()
            old = old_row["category"] if old_row else None
            conn.execute("""INSERT INTO overrides(occurrence_id,category,updated_at) VALUES (?,?,?)
                     ON CONFLICT(occurrence_id) DO UPDATE SET category=excluded.category, updated_at=excluded.updated_at""",
                     (occurrence_key, category, utc_now()))
        conn.execute("""INSERT INTO decisions(recorded_at,scope,occurrence_id,word_key,previous_value,new_value,note)
                        VALUES (?,?,?,?,?,?,?)""",
                     (utc_now(), scope, occurrence_key, word_key, old, category, note))


def update_canonical(db_path: Path, word_key: str, canonical_word: str):
    canonical = safe_canonical(canonical_word)
    with db_session(db_path) as conn:
        old = conn.execute("SELECT canonical_word FROM dictionary WHERE word_key=?", (word_key,)).fetchone()
        if old is None:
            raise KeyError(word_key)
        conn.execute("UPDATE dictionary SET canonical_word=?,updated_at=? WHERE word_key=?",
                     (canonical, utc_now(), word_key))
        conn.execute("""INSERT INTO decisions(recorded_at,scope,occurrence_id,word_key,previous_value,new_value,note)
                        VALUES (?,?,?,?,?,?,?)""",
                     (utc_now(), "canonical", None, word_key, old["canonical_word"], canonical, ""))


def clear_occurrence_override(db_path: Path, occurrence_key: str):
    with db_session(db_path) as conn:
        row = conn.execute("SELECT word_key FROM occurrences WHERE occurrence_id=?", (occurrence_key,)).fetchone()
        if row is None:
            raise KeyError(occurrence_key)
        prev = conn.execute("SELECT category FROM overrides WHERE occurrence_id=?", (occurrence_key,)).fetchone()
        conn.execute("DELETE FROM overrides WHERE occurrence_id=?", (occurrence_key,))
        if prev:
            conn.execute("""INSERT INTO decisions(recorded_at,scope,occurrence_id,word_key,previous_value,new_value,note)
                            VALUES (?,?,?,?,?,?,?)""",
                         (utc_now(), "clear_override", occurrence_key, row["word_key"], prev["category"], "INHERIT", ""))


def all_occurrences(db_path: Path) -> list[dict[str, Any]]:
    with db_session(db_path) as conn:
        rows = conn.execute("""
            SELECT o.occurrence_id,o.image_id,o.image_name,o.model,o.position,
                   o.raw_word,o.word_key,o.five_words,o.image_md5,
                   d.canonical_word, d.category AS global_category,
                   x.category AS override_category,
                   COALESCE(x.category,d.category,'Unassigned') AS category,
                   CASE WHEN x.category IS NOT NULL THEN 'Occurrence'
                        WHEN d.category IS NOT NULL THEN 'All' ELSE 'Pending' END AS source
            FROM occurrences o
            JOIN dictionary d ON d.word_key=o.word_key
            LEFT JOIN overrides x ON x.occurrence_id=o.occurrence_id
            ORDER BY o.image_id, CASE o.model WHEN 'gemini' THEN 0 WHEN 'qwen' THEN 1
                        WHEN 'llama' THEN 2 ELSE 3 END, o.position
        """).fetchall()
        return [dict(row) for row in rows]


def get_dictionary(db_path: Path) -> list[dict[str, Any]]:
    with db_session(db_path) as conn:
        rows = conn.execute("""
            SELECT d.word_key, d.canonical_word, d.category,
                   COUNT(o.occurrence_id) AS frequency,
                   COUNT(DISTINCT o.image_id) AS image_count
            FROM dictionary d JOIN occurrences o ON o.word_key=d.word_key
            GROUP BY d.word_key, d.canonical_word, d.category
            ORDER BY frequency DESC,d.word_key ASC
        """).fetchall()
        return [dict(r) for r in rows]


def get_decision_log(db_path: Path) -> list[dict[str, Any]]:
    with db_session(db_path) as conn:
        return [dict(x) for x in conn.execute("SELECT * FROM decisions ORDER BY decision_id")]


def progress(rows: list[dict[str, Any]]) -> dict[str, int]:
    assigned = sum(r["category"] != "Unassigned" for r in rows)
    return {"assigned": assigned, "total": len(rows), "remaining": len(rows) - assigned}


def image_coverage(rows: list[dict[str, Any]]) -> tuple[list[dict], list[dict]]:
    """One record per image x model, and per image aggregated across models.

    Category coverage only definitive if all words in that unit are classified.
    """
    by_pair: dict[tuple[int, str], list[dict]] = {}
    by_image: dict[int, list[dict]] = {}
    for r in rows:
        by_pair.setdefault((r["image_id"], r["model"]), []).append(r)
        by_image.setdefault(r["image_id"], []).append(r)

    def summarize(items: list[dict], base: dict) -> dict:
        pending = sum(r["category"] == "Unassigned" for r in items)
        result = dict(base)
        result["descriptors"] = len(items)
        result["unassigned"] = pending
        result["complete"] = pending == 0
        present = set(r["category"] for r in items if r["category"] != "Unassigned")
        result["known_appearance_categories"] = len(set(APPEARANCE) & present)
        result["appearance_count"] = len(set(APPEARANCE) & present) if not pending else None
        for cat in CATEGORIES:
            result[cat] = ", ".join(r["raw_word"] for r in items if r["category"] == cat)
            result[f"has_{cat}"] = (cat in present) if not pending else None
        return result

    pairs = [summarize(items, {"image_id": key[0], "image_name": items[0]["image_name"], "model": key[1]})
             for key, items in sorted(by_pair.items(), key=lambda pair: (pair[0][0], MODELS.index(pair[0][1])))]
    images = [summarize(items, {"image_id": key, "image_name": items[0]["image_name"]})
              for key, items in sorted(by_image.items())]
    return pairs, images


def coverage_distribution(items: list[dict]) -> list[dict]:
    total = len(items)
    complete = sum(bool(r["complete"]) for r in items)
    result = []
    for n in (4, 3, 2, 1, 0):
        count = sum(r["complete"] and r["appearance_count"] == n for r in items)
        result.append({
            "appearance_categories": n,
            "image_count": count,
            "percent_all_images": round(100 * count / total, 3) if total else 0,
            "percent_classified_images": round(100 * count / complete, 3) if complete else None,
            "total_images": total,
            "fully_classified_images": complete,
            "incomplete_images": total - complete,
        })
    return result
