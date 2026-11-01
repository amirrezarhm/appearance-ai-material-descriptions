"""Offline tests: dataset integrity, scope precedence, coverage, and reporting."""
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from openpyxl import Workbook

from src.core import (APPEARANCE, MODELS, all_occurrences, clear_occurrence_override,
                      coverage_distribution, image_coverage, initialize, load_word_rows,
                      occurrence_id, progress, save_decision, update_canonical)
from src.reports import create_reports


def example_excel(path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Word_Level_1800"
    ws.append(["Word Level 1800"])
    ws.append(["Generated sample for tests"])
    ws.append(["row_id", "image_name", "model_key", "position", "word", "five_words", "image_md5"])
    samples = ("clear", "blue", "smooth", "round", "glass")
    for image in range(1, 91):
        for model in MODELS:
            for pos, word in enumerate(samples, 1):
                ws.append([image, f"{image:03d}.jfif", model, pos, word,
                           ", ".join(samples), f"mock-md5-{image:03d}"])
    wb.save(path)


class PipelineTests(unittest.TestCase):
    def setUp(self):
        self.temp = TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.excel = self.root / "source.xlsx"
        self.db = self.root / "classification.sqlite3"
        example_excel(self.excel)

    def init(self):
        return initialize(self.excel, self.db)

    def test_source_exactly_1800_words(self):
        rows, digest = load_word_rows(self.excel)
        self.assertEqual(len(rows), 1800)
        self.assertEqual(len(digest), 64)
        self.assertEqual({r["image_id"] for r in rows}, set(range(1, 91)))

    def test_idempotent_init(self):
        self.assertEqual(self.init()["unique_words"], 5)
        self.assertEqual(self.init()["words"], 1800)
        self.assertEqual(len(all_occurrences(self.db)), 1800)

    def test_global_scope_and_occurrence_exception(self):
        self.init()
        save_decision(self.db, occurrence_key="001:gemini:1", category="Transparency", scope="all")
        words = [x for x in all_occurrences(self.db) if x["word_key"] == "clear"]
        self.assertEqual(len(words), 360)
        self.assertEqual({x["category"] for x in words}, {"Transparency"})
        save_decision(self.db, occurrence_key="001:qwen:1", category="Other", scope="occurrence")
        current = {x["occurrence_id"]: x for x in all_occurrences(self.db)}
        self.assertEqual(current["001:qwen:1"]["category"], "Other")
        self.assertEqual(current["001:gemini:1"]["category"], "Transparency")
        clear_occurrence_override(self.db, "001:qwen:1")
        current = {x["occurrence_id"]: x for x in all_occurrences(self.db)}
        self.assertEqual(current["001:qwen:1"]["category"], "Transparency")

    def test_incomplete_not_counted_as_no_appearance(self):
        self.init()
        save_decision(self.db, occurrence_key="001:gemini:1", category="Transparency", scope="all")
        pairs, images = image_coverage(all_occurrences(self.db))
        self.assertEqual(len(pairs), 360)
        self.assertEqual(len(images), 90)
        self.assertIsNone(images[0]["appearance_count"])
        self.assertFalse(images[0]["complete"])
        dist = coverage_distribution(images)
        self.assertEqual(dist[0]["fully_classified_images"], 0)
        self.assertIsNone(dist[0]["percent_classified_images"])

    def test_4_category_image_coverage(self):
        self.init()
        classes = {1: "Transparency", 2: "Color", 3: "Texture", 4: "Shape", 5: "Glossiness"}
        for pos, category in classes.items():
            save_decision(self.db, occurrence_key=f"001:gemini:{pos}", category=category, scope="all")
        pairs, images = image_coverage(all_occurrences(self.db))
        self.assertEqual(images[0]["appearance_count"], 4)
        self.assertEqual(pairs[0]["appearance_count"], 4)
        self.assertEqual(coverage_distribution(images)[0]["image_count"], 90)
        self.assertEqual(progress(all_occurrences(self.db))["remaining"], 0)

    def test_can_change_canonical_without_losing_raw(self):
        self.init()
        update_canonical(self.db, "clear", "clearer")
        row = next(x for x in all_occurrences(self.db) if x["word_key"] == "clear")
        self.assertEqual(row["raw_word"], "clear")
        self.assertEqual(row["canonical_word"], "clearer")

    def test_change_input_refused_without_erasing_labels(self):
        self.init()
        save_decision(self.db, occurrence_key="001:gemini:1", category="Transparency", scope="all")
        from openpyxl import load_workbook
        wb = load_workbook(self.excel)
        ws = wb["Word_Level_1800"]
        ws["E4"] = "altered"
        wb.save(self.excel)
        with self.assertRaisesRegex(RuntimeError, "changed"):
            self.init()
        row = next(x for x in all_occurrences(self.db) if x["occurrence_id"] == "001:gemini:1")
        self.assertEqual(row["category"], "Transparency")

    def test_report_generates_excel_csv_and_png(self):
        self.init()
        save_decision(self.db, occurrence_key="001:gemini:1", category="Transparency", scope="all")
        result = create_reports(self.db, self.root / "output")
        folder = Path(result["output"])
        self.assertTrue(Path(result["excel"]).exists())
        self.assertEqual(result["charts"], 14)
        self.assertTrue((folder / "csv/Image_Summary_90.csv").exists())
        self.assertTrue((folder / "database_snapshot.sqlite3").exists())
        from openpyxl import load_workbook
        wb = load_workbook(result["excel"], read_only=True)
        self.assertEqual(wb["Word_Instances_1800"].max_row, 1801)
        self.assertEqual(wb["Image_Summary_90"].max_row, 91)
        wb.close()


if __name__ == "__main__":
    unittest.main()
