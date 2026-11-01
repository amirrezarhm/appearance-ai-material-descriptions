"""Repeatable local reports. Original Excel and user labels remain unchanged."""
from __future__ import annotations
from collections import Counter, defaultdict
from datetime import datetime, timezone
from pathlib import Path
import csv
import json
import shutil

from .core import (APPEARANCE, CATEGORIES, MODELS, all_occurrences, coverage_distribution,
                   get_decision_log, get_dictionary, image_coverage, progress)


def _save_csv(path: Path, rows: list[dict], columns: list[str]):
    with path.open("w", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=columns, extrasaction="ignore")
        w.writeheader()
        for row in rows:
            w.writerow({k: row.get(k) for k in columns})


def _matrix(counter: Counter, *, column_name: str) -> list[dict]:
    keys = sorted({key for key, model in counter})
    result = []
    for key in keys:
        vals = {model: int(counter[(key, model)]) for model in MODELS}
        result.append({column_name: key, **vals, "Total": sum(vals.values())})
    return sorted(result, key=lambda r: (-r["Total"], r[column_name]))


def _model_long(counter: Counter, name: str, denom: dict[str, int]) -> list[dict]:
    rows = []
    for (word, model), count in counter.items():
        rows.append({"model": model, name: word, "count": count,
                     "percent_of_descriptors": round(100 * count / denom[model], 4) if denom[model] else 0})
    return sorted(rows, key=lambda r: (MODELS.index(r["model"]), -r["count"], r[name]))


def _figures(out_dir: Path, raw_mat: list[dict], clean_mat: list[dict],
             raw_long: list[dict], clean_long: list[dict], cats: list[dict],
             distribution: list[dict], model_dist: list[dict]):
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import numpy as np

    images = []
    def save(fig, name):
        path = out_dir / f"{name}.png"
        fig.savefig(path, dpi=190, bbox_inches="tight")
        plt.close(fig)
        images.append(path)

    def grouped_word(matrix, name, title):
        top = sorted(matrix[:20], key=lambda x: x["Total"])
        if not top:
            return
        y = np.arange(len(top), dtype=float)
        fig, ax = plt.subplots(figsize=(12, max(7, len(top) * .45)))
        for i, model in enumerate(MODELS):
            ax.barh(y + (i - 1.5) * .19, [r[model] for r in top], height=.18, label=model.title())
        ax.set_yticks(y, [r.get("word", r.get("canonical_word")) for r in top])
        ax.set_xlabel("Number of descriptors")
        ax.set_title(title)
        ax.grid(axis="x", alpha=.16)
        ax.legend(ncol=4, loc="lower right")
        fig.tight_layout()
        save(fig, name)

    grouped_word(raw_mat, "01_Raw_Top20_4Models", "Top 20 raw words - Four-model comparison")
    grouped_word(clean_mat, "02_Cleaned_Top20_4Models", "Top 20 canonical words - Four-model comparison")

    for data, prefix, col in ((raw_long, "03_Raw", "raw_word"), (clean_long, "04_Cleaned", "canonical_word")):
        for model in MODELS:
            top = sorted((x for x in data if x["model"] == model), key=lambda x: x["count"], reverse=True)[:15]
            top = top[::-1]
            fig, ax = plt.subplots(figsize=(9, 6))
            ax.barh([x[col] for x in top], [x["count"] for x in top])
            ax.set_title(f"{model.title()} - Top 15 {'raw' if col == 'raw_word' else 'canonical'} words")
            ax.set_xlabel("Occurrences")
            fig.tight_layout()
            save(fig, f"{prefix}_{model}_Top15")

    for measure, name, ylabel in (("count", "05_Category_Counts", "Descriptors"),
                                 ("percent_of_model", "06_Category_Percent", "Share of model descriptors (%)")):
        fig, ax = plt.subplots(figsize=(11, 5.5))
        x = np.arange(len(CATEGORIES), dtype=float)
        width = .19
        for idx, model in enumerate(MODELS):
            sub = {r["category"]: r for r in cats if r["model"] == model}
            ax.bar(x + (idx - 1.5) * width, [sub[c][measure] for c in CATEGORIES], width, label=model.title())
        ax.set_xticks(x, CATEGORIES, rotation=25, ha="right")
        ax.set_ylabel(ylabel)
        ax.set_title("Seven descriptor categories by model")
        ax.legend(ncol=4)
        fig.tight_layout()
        save(fig, name)

    for rows, name, title in ((distribution, "07_Image_Appearance_Coverage", "Appearance coverage per image (4-model union)"),
                              (model_dist, "08_ImageModel_Appearance_Coverage", "Appearance coverage per image-model response")):
        fig, ax = plt.subplots(figsize=(8.5, 4.6))
        ax.bar([str(r["appearance_categories"]) for r in rows], [r["image_count"] for r in rows])
        ax.set_xlabel("Number of appearance categories (Color, Transparency, Glossiness, Texture)")
        ax.set_ylabel("Fully classified units")
        ax.set_title(title)
        fig.text(.1, .01, f"Incomplete units excluded: {rows[0]['incomplete_images']}", fontsize=9)
        fig.tight_layout(rect=[0, .04, 1, 1])
        save(fig, name)
    return images


def _xlsx(out_path: Path, tables: dict[str, tuple[list[dict], list[str]]], metadata: dict):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.chart import BarChart, Reference
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    info = wb.active
    info.title = "Read_Me"
    info.append(["Project", "Appearance AI - 90 images / seven categories"])
    for key, value in metadata.items():
        info.append([key, value])
    info.append(["Appearance categories", ", ".join(APPEARANCE)])
    info.append(["Other categories", "Shape, Material, Other"])
    info.append(["Coverage rule", "Count unique appearance categories; only completed units have a final count"])
    info.append(["Image unit", "90 images; union of Gemini/Qwen/Llama/Mistral responses"])
    info.append(["Image-model unit", "360 responses; five descriptors per response"])
    info.append(["Missing labels", "Unassigned, never counted as Other or as category absence"])
    info.append(["Percentage bases", "All units and fully classified units shown separately"])
    info.append(["No artificial intelligence requests", "All labels are supplied by the researcher"])

    for name, (rows, cols) in tables.items():
        ws = wb.create_sheet(name)
        ws.append(cols)
        for row in rows:
            ws.append([row.get(col) for col in cols])
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{max(2, ws.max_row)}"
        for idx, col in enumerate(cols, 1):
            samples = [str(r.get(col) or "") for r in rows[:120]]
            length = min(42, max(12, len(col) + 2, max((len(v) for v in samples), default=0) + 1))
            ws.column_dimensions[get_column_letter(idx)].width = length
        for cell in ws[1]:
            cell.fill = PatternFill("solid", fgColor="22324C")
            cell.font = Font(bold=True, color="FFFFFF")
            cell.alignment = Alignment(wrap_text=True, vertical="center")
        ws.row_dimensions[1].height = 30

    info.column_dimensions["A"].width = 29
    info.column_dimensions["B"].width = 92
    for cell in info[1]:
        cell.fill = PatternFill("solid", fgColor="22324C")
        cell.font = Font(bold=True, color="FFFFFF")

    charts_ws = wb.create_sheet("Embedded_Charts")
    charts_ws.append(["Category", *MODELS])
    cat_lookup = {(r["category"], r["model"]): r["count"] for r in tables["Category_Model"][0]}
    for cat in CATEGORIES:
        charts_ws.append([cat, *[cat_lookup.get((cat, m), 0) for m in MODELS]])
    chart = BarChart()
    chart.type = "col"
    chart.grouping = "clustered"
    chart.title = "Category usage: four models"
    chart.y_axis.title = "Count"
    chart.x_axis.title = "Category"
    chart.add_data(Reference(charts_ws, min_col=2, max_col=5, min_row=1, max_row=8), titles_from_data=True)
    chart.set_categories(Reference(charts_ws, min_col=1, min_row=2, max_row=8))
    chart.width = 24
    chart.height = 11
    charts_ws.add_chart(chart, "A11")
    for col in "ABCDE":
        charts_ws.column_dimensions[col].width = 20
    wb.save(out_path)


def create_reports(db_path: Path, output_root: Path) -> dict:
    db_path = Path(db_path)
    if not db_path.is_file():
        raise FileNotFoundError(f"Database has not been initialized: {db_path}")
    all_rows = all_occurrences(db_path)
    if len(all_rows) != 1800:
        raise ValueError(f"Unexpected row count: {len(all_rows)}")
    pairs, images = image_coverage(all_rows)
    if len(pairs) != 360 or len(images) != 90:
        raise ValueError("Image/model coverage is incomplete in the source data.")
    p = progress(all_rows)
    model_totals = Counter(r["model"] for r in all_rows)
    raw_ctr = Counter((r["word_key"], r["model"]) for r in all_rows)
    clean_ctr = Counter((r["canonical_word"], r["model"]) for r in all_rows)
    raw_m = _matrix(raw_ctr, column_name="word")
    clean_m = _matrix(clean_ctr, column_name="canonical_word")
    raw_l = _model_long(raw_ctr, "raw_word", model_totals)
    clean_l = _model_long(clean_ctr, "canonical_word", model_totals)
    c_ctr = Counter((r["category"], r["model"]) for r in all_rows)
    categories = [*CATEGORIES, "Unassigned"]
    cat_models = [
        {"model": m, "category": c, "count": c_ctr[(c, m)],
         "percent_of_model": round(100 * c_ctr[(c, m)] / model_totals[m], 4)}
        for m in MODELS for c in categories
    ]
    appearance_dist = coverage_distribution(images)
    pair_dist = coverage_distribution(pairs)
    dictionary = get_dictionary(db_path)
    journal = get_decision_log(db_path)

    top_by_cat = []
    group_freq = Counter((r["category"], r["canonical_word"], r["model"]) for r in all_rows)
    for category in categories:
        for model in MODELS:
            sub = [(word, count) for (cat, word, m), count in group_freq.items() if cat == category and m == model]
            for rank, (word, count) in enumerate(sorted(sub, key=lambda a: (-a[1], a[0]))[:20], 1):
                top_by_cat.append({"category": category, "model": model, "rank": rank, "canonical_word": word, "count": count})

    descriptors = ["occurrence_id","image_id","image_name","model","position","raw_word","word_key",
                   "canonical_word","category","source","global_category","override_category","five_words","image_md5"]
    pair_cols = ["image_id","image_name","model","descriptors","unassigned","complete","known_appearance_categories",
                 "appearance_count", *CATEGORIES, *[f"has_{c}" for c in CATEGORIES]]
    image_cols = [c for c in pair_cols if c != "model"]
    coverage_cols = ["appearance_categories","image_count","percent_all_images","percent_classified_images",
                     "total_images","fully_classified_images","incomplete_images"]
    tables = {
        "Word_Dictionary": (dictionary, ["word_key","canonical_word","category","frequency","image_count"]),
        "Word_Instances_1800": (all_rows, descriptors),
        "Image_Model_360": (pairs, pair_cols),
        "Image_Summary_90": (images, image_cols),
        "Coverage_Images": (appearance_dist, coverage_cols),
        "Coverage_Model_Responses": (pair_dist, coverage_cols),
        "Raw_Frequency": (raw_l, ["model","raw_word","count","percent_of_descriptors"]),
        "Clean_Frequency": (clean_l, ["model","canonical_word","count","percent_of_descriptors"]),
        "Raw_Model_Matrix": (raw_m, ["word",*MODELS,"Total"]),
        "Clean_Model_Matrix": (clean_m, ["canonical_word",*MODELS,"Total"]),
        "Category_Model": (cat_models, ["model","category","count","percent_of_model"]),
        "Top_Words_By_Category": (top_by_cat, ["category","model","rank","canonical_word","count"]),
        "Needs_Review": ([r for r in all_rows if r["category"] == "Unassigned"], descriptors),
        "Decision_History": (journal, ["decision_id","recorded_at","scope","occurrence_id","word_key","previous_value","new_value","note"]),
    }
    output_root = Path(output_root)
    stamp = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S_%f")
    target = output_root / f"analysis_{stamp}"
    charts = target / "charts"
    csv_dir = target / "csv"
    charts.mkdir(parents=True, exist_ok=False)
    csv_dir.mkdir(parents=True)
    for name, (records, cols) in tables.items():
        _save_csv(csv_dir / (name + ".csv"), records, cols)
    metadata = {
        "Generated UTC": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "Dataset": "Dataset 2 - 90 images only",
        "Total images": 90,
        "Total model responses": 360,
        "Total word occurrences": p["total"],
        "Classified occurrences": p["assigned"],
        "Unassigned occurrences": p["remaining"],
        "Fully classified images": sum(x["complete"] for x in images),
        "Fully classified responses": sum(x["complete"] for x in pairs),
    }
    workbook = target / "Appearance_AI_90_Seven_Categories_Analysis.xlsx"
    _xlsx(workbook, tables, metadata)
    pngs = _figures(charts, raw_m, clean_m, raw_l, clean_l, cat_models, appearance_dist, pair_dist)
    shutil.copy2(db_path, target / "database_snapshot.sqlite3")
    (target / "summary.json").write_text(json.dumps({**metadata,
        "coverage_images": appearance_dist,"coverage_model_responses": pair_dist}, indent=2), encoding="utf-8")
    return {"output": str(target), "excel": str(workbook), "charts": len(pngs), **metadata}
