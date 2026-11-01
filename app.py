"""Local research UI for human-supervised classification of 1,800 descriptors."""
from __future__ import annotations
from collections import Counter
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile
import re

import streamlit as st
from PIL import Image

from src.core import (APPEARANCE, CATEGORIES, MODELS, all_occurrences,
                      clear_occurrence_override, get_dictionary, initialize,
                      is_initialized, progress, save_decision, update_canonical)
from src.reports import create_reports

ROOT = Path(__file__).resolve().parent
DATASET = ROOT / "dataset_2"
INPUT_XLSX = DATASET / "input/Appearance_Part2_90_FINAL.xlsx"
IMAGES_ZIP = DATASET / "input/images2.zip"
DB = DATASET / "database/classifications.sqlite3"
OUTPUT = DATASET / "output"

st.set_page_config(page_title="Appearance AI - Dataset 90", layout="wide")
st.title("Appearance AI | Dataset 90")
st.caption("Human-supervised descriptor classification | Seven categories | Offline | No external APIs")

if not is_initialized(DB):
    st.warning("Classification database is not initialized. Add the Excel file and click Initialize.")
    st.code(str(INPUT_XLSX))
    if st.button("Initialize Dataset 90"):
        try:
            result = initialize(INPUT_XLSX, DB)
        except Exception as exc:
            st.error(f"Initialization failed: {exc}")
        else:
            st.success(f"Database ready: {result['words']} descriptors, {result['unique_words']} unique words.")
            st.rerun()
    st.stop()

@st.cache_data(show_spinner=False)
def image_from_zip(zip_path: str, image_name: str):
    if not Path(zip_path).is_file():
        return None
    with ZipFile(zip_path) as zf:
        matches = [name for name in zf.namelist() if Path(name).name.lower() == image_name.lower()]
        if len(matches) != 1:
            return None
        data = zf.read(matches[0])
    with Image.open(BytesIO(data)) as pic:
        return pic.convert("RGB").copy()


def show_case(row):
    left, right = st.columns([1, 1.3])
    with left:
        image = image_from_zip(str(IMAGES_ZIP), row["image_name"])
        if image is None:
            st.info(f"Image unavailable in images2.zip: {row['image_name']}")
        else:
            st.image(image, caption=f"Image {row['image_name']}", use_container_width=True)
    with right:
        st.subheader(row["raw_word"])
        st.write(f"**Image:** `{row['image_name']}`")
        st.write(f"**Model:** `{row['model']}`")
        st.write(f"**Word position:** `{row['position']}/5`")
        st.write(f"**Original five-word response:** `{row['five_words'] or '(unavailable)'}`")
        st.write(f"**Existing category:** `{row['category']}` (source: {row['source']})")
        if row["override_category"]:
            st.info(f"This occurrence overrides its global label: {row['override_category']}")

rows = all_occurrences(DB)
status = progress(rows)
summary_col1, summary_col2, summary_col3 = st.columns(3)
summary_col1.metric("Classified descriptors", f"{status['assigned']}/1800")
summary_col2.metric("Pending descriptors", status["remaining"])
summary_col3.metric("Unique vocabulary", len(get_dictionary(DB)))
st.progress(status["assigned"] / max(1, status["total"]))

classify_tab, review_tab, report_tab = st.tabs(["Classify words", "Review / Edit decisions", "Export reports"])

with classify_tab:
    st.markdown("**Choose an image/model context, then classify the descriptor.**")
    c1, c2, c3 = st.columns(3)
    with c1:
        model_filter = st.selectbox("Model", ["All", *MODELS], key="filter_model")
    with c2:
        img_filter = st.selectbox("Image", ["All", *[f"{n:03d}" for n in range(1, 91)]], key="filter_image")
    with c3:
        search = st.text_input("Search word", "", key="filter_word").lower().strip()

    pending = [r for r in rows if r["category"] == "Unassigned"
               and (model_filter == "All" or r["model"] == model_filter)
               and (img_filter == "All" or r["image_id"] == int(img_filter))
               and (not search or search in r["word_key"])]
    st.write(f"**Pending in current filter:** {len(pending)}")
    if not pending:
        st.success("No pending descriptors match your filters. Use Review / Edit to change previous decisions.")
    else:
        lookup = {r["occurrence_id"]: r for r in pending}
        choice = st.selectbox(
            "Select descriptor occurrence",
            options=list(lookup),
            format_func=lambda key: (f"{lookup[key]['image_name']} | {lookup[key]['model']} | "
                                     f"{lookup[key]['position']}/5 | {lookup[key]['raw_word']}")
        )
        selected = lookup[choice]
        show_case(selected)
        word_frequency = sum(r["word_key"] == selected["word_key"] for r in rows)
        st.info(f"'{selected['word_key']}' appears {word_frequency} time(s) in this dataset.")
        with st.form("save_classification", clear_on_submit=False):
            category = st.selectbox("Final category", CATEGORIES)
            scope = st.radio("Apply to", ["All occurrences of this word", "Only this image + model + word position"])
            canonical = st.text_input(
                "Canonical word (optional spelling standardization)",
                value=selected["canonical_word"],
                help="Only saved when applying to all occurrences. Does not change original words."
            )
            note = st.text_input("Research note (optional)")
            submit = st.form_submit_button("Save classification", type="primary")
        if submit:
            try:
                save_decision(
                    DB,
                    occurrence_key=selected["occurrence_id"],
                    category=category,
                    scope="all" if scope.startswith("All") else "occurrence",
                    canonical_word=canonical if scope.startswith("All") else None,
                    note=note,
                )
                st.success("Classification stored in the Dataset 90 database.")
                st.rerun()
            except Exception as exc:
                st.error(f"Could not save decision: {exc}")

with review_tab:
    st.subheader("Review vocabulary and contextual exceptions")
    dictionary = get_dictionary(DB)
    st.dataframe(dictionary, use_container_width=True, hide_index=True)
    selected_word = st.selectbox("Edit word", [r["word_key"] for r in dictionary], key="edit_word")
    word_item = next(x for x in dictionary if x["word_key"] == selected_word)
    candidates = [r for r in rows if r["word_key"] == selected_word]
    st.write(f"**Occurrences:** {len(candidates)} | **Global category:** {word_item['category'] or 'Unassigned'}")
    with st.form("edit_global"):
        edit_category = st.selectbox("Global category", CATEGORIES,
                        index=CATEGORIES.index(word_item["category"]) if word_item["category"] in CATEGORIES else 0)
        edit_canonical = st.text_input("Canonical word", value=word_item["canonical_word"])
        note_edit = st.text_input("Reason for change (optional)")
        apply_global = st.form_submit_button("Save global label + canonical word")
    if apply_global:
        try:
            save_decision(DB, occurrence_key=candidates[0]["occurrence_id"],
                          category=edit_category, scope="all", canonical_word=edit_canonical, note=note_edit)
            st.success("Global label updated.")
            st.rerun()
        except Exception as exc:
            st.error(str(exc))
    selected_instance = st.selectbox(
        "View / edit one occurrence",
        options=[r["occurrence_id"] for r in candidates],
        format_func=lambda key: next(f"{r['image_name']} | {r['model']} | {r['position']}/5 | "
                                     f"{r['category']}" for r in candidates if r["occurrence_id"] == key)
    )
    record = next(r for r in candidates if r["occurrence_id"] == selected_instance)
    show_case(record)
    col_a, col_b = st.columns(2)
    with col_a:
        override_cat = st.selectbox("Occurrence-specific category", CATEGORIES, key="override_cat")
        if st.button("Save occurrence override"):
            save_decision(DB, occurrence_key=selected_instance, category=override_cat, scope="occurrence")
            st.rerun()
    with col_b:
        if st.button("Remove occurrence override", disabled=not bool(record["override_category"])):
            clear_occurrence_override(DB, selected_instance)
            st.rerun()

with report_tab:
    st.subheader("Scientific reports and charts")
    st.write("Image-level coverage uses the union of four model descriptions (20 words per image).")
    st.write("Image-model coverage uses five descriptors per model, per image.")
    if status["remaining"]:
        st.warning(f"{status['remaining']} descriptors are still Unassigned. "
                   "Coverage percentages will use only fully classified images and will also report incomplete cases.")
    if st.button("Generate Excel, CSV, and charts", type="primary"):
        try:
            with st.spinner("Creating local reports..."):
                result = create_reports(DB, OUTPUT)
            st.success("Reports created.")
            st.json(result)
            with open(result["excel"], "rb") as fh:
                st.download_button("Download Excel", fh.read(), file_name=Path(result["excel"]).name,
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as exc:
            st.error(f"Report generation failed: {exc}")
