"""Windows/VS Code CLI for the 90-image research dataset."""
from __future__ import annotations
import argparse
from pathlib import Path
import json

from src.core import initialize, is_initialized, all_occurrences, progress
from src.reports import create_reports

ROOT = Path(__file__).resolve().parent
DATASET = ROOT / "dataset_2"
EXCEL = DATASET / "input/Appearance_Part2_90_FINAL.xlsx"
DB = DATASET / "database/classifications.sqlite3"
OUT = DATASET / "output"


def main():
    parser = argparse.ArgumentParser(description="Appearance AI - Dataset 90 only")
    parser.add_argument("action", choices=["init", "status", "report"])
    args = parser.parse_args()

    if args.action == "init":
        print(json.dumps(initialize(EXCEL, DB), indent=2))
    elif args.action == "status":
        if not is_initialized(DB):
            parser.error("Database missing. Run: python main.py init")
        print(json.dumps(progress(all_occurrences(DB)), indent=2))
    elif args.action == "report":
        if not is_initialized(DB):
            parser.error("Database missing. Run: python main.py init")
        print(json.dumps(create_reports(DB, OUT), indent=2))


if __name__ == "__main__":
    main()
