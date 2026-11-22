"""Reconcile published complaint and feedback counts against the source workbook."""

import argparse
from pathlib import Path
import re
import warnings

from openpyxl import load_workbook


SOURCE = Path(__file__).with_name(
    "2022-2024-public-complaints-central-adelaide-local-health-network.xlsx"
)
FEEDBACK_LABELS = (
    "Number of positive feedback comments",
    "Number of negative feedback comments",
    "Total number of feedback comments",
)


def count_value(value, location):
    if type(value) not in (int, float) or value < 0 or value % 1:
        raise ValueError(f"{location}: expected a non-negative whole count, got {value!r}")
    return int(value)


def audit_sheet(sheet):
    rows = list(sheet.iter_rows(min_col=1, max_col=4, values_only=True))
    headers = {row[0]: index for index, row in enumerate(rows)}
    start = headers["Complaint Categories"] + 1
    end = headers["Additional Metrics"]
    counts = []
    subcategories = set()
    not_available = empty = 0
    published_total = None

    for row_number, (category, subcategory, description, value) in enumerate(
        rows[start:end], start + 1
    ):
        location = f"{sheet.title}, row {row_number}"
        if description == "TOTAL":
            published_total = count_value(value, location)
        elif category and subcategory:
            key = (category, subcategory)
            if key in subcategories:
                raise ValueError(f"{location}: duplicate complaint subcategory {key!r}")
            subcategories.add(key)
            if value == "N/A":
                not_available += 1
            elif value is None:
                empty += 1
            else:
                counts.append(count_value(value, location))

    if published_total is None or not subcategories:
        raise ValueError(f"{sheet.title}: complaint rows or published total missing")

    feedback = {}
    for label, value, *_ in rows[end + 1:]:
        if label in FEEDBACK_LABELS:
            if label in feedback:
                raise ValueError(f"{sheet.title}: duplicate feedback metric {label!r}")
            feedback[label] = count_value(value, f"{sheet.title}, {label}")
    positive, negative, feedback_total = (feedback[label] for label in FEEDBACK_LABELS)

    return {
        "Complaint subcategories": len(subcategories),
        "Numeric complaint counts": len(counts),
        "Complaint counts marked N/A": not_available,
        "Blank complaint counts": empty,
        "Published complaint total": published_total,
        "Sum of numeric complaint counts": sum(counts),
        "Complaint total difference": published_total - sum(counts),
        "Positive feedback comments": positive,
        "Negative feedback comments": negative,
        "Positive and negative feedback comments combined": positive + negative,
        "Published feedback-comment total": feedback_total,
        "Feedback total difference": feedback_total - (positive + negative),
    }


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("workbook", type=Path, nargs="?", default=SOURCE)
    args = parser.parse_args()

    with warnings.catch_warnings():
        # The audit reads cell values only and never saves the source workbook.
        warnings.filterwarnings("ignore", message="Data Validation extension is not supported.*")
        workbook = load_workbook(args.workbook, read_only=True, data_only=True)
        try:
            sheets = [
                sheet for sheet in workbook.worksheets
                if re.fullmatch(r"\d{4}-\d{2} Complaints", sheet.title)
            ]
            if not sheets:
                raise ValueError("No annual complaints sheets found.")
            for sheet in sorted(sheets, key=lambda item: item.title):
                print(sheet.title)
                for label, value in audit_sheet(sheet).items():
                    print(f"  {label}: {value:,}")
                print()
        finally:
            workbook.close()


if __name__ == "__main__":
    main()
