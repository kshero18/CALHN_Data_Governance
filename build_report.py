"""Regenerate the Power BI summary and its PDF companion for the bundled workbook."""
from copy import deepcopy
import json
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile

from openpyxl import load_workbook
from reportlab.lib.colors import HexColor
from reportlab.pdfgen.canvas import Canvas

from audit_complaints import SOURCE, audit_sheet

ROOT = Path(__file__).resolve().parent
PBIX = ROOT / "CALHN_Data_Governance.pbix"
YEARS = ("2022-23", "2023-24")
GROUPS = {
    "Complaint counts": [
        "Complaint subcategories", "Numeric complaint counts", "Complaint counts marked N/A",
        "Blank complaint counts", "Sum of numeric complaint counts", "Published complaint total",
        "Complaint total difference",
    ],
    "Feedback comments": [
        "Positive feedback comments", "Negative feedback comments",
        "Positive and negative feedback comments combined", "Published feedback-comment total",
        "Feedback total difference",
    ],
}
TITLE = "CALHN public complaints and feedback"
NOTES = [
    "Seven subcategories carry the source label N/A in each year. Their meaning is unspecified.",
    "The published 2022-23 feedback total exceeds positive and negative comments combined by 210.",
    "The source gives no breakdown of this difference. Complaint counts and feedback comments are separate measures.",
]


def textbox(name, x, y, width, height, lines, size, bold=False):
    position = dict(x=x, y=y, width=width, height=height, z=y + x, tabOrder=y + x)
    paragraphs = [{"textRuns": [{"value": line, "textStyle": {
        "fontSize": f"{size}pt", "fontWeight": "bold" if bold else "normal",
    }}]} for line in lines]
    config = {"name": name, "layouts": [{"id": 0, "position": position}],
              "singleVisual": {"visualType": "textbox", "objects": {
                  "general": [{"properties": {"paragraphs": paragraphs}}]}}}
    return {**position, "filters": "[]", "config": json.dumps(config)}


def write_power_bi(results, target):
    with ZipFile(PBIX) as source, ZipFile(target, "w") as output:
        layout = json.loads(source.read("Report/Layout").decode("utf-16le"))
        summary = deepcopy(layout["sections"][0])
        visuals = [textbox("auditTitle", 20, 10, 1220, 62, [TITLE], 29, True)]
        y = 80
        for group, labels in GROUPS.items():
            visuals.append(textbox("heading" + str(y), 20, y, 680, 44, [group], 20, True))
            for index, year in enumerate(YEARS):
                x = 790 + index * 220
                visuals.append(textbox("year" + str(x + y), x, y, 200, 44, [year], 20, True))
            y += 44
            visuals.append(textbox("labels" + str(y), 20, y, 750, 32 * len(labels), labels, 16))
            for index, year in enumerate(YEARS):
                x = 790 + index * 220
                values = [f"{results[year][label]:,}" for label in labels]
                visuals.append(textbox("values" + str(x + y), x, y, 190, 32 * len(labels), values, 16))
            y += 32 * len(labels)
        visuals.append(textbox("auditNotes", 20, y, 1230, 135, NOTES, 14))
        summary.update(displayName="Annual reconciliation", ordinal=0, visualContainers=visuals)
        detail = next(page for page in layout["sections"] if page["displayName"] == "Complaint Distribution")
        detail["ordinal"] = 1
        layout["sections"] = [summary, detail]
        config = json.loads(layout["config"])
        config["activeSectionIndex"] = 0
        layout["config"] = json.dumps(config)
        for entry in source.infolist():
            if entry.filename == "SecurityBindings":
                continue  # The original binding belongs to the previous report layout.
            data = source.read(entry)
            if entry.filename == "Report/Layout":
                data = json.dumps(layout, ensure_ascii=False).encode("utf-16le")
            elif entry.filename == "[Content_Types].xml":
                data = data.replace(b'<Override PartName="/SecurityBindings" ContentType="" />', b'')
            output.writestr(entry, data)


def write_pdf(results, target):
    pdf = Canvas(str(target), pagesize=(800, 720))
    pdf.setTitle(TITLE)
    pdf.setFillColor(HexColor("#153846"))
    pdf.setFont("Helvetica-Bold", 25)
    pdf.drawString(40, 665, TITLE)
    pdf.setFont("Helvetica", 12)
    pdf.drawString(40, 639, "Independent review of the published annual aggregates")
    y = 598
    for group, labels in GROUPS.items():
        pdf.setFont("Helvetica-Bold", 14)
        pdf.drawString(40, y, group)
        for index, year in enumerate(YEARS):
            pdf.drawRightString(585 + index * 150, y, year)
        y -= 24
        for label in labels:
            pdf.setFont("Helvetica-Bold" if label.endswith("difference") else "Helvetica", 12)
            pdf.drawString(40, y, label)
            for index, year in enumerate(YEARS):
                pdf.drawRightString(585 + index * 150, y, f"{results[year][label]:,}")
            y -= 23
        y -= 21
    pdf.setFont("Helvetica", 11)
    for note in NOTES:
        pdf.drawString(40, y, note)
        y -= 19
    pdf.drawString(40, 45, "Source: Central Adelaide Local Health Network annual public complaints workbook, 2022-23 and 2023-24.")
    pdf.save()


def main():
    workbook = load_workbook(SOURCE, read_only=True, data_only=True)
    try:
        results = {year: audit_sheet(workbook[f"{year} Complaints"]) for year in YEARS}
    finally:
        workbook.close()
    with TemporaryDirectory(dir=ROOT) as directory:
        report = Path(directory) / PBIX.name
        preview = report.with_suffix(".pdf")
        write_power_bi(results, report)
        write_pdf(results, preview)
        report.replace(PBIX)
        preview.replace(PBIX.with_suffix(".pdf"))


if __name__ == "__main__":
    main()
