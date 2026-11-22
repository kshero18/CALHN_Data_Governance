# CALHN public complaints data quality review

This independent portfolio project examines two years of published aggregate complaints and feedback data from Central Adelaide Local Health Network (CALHN). It asks which comparisons the workbook supports and which need more context.

[View the one-page results (PDF)](CALHN_Data_Governance.pdf) · [Download the Power BI report](CALHN_Data_Governance.pbix)

## Findings

| Check | 2022–23 | 2023–24 |
| --- | ---: | ---: |
| Complaint subcategories in the source | 16 | 16 |
| Numeric counts | 9 | 9 |
| Counts marked `N/A` in the source | 7 | 7 |
| Published complaint total | 1,085 | 1,650 |
| Sum of numeric subcategory counts | 1,085 | 1,650 |
| Positive and negative feedback comments combined | 3,367 | 6,704 |
| Published feedback-comment total | 3,577 | 6,704 |
| Published feedback total minus combined counts | 210 | 0 |

- **Complaint totals reconcile.** In both years the numeric subcategory counts add up to the published complaints total. The source does not explain what `N/A` means for the seven subcategories. Keep that label and avoid treating it as zero when comparing individual subcategories.
- **One feedback total needs clarification.** In 2022–23, adding the published positive and negative feedback counts gives 3,367, while the published total is 3,577. The workbook and annual report both show those values. The available figures leave a difference of 210 without explaining it. Keep the published figures and flag the difference before calculating a positive or negative share of all feedback.
- **These are aggregate counts.** The workbook supports comparisons of reported counts for subcategories with numeric values in both years. It does not provide activity denominators or individual records, so changes in counts cannot establish changes in complaint rates or service quality.

## Reproduce the checks

```sh
python -m pip install openpyxl
python audit_complaints.py
```

The script reads the workbook included in this repository. It separates the complaint table from the additional metrics, checks for duplicate subcategories within each year, distinguishes numeric counts from `N/A` and empty cells, and reconciles totals. It prints the results without modifying the workbook.

## Dashboard

The opening Power BI page shows the annual reconciliation side by side. It is a fixed summary generated from the workbook. The second page retains the interactive complaint distribution by category and year. The PDF provides a one-page copy of the reconciliation for readers without Power BI Desktop.

To rebuild the summary page and PDF from the bundled two-year workbook, install `openpyxl` and `reportlab`, then run `python build_report.py`. This updates the summary and PDF. The interactive detail page uses the original embedded complaint-count model.

### Earlier analysis

The original CSV extracts and model contain the prototype calculations. That extraction mixed 14 source-labelled `N/A` counts with 10 summary rows outside the complaint table, producing a misleading 57.14% missing-count figure. The former 88.57% quality score inherited this classification problem. The revised dashboard uses the source-level checks above.

## Public sources

- [CALHN data catalogue](https://data.sa.gov.au/data/dataset/central-adelaide-local-health-network-calhn)
- [CALHN 2022–23 annual report, pp. 50–51](https://www.sahealth.sa.gov.au/wps/wcm/connect/58393db8-b1ff-4e6a-956b-c0536761ad68/CALHN%2BAnnual%2BReport%2B2022-2023.pdf?MOD=AJPERES)
- [CALHN 2023–24 annual report, pp. 66–67](https://www.centraladelaide.health.sa.gov.au/wp-content/uploads/2024/12/CALHN-2023-24-Annual-Report.pdf)
