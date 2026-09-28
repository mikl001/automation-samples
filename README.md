# Portfolio demos (sample projects)

Demo work for the Fiverr gigs. All data is fictional ("Bean & Leaf", a made-up coffee shop); label it as a sample project wherever it is shown.

| Demo | What it shows | Run |
|---|---|---|
| `excel-sales-report/` | 12 messy monthly CSV exports → Power Query cleanup (blank/duplicate lines, product and country spelling) → dashboard → VBA button "Refresh & export PDF" | `python gen_data.py` (sample data), `python build_workbook.py` (builds `Sales_Report_Automation.xlsm` through a private Excel instance) |
| `python-report-automation/` | Same exports → `build_report.py` (standard library + xlsxwriter) → formatted Excel report with charts and a data-quality log in 0.25 s | `python build_report.py --input ../excel-sales-report/data` |
| `sheets-apps-script/Code.gs` | Apps Script sample (import & clean CSVs from a Drive folder, custom menu, rows padded to one width for `setValues`) — excerpt on `gig-images/sheets-2-script.png`. **Not yet run in a live Google account**: the Sheets gig stays unpublished until it is tested with a separate test Google account (`clasp login` by the owner) | paste into Extensions › Apps Script, set script property `EXPORT_FOLDER_ID` |

Both pipelines produce the same result (4,901 raw lines → 4,792 clean), which cross-checks the cleanup logic.

- `render_shots.py` — renders the dashboards to `shots/*.png` (Excel opened read-only, macros forced off).
- `gig_images.py` — builds the 1280×769 gallery images in `gig-images/` with headless Edge. Honesty rules from `hq/offers/FIVERR-GIGS.md`: no fake badges or reviews, no Google/Microsoft logos, real numbers only.

`build_workbook.py` enables "Trust access to the VBA project object model" only while it injects the macro module and restores the previous registry value in a `finally` block.
