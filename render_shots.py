"""Render demo outputs to PNG for the Fiverr gallery (private Excel instance, read-only)."""
from pathlib import Path

import fitz  # PyMuPDF
import pythoncom
import win32com.client as win32

ROOT = Path(__file__).parent.resolve()
SHOTS = ROOT / "shots"


def pdf_to_png(pdf: Path, png: Path, width_px: int = 2200) -> None:
    doc = fitz.open(pdf)
    page = doc[0]
    zoom = width_px / page.rect.width
    page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False).save(png)
    doc.close()


def sheet_to_pdf(xl, book: Path, sheet: str, pdf: Path) -> None:
    wb = xl.Workbooks.Open(str(book), 0, True)  # no link updates, read-only
    try:
        wb.Worksheets(sheet).ExportAsFixedFormat(0, str(pdf))
    finally:
        wb.Close(False)


def main() -> None:
    SHOTS.mkdir(exist_ok=True)
    excel_pdf = sorted((ROOT / "excel-sales-report").glob("Sales_Report_*.pdf"))[-1]
    pdf_to_png(excel_pdf, SHOTS / "excel_dashboard.png")
    pythoncom.CoInitialize()
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible, xl.DisplayAlerts = False, False
    xl.AutomationSecurity = 3  # never run macros while rendering
    try:
        py_pdf = SHOTS / "python_summary.pdf"
        sheet_to_pdf(xl, ROOT / "python-report-automation" / "Sales_Report_2025.xlsx", "Summary", py_pdf)
        pdf_to_png(py_pdf, SHOTS / "python_summary.png")
    finally:
        xl.Quit()
    for p in sorted(SHOTS.glob("*.png")):
        print(p.name, p.stat().st_size // 1024, "KB")


if __name__ == "__main__":
    main()
