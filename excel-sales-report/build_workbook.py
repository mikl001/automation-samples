"""Build Sales_Report_Automation.xlsm: Power Query cleanup + dashboard + VBA refresh/export.

Runs a private, invisible Excel instance through COM. VBA injection needs
"Trust access to the VBA project object model"; the script enables it only for
the duration of the build and restores the previous registry value afterwards.
"""
import time
import winreg
from pathlib import Path

import pythoncom
import win32com.client as win32

ROOT = Path(__file__).parent.resolve()
DATA = ROOT / "data"
OUT = ROOT / "Sales_Report_Automation.xlsm"

VBOM_KEY = r"Software\Microsoft\Office\16.0\Excel\Security"

SLATE, INK, MUTED, LINE, PAPER, WHITE = "2F3E46", "1F2A30", "6B7780", "DADFE3", "F4F5F7", "FFFFFF"
COPPER, SAGE, SAND, CLAY = "C8773A", "6E8B74", "D9B26F", "9C5B4F"

PRODUCTS = [
    ("Ethiopia Yirgacheffe 250g", "CB-ETH-250", "Coffee Beans", 12.50),
    ("Colombia Huila 250g", "CB-COL-250", "Coffee Beans", 11.00),
    ("Brazil Santos 1kg", "CB-BRA-1KG", "Coffee Beans", 29.00),
    ("House Espresso Blend 500g", "CB-ESP-500", "Coffee Beans", 18.50),
    ("Japanese Sencha 100g", "TE-SEN-100", "Tea", 9.50),
    ("Earl Grey Supreme 100g", "TE-EGR-100", "Tea", 7.90),
    ("Ceremonial Matcha 30g", "TE-MAT-030", "Tea", 21.00),
    ("V60 Dripper Ceramic", "EQ-V60-02", "Equipment", 27.00),
    ("Hand Grinder Steel Burr", "EQ-GRD-HND", "Equipment", 64.00),
    ("Gooseneck Kettle 1L", "EQ-KTL-GOS", "Equipment", 49.00),
    ("Paper Filters 100pcs", "AC-FLT-100", "Accessories", 5.50),
    ("Stoneware Mug 350ml", "AC-MUG-CER", "Accessories", 14.00),
]
COUNTRY_MAP = [
    ("NL", "Netherlands"), ("NETHERLANDS", "Netherlands"), ("NEDERLAND", "Netherlands"),
    ("BE", "Belgium"), ("BELGIUM", "Belgium"), ("BELGIE", "Belgium"), ("BELGIË", "Belgium"),
    ("DE", "Germany"), ("GERMANY", "Germany"), ("DEUTSCHLAND", "Germany"),
    ("FR", "France"), ("FRANCE", "France"),
]

M_RAW = r'''
let
    FolderPath = Excel.CurrentWorkbook(){[Name="DataFolder"]}[Content]{0}[Column1],
    Files = Folder.Files(FolderPath),
    CsvFiles = Table.SelectRows(Files, each Text.Lower([Extension]) = ".csv"),
    WithRows = Table.AddColumn(CsvFiles, "Rows", each Table.PromoteHeaders(
        Csv.Document([Content], [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
        [PromoteAllScalars = true])),
    Slim = Table.RenameColumns(Table.SelectColumns(WithRows, {"Name", "Rows"}), {{"Name", "SourceFile"}}),
    Expanded = Table.ExpandTableColumn(Slim, "Rows",
        {"OrderID", "OrderDate", "Country", "Channel", "Product", "Category", "Qty", "UnitPrice", "DiscountPct"})
in
    Expanded
'''

M_SALES = r'''
let
    Normalize = (t) => Text.Lower(Text.Combine(List.Select(Text.Split(Text.Trim(if t = null then "" else t), " "), each _ <> ""), " ")),
    NonBlank = Table.SelectRows(RawCombined, each [OrderID] <> null and Text.Trim([OrderID]) <> ""),
    Deduped = Table.Distinct(NonBlank),
    WithKeys = Table.AddColumn(
        Table.AddColumn(Deduped, "ProductKey", each Normalize([Product]), type text),
        "CountryKey", each Text.Upper(Text.Trim(if [Country] = null then "" else [Country])), type text),
    Products = Table.AddColumn(Excel.CurrentWorkbook(){[Name="Products"]}[Content], "ProductKey", each Normalize([Product]), type text),
    CountryMap = Table.TransformColumns(Excel.CurrentWorkbook(){[Name="CountryMap"]}[Content], {{"Raw", each Text.Upper(Text.Trim(_)), type text}}),
    JoinP = Table.NestedJoin(WithKeys, {"ProductKey"}, Products, {"ProductKey"}, "P", JoinKind.LeftOuter),
    ExpP = Table.ExpandTableColumn(JoinP, "P", {"SKU", "Product", "Category"}, {"SKU", "ProductClean", "CategoryClean"}),
    JoinC = Table.NestedJoin(ExpP, {"CountryKey"}, CountryMap, {"Raw"}, "C", JoinKind.LeftOuter),
    ExpC = Table.ExpandTableColumn(JoinC, "C", {"Country"}, {"CountryClean"}),
    Typed = Table.TransformColumnTypes(ExpC,
        {{"OrderDate", type date}, {"Qty", Int64.Type}, {"UnitPrice", type number}, {"DiscountPct", type number}}, "en-US"),
    WithRevenue = Table.AddColumn(Typed, "Revenue", each Number.Round([Qty] * [UnitPrice] * (1 - [DiscountPct]), 2), type number),
    WithMonth = Table.AddColumn(WithRevenue, "Month", each Date.StartOfMonth([OrderDate]), type date),
    Picked = Table.SelectColumns(WithMonth, {"OrderID", "OrderDate", "Month", "CountryClean", "Channel", "SKU",
        "ProductClean", "CategoryClean", "Qty", "UnitPrice", "DiscountPct", "Revenue", "SourceFile"}),
    Renamed = Table.RenameColumns(Picked, {{"CountryClean", "Country"}, {"ProductClean", "Product"}, {"CategoryClean", "Category"}}),
    Final = Table.TransformColumns(Renamed, {{"Channel", each Text.Proper(Text.Trim(_)), type text}})
in
    Final
'''

M_STATS = r'''
let
    NonBlank = Table.SelectRows(RawCombined, each [OrderID] <> null and Text.Trim([OrderID]) <> ""),
    Stats = #table(type table [Metric = text, Value = Int64.Type], {
        {"Export files loaded", List.Count(List.Distinct(RawCombined[SourceFile]))},
        {"Raw lines read", Table.RowCount(RawCombined)},
        {"Blank lines removed", Table.RowCount(RawCombined) - Table.RowCount(NonBlank)},
        {"Duplicate lines removed", Table.RowCount(NonBlank) - Table.RowCount(Table.Distinct(NonBlank))},
        {"Unrecognised products", Table.RowCount(Table.SelectRows(Sales, each [SKU] = null))},
        {"Unrecognised countries", Table.RowCount(Table.SelectRows(Sales, each [Country] = null))},
        {"Clean order lines", Table.RowCount(Sales)}
    })
in
    Stats
'''

VBA = r'''Option Explicit

' Button on the Dashboard: refresh everything and save a dated PDF next to the workbook.
Public Sub RefreshAndExport()
    Dim started As Double: started = Timer
    Dim pdfPath As String
    On Error GoTo Failed
    pdfPath = DoRefreshAndExport()
    MsgBox "Report refreshed in " & Format(Timer - started, "0.0") & " s." & vbCrLf & vbCrLf & _
           "PDF saved to:" & vbCrLf & pdfPath, vbInformation, "Sales report"
    Exit Sub
Failed:
    Application.ScreenUpdating = True
    MsgBox "Refresh failed: " & Err.Description & vbCrLf & _
           "Check the data folder on the Settings sheet.", vbExclamation, "Sales report"
End Sub

' Does the work without any dialogs, so it can also run from a scheduler or script.
Public Function DoRefreshAndExport() As String
    Dim pdfPath As String
    Application.ScreenUpdating = False
    ThisWorkbook.RefreshAll
    Application.CalculateUntilAsyncQueriesDone
    Application.CalculateFull
    ThisWorkbook.Names("LastRefresh").RefersToRange.Value = Now
    pdfPath = ThisWorkbook.Path & Application.PathSeparator & "Sales_Report_" & Format(Now, "yyyy-mm-dd_hhnn") & ".pdf"
    ThisWorkbook.Worksheets("Dashboard").ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    Application.ScreenUpdating = True
    DoRefreshAndExport = pdfPath
End Function

' Settings sheet: point the report at another export folder.
Public Sub PickDataFolder()
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select the folder with the monthly CSV exports"
        If .Show = -1 Then
            ThisWorkbook.Worksheets("Settings").Range("DataFolder").Value = .SelectedItems(1) & Application.PathSeparator
        End If
    End With
End Sub
'''


def bgr(hex_rgb: str) -> int:
    """Excel wants colours as 0xBBGGRR."""
    r, g, b = int(hex_rgb[0:2], 16), int(hex_rgb[2:4], 16), int(hex_rgb[4:6], 16)
    return r + (g << 8) + (b << 16)


def set_vbom(value):
    """Set AccessVBOM; returns the previous value (None if it did not exist)."""
    key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, VBOM_KEY)
    try:
        prev = winreg.QueryValueEx(key, "AccessVBOM")[0]
    except FileNotFoundError:
        prev = None
    if value is None:
        try:
            winreg.DeleteValue(key, "AccessVBOM")
        except FileNotFoundError:
            pass
    else:
        winreg.SetValueEx(key, "AccessVBOM", 0, winreg.REG_DWORD, value)
    winreg.CloseKey(key)
    return prev


def write_table(ws, top_left, headers, rows, name, style="TableStyleLight1"):
    r0, c0 = top_left
    ws.Range(ws.Cells(r0, c0), ws.Cells(r0, c0 + len(headers) - 1)).Value = headers
    if rows:
        ws.Range(ws.Cells(r0 + 1, c0), ws.Cells(r0 + len(rows), c0 + len(headers) - 1)).Value = rows
    rng = ws.Range(ws.Cells(r0, c0), ws.Cells(r0 + len(rows), c0 + len(headers) - 1))
    lo = ws.ListObjects.Add(1, rng, None, 1)  # xlSrcRange, has headers
    lo.Name = name
    lo.TableStyle = style
    return lo


def load_query(ws, query, dest, table_name):
    src = ("OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;"
           f"Location={query};Extended Properties=\"\"")
    lo = ws.ListObjects.Add(SourceType=0, Source=src, Destination=ws.Range(dest))
    qt = lo.QueryTable
    qt.CommandType = 2  # xlCmdSql
    qt.CommandText = f"SELECT * FROM [{query}]"
    qt.BackgroundQuery = False
    qt.RefreshStyle = 1  # xlInsertDeleteCells
    qt.AdjustColumnWidth = True
    qt.PreserveColumnInfo = True
    lo.Name = table_name
    qt.Refresh(False)
    return lo


def tile(ws, col, label, formula, fmt, caption, dynamic=False):
    c1, c3 = ws.Cells(5, col), ws.Cells(8, col + 2)
    box = ws.Range(c1, c3)
    box.Interior.Color = bgr(WHITE)
    for edge in (7, 8, 9, 10):  # left, top, bottom, right
        b = box.Borders(edge)
        b.LineStyle = 1
        b.Color = bgr(LINE)
    lab = ws.Range(ws.Cells(5, col), ws.Cells(5, col + 2))
    lab.Merge()
    lab.Value = label.upper()
    lab.Font.Size, lab.Font.Color, lab.Font.Bold = 8, bgr(MUTED), True
    lab.IndentLevel = 1
    val = ws.Range(ws.Cells(6, col), ws.Cells(7, col + 2))
    val.Merge()
    if dynamic:
        val.Cells(1, 1).Formula2 = formula
    else:
        val.Cells(1, 1).Formula = formula
    val.NumberFormat = fmt
    val.Font.Size, val.Font.Bold, val.Font.Color = 20, True, bgr(INK)
    val.HorizontalAlignment, val.VerticalAlignment = -4131, -4108  # left, center
    val.IndentLevel = 1
    cap = ws.Range(ws.Cells(8, col), ws.Cells(8, col + 2))
    cap.Merge()
    if caption.startswith("="):
        cap.Cells(1, 1).Formula2 = caption
    else:
        cap.Value = caption
    cap.Font.Size, cap.Font.Color = 8, bgr(MUTED)
    cap.IndentLevel = 1


def style_chart(ch, title, number_format=None):
    ch.HasTitle = True
    ch.ChartTitle.Text = title
    ch.ChartTitle.Format.TextFrame2.TextRange.Font.Size = 11
    ch.ChartTitle.Format.TextFrame2.TextRange.Font.Bold = True
    ch.ChartTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = bgr(INK)
    ch.ChartArea.Format.Line.Visible = False
    ch.ChartArea.Format.Fill.ForeColor.RGB = bgr(WHITE)
    ch.ChartArea.Format.TextFrame2.TextRange.Font.Size = 9
    ch.ChartArea.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = bgr(MUTED)
    ch.HasLegend = False
    if number_format:
        ax = ch.Axes(2)
        ax.TickLabels.NumberFormat = number_format
        ax.MajorGridlines.Format.Line.ForeColor.RGB = bgr(LINE)
        ax.Format.Line.Visible = False


def add_chart(ws, calc, anchor, span, kind, title, cat_rng, val_rng, colors, fmt="€#,##0", labels=False):
    left, top = ws.Range(anchor).Left, ws.Range(anchor).Top
    width = ws.Range(span).Width
    height = ws.Range(span).Height
    ch = ws.ChartObjects().Add(left, top, width, height).Chart
    ch.ChartType = kind
    s = ch.SeriesCollection().NewSeries()
    s.Name = title
    s.Values = calc.Range(val_rng)
    s.XValues = calc.Range(cat_rng)
    style_chart(ch, title, None if kind == -4120 else fmt)
    if kind == -4120:  # doughnut
        ch.HasLegend = True
        ch.Legend.Position = -4107  # bottom
        ch.ChartGroups(1).DoughnutHoleSize = 58
        for i, c in enumerate(colors, start=1):
            s.Points(i).Format.Fill.ForeColor.RGB = bgr(c)
        s.HasDataLabels = True
        dl = s.DataLabels()
        dl.ShowValue, dl.ShowPercentage, dl.ShowCategoryName = False, True, False
        dl.NumberFormat = "0%"
        dl.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = bgr(WHITE)
        dl.Format.TextFrame2.TextRange.Font.Bold = True
    else:
        s.Format.Fill.ForeColor.RGB = bgr(colors[0])
        ch.ChartGroups(1).GapWidth = 60
        if labels:
            s.HasDataLabels = True
            dl = s.DataLabels()
            dl.NumberFormat = fmt
            dl.Position = 3  # inside end: long bars never push the label past the plot edge
            dl.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = bgr(WHITE)
            dl.Format.TextFrame2.TextRange.Font.Bold = True
    if kind == 57:  # bar: biggest on top
        ch.Axes(1).ReversePlotOrder = True
        ch.Axes(2).Delete()
        ch.Axes(2).MajorGridlines.Delete() if ch.Axes(2).HasMajorGridlines else None
    return ch


def build():
    prev = set_vbom(1)
    xl = None
    try:
        xl = win32.gencache.EnsureDispatch(win32.DispatchEx("Excel.Application"))
        xl.Visible = False
        xl.DisplayAlerts = False
        xl.ScreenUpdating = False
        wb = xl.Workbooks.Add()
        while wb.Worksheets.Count < 5:
            wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        dash, data, settings, lists, calc = (wb.Worksheets(i) for i in range(1, 6))
        for ws, name in ((dash, "Dashboard"), (data, "Data"), (settings, "Settings"), (lists, "Lists"), (calc, "Calc")):
            ws.Name = name
        xl.ActiveWindow.DisplayGridlines = False

        # Settings
        settings.Range("B2").Value = "Settings"
        settings.Range("B2").Font.Size, settings.Range("B2").Font.Bold = 16, True
        settings.Range("B4").Value = "Data folder"
        settings.Range("C4").Value = str(DATA) + "\\"
        wb.Names.Add(Name="DataFolder", RefersTo="=Settings!$C$4")
        settings.Range("B6").Value = ("Drop each month's webshop export (CSV) into this folder, "
                                      "then press 'Refresh & export PDF' on the Dashboard.")
        settings.Range("B6").Font.Color = bgr(MUTED)
        settings.Columns("B").ColumnWidth = 14
        settings.Columns("C").ColumnWidth = 70
        btn = settings.Shapes.AddShape(5, settings.Range("E4").Left + 6, settings.Range("E4").Top - 3, 90, 22)
        btn.TextFrame2.TextRange.Text = "Browse..."
        btn.Fill.ForeColor.RGB, btn.Line.Visible = bgr(SLATE), False
        btn.TextFrame2.TextRange.Font.Size = 10
        btn.OnAction = "PickDataFolder"

        # Master data
        write_table(lists, (1, 1), ["Product", "SKU", "Category", "ListPrice"], [list(p) for p in PRODUCTS], "Products")
        write_table(lists, (1, 6), ["Raw", "Country"], [list(c) for c in COUNTRY_MAP], "CountryMap")
        lists.Columns("A:G").AutoFit()

        # Power Query
        wb.Queries.FastCombine = True
        wb.Queries.Add("RawCombined", M_RAW, "All CSV exports in the data folder, combined as-is")
        wb.Queries.Add("Sales", M_SALES, "Cleaned, de-duplicated, standardised order lines")
        wb.Queries.Add("LoadStats", M_STATS, "Data-quality counters for the last refresh")
        sales = load_query(data, "Sales", "A1", "Sales")
        sales.TableStyle = "TableStyleLight9"
        for col, fmt in (("OrderDate", "yyyy-mm-dd"), ("Month", "mmm yyyy"), ("UnitPrice", "€#,##0.00"),
                         ("DiscountPct", "0%"), ("Revenue", "€#,##0.00")):
            sales.ListColumns(col).DataBodyRange.NumberFormat = fmt
        data.Columns("A:M").AutoFit()
        load_query(calc, "LoadStats", "M1", "LoadStats")

        # Calc sheet: chart feeds (dynamic arrays)
        calc.Range("A1:K1").Value = ["Month", "Revenue", "", "Category", "Revenue", "", "Country", "Revenue", "", "Channel", "Revenue"]
        calc.Range("A2").Formula2 = "=SORT(UNIQUE(Sales[Month]))"
        calc.Range("B2").Formula2 = "=SUMIFS(Sales[Revenue],Sales[Month],A2#)"
        for key, c in (("Category", "D"), ("Country", "G"), ("Channel", "J")):
            nxt = chr(ord(c) + 1)
            calc.Range(f"{c}2").Formula2 = (f"=LET(k,UNIQUE(Sales[{key}]),r,SUMIFS(Sales[Revenue],Sales[{key}],k),SORTBY(k,r,-1))")
            calc.Range(f"{nxt}2").Formula2 = f"=SUMIFS(Sales[Revenue],Sales[{key}],{c}2#)"
        calc.Range("A2:A13").NumberFormat = "mmm"
        calc.Columns("A:N").AutoFit()

        # Dashboard layout
        dash.Activate()
        xl.ActiveWindow.DisplayGridlines = False
        xl.ActiveWindow.DisplayHeadings = False
        dash.Cells.Font.Name = "Segoe UI"
        dash.Range("A1:Q42").Interior.Color = bgr(PAPER)
        dash.Columns("A").ColumnWidth = 2
        dash.Columns("B:P").ColumnWidth = 10.5
        dash.Columns("Q").ColumnWidth = 2
        dash.Rows(1).RowHeight = 10
        dash.Rows(2).RowHeight = 30
        band = dash.Range("A1:Q3")
        band.Interior.Color = bgr(SLATE)
        dash.Range("B2").Value = "Bean & Leaf  |  Sales report 2025"
        dash.Range("B2").Font.Size, dash.Range("B2").Font.Bold, dash.Range("B2").Font.Color = 18, True, bgr(WHITE)
        dash.Range("B3").Formula2 = '="Built automatically from "&INDEX(LoadStats[Value],1)&" monthly webshop exports  ·  last refresh "&TEXT(LastRefresh,"dd mmm yyyy hh:mm")'
        dash.Range("B3").Font.Size, dash.Range("B3").Font.Color = 9, bgr("C9D1D6")
        calc.Range("P1").Value = "Last refresh"
        calc.Range("P2").Value = time.strftime("%Y-%m-%d %H:%M")
        calc.Range("P2").NumberFormat = "yyyy-mm-dd hh:mm"
        wb.Names.Add(Name="LastRefresh", RefersTo="=Calc!$P$2")
        dash.Rows(3).RowHeight = 20
        dash.Rows(4).RowHeight = 12
        b = dash.Shapes.AddShape(5, dash.Range("N2").Left, dash.Range("N2").Top + 3, dash.Range("N2:P2").Width, 26)
        b.TextFrame2.TextRange.Text = "Refresh & export PDF"
        b.TextFrame2.TextRange.Font.Size, b.TextFrame2.TextRange.Font.Bold = 10, True
        b.Fill.ForeColor.RGB, b.Line.Visible = bgr(COPPER), False
        b.OnAction = "RefreshAndExport"

        tile(dash, 2, "Revenue", "=SUM(Sales[Revenue])", "€#,##0", "net of discounts")
        tile(dash, 5, "Orders", "=ROWS(UNIQUE(Sales[OrderID]))", "#,##0", "unique order IDs", dynamic=True)
        tile(dash, 8, "Avg order value", "=B6/E6", "€#,##0.00", "revenue ÷ orders")
        tile(dash, 11, "Units sold", "=SUM(Sales[Qty])", "#,##0", "items shipped")
        tile(dash, 14, "Best month", '=TEXT(INDEX(Calc!A2#,MATCH(MAX(Calc!B2#),Calc!B2#,0)),"mmmm")', "@",
             '=TEXT(MAX(Calc!B2#),"€#,##0")&" revenue"', dynamic=True)
        for r in (9, 25):
            dash.Rows(r).RowHeight = 10

        add_chart(dash, calc, "B10", "B10:J24", 51, "Revenue by month", "A2:A13", "B2:B13", [SLATE])
        add_chart(dash, calc, "K10", "K10:P24", 57, "Revenue by category", "D2:D5", "E2:E5", [COPPER], labels=True)
        add_chart(dash, calc, "B26", "B26:F40", 57, "Revenue by country", "G2:G5", "H2:H5", [SAGE], labels=True)
        add_chart(dash, calc, "G26", "G26:J40", -4120, "Sales channel mix", "J2:J4", "K2:K4", [SLATE, COPPER, SAND])

        # Top products + data quality panel
        panel = dash.Range("K26:P40")
        panel.Interior.Color = bgr(WHITE)
        dash.Range("K26").Value = "Top 5 products"
        dash.Range("K34").Value = "Data quality (last refresh)"
        for cell in ("K26", "K34"):
            dash.Range(cell).Font.Bold, dash.Range(cell).Font.Size, dash.Range(cell).Font.Color = True, 11, bgr(INK)
            dash.Range(cell).IndentLevel = 1
        dash.Range("K27").Formula2 = ("=LET(p,UNIQUE(Sales[Product]),r,SUMIFS(Sales[Revenue],Sales[Product],p),"
                                      "TAKE(SORTBY(p,r,-1),5))")
        dash.Range("O27").Formula2 = "=SUMIFS(Sales[Revenue],Sales[Product],K27#)"
        dash.Range("K27:K31").IndentLevel = 1
        dash.Range("O27:O31").NumberFormat = "€#,##0"
        dash.Range("K35").Formula2 = "=TAKE(LoadStats[Metric],6)"
        dash.Range("O35").Formula2 = "=TAKE(LoadStats[Value],6)"
        dash.Range("K35:K40").IndentLevel = 1
        dash.Range("O35:O40").NumberFormat = "#,##0"
        for rng in ("K27:P31", "K35:P40"):
            dash.Range(rng).Font.Size = 9
            dash.Range(rng).Font.Color = bgr(INK)
        dash.Range("O27:P31").HorizontalAlignment = -4152
        dash.Range("O35:P40").HorizontalAlignment = -4152

        ps = dash.PageSetup
        ps.PrintArea = "$A$1:$Q$41"
        ps.Orientation = 2
        ps.Zoom = False
        ps.FitToPagesWide = 1
        ps.FitToPagesTall = 1
        ps.CenterHorizontally = True
        for side in ("LeftMargin", "RightMargin", "TopMargin", "BottomMargin"):
            setattr(ps, side, xl.InchesToPoints(0.25))

        # VBA
        mod = wb.VBProject.VBComponents.Add(1)  # standard module
        mod.Name = "modReport"
        mod.CodeModule.AddFromString(VBA)

        for ws in (data, settings, lists, calc):
            ws.Activate()
            xl.ActiveWindow.DisplayGridlines = False
        dash.Activate()
        dash.Range("A1").Select()
        calc.Visible = 0  # xlSheetHidden

        wb.SaveAs(str(OUT), 52)
        pdf = xl.Run("DoRefreshAndExport")
        wb.Save()
        stats = [(calc.Cells(r, 13).Value, calc.Cells(r, 14).Value) for r in range(2, 9)]
        kpis = {n: dash.Range(a).Text for n, a in (("revenue", "B6"), ("orders", "E6"), ("aov", "H6"),
                                                    ("units", "K6"), ("best", "N6"))}
        wb.Close(SaveChanges=False)
        return pdf, stats, kpis
    finally:
        if xl is not None:
            xl.Quit()
        set_vbom(prev)


if __name__ == "__main__":
    pythoncom.CoInitialize()
    t = time.time()
    pdf, stats, kpis = build()
    print("built", OUT.name, f"in {time.time() - t:.1f}s")
    print("pdf:", pdf)
    for m, v in stats:
        print(f"  {m}: {v}")
    print("kpis:", kpis)
