"""Render Fiverr gallery images (1280x769) from HTML with headless Edge.

Rules (from hq/offers/FIVERR-GIGS.md): before -> after, big niche title,
no fake badges/reviews, no Google or Microsoft logos. Demo work is labelled as such.
"""
import csv
import html
import subprocess
from pathlib import Path

ROOT = Path(__file__).parent.resolve()
OUT = ROOT / "gig-images"
SHOTS = ROOT / "shots"
DATA = ROOT / "excel-sales-report" / "data"
EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"

CSS = """
*{box-sizing:border-box;margin:0;padding:0}
body{width:1280px;height:769px;overflow:hidden;font-family:'Segoe UI',system-ui,sans-serif;background:#F4F5F7;color:#1F2A30}
.top{height:210px;background:#2F3E46;color:#fff;padding:44px 56px 0}
.top h1{font-size:64px;font-weight:800;letter-spacing:-1px;line-height:1.05}
.top p{font-size:28px;color:#C9D1D6;margin-top:14px}
.row{display:flex;gap:28px;align-items:center;padding:34px 44px 0}
.card{background:#fff;border:1px solid #DADFE3;border-radius:14px;padding:18px;position:relative}
.tag{position:absolute;top:-16px;left:18px;font-size:15px;font-weight:800;letter-spacing:2px;padding:5px 12px;border-radius:6px;color:#fff}
.bad{background:#B4452F}.good{background:#3F7A55}
.arrow{font-size:64px;color:#C8773A;font-weight:900}
pre{font-family:Consolas,'Cascadia Mono',monospace;font-size:15px;line-height:1.55;white-space:pre}
.hl{background:#F8D9D2;color:#8E2F1D;border-radius:3px}
.muted{color:#6B7780}
.shot{display:block;border-radius:8px;border:1px solid #DADFE3}
.note{position:absolute;bottom:14px;right:18px;font-size:13px;color:#6B7780}
.caption{padding:26px 56px 0;font-size:26px;line-height:1.35}
.caption b{color:#C8773A}
.grid{border-collapse:collapse;font-size:15px}
.grid td,.grid th{border:1px solid #D5DADF;padding:5px 10px;text-align:left}
.grid th{background:#EEF1F3;color:#6B7780;font-weight:600}
.stat{font-size:22px;line-height:1.9}
.stat b{display:inline-block;min-width:110px;text-align:right;margin-right:14px;color:#2F3E46}
.code{background:#1F2A30;color:#E6EDF3;border-radius:12px;padding:20px 24px}
.code .k{color:#E0A36B}.code .c{color:#8FA3AE}.code .s{color:#A8D0A0}
"""


def page(body: str) -> str:
    return f"<!doctype html><html><head><meta charset='utf-8'><style>{CSS}</style></head><body>{body}</body></html>"


def raw_sample() -> str:
    """Real lines from the January export, with the problems highlighted."""
    with (DATA / "orders_2025-01.csv").open(encoding="utf-8", newline="") as f:
        rows = list(csv.reader(f))
    header, body = rows[0], rows[1:]
    picks, seen = [], set()
    def issue(r):
        if not any(r):
            return "blank"
        if r[4] != r[4].strip() or "  " in r[4] or r[4].islower() or r[4].isupper():
            return "product"
        if r[2].strip() not in ("Netherlands", "Belgium", "Germany", "France"):
            return "country"
        return None
    dup = next(r for i, r in enumerate(body) if any(r) and r in body[i + 1:])
    clean = [r for r in body if any(r) and issue(r) is None]
    for r in body:
        kind = issue(r)
        if kind and kind not in ("blank",) and sum(k == kind for _, k in picks) < 3:
            picks.append((r, kind))
        if len(picks) == 6:
            break
    order = [(clean[0], None), picks[0], (dup, "dup"), (dup, "dup"), picks[1], (clean[1], None),
             picks[2], (["", "", "", "", "", ""], "blank"), picks[3], picks[4], (clean[2], None), picks[5]]
    cols = (0, 2, 4)  # OrderID, Country, Product: the columns where the mess is visible
    lines = ["<span class='muted'>" + ",".join(header[c] for c in cols) + ",…</span>"]
    for r, kind in order:
        cells = [html.escape(r[c]) for c in cols]
        if kind == "product":
            cells[2] = f"<span class='hl'>\"{html.escape(r[4])}\"</span>"
        elif kind == "country":
            cells[1] = f"<span class='hl'>{html.escape(r[2])}</span>"
        text = ",".join(cells) + ",…"
        if kind == "blank":
            text = "<span class='hl'>,,,,,,,,</span>"
        elif kind == "dup":
            text = f"<span class='hl'>{text}</span>"
        lines.append(text)
    return "\n".join(lines)


def before_after(title, subtitle, before_html, after_img, after_w=580, note="sample project"):
    return page(f"""
<div class=top><h1>{title}</h1><p>{subtitle}</p></div>
<div class=row>
  <div class=card style='width:520px;height:470px'><span class='tag bad'>BEFORE</span>
    <div style='overflow:hidden;height:432px'><pre style='margin-top:14px;font-size:16px;line-height:1.6'>{before_html}</pre></div></div>
  <div class=arrow>&rarr;</div>
  <div class=card style='width:{after_w + 36}px'><span class='tag good'>AFTER</span>
    <img class=shot src='{after_img}' style='width:{after_w}px;margin-top:10px'><div class=note>{note}</div></div>
</div>""")


def render(name: str, doc: str) -> None:
    src = OUT / f"{name}.html"
    src.write_text(doc, encoding="utf-8")
    png = OUT / f"{name}.png"
    subprocess.run([EDGE, "--headless=new", "--disable-gpu", "--hide-scrollbars", "--force-device-scale-factor=1",
                    "--window-size=1280,769", f"--screenshot={png}", src.as_uri()], check=True,
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=60)
    print(png.name, png.stat().st_size // 1024, "KB")


def main() -> None:
    OUT.mkdir(exist_ok=True)
    raw = raw_sample()
    dash = (SHOTS / "excel_dashboard.png").as_uri()
    pyrep = (SHOTS / "python_summary.png").as_uri()

    render("excel-1-main", before_after("Excel VBA &amp; Power Query", "Messy exports &rarr; clean data &rarr; one-click report",
                                        raw, dash, 580))
    render("excel-2-dashboard", page(f"""
<div class=top style='height:150px'><h1 style='font-size:46px'>12 exports &rarr; 1 click</h1>
<p style='font-size:24px'>Sample project &middot; Power Query + VBA + dashboard</p></div>
<div style='padding:22px 56px 0;display:flex;gap:30px'>
<img class=shot src='{dash}' style='width:760px'>
<div class=caption style='padding:0;font-size:24px'>
<p>Power Query merges and cleans <b>12 monthly CSV exports</b>.</p><br>
<p>One button refreshes everything and saves a <b>dated PDF</b>.</p><br>
<p class=muted style='font-size:21px'>&asymp; 2&ndash;3 hours of copy-paste a month by hand (our estimate).</p></div></div>"""))
    render("excel-3-quality", page(f"""
<div class=top style='height:150px'><h1 style='font-size:46px'>What the automation fixes</h1>
<p style='font-size:24px'>Sample project &middot; every refresh, automatically</p></div>
<div style='display:flex;gap:40px;padding:40px 56px 0'>
<div class=stat><p><b>4,901</b>raw lines read</p><p><b>20</b>blank lines removed</p><p><b>89</b>duplicate lines removed</p>
<p><b>1,498</b>product names corrected</p><p><b>3,604</b>country values standardised</p><p><b>4,792</b>clean order lines</p></div>
<pre class=code style='font-size:15px'><span class=c>' Does the work without dialogs, so it can also run from a scheduler.</span>
<span class=k>Public Function</span> DoRefreshAndExport() <span class=k>As String</span>
    ThisWorkbook.RefreshAll
    Application.CalculateUntilAsyncQueriesDone
    Application.CalculateFull
    pdfPath = ThisWorkbook.Path &amp; <span class=s>"\\Sales_Report_"</span> &amp; _
              Format(Now, <span class=s>"yyyy-mm-dd_hhnn"</span>) &amp; <span class=s>".pdf"</span>
    ThisWorkbook.Worksheets(<span class=s>"Dashboard"</span>).ExportAsFixedFormat _
        Type:=xlTypePDF, Filename:=pdfPath
    DoRefreshAndExport = pdfPath
<span class=k>End Function</span></pre></div>
<div class=caption style='padding-top:40px'>Runs in desktop Excel &middot; <b>no add-ins</b> &middot; <b>no locked macros</b>
<span class=muted>&nbsp;&middot; you get the file and a short guide</span></div>"""))

    render("python-1-main", before_after("Python Automation", "A folder of messy files &rarr; a formatted Excel report in 0.3 s",
                                         raw, pyrep, 580))
    render("python-2-run", page(f"""
<div class=top style='height:150px'><h1 style='font-size:46px'>One command, done</h1>
<p style='font-size:24px'>Sample project &middot; standard library + xlsxwriter, fully commented</p></div>
<div style='display:flex;gap:30px;padding:30px 56px 0'>
<pre class=code style='font-size:16px;width:640px'><span class=c>$ python build_report.py --input exports/</span>
Read 12 files, 4,901 lines
  removed 20 blank and 89 duplicate lines
  corrected 1498 product names,
  standardised 3604 country values
  unknown products: 0, unknown countries: 0
Wrote Sales_Report_2025.xlsx:
  4,792 clean lines in 0.25 s</pre>
<img class=shot src='{pyrep}' style='width:470px;align-self:flex-start'></div>
<div class=caption style='padding-top:34px'>You get: <b>the script</b> &middot; <b>requirements.txt</b> &middot; <b>a step-by-step guide</b>
<span class=muted>&nbsp;&middot; your files never leave your computer</span></div>"""))

    sheets_before = """<table class=grid><tr><th></th><th>A</th><th>B</th><th>C</th></tr>
<tr><th>1</th><td>Order</td><td>Country</td><td>Product</td></tr>
<tr><th>2</th><td>1002</td><td><span class=hl>netherlands</span></td><td><span class=hl>&nbsp;v60 dripper </span></td></tr>
<tr><th>3</th><td><span class=hl>1002</span></td><td><span class=hl>netherlands</span></td><td><span class=hl>&nbsp;v60 dripper </span></td></tr>
<tr><th>4</th><td>1003</td><td><span class=hl>DE</span></td><td><span class=hl>EARL GREY 100G</span></td></tr>
<tr><th>5</th><td></td><td></td><td></td></tr>
<tr><th>6</th><td>1004</td><td><span class=hl>België</span></td><td>Matcha 30g</td></tr>
<tr><th>7</th><td><span class=hl>#REF!</span></td><td></td><td></td></tr></table>"""
    sheets_after = """<table class=grid><tr><th></th><th>A</th><th>B</th><th>C</th><th>D</th></tr>
<tr><th>1</th><td>Order</td><td>Country</td><td>Product</td><td>Revenue</td></tr>
<tr><th>2</th><td>1002</td><td>Netherlands</td><td>V60 Dripper</td><td>€27.00</td></tr>
<tr><th>3</th><td>1003</td><td>Germany</td><td>Earl Grey 100g</td><td>€7.90</td></tr>
<tr><th>4</th><td>1004</td><td>Belgium</td><td>Matcha 30g</td><td>€21.00</td></tr></table>
<div style='margin-top:18px;font-size:15px;color:#6B7780'>Menu <b style='color:#1F2A30'>Automation &rsaquo; Import &amp; clean exports</b></div>
<div style='margin-top:12px;display:flex;gap:10px;align-items:flex-end;height:90px'>
<div style='width:46px;height:60px;background:#2F3E46'></div><div style='width:46px;height:44px;background:#2F3E46'></div>
<div style='width:46px;height:78px;background:#2F3E46'></div><div style='width:46px;height:88px;background:#C8773A'></div></div>"""
    render("sheets-1-main", page(f"""
<div class=top><h1>Google Sheets Automation</h1><p>Formulas &middot; Apps Script &middot; dashboards</p></div>
<div class=row>
  <div class=card style='width:520px;height:470px'><span class='tag bad'>BEFORE</span><div style='margin-top:16px'>{sheets_before}</div></div>
  <div class=arrow>&rarr;</div>
  <div class=card style='width:560px;height:470px'><span class='tag good'>AFTER</span><div style='margin-top:16px'>{sheets_after}</div>
  <div class=note>illustration</div></div>
</div>"""))
    render("sheets-2-script", page(r"""
<div class=top style='height:150px'><h1 style='font-size:46px'>One menu click imports &amp; cleans</h1>
<p style='font-size:24px'>Apps Script sample &middot; readable code you own</p></div>
<pre class=code style='margin:22px 56px 0;font-size:15px;line-height:1.5'><span class=c>/** Reads every CSV in the export folder, drops blank and duplicate lines, writes one clean table. */</span>
<span class=k>function</span> importExports() {
  <span class=k>const</span> sheet = book.getSheetByName(<span class=s>'Clean'</span>) || book.insertSheet(<span class=s>'Clean'</span>);
  <span class=k>const</span> files = DriveApp.getFolderById(folderId).getFilesByType(MimeType.CSV);
  <span class=k>while</span> (files.hasNext()) {
    <span class=k>const</span> [head, ...lines] = Utilities.parseCsv(files.next().getBlob().getDataAsString(<span class=s>'UTF-8'</span>));
    <span class=k>for</span> (<span class=k>const</span> line <span class=k>of</span> lines) {
      <span class=k>const</span> key = line.join(<span class=s>'\u0001'</span>);
      <span class=k>if</span> (!line[0] || seen.has(key)) <span class=k>continue</span>;   <span class=c>// blank or duplicate line</span>
      seen.add(key);
      rows.push(line.map((v) =&gt; String(v).replace(<span class=s>/\s+/g</span>, <span class=s>' '</span>).trim()));
    }
  }
  <span class=c>// setValues() needs a rectangle: pad short rows, trim long ones to one width.</span>
  sheet.getRange(2, 1, rows.length, width).setValues(rows.map(pad));
  book.toast(<span class=s>`${rows.length} clean rows imported`</span>, <span class=s>'Automation'</span>);
}</pre>
<div class=caption style='padding-top:14px;font-size:20px'><span class=muted>Excerpt &middot; full file: sheets-apps-script/Code.gs</span></div>"""))


if __name__ == "__main__":
    main()
