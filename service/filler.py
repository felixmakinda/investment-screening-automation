"""Fills AHL's template, then reads it back to prove nothing else changed."""
import shutil, subprocess, tempfile, warnings, zipfile
from pathlib import Path
import openpyxl
from lxml import etree
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from checks import country, money
from config import TEMPLATE
warnings.simplefilter("ignore")
FO, TMS = "Financial Overview", "TMS Assessment"
WRAP = Alignment(wrap_text=True, vertical="top")

def plan(run, K, cfg, info, approvals):
    """Every cell to write: {(sheet, cell): value}, plus a log of where each figure came from."""
    w, log, ex = {}, [], []
    years, facts = run["years"][-4:], run["facts"]
    cols = dict(zip(years, "CDEF")); F = lambda f: facts.get(f, {}).get("value")
    put = lambda sh, c, v: w.__setitem__((sh, c), v)
    val = lambda st, n, y: (K.get((st, n)) or {"values": {}})["values"].get(y)
    ccy = run["unit"][0] if run["unit"] else None
    if ccy: put(FO, "B5", ccy)
    for y, c in cols.items():
        put(FO, f"{c}5", int(y[:4])); put(FO, f"{c}66", int(y[:4]))
        if ccy == "USD": put(FO, f"{c}93", 1); put(FO, f"{c}94", 1)
        sd = val("PL", "Selling & distribution expenses", y)
        if sd is not None: put(FO, f"{c}121", -sd)
        for r in cfg["rows"]:
            parts = [(l, val(r["st"], l, y), K.get((r["st"], l))) for l in r["lines"]]
            got = [p for p in parts if p[1] is not None]
            if len(got) < len(parts): ex.append(("V06", f"{r['label']} {y}: missing {', '.join(p[0] for p in parts if p[1] is None)}."))
            if len(got) == len(parts):
                put(FO, f"{c}{r['row']}", sum(p[1] for p in got))
                log.append((FO, f"{c}{r['row']}", f"{r['label']} {y}", sum(p[1] for p in got), "Accounts", ", ".join(sorted({f"p{p[2]['page']}" for p in got})), " + ".join(f"{p[0]} [{p[1]:,}]" for p in got), "", "", "Read directly, no AI" + (f". {r['note']}" if r["note"] else "")))
    prior = run["years"][-5] if len(run["years"]) > 4 else None
    if prior: put(FO, "C82", val("BS", "Cash & cash equivalents", prior))
    put(FO, "E82", "=D83"); put(FO, "F82", "=E83")
    for a in "IJKL": put(FO, f"{a}28", f"={a}5")
    if F("debt_usd"): put(FO, "L99", F("debt_usd") / 1000)
    # screening
    ids = {i for i, _ in run["exceptions"]}; note = lambda i: f" See {i}." if i in ids else ""
    put(TMS, "B5", F("company_name")); put(TMS, "B8", F("raise_total_usd"))
    cty = country(F("country_main") or "", info["tier"]); put(TMS, "B12", cty)
    if F("country_main") and not cty: ex.append(("V07", f"B12: {F('country_main')} is not in AHL's country list."))
    com = {}
    if F("debt_usd") and F("raise_total_usd"): com["C8"] = f"{money(F('debt_usd'))} senior secured debt" + (f" plus {money(F('raise_total_usd')-F('debt_usd'))} equity." if F("raise_total_usd") > F("debt_usd") else ".") + note("E01")
    com["C12"] = " ".join(filter(None, [f"Countries today: {F('countries_current')}." if F("countries_current") else "", f"Planned: {F('country_planned')}." if F("country_planned") else ""])) + note("E10")
    for c, s in run["screening"].items():
        if c not in ("B5", "B8", "B12"):
            put(TMS, c, s["value"])
            if int(c[1:]) >= 8: com[f"C{c[1:]}"] = "Suggested. " + s["comment"] + (note("E20") + note("E02") if c == "B15" else note("E19") if c == "B13" else "")
    for c, v in com.items():
        if v.strip(): put(TMS, c, v)
    for c, s in run["scores"].items():
        put(TMS, c, s["score"]); put(TMS, f"G{c[1:]}", s["comment"] + (note("E11") if c == "D36" else note("E06") if c == "D35" else ""))
    for s in cfg["scores"]:
        if s["cell"] not in run["scores"]: put(TMS, f"G{s['cell'][1:]}", "No information found. Left blank for a person to complete.")
    for key, v in (approvals or {}).items():            # reviewer's decisions win
        if key in {x["cell"] for x in cfg["screening"] + cfg["scores"]}:
            put(TMS, key, v); log.append((TMS, key, "Reviewer's decision", v, "Reviewer", "", "", "", "", "Changed or confirmed by reviewer"))
    for c, s in run["scores"].items(): log.append((TMS, c, next(x["label"] for x in cfg["scores"] if x["cell"] == c), s["score"], "Suggested" if s["comment"].startswith("Sug") else "Calculated", "", s["comment"], "", "", ", ".join(s["evidence"])))
    for c, s in run["screening"].items(): log.append((TMS, c, next(x["label"] for x in cfg["screening"] if x["cell"] == c), s["value"], "Suggested", "", s["comment"], "", "", ", ".join(s["evidence"])))
    for f, v in facts.items(): log.append(("", "", cfg["memo_fields"][f]["question"], v["value"], "Memo", f"p{v['page']}" if v["page"] else "", v["quote"], v["reader_a"], v["reader_b"], f"{v['status']}: {v['reason']}"))
    return w, log, ex

def write(w, out, budget=False):
    wb = openpyxl.load_workbook(TEMPLATE); f, t = wb[FO], wb[TMS]
    # dropdowns openpyxl cannot read, added back exactly
    for ws, c, ref in ((f, "B5", "'Currency Data'!$B$4:$B$17"), (t, "B12", "'Country Gating Screen'!$C$9:$C$62")):
        d = DataValidation(type="list", formula1=ref, allow_blank=True); ws.add_data_validation(d); d.add(c)
    cf = ConditionalFormattingList()
    for x in t.conditional_formatting:
        for rule in x.rules: cf.add("B9:B15" if str(x.sqref) == "B9:C15" else str(x.sqref), rule)
    t.conditional_formatting = cf
    for (sh, c), v in w.items():
        cell = wb[sh][c]; own = cell.value if sh == TMS and c[0] == "G" else None
        if v is None: continue
        cell.value = f"{v} [AHL note kept: {own}]" if own else v
        if isinstance(cell.value, str) and cell.value.startswith("=") and (sh, c) not in {(FO, x) for x in ("E82", "F82", "I28", "J28", "K28", "L28")}: cell.data_type = "s"   # text from a document is never a formula
        if sh == TMS and c[0] in "CG" or c == "B3":
            cell.alignment = WRAP; cell.number_format = "General"; r = cell.row
            t.row_dimensions[r].height = max(t.row_dimensions[r].height or 15, 15 * (1 + len(str(cell.value)) // (80 if c[0] == "G" else 62)))
    for c in ("I28", "J28", "K28", "L28"): f[c].number_format = "0"
    if budget:
        for c in ("F5", "F28", "F66", "L5", "L28", "L66"): f[c].number_format = '0"F"'
    wb.save(out)

def add_tabs(path, log, exceptions, checks):
    wb = openpyxl.load_workbook(path)
    blue, band = PatternFill("solid", fgColor="33305C"), {"High": "F8CBAD", "Medium": "FFE699", "Low": "E2EFDA"}
    if not any(r[8] for r in log): log, second = [r[:8] + r[9:] for r in log], False      # no second reader was used
    else: second = True
    one = lambda x: x if second else x[:8] + x[9:]
    for title, head, rows, widths in (("Exceptions", ["ID", "Severity", "Issue", "Evidence (accounts in USD '000)", "Why it matters", "What was done", "Question for the company"], exceptions, [7, 10, 34, 70, 46, 42, 60]),
        ("Extraction Log", one(["Sheet", "Cell", "Field", "Value", "Source", "Page", "Quote or workings", "AI answer", "Second reader", "Status"]), log, one([18, 7, 40, 34, 11, 8, 70, 30, 22, 50])),
        ("Checks", ["Check", "Result", "Detail"], checks, [44, 16, 110])):
        ws = wb.create_sheet(title); ws.append(head)
        for c in ws[1]: c.font, c.fill = Font(bold=True, color="FFFFFF"), blue
        for r in rows:
            ws.append(list(r)); n = ws.max_row
            for c in ws[n]:
                if isinstance(c.value, str) and c.value.startswith("="): c.data_type = "s"   # text, never a formula
            ws.row_dimensions[n].height = 14.5 * max(1, *(-(-len(str(x or "")) * 1.15 // wd) for x, wd in zip(r, widths)))
        for row in ws.iter_rows(min_row=2):
            for c in row:
                c.alignment = WRAP
                if c.value in band and c.column == 2: c.fill = PatternFill("solid", fgColor=band[c.value])
        for i, wd in enumerate(widths): ws.column_dimensions["ABCDEFGHIJ"[i]].width = wd
        ws.freeze_panes = "A2"; ws.auto_filter.ref = ws.dimensions
    wb.save(path)

def recalc(path):
    """Calculates every formula with a private LibreOffice, so an open copy is not disturbed."""
    tmp = Path(tempfile.mkdtemp())
    try:
        subprocess.run(["soffice", f"-env:UserInstallation=file://{tmp}/profile", "--headless", "--convert-to", "xlsx", "--outdir", str(tmp / "o"), str(path)], capture_output=True, timeout=180, check=True)
        return openpyxl.load_workbook(tmp / "o" / Path(path).name, data_only=True)
    finally: shutil.rmtree(tmp, ignore_errors=True)

def store_results(path, calc):
    """Saves calculated results in the file so previews and phones show numbers, not blanks."""
    ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}; M = "{%s}" % ns["m"]
    names = openpyxl.load_workbook(path, read_only=True).sheetnames; tmp = Path(str(path) + ".tmp")
    with zipfile.ZipFile(path) as zi, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
        for it in zi.infolist():
            data = zi.read(it.filename)
            if it.filename.startswith("xl/worksheets/sheet"):
                ws = calc[names[int("".join(x for x in it.filename.split("/")[-1] if x.isdigit())) - 1]]; root = etree.fromstring(data)
                for c in root.iterfind(".//m:c[m:f]", ns):
                    v = ws[c.get("r")].value; e = c.find("m:v", ns)
                    if e is None: e = etree.SubElement(c, M + "v")
                    if v is None: c.remove(e); c.attrib.pop("t", None); continue
                    if isinstance(v, str): c.set("t", "e" if v.startswith("#") else "str"); e.text = v
                    else: c.attrib.pop("t", None); e.text = repr(v) if not isinstance(v, bool) else str(int(v))
                data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            zo.writestr(it, data)
    tmp.replace(path)

def read_back(path, w, calc):
    """Confirms what was written, that nothing else changed, and that no formula is broken."""
    a, b, problems = openpyxl.load_workbook(TEMPLATE), openpyxl.load_workbook(path), []
    for ws in a:
        for row in b[ws.title].iter_rows():
            for c in row:
                if (ws.title, c.coordinate) not in w and c.value != ws[c.coordinate].value: problems.append(f"{ws.title}!{c.coordinate} changed unexpectedly")
    for (sh, c), v in w.items():
        got = b[sh][c].value
        if v is not None and got != v and not str(got).startswith(str(v)): problems.append(f"{sh}!{c}: wrote {v}, found {got}")
    # the template itself divides by the number of scores, so an empty group shows an error there
    empty = any(not any((TMS, f"D{r}") in w and w[(TMS, f"D{r}")] for r in g) for g in (range(23, 30), range(32, 39), range(41, 45)))
    for ws in a:
        for row in calc[ws.title].iter_rows():
            for c in row:
                if empty and ws.title == TMS and c.coordinate in ("D18", "D22", "D31", "D40", "D47"): continue
                if isinstance(c.value, str) and c.value.startswith("#"): problems.append(f"{ws.title}!{c.coordinate} shows {c.value}")
    want = {ws.title: sum(len(d.sqref.ranges) for d in ws.data_validations.dataValidation) for ws in b}
    if want[FO] < 1 or want[TMS] < 6: problems.append("A dropdown is missing")
    return problems

def balance_checks(calc, years, K):
    f, t, out = calc[FO], calc[TMS], []
    for y, c in zip(years[-4:], "CDEF"):
        out.append((f"Balance sheet balances, {y}", "Pass" if abs(f[f"{c}58"].value or 0) < 1 else "FAIL", f"Difference {f[f'{c}58'].value:,.0f}"))
    for y, c in zip(years[-4:], "CDEF"):
        for n, r, src in (("EBITDA", 16, "EBITDA"), ("PBT", 22, "Profit before tax"), ("PAT", 24, "Profit for the year (PAT)")):
            s = K[("PL", src)]["values"][y]; d = f[f"{c}{r}"].value - s
            out.append((f"{n} agrees with accounts, {y}", "Pass" if abs(d) < 1 else "Exception", f"Template {f[f'{c}{r}'].value:,.0f}, accounts {s:,.0f}, difference {d:,.0f}" + (f" ({fi:,.0f} of it is finance income, which the template counts in EBITDA)" if n == "EBITDA" and (fi := K[("PL", "Finance income (interest on cash)")]["values"][y] or 0) else "")))
        out.append((f"Cash flow explains cash, {y}", "Pass" if abs(f[f"{c}85"].value) < 1 else "Exception", f"Cumulative difference {f[f'{c}85'].value:,.0f}. Not forced to balance."))
    num = lambda v: v if isinstance(v, (int, float)) else None
    return out, dict(score=num(t["D47"].value), team=num(t["D22"].value), model=num(t["D31"].value), scale=num(t["D40"].value), risk=t["B14"].value, coverage=f["L104"].value, debt_equity=f["E88"].value)
