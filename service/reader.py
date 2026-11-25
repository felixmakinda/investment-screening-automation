"""Reads PDFs on this machine. No AI is used here."""
import re, subprocess
from pathlib import Path

class BadDocument(ValueError): pass

TOK = r"(?:\([\d,.]+\)|\([—–−-]\)|[—–−]|\d[\d,]*(?:\.\d+)?%?)"

def pdf_pages(path):
    if Path(path).read_bytes()[:5] != b"%PDF-": raise BadDocument(f"{Path(path).name} is not a PDF")
    r = subprocess.run(["pdftotext", "-layout", str(path), "-"], capture_output=True, text=True, timeout=60)
    if r.returncode: raise BadDocument(f"{Path(path).name} could not be read (damaged or password protected)")
    pages = r.stdout.split("\f")
    if pages and not pages[-1].strip(): pages.pop()
    if sum(len(p.strip()) for p in pages) < 300: raise BadDocument(f"{Path(path).name} has no readable text (it may be a scan)")
    return pages

def number(tok):
    """'—' means nil (0). '(−)' means the figure was not printed (None)."""
    if tok in ("—", "(—)", "–"): return 0
    if tok in ("−", "(−)", "(-)"): return None
    v = float(re.sub(r"[^\d.]", "", tok))
    if tok.endswith("%"): return v / 100
    v = int(v) if v == int(v) else v
    return -v if tok.startswith("(") else v

def statements(pages):
    """Returns {'unit', 'years', 'lines': [{st,label,page,values}]}"""
    st, years, unit, lines, pend, open_row = None, [], None, [], "", None
    for pn, page in enumerate(pages, 1):
        for raw in page.split("\n"):
            s = raw.strip()
            if not s: continue
            if re.match(r"^(?:[A-Z]\.\s*)?(?:abbreviated\s+)?notes\b", s, re.I): st = None
            head = re.search(r"statement of|income statement|balance sheet|statement of cash", s, re.I)
            if head and not re.search(TOK + r"\s+" + TOK, s):
                st = "PL" if re.search(r"profit|income", s, re.I) else "BS" if re.search(r"position|balance", s, re.I) else "CF" if re.search(r"cash", s, re.I) else st
                m = re.search(r"\(([A-Z]{3})\s*[’']000\)", s)
                if m: unit = (m.group(1), 1000)
                pend, open_row = "", None; continue
            ys = re.findall(r"\b(20\d\d)F?\b", s)
            if len(ys) >= 3 and not re.search(r"\d,\d{3}", s):
                years = [y + ("F" if y + "F" in s else "") for y in ys]; continue
            if not st or not years: continue
            n = len(years)
            m = re.match(rf"^(.*?)\s*((?:(?<![\d,(]){TOK}\s+){{{n-1}}}{TOK})$", s)
            if m:
                label = m.group(1).strip()
                if not label: label = pend
                row = dict(st=st, label=label, page=pn, values=dict(zip(years, [number(t) for t in re.findall(TOK, m.group(2))])))
                lines.append(row); open_row = row if label.count("(") > label.count(")") else None; pend = ""
            elif open_row is not None:
                open_row["label"] += " " + s
                if open_row["label"].count("(") <= open_row["label"].count(")"): open_row = None
            else: pend = s
    return dict(unit=unit, years=years, lines=lines)

key = lambda t: re.sub(r"[^a-z0-9]+", "", t.lower().replace("’", "'"))
base = lambda t: key(re.sub(r"\([^)]*\)", "", t))

def recognise(lines, names):
    """Matches printed labels to known lines. Returns ({(st,name): row}, [unrecognised])."""
    out, unknown = {}, []
    for row in lines:
        hit = next((n for n in names if n["st"] == row["st"] and any(key(row["label"]) == key(x) for x in [n["name"], *n["alts"]])), None) \
           or next((n for n in names if n["st"] == row["st"] and base(row["label"]) and not row["label"].startswith("(") and any(base(row["label"]) == base(x) for x in [n["name"], *n["alts"]])), None)
        if hit and (row["st"], hit["name"]) not in out: out[(row["st"], hit["name"])] = row
        else: unknown.append(row)
    return out, unknown

def memo_table(pages, columns):
    """Summary table in the memo: rows that start with a year."""
    out = {}
    for pn, page in enumerate(pages, 1):
        for raw in page.split("\n"):
            m = re.match(rf"^\s*(20\d\d)(F?)\s+((?:{TOK}\s+){{{len(columns)-1}}}{TOK})\s*$", raw)
            if m: out[m.group(1) + m.group(2)] = dict(zip(columns, [number(t) for t in re.findall(TOK, m.group(3))]), page=pn)
    return out

def suspicious(pages):
    pat = r"(ignore|disregard)[^.]{0,40}(instruction|prompt|above|previous)|as an ai\b|language model|\bscore (this|the|us|it)\b[^.]{0,40}\b(3|three|high|highest|maximum)\b|(give|assign|award)[^.]{0,40}\b(top|highest|maximum|full)\b[^.]{0,20}\b(score|rating)|system prompt"
    return [(pn, re.sub(r"\s+", " ", m.group(0))) for pn, p in enumerate(pages, 1) for m in re.finditer(pat, re.sub(r"\s+", " ", p), re.I)]
