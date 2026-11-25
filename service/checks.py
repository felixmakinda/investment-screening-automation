"""Plain rules that decide what is trusted. No AI is used here."""
import json, re, unicodedata
import openpyxl
from config import TEMPLATE

norm = lambda s: re.sub(r"[^a-z0-9%$.]+", "", str(s).lower())
plain = lambda s: re.sub(r"[^a-z]", "", unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode().lower())
WORDS = {w: i for i, w in enumerate("zero one two three four five six seven eight nine ten eleven twelve".split())} | {"fifteen": 15, "twenty": 20, "thirty": 30, "forty": 40, "fifty": 50, "sixty": 60, "seventy": 70, "eighty": 80, "ninety": 90}
MULT = {"k": 1e3, "thousand": 1e3, "m": 1e6, "million": 1e6, "bn": 1e9, "billion": 1e9}
money = lambda v: f"${v:,.0f}" if abs(v) < 1e6 else f"${v/1e6:,.1f}m".replace(".0m", "m")
pct = lambda v: f"{v:.0%}" if abs(v * 100 - round(v * 100)) < .05 else f"{v:.1%}"

def tidy(t): return re.sub(r"\s*[\u2014\u2013]\s+", ", ", str(t or "")).strip()

def parse_json(x):
    """Accepts a dict or an AI's raw text, with or without code fences."""
    if isinstance(x, dict): return x
    m = re.search(r"\{.*\}", str(x or ""), re.S)
    try: return json.loads(m.group(0)) if m else None
    except json.JSONDecodeError: return None

def numbers(text):
    t, out = str(text).lower(), set()
    for m in re.finditer(r"(\d[\d,]*\.?\d*)\s*(%|percent|thousand|million|billion|k\b|m\b|bn\b)?", t):
        v = float(m.group(1).replace(",", "").rstrip(".")); u = m.group(2)
        out |= {v, v / 100} if u in ("%", "percent") else {v * MULT.get(u, 1)}
    for w in re.findall(r"[a-z]+", t):
        if w in WORDS: out |= {float(WORDS[w]), WORDS[w] / 100 if "percent" in t else float(WORDS[w])}
    return out

def typed(value, kind):
    if value is None or str(value).strip() == "" or str(value).strip().lower() in ("null", "none", "n/a", "not stated"): return None
    if kind in ("number", "percent"):
        n = sorted(numbers(value), key=abs)
        if not n: raise ValueError("not a number")
        v = float(value) if isinstance(value, (int, float)) else n[0] if kind == "percent" else n[-1]
        return v / 100 if kind == "percent" and v > 1 else int(v) if kind == "number" and v == int(v) else v
    return re.sub(r"\s+", " ", str(value)).strip()

def verify(ans, field, pages):
    """One reader's answer. Returns (status, value, quote, page, reason)."""
    try: v = typed(ans.get("value"), field["type"])
    except ValueError: return "blocked", None, ans.get("quote"), ans.get("page"), "V02: value is not the expected type"
    if v is None: return "blank", None, None, None, "Not stated"
    q = str(ans.get("quote") or "").strip()
    if len(norm(q)) < 6: return "blocked", v, q, ans.get("page"), "V01: no quote given"
    found = [i for i, p in enumerate(pages, 1) if norm(q) in norm(p)]
    if not found: return "blocked", v, q, ans.get("page"), "V01: quote is not in the memo"
    page = ans.get("page") if ans.get("page") in found else found[0]
    if field["type"] in ("number", "percent"):
        if not any(abs(v - n) <= 1e-6 * max(1, abs(v)) for n in numbers(q)): return "blocked", v, q, page, "V02: number is not in its quote"
    else:
        words = re.findall(r"[^\W\d_][\w'’-]*", v)
        names = [w for i, w in enumerate(words) if w[0].isupper() and w not in ("Yes", "No", "None") and (i > 0 or len(words) <= 3 and all(x[0].isupper() for x in words))]
        miss = [w for w in names if plain(w) not in plain(q)] + [d for d in re.findall(r"\d+", v) if d not in re.sub(r"[,]", "", q)]
        if miss: return "blocked", v, q, page, f"V02: not in its quote: {', '.join(miss)}"
    return "accepted", v, q, page, ""

def agree(a, b, kind):
    if kind in ("number", "percent"): return abs(a[1] - b[1]) <= 1e-6 * max(1, abs(a[1]))
    x, y = norm(a[1]), norm(b[1])
    if x == y or x in y or y in x: return True
    p, q = norm(a[2]), norm(b[2])
    wa, wb = set(re.findall(r"\w+", a[1].lower())), set(re.findall(r"\w+", b[1].lower()))
    return p in q or q in p or len(wa & wb) / max(1, len(wa | wb)) >= .4

def sentence(quote, page_text):
    """The whole memo sentence around a verified quote, so a rule never depends on how much the AI quoted."""
    return " ".join(s for s in re.split(r"(?<=[.!?])\s+", re.sub(r"\s+", " ", page_text)) if quote and norm(quote) in norm(s)) or quote

def combine(reader_a, reader_b, fields, pages):
    get = lambda r: {x.get("field_id"): x for x in (parse_json(r) or {}).get("answers", []) if isinstance(x, dict)}
    A, B = get(reader_a), get(reader_b); both = bool(A) and bool(B); facts, ex = {}, []
    second = reader_b is not None      # a second reader is optional
    if not A and not B: raise ValueError("Neither AI reader returned usable answers")
    if second and not both: ex.append(("V09", f"Reader {'A' if A else 'B'} answered; the other did not."))
    elif both:
        given = {n: sum(str((r.get(i) or {}).get("value") or "").strip().lower() not in ("", "null", "none", "n/a", "not stated") for i in fields) for n, r in (("A", A), ("B", B))}
        low = min(given, key=given.get)
        if max(given.values()) - given[low] >= 10: ex.append(("V09", f"Reader {low} answered {given[low]} of {len(fields)} questions; the other answered {max(given.values())}."))
    for fid, f in fields.items():
        res = {n: verify(r[fid], f, pages) if fid in r else ("missing", None, None, None, "Not answered") for n, r in (("a", A), ("b", B)) if r}
        ok = [x for x in res.values() if x[0] == "accepted"]; bad = [x for x in res.values() if x[0] == "blocked"]
        out = dict(value=None, quote=None, page=None, status="blank", reason="Not stated", reader_a=A.get(fid, {}).get("value"), reader_b=B.get(fid, {}).get("value"))
        take = lambda x, st, why: out.update(value=x[1], quote=x[2], page=x[3], status=st, reason=why)
        if len(ok) == 2 and agree(*ok, f["type"]): take(ok[0], "accepted", "Both readers agree; quote verified")
        elif len(ok) == 2:
            out.update(status="review", reason="V03: readers disagree"); ex.append(("V03", f"{f['question']}: '{ok[0][1]}' (memo p{ok[0][3]}) or '{ok[1][1]}' (memo p{ok[1][3]})."))
        elif len(ok) == 1 and any(x[0] == "blank" for x in res.values()):
            out.update(status="review", reason="V03: one reader found an answer, the other says it is not stated"); ex.append(("V03", f"{f['question']}: one reader answered '{ok[0][1]}' (memo p{ok[0][3]}: \"{ok[0][2][:120]}\"); the other says it is not stated."))
        elif len(ok) == 1: take(ok[0], "accepted", "Quote verified in the memo" if not second else "One reader; quote verified" + ("; the other was blocked" if bad else ""))
        elif bad:
            out.update(status="blocked", reason=bad[0][4]); ex.append((bad[0][4][:3], f"{f['question']}: '{bad[0][1]}'. {bad[0][4][5:]}."))
        if f["required"] and out["value"] is None: ex.append(("V04", f["question"]))
        out["sentence"] = sentence(out["quote"], pages[out["page"] - 1]) if out["page"] else None
        facts[fid] = out
    return facts, ex

def template_info():
    wb = openpyxl.load_workbook(TEMPLATE, data_only=True)
    c, g, t = wb["Currency Data"], wb["Country Gating Screen"], wb["TMS Assessment"]
    yrs = {i: c.cell(3, i).value for i in (3, 4, 5, 6)}
    return dict(fx={c.cell(r, 2).value: {yrs[i]: c.cell(r, i).value for i in yrs if c.cell(r, i).value} for r in range(4, 18)},
                tier={g.cell(r, 3).value: g.cell(r, 4).value for r in range(9, 63) if g.cell(r, 3).value},
                guide={f"D{r}": f"{t.cell(r,1).value}: {t.cell(r,2).value} Guide: {t.cell(r,3).value}" for r in [*range(23, 30), *range(32, 39), *range(41, 45)]})

def country(name, tier): return next((k for k in tier if plain(k) == plain(name)), None)

def run_rules(K, years, facts, memo_tbl, pages_fin, cfg, info):
    S, ex = cfg["set"], []
    F = lambda f: facts.get(f, {}).get("value")
    V = lambda st, n, y: (K.get((st, n)) or {"values": {}})["values"].get(y)
    Z = lambda st, n, y: V(st, n, y) or 0
    add = lambda i, e: ex.append((i, e))
    P, N = list(zip(years, years[1:])), lambda v: f"{v:,.0f}"
    actual = [y for y in years if not y.endswith("F")]; last = actual[-1]; fc = [y for y in years if y.endswith("F")]
    L = dict(cur="Loan receivables – current", nc="Loan receivables – non-current", st="Borrowings – short-term", lt="Borrowings – long-term")
    book = {y: Z("BS", L["cur"], y) + Z("BS", L["nc"], y) for y in years}; debt = {y: Z("BS", L["st"], y) + Z("BS", L["lt"], y) for y in years}
    eq = {y: Z("BS", "Total equity", y) for y in years}; rev = {y: Z("PL", "Revenue", y) for y in years}; pat = {y: Z("PL", "Profit for the year (PAT)", y) for y in years}
    # E01 raise
    tot, deb, intro = F("raise_total_usd"), F("debt_usd"), F("raise_intro_usd")
    iq = str(facts.get("raise_intro_usd", {}).get("quote") or "").lower()
    if intro and tot and intro != tot: add("E01", f"Memo p{facts['raise_intro_usd']['page']} says {money(intro)}; p{facts['raise_total_usd']['page']} says {money(tot)}.")
    elif deb and tot and deb < tot and "debt" in iq and "equity" not in iq: add("E01", f"Memo p{facts['raise_intro_usd']['page']} describes all {money(intro)} as debt; p{facts['debt_usd']['page']} gives {money(deb)} of debt, leaving {money(tot-deb)} of equity.")
    own = [F(x) for x in ("own_founders", "own_investor", "own_staff")]
    if all(own) and abs(sum(own) - 1) > .005: add("V10", f"Ownership shares add to {pct(sum(own))}.")
    # E02
    ten = F("loan_tenor_months")
    if ten and ten <= 12 and Z("BS", L["nc"], last): add("E02", f"${Z('BS',L['nc'],last):,.0f}k of ${book[last]:,.0f}k ({Z('BS',L['nc'],last)/book[last]:.0%}) of {last} loans are non-current, though loans last about {ten} months.")
    # E03 currency
    cty = country(F("country_main"), info["tier"]) if F("country_main") else None
    ccy = dict(p.split(":") for p in re.split(r";\s*", S["currency_by_country"])).get(cty); fx = info["fx"].get(ccy, {}); bits = []
    a, b = fx.get(int(last) - 1), fx.get(int(last))
    if a and b and abs(1 - a / b) > S["fx_move_flag"] and abs(Z("PL", "Net foreign exchange (loss)/gain", last)) < .005 * rev[last]: bits.append(f"{ccy} moved from {a:,.0f} to {b:,.0f} per USD in {last} (AHL's rates), yet the accounts show an FX result of {N(Z('PL','Net foreign exchange (loss)/gain',last))}.")
    if b and F("memo_fx_rate") and abs(F("memo_fx_rate") / b - 1) > S["fx_rate_gap_flag"]: bits.append(f"Memo p{facts['memo_fx_rate']['page']} uses {F('memo_fx_rate'):,.0f}; AHL's {last} year-end rate is {b:,.0f}.")
    if bits: add("E03", " ".join(bits))
    # E04, E15 cash flow
    op = ["Profit before tax", "Depreciation & amortization", "Net finance costs (paid)", "ECL impairment (non-cash)", "(Increase)/decrease in trade receivables", "(Increase)/decrease in inventories", "(Increase)/decrease in loan receivables (net)", "Increase/(decrease) in payables & other", "Income taxes paid"]
    inv, fin = ["Purchase of PPE", "Purchase of intangibles"], ["Net proceeds from borrowings", "Proceeds from share issues", "Lease payments (principal)"]
    part = lambda ls, y: sum(Z("CF", l, y) for l in ls); gaps, st = [], []
    for p, y in P:
        net, mv = part(op + inv + fin, y), Z("BS", "Cash & cash equivalents", y) - Z("BS", "Cash & cash equivalents", p)
        if abs(net - mv) > 1: gaps.append(f"{y}: lines give {N(net)}, cash moved {N(mv)}")
    ocf = "Net cash (used in)/from operating activities"
    for y in years:
        v = V("CF", ocf, y)
        if v is None: st.append(f"{y} operating total not printed")
        elif abs(v - part(op, y)) > 1: st.append(f"{y} operating total {N(v)}, lines give {N(part(op, y))}")
    if gaps or st: add("E04", ". ".join(filter(None, ["; ".join(gaps), "; ".join(st)])) + ".")
    m = []
    if any(Z("CF", "Net finance costs (paid)", y) < 0 for y in years): m.append("Finance costs are deducted from profit that already includes them")
    for p, y in P:
        for l, bs in (("(Increase)/decrease in inventories", "Inventories"), ("(Increase)/decrease in trade receivables", "Trade & other receivables")):
            d = Z("BS", bs, p) - Z("BS", bs, y)
            if d * Z("CF", l, y) < 0: m.append(f"{y} {bs.lower()} {'fell' if d > 0 else 'rose'} by {N(abs(d))} on the balance sheet, which should {'add' if d > 0 else 'use'} cash, but the cash flow shows the opposite")
    if m: add("E15", "; ".join(m) + ".")
    # E05 profit footing
    lines = ["Revenue", "Cost of sales", "Other operating income", "Selling & distribution expenses", "General & administrative expenses", "Technology & platform costs", "Expected credit loss (ECL) impairment"]; foot = []
    for y in years:
        e = sum(Z("PL", l, y) for l in lines); d = e - Z("PL", "EBITDA", y)
        if abs(d) > 1: foot.append(f"{y} EBITDA: lines give {N(e)}, stated {N(Z('PL','EBITDA',y))} (gap {N(d)}" + (", equal to other operating income)" if d == Z("PL", "Other operating income", y) else ")"))
        eb = Z("PL", "EBITDA", y) + Z("PL", "Depreciation (PPE & ROU)", y) + Z("PL", "Amortization (software & intangibles)", y)
        pb = eb + sum(Z("PL", l, y) for l in ("Finance income (interest on cash)", "Finance costs (borrowings, leases)", "Net foreign exchange (loss)/gain"))
        for nm, calc in (("Operating profit (EBIT)", eb), ("Profit before tax", pb), ("Profit for the year (PAT)", pb + Z("PL", "Income tax expense", y))):
            if abs(calc - Z("PL", nm, y)) > 1: foot.append(f"{y} {nm}: lines give {N(calc)}, stated {N(Z('PL', nm, y))}")
    if foot: add("E05", "; ".join(foot) + ".")
    # E06, E17
    txt = re.sub(r"\s+", " ", " ".join(pages_fin)); r = re.search(r"interest and fees[^()]*\((\d+)\s*[–-]\s*(\d+)%", txt, re.I); lo, hi = (int(r.group(1)) / 100, int(r.group(2)) / 100) if r else (1, 1)
    prev = years[years.index(last) - 1]; avg = (book[prev] + book[last]) / 2
    if avg and lo * rev[last] / avg > S["max_plausible_yield"]: add("E06", f"Interest and fees are {pct(lo)}-{pct(hi)} of revenue: ${lo*rev[last]:,.0f}k-${hi*rev[last]:,.0f}k on a ${avg:,.0f}k average {last} loan book, a yield of {lo*rev[last]/avg:.0%}-{hi*rev[last]/avg:.0%}.")
    if F("farmers_served") and book[last] * 1000 / F("farmers_served") < S["min_loan_per_client_usd"]:
        c = re.search(r"commissions[^()]*\((\d+)\s*[–-]\s*(\d+)%", txt, re.I); t = F("produce_tons")
        add("E17", f"${book[last]:,.0f}k loan book / {F('farmers_served'):,.0f} clients = ${book[last]*1000/F('farmers_served'):,.0f} each." + (f" Commissions on {t:,.0f} tons = ${int(c.group(1))*10*rev[last]/t:,.2f}-${int(c.group(2))*10*rev[last]/t:,.2f} a ton." if c and t else ""))
    # E07, E08, E09, E13
    liab = ["Borrowings – short-term", "Trade payables & accruals", "Current tax liabilities", "Lease liabilities – short-term", "Other current liabilities (deferred rev., VAT)"]
    ca = ["Cash & cash equivalents", "Trade & other receivables", L["cur"], L["nc"], "Inventories", "Other current assets (prepaids, VAT)"]; b7, b9 = [], []
    for y in years:
        if abs(Z("BS", "Total assets", y) - Z("BS", "Total equity & liabilities", y)) > 1: b7.append(f"{y}: assets {N(Z('BS','Total assets',y))}, equity and liabilities {N(Z('BS','Total equity & liabilities',y))}")
        s = sum(Z("BS", l, y) for l in liab)
        if abs(s - Z("BS", "Total current liabilities", y)) > 1: b7.append(f"{y} current liabilities: lines give {N(s)}, stated {N(Z('BS','Total current liabilities',y))}")
        s, t = sum(Z("BS", l, y) for l in ca), Z("BS", "Total current assets", y)
        if abs(s - t) > 1: b9.append(f"{y}: stated {N(t)}, lines {N(s)} ({N(s-Z('BS',L['nc'],y))} without non-current loans)")
    if b7: add("E07", "; ".join(b7) + ".")
    if b9: add("E09", "; ".join(b9) + ".")
    rr = [f"{y}: {N(Z('BS','Retained earnings',p))} + {N(pat[y])} profit - {N(abs(Z('CF','Dividends paid',y)))} dividends = {N(Z('BS','Retained earnings',p)+pat[y]-abs(Z('CF','Dividends paid',y)))}, stated {N(Z('BS','Retained earnings',y))}" for p, y in P if abs(Z("BS", "Retained earnings", p) + pat[y] - abs(Z("CF", "Dividends paid", y)) - Z("BS", "Retained earnings", y)) > 1]
    if rr: add("E08", "; ".join(rr) + ".")
    ppe = "Property, plant & equipment (net)"
    w = [f"{y}: {N(Z('BS',ppe,p))} + {N(-Z('CF','Purchase of PPE',y))} bought - {N(-Z('PL','Depreciation (PPE & ROU)',y))} depreciation, but stated {N(Z('BS',ppe,y))}" for p, y in P if Z("BS", ppe, p) and (Z("BS", ppe, p) - Z("CF", "Purchase of PPE", y) + Z("PL", "Depreciation (PPE & ROU)", y) - Z("BS", ppe, y)) / Z("BS", ppe, p) > S["ppe_gap_flag"]]
    if w: add("E13", "; ".join(w) + ".")
    # E10
    bad = [f"{c}: memo {t}, AHL {info['tier'][k]}" for c, t in re.findall(r"([^:;]+):\s*(Tier \d)", F("tiers_stated") or "") if (k := country(c, info["tier"])) and info["tier"][k] != t]
    if bad: add("E10", "; ".join(b.strip() for b in bad) + ".")
    # E11, E12, E14, E16
    cap = pat[last] / (debt[last] + eq[last]) if debt[last] + eq[last] else 0
    if F("net_margin_claim_usd") and abs(F("net_margin_claim_usd") / 1e6 - cap) > .02: add("E11", f"Memo p{facts['net_margin_claim_usd']['page']} claims {money(F('net_margin_claim_usd'))} per $1m. {last} profit after tax / (debt + equity) = {N(pat[last])} / {N(debt[last]+eq[last])} = ${cap*1e6:,.0f} per $1m.")
    tx = {y: -Z("PL", "Income tax expense", y) / Z("PL", "Profit before tax", y) for y in years if Z("PL", "Profit before tax", y) > 0}
    if tx and max(tx.values()) - min(tx.values()) > S["tax_swing_flag"]: add("E12", "Effective rate: " + ", ".join(f"{y} {v:.0%}" for y, v in tx.items()) + ".")
    for y in fc:
        new = (Z("CF", "Net proceeds from borrowings", y) + Z("CF", "Proceeds from share issues", y)) * 1000
        if tot and new < .5 * tot: add("E14", f"{y} includes {money(new)} of new borrowing and shares, against a {money(tot)} raise." + (f" Debt to equity would be {(debt[last]+deb/1000)/eq[last]:.2f}x on {last} equity before any new equity." if deb else ""))
    ecl = lambda y: -Z("PL", "Expected credit loss (ECL) impairment", y)
    fl = [f"{y}: charge {N(ecl(p))} to {N(ecl(y))} while the loan book grew {book[y]/book[p]-1:.0%}; charge / book {ecl(p)/book[p]:.1%} to {ecl(y)/book[y]:.1%}" for p, y in P if y in actual and book[p] and ecl(p) and book[y] / book[p] - 1 > S["ecl_book_growth_flag"] and ecl(y) / ecl(p) - 1 < .05]
    if fl: add("E16", "; ".join(fl) + ".")
    # E18 to E22
    if F("women_share") and F("youth_share") and abs(F("women_share") + F("youth_share") - 1) < 1e-9: add("E18", f"{pct(F('women_share'))} women + {pct(F('youth_share'))} youth = 100%.")
    if not F("founder_names"): add("E19", "No founder is named in the memo.")
    if debt[last] and not F("existing_lenders"): add("E20", f"${debt[last]:,.0f}k of borrowings at end {last}; lenders, security and ranking not disclosed.")
    cq = str(facts.get("cfo", {}).get("sentence") or "")
    firm = next((w for w in re.findall(r"[A-Z][\w]+", F("auditor") or "") if w not in ("IFRS", "GAAP", "IAS") and w in cq), None)
    if firm: add("E21", f"Auditor is {firm}; memo p{facts['cfo']['page']} links the CFO to {firm}.")
    if not re.search(r"independent auditor|auditor.s report|in our opinion|true and fair", txt, re.I): add("E22", "The accounts contain no auditor's report." + (f" {', '.join(fc)} is a budget." if fc else "") + (f" Memo is dated {F('memo_date')}." if F("memo_date") else ""))
    # V10 memo table
    acc = lambda n, y: dict(Debt=debt[y], **{"Loan book": book[y]}).get(n, next((V(s, n, y) for s in ("PL", "BS") if (s, n) in K), None))
    df = [f"{y} {n}: memo {v:,}, accounts {acc(n,y):,}" for y, row in memo_tbl.items() if y in years for n, v in row.items() if n != "page" and v is not None and acc(n, y) is not None and abs(v - acc(n, y)) > .006]
    if df: add("V10", "; ".join(df) + ".")
    # calculated scores
    tri = lambda v, one, three: 1 if v < one else 3 if v > three else 2
    g = [rev[y] / rev[p] - 1 for p, y in P if y in actual and rev[p]][-3:]; avgg = sum(g) / len(g) if g else None
    gm = 1 + Z("PL", "Cost of sales", last) / rev[last] if rev[last] else 0; pbt = Z("PL", "Profit before tax", last); de = debt[last] / eq[last] if eq[last] else 9
    calc = {"D33": (tri(avgg, S["growth_score_1"], S["growth_score_3"]), f"Growth {', '.join(f'{x:.1%}' for x in g)}: average {avgg:.1%}.") if g else (None, ""),
            "D34": (1 if pbt < S["pbt_score_1"] else 3 if pbt > S["pbt_score_3"] else 2, f"Audited {last} profit before tax ${pbt:,.0f}k."),
            "D35": (tri(gm, S["gm_score_1"], S["gm_score_3"]), f"{last} gross margin {gm:.0%}."),
            "D43": (3 if pat[last] > 0 and de < S["max_debt_equity"] and all(Z("BS", "Retained earnings", y) > Z("BS", "Retained earnings", p) for p, y in P if y in actual) else 1 if pat[last] <= 0 else 2, f"Profitable; retained earnings growing; {last} debt to equity {de:.2f}x." if pat[last] > 0 else f"Loss-making in {last}.")}
    fired = {i for i, _ in ex}
    cautious = {}
    if "E11" in fired: cautious["D36"] = tri(cap, S["cap_eff_score_1"], S["cap_eff_score_3"])
    if "E06" in fired and calc["D35"][0] == 3: cautious["D35"] = 2
    return ex, calc, cautious

OPTIONS = {"B9": ["Yes", "No", "Maybe"], "B10": ["Yes, Adaptation", "Yes, Mitigation", "Maybe", "No"], "B11": ["Yes", "No", "Maybe"], "B13": ["Yes", "No", "Maybe"], "B15": ["Yes", "No", "Maybe"]}

def check_assessment(raw, facts, calc, cautious, cfg):
    """Keeps an AI suggestion only if it is allowed and backed by verified facts."""
    d = parse_json(raw) or {}; ex, scr, sco = [], {}, {}
    known = lambda ev: [e for e in ev or [] if facts.get(e, {}).get("value") is not None]
    for s in d.get("screening", []):
        c, v = s.get("cell"), s.get("value")
        if c not in {x["cell"] for x in cfg["screening"] if x["by"] in ("suggested", "memo")} or v in (None, ""): continue
        if c in OPTIONS and v not in OPTIONS[c]: ex.append(("V07", f"{c}: '{v}' is not an allowed option.")); continue
        if not known(s.get("evidence")) and v != "Maybe": ex.append(("V07", f"{c}: no verified evidence.")); continue
        scr[c] = dict(value=tidy(v)[:400], comment=tidy(s.get("reason"))[:300], evidence=known(s.get("evidence")))
    for s in d.get("scores", []):
        c, v = s.get("cell"), s.get("score")
        if c not in {x["cell"] for x in cfg["scores"]} or c in calc or v is None: continue
        if v not in (1, 2, 3): ex.append(("V07", f"{c}: score {v} is not allowed.")); continue
        if not known(s.get("evidence")): ex.append(("V07", f"{c}: no verified evidence, left blank.")); continue
        sco[c] = dict(score=v, comment="Suggested. " + tidy(s.get("reason"))[:300], evidence=known(s.get("evidence")), cautious=cautious.get(c, 2 if v == 3 and s.get("support") == "general" else v))
    for c, (v, why) in calc.items():
        if v: sco[c] = dict(score=v, comment="Calculated. " + why, evidence=["accounts"], cautious=cautious.get(c, v))
    for x in d.get("contradictions", [])[:5]:
        e = known(x.get("evidence"))
        if len(e) >= 2: ex.append(("A01", f"{tidy(x.get('issue'))[:200].rstrip('. ')}. " + " / ".join(f"p{facts[i]['page']}: \"{str(facts[i]['quote'])[:120]}\"" for i in e[:2])))
    return scr, sco, ex
