"""Compares a saved run with the answer key. Usage: python measure.py <run folder>"""
import json, re, sys, warnings
from pathlib import Path
import openpyxl
warnings.simplefilter("ignore")
HERE = Path(__file__).resolve().parent.parent / "private"; A = Path(sys.argv[1]); run = json.loads((A / "run.json" if A.is_dir() else A).read_text())
def tab(name, first):
    on, out = False, []
    for r in openpyxl.load_workbook(HERE / "answer_key.xlsx", data_only=True)[name].iter_rows(values_only=True):
        if on and r[0] is not None: out.append(r)
        elif r[0] == first: on = True
    return out
print("models", run["models"], run.get("assess_model")); ok, look = 0, []
for fid, q, typ, exp, page, quote in tab("Memo answers", "field_id"):
    f = run["facts"][fid]; v = f["value"]; exp = None if exp == "(not stated)" else exp
    same = v == exp or (v is not None and exp is not None and (abs(float(v) - float(exp)) < 1e-9 if typ in ("number", "percent") else re.sub(r"\W", "", str(v).lower()) == re.sub(r"\W", "", str(exp).lower())))
    if same: ok += 1
    else: look.append((fid, typ, f["status"], exp, v, f["reason"], f["reader_a"], f["reader_b"], f["quote"]))
print(f"exact match {ok}/{len(run['facts'])}")
for x in look: print(" ", x[0], "|", x[2], "\n     key:", x[3], "\n     got:", x[4], "\n     A:", x[6], "\n     B:", x[7], "\n     ", x[5])
print("SCORES (cell: key/got)"); d = {r[0]: r[2] for r in tab("Scores", "cell")}
print({c: (d[c], run["scores"].get(c, {}).get("score")) for c in d if str(d[c]).replace("(blank)", "None") != str(run["scores"].get(c, {}).get("score"))} or "all equal")
print("SCREENING", {c: s["value"][:70] for c, s in run["screening"].items()})
print("OTHER EXCEPTIONS", [(i, e[:200]) for i, e in run["exceptions"] + run.get("assess_exceptions", []) if i[0] in "VA"])
