"""Loads the rules AHL staff can edit in field_map.xlsx."""
import hashlib, os
from pathlib import Path
import openpyxl

BASE = Path(os.getenv("APP_BASE", Path(__file__).resolve().parent.parent))
FIELD_MAP = Path(os.getenv("FIELD_MAP", BASE / "config" / "field_map.xlsx"))
TEMPLATE = Path(os.getenv("TEMPLATE", BASE.parent / "2025.12_Screening Template vF.xlsx"))
RUNS = Path(os.getenv("RUNS", BASE / "runs"))

def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()

def load():
    wb = openpyxl.load_workbook(FIELD_MAP, data_only=True)
    def rows(tab, first):
        out, on = [], False
        for r in wb[tab].iter_rows(values_only=True):
            if on and r[0] is not None: out.append(r)
            elif r[0] == first: on = True
        return out
    split = lambda t, sep=";": [x.strip() for x in str(t or "").split(sep) if x.strip()]
    return {
        "version": sha(FIELD_MAP)[:10],
        "memo_fields": {r[0]: dict(question=r[1], type=r[2], required=r[3] == "yes") for r in rows("Memo Fields", "field_id")},
        "rows": [dict(row=int(r[0]), label=r[1], st=r[2], lines=split(r[3], " + "), note=r[6] or "") for r in rows("Financial Rows", "template_row") if r[2]],
        "blank_rows": {int(r[0]): (r[1], r[6]) for r in rows("Financial Rows", "template_row") if not r[2]},
        "screening": [dict(cell=r[0], label=r[1], by=r[2], sources=split(r[3])) for r in rows("Assessment", "cell")],
        "scores": [dict(cell=r[0], label=r[1], by=r[2]) for r in rows("Scores", "cell")],
        "names": [dict(st=r[0], name=r[1], kind=r[2], alts=split(r[3])) for r in rows("Line Names", "statement")],
        "rules": {r[0]: dict(severity=r[1], issue=r[2], why=r[3] or "", action=r[4] or "", question=r[5] or "") for r in rows("Exception Rules", "id")},
        "set": {r[0]: r[1] for r in rows("Settings", "name")},
        "prompts": {r[0]: dict(version=r[1], text=r[2]) for r in rows("Prompts", "name")},
        "changes": rows("Template Changes", "id"),
    }
