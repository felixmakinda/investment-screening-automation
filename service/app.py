"""AHL screening service. Reads documents, checks answers, fills the template. Contains no AI and no API keys."""
import hashlib, json, os, re, shutil
from datetime import datetime, timezone
from fastapi import Depends, FastAPI, File, Header, HTTPException, UploadFile
from fastapi.responses import FileResponse
from pydantic import BaseModel
import checks, config, filler, reader

app = FastAPI(title="AHL screening service", version="1.0")
MAX = 10 * 1024 * 1024

def allowed(x_token: str | None = Header(None)):
    if os.getenv("SERVICE_TOKEN") and x_token != os.environ["SERVICE_TOKEN"]: raise HTTPException(401, "Wrong or missing service token")

def folder(run_id):
    if not re.fullmatch(r"\d{8}-\d{6}-[0-9a-f]{8}", run_id or ""): raise HTTPException(400, "Invalid run_id")
    d = config.RUNS / run_id
    if not d.is_dir(): raise HTTPException(404, "Run not found")
    return d
def save(d, run): (d / "run.json").write_text(json.dumps(run, indent=1, ensure_ascii=False, default=str))
def load(d): return json.loads((d / "run.json").read_text())
def event(d, what, **kw):
    with (d / "log.jsonl").open("a") as f: f.write(json.dumps(dict(time=datetime.now(timezone.utc).isoformat(timespec="seconds"), event=what, **kw), default=str) + "\n")
def documents(d, cfg):
    """Re-reads the stored originals each time, so nothing in between can alter the figures."""
    fin = reader.pdf_pages(d / "accounts.pdf"); st = reader.statements(fin); K, unknown = reader.recognise(st["lines"], cfg["names"])
    return reader.pdf_pages(d / "memo.pdf"), fin, st, K, unknown
def ai_request(prompt, content, name, schema):
    return dict(messages=[dict(role="system", content=prompt["text"]), dict(role="user", content=content)], response_format=dict(type="json_schema", json_schema=dict(name=name, strict=True, schema=schema)))
obj = lambda **p: dict(type="object", properties=p, required=list(p), additionalProperties=False)
arr = lambda items: dict(type="array", items=items)
NULLSTR = {"type": ["string", "null"]}

@app.get("/health")
def health():
    ok = dict(template=config.TEMPLATE.exists(), field_map=config.FIELD_MAP.exists(), pdftotext=bool(shutil.which("pdftotext")), soffice=bool(shutil.which("soffice")))
    if not all(ok.values()): raise HTTPException(503, ok)
    return dict(status="ok", field_map=config.load()["version"], template=config.sha(config.TEMPLATE)[:10])

@app.post("/read", dependencies=[Depends(allowed)])
async def read(memo: UploadFile = File(...), accounts: UploadFile = File(...)):
    cfg, data = config.load(), {}
    for n, f in (("memo", memo), ("accounts", accounts)):
        data[n] = await f.read(MAX + 1)
        if len(data[n]) > MAX: raise HTTPException(413, f"{n} is larger than 10 MB")
    files = {n: hashlib.sha256(b).hexdigest() for n, b in data.items()}; mark = hashlib.sha256("".join(files.values()).encode()).hexdigest()[:8]
    earlier = sorted(p.name for p in config.RUNS.glob(f"*-{mark}")) if config.RUNS.exists() else []
    rid = f"{datetime.now():%Y%m%d-%H%M%S}-{mark}"; d = config.RUNS / rid; d.mkdir(parents=True, exist_ok=True)
    for n, b in data.items(): (d / f"{n}.pdf").write_bytes(b)
    try:
        pages, fin, st, K, unknown = documents(d, cfg)
        if not K: raise reader.BadDocument("No financial statement lines were recognised. Are the two files swapped?")
    except reader.BadDocument as e:
        shutil.rmtree(d); raise HTTPException(422, str(e))
    ex = [("V05", f"{u['st']} p{u['page']}: '{u['label']}'.") for u in unknown] + [("V08", f"Memo p{p}: \"{t}\".") for p, t in reader.suspicious(pages)]
    if earlier: ex.append(("V11", f"Earlier run: {earlier[-1]}."))
    if st["unit"] != ("USD", 1000): ex.append(("V06", f"Accounts unit is {st['unit']}; only USD '000 is supported today."))
    save(d, dict(run_id=rid, files=files, years=st["years"], unit=st["unit"], read_exceptions=ex, field_map=cfg["version"]))
    event(d, "read", lines=len(K), unrecognised=len(unknown), memo_pages=len(pages), **files)
    qs = "\n".join(f"- {k}: {v['question']} ({v['type']})" for k, v in cfg["memo_fields"].items())
    text = "\n".join(f"=== PAGE {i} ===\n{p}" for i, p in enumerate(pages, 1))
    schema = obj(answers=arr(obj(field_id=dict(type="string", enum=list(cfg["memo_fields"])), value=NULLSTR, quote=NULLSTR, page={"type": ["integer", "null"]})))
    return dict(run_id=rid, memo_pages=len(pages), account_lines=len(K), years=st["years"], warnings=[t for _, t in ex],
                ai_request=ai_request(cfg["prompts"]["extract"], f"Answer every field once.\n{qs}\n\n<memo>\n{text}\n</memo>", "memo_answers", schema))

class Readers(BaseModel):
    run_id: str
    reader_a: dict | str | None = None
    reader_b: dict | str | None = None
    model_a: str = ""
    model_b: str = ""

@app.post("/validate", dependencies=[Depends(allowed)])
def validate(r: Readers):
    d, cfg = folder(r.run_id), config.load(); run, info = load(d), checks.template_info()
    pages, fin, st, K, _ = documents(d, cfg)
    try: facts, ex = checks.combine(r.reader_a, r.reader_b, cfg["memo_fields"], pages)
    except ValueError as e: raise HTTPException(422, str(e))
    table = reader.memo_table(pages, [x.strip() for x in cfg["set"]["memo_table_columns"].split(";")])
    rules, calc, cautious = checks.run_rules(K, st["years"], facts, table, fin, cfg, info)
    run.update(facts=facts, exceptions=run["read_exceptions"] + ex + rules, calc=calc, cautious=cautious, models=[r.model_a, r.model_b], screening={}, scores={}); save(d, run)
    count = {s: sum(f["status"] == s for f in facts.values()) for s in ("accepted", "blank", "blocked", "review")}
    event(d, "validate", models=[r.model_a, r.model_b], exceptions=len(run["exceptions"]), **count)
    known = "\n".join(f"- {k}: {v['value']} (memo p{v['page']}: \"{v['quote']}\")" for k, v in facts.items() if v["value"] is not None)
    none = ", ".join(k for k, v in facts.items() if v["value"] is None)
    guides = "\n".join(f"- {s['cell']} {info['guide'][s['cell']]}" for s in cfg["scores"] if s["cell"] not in calc)
    screens = "\n".join(f"- {s['cell']} {s['label']}" + (f" (choose one: {' / '.join(checks.OPTIONS[s['cell']])})" if s["cell"] in checks.OPTIONS else " (short text)") for s in cfg["screening"] if s["by"] in ("suggested", "memo") and s["cell"] not in ("B5", "B8", "B12"))
    issues = "\n".join(f"- {i}: {e}" for i, e in rules)
    ids = dict(type="string", enum=list(facts)); cells = [s["cell"] for s in cfg["scores"] if s["cell"] not in calc]
    schema = obj(screening=arr(obj(cell=dict(type="string"), value=dict(type="string"), reason=dict(type="string"), evidence=arr(ids))),
                 scores=arr(obj(cell=dict(type="string", enum=cells), score={"type": ["integer", "null"]}, reason=dict(type="string"), support=dict(type="string", enum=["specific", "general"]), evidence=arr(ids))),
                 contradictions=arr(obj(issue=dict(type="string"), evidence=arr(ids))))
    return dict(run_id=r.run_id, counts=count, exceptions=len(run["exceptions"]), high=sum(cfg["rules"][i]["severity"] == "High" for i, _ in run["exceptions"]),
                ai_request=ai_request(cfg["prompts"]["assess"], f"VERIFIED FACTS\n{known}\n\nNOT STATED: {none}\n\nISSUES ALREADY FOUND\n{issues}\n\nSCREENING FIELDS\n{screens}\n\nSCORES\n{guides}", "assessment", schema))

class Assessment(BaseModel):
    run_id: str
    assessment: dict | str | None = None
    model: str = ""

@app.post("/assess", dependencies=[Depends(allowed)])
def assess(a: Assessment):
    d, cfg = folder(a.run_id), config.load(); run = load(d)
    if "facts" not in run: raise HTTPException(409, "Run /validate first")
    scr, sco, ex = checks.check_assessment(a.assessment, run["facts"], run["calc"], run["cautious"], cfg)
    run.update(screening=scr, scores=sco, assess_exceptions=ex, assess_model=a.model); save(d, run)
    event(d, "assess", model=a.model, scores=len(sco), screening=len(scr), rejected=sum(i == "V07" for i, _ in ex))
    lab = {x["cell"]: x["label"] for x in cfg["scores"] + cfg["screening"]}
    return dict(run_id=a.run_id, review=[dict(cell=c, label=lab[c], suggested=s.get("score", s.get("value")), why=s["comment"]) for c, s in {**scr, **sco}.items()],
                blank=[dict(cell=x["cell"], label=x["label"]) for x in cfg["scores"] + cfg["screening"] if x["cell"] not in {**scr, **sco} and (x["cell"][0] == "D" or x["by"] == "human")],
                exceptions=listed(run, cfg))

def listed(run, cfg):
    order, out = ("High", "Medium", "Low"), []
    for i, e in run["exceptions"] + run.get("assess_exceptions", []):
        r = cfg["rules"][i]; out.append((i, r["severity"], r["issue"], e, r["why"], r["action"], r["question"]))
    return sorted(dict.fromkeys(out), key=lambda x: (order.index(x[1]), x[0]))

class Fill(BaseModel):
    run_id: str
    approvals: dict[str, int | str | None] = {}
    reviewer: str = ""

@app.post("/fill", dependencies=[Depends(allowed)])
def fill(f: Fill):
    d, cfg = folder(f.run_id), config.load(); run, info = load(d), checks.template_info()
    if "facts" not in run: raise HTTPException(409, "Run /validate first")
    cells = {x["cell"] for x in cfg["scores"]}
    for c, v in f.approvals.items():
        if c in cells and v not in (1, 2, 3, None) or c in checks.OPTIONS and v not in checks.OPTIONS[c] + [None]: raise HTTPException(422, f"{c}: {v} is not allowed")
    _, _, st, K, _ = documents(d, cfg)
    w, log, ex = filler.plan(run, K, cfg, info, f.approvals)
    if ex: run["exceptions"] += ex
    name = re.sub(r"[^\w -]+", "", str(run["facts"]["company_name"]["value"] or "Company")).strip() or "Company"
    out = d / f"{name} - Screening Template.xlsx"
    filler.write(w, out, budget=run["years"][-1].endswith("F")); calc = filler.recalc(out)
    problems = filler.read_back(out, w, calc)
    if problems: event(d, "fill_stopped", problems=problems[:20]); out.unlink(); raise HTTPException(409, dict(message="Read-back failed. The file was not released.", problems=problems[:20]))
    bal, res = filler.balance_checks(calc, run["years"], K)
    groups = [(.35, range(23, 30)), (.5, range(32, 39)), (.15, range(41, 45))]
    final = {c: f.approvals.get(c, s["score"]) for c, s in run["scores"].items()} | {c: v for c, v in f.approvals.items() if c in cells}
    caut = {c: min(v, run["scores"].get(c, {}).get("cautious", v)) for c, v in final.items() if v}
    cs = sum(wgt * sum(caut.get(f"D{r}", 0) for r in rg) / (3 * max(1, sum(f"D{r}" in caut for r in rg))) for wgt, rg in groups) if res["score"] else None
    status = f"APPROVED by {f.reviewer} on {datetime.now():%d %b %Y %H:%M}" if f.reviewer.strip() else "DRAFT: not yet approved by a person"
    cnt = {s: sum(x["status"] == s for x in run["facts"].values()) for s in ("accepted", "blank", "blocked", "review")}
    ch = [("Status", "", status), ("Run", f.run_id, f"Field map {cfg['version']}; template {config.sha(config.TEMPLATE)[:10]}; prompts {cfg['prompts']['extract']['version']}/{cfg['prompts']['assess']['version']}"),
          ("AI models", "", f"Reader {' and '.join(m for m in run['models'] if m)}; scores {run.get('assess_model','')}"), ("Score", f"{res['score']:.1%}" if res["score"] else "Not available", f"Cautious case {cs:.1%}. Suggestions until approved." if res["score"] else "Too few scores. Complete the blank scores first."),
          ("Memo answers", "", f"{cnt['accepted']} accepted with verified quotes; {cnt['blank']} not stated; {cnt['blocked']} blocked; {cnt['review']} need a person"),
          ("Read-back", "Pass", f"{sum(v is not None for v in w.values())} cells confirmed; every other cell unchanged; no formula errors"), *bal,
          *[("Template change", c[0], f"{c[1]} {c[2]}: {c[3]}. {c[4]}") for c in cfg["changes"]]]
    exl = listed(run, cfg); filler.add_tabs(out, log, exl, ch); filler.store_results(out, filler.recalc(out))
    run.update(output=out.name, status=status, results=res, cautious_score=cs); save(d, run); event(d, "fill", status=status, reviewer=f.reviewer, score=res["score"], sha=config.sha(out))
    return dict(run_id=f.run_id, status=status, file=out.name, download=f"/files/{f.run_id}", results=res, cautious_score=cs, exceptions=len(exl), high=sum(x[1] == "High" for x in exl))

@app.get("/files/{run_id}", dependencies=[Depends(allowed)])
def files(run_id: str):
    d = folder(run_id); run = load(d)
    if "output" not in run: raise HTTPException(404, "Not filled yet")
    return FileResponse(d / run["output"], filename=run["output"], media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
