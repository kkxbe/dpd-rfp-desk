"""DPD proposal builder (GitHub Actions). Builds every proposal the RFP Desk has requested, without Kyle's computer.

For each requested group: download the proposal-generator code from the SharePoint toolkit folder, the group's calculator
(06 Decision), SBCs (01) and Benefit RE quote (05); recalculate the calculator; build with make_proposal.py; QA; save the
draft PDF to the group's 07 folder; mark the desk 'ready' (with the link) or 'failed' (with the reason).
Nothing proprietary lives in this repo: the generator code, Kerix content and claims stories are read from SharePoint at run time.

Secrets (GitHub > Settings > Secrets and variables > Actions): MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET, SUPABASE_SERVICE_ROLE_KEY.
Optional input FORCE_GROUPS: comma-separated desk group names to build now whatever their status.
"""
import os, sys, re, json, glob, shutil, subprocess, datetime as dt, urllib.parse, urllib.request

SB_URL = "https://uwcciwmsjzaslxhykyqc.supabase.co"
SP_HOST, SP_SITE, SP_LIB = "directpaydesigns432.sharepoint.com", "/sites/DirectPayDesignsDocuments", "03_Clients  INTERNAL"
SP_ROOT = "2026 DPD RFPs/Level-Funded"
TOOLKIT = "_DPD RFP Toolkit/5 System - proposal generator (do not edit)"
WORK = os.path.abspath("work")
env = os.environ.get
from zoneinfo import ZoneInfo
MT = ZoneInfo("America/Denver")


def http(method, url, headers=None, data=None, raw=False):
    req = urllib.request.Request(url, method=method, headers=headers or {}, data=data)
    try:
        with urllib.request.urlopen(req, timeout=120) as r:
            body = r.read()
            return body if raw else (json.loads(body) if body else None)
    except urllib.error.HTTPError as e:
        raise RuntimeError(f"{method} {url.split('?')[0]} -> {e.code}: {e.read()[:300]!r}")

# ---------------------------------------------------------------- Supabase (service role)
SBH = {"apikey": env("SUPABASE_SERVICE_ROLE_KEY", ""), "Authorization": f"Bearer {env('SUPABASE_SERVICE_ROLE_KEY', '')}", "Content-Type": "application/json"}
def sb(method, path, body=None, prefer="return=representation"):
    return http(method, f"{SB_URL}/rest/v1/{path}", {**SBH, "Prefer": prefer}, json.dumps(body).encode() if body is not None else None)

def claim(gid, ok_from):
    """Atomically move a request to 'building' so the laptop job and this one never build the same group twice."""
    flt = f"id=eq.{gid}" + (f"&proposal_status=eq.{ok_from}" if ok_from else "")
    rows = sb("PATCH", f"groups?{flt}", {"proposal_status": "building", "proposal_note": "Claude is building the draft (cloud builder)"})
    return bool(rows)

def finish(gid, status, note, url=None, change=None):
    body = {"proposal_status": status, "proposal_note": note}
    if url: body["proposal_url"] = url
    sb("PATCH", f"groups?id=eq.{gid}", body)
    sb("POST", "activity", {"who": "Claude (proposal)", "group_id": gid, "change": change or note}, prefer="return=minimal")

# ---------------------------------------------------------------- Microsoft Graph (app-only)
TOKEN = None
def tok():
    global TOKEN
    if not TOKEN:
        d = urllib.parse.urlencode({"client_id": env("MS_CLIENT_ID"), "client_secret": env("MS_CLIENT_SECRET"),
                                    "scope": "https://graph.microsoft.com/.default", "grant_type": "client_credentials"}).encode()
        TOKEN = http("POST", f"https://login.microsoftonline.com/{env('MS_TENANT_ID')}/oauth2/v2.0/token",
                     {"Content-Type": "application/x-www-form-urlencoded"}, d)["access_token"]
    return TOKEN
def g(path, method="GET", data=None, ctype=None, raw=False):
    h = {"Authorization": f"Bearer {tok()}"}
    if ctype: h["Content-Type"] = ctype
    return http(method, "https://graph.microsoft.com/v1.0" + path, h, data, raw)
enc = lambda p: "/".join(urllib.parse.quote(x, safe="") for x in p.split("/"))
DRIVE = None
def drive():
    global DRIVE
    if not DRIVE:
        site = g(f"/sites/{SP_HOST}:{SP_SITE}")
        ds = g(f"/sites/{site['id']}/drives")["value"]
        DRIVE = next(d["id"] for d in ds if d["name"] == SP_LIB or urllib.parse.unquote(d["webUrl"]).endswith("/" + SP_LIB))
    return DRIVE
def item(rel): return f"/drives/{drive()}/root:/{enc(SP_ROOT + '/' + rel)}"
def children(rel):
    try: return g(item(rel) + ":/children?$top=500&$select=name,size,folder,file,webUrl")["value"]
    except RuntimeError as e:
        if "404" in str(e): return None
        raise
def download(rel, dest):
    os.makedirs(os.path.dirname(dest), exist_ok=True)
    open(dest, "wb").write(g(item(rel) + ":/content", raw=True))
def download_tree(rel, dest):
    for c in children(rel) or []:
        if c.get("folder"): download_tree(f"{rel}/{c['name']}", os.path.join(dest, c["name"]))
        elif c["name"] != "desktop.ini": download(f"{rel}/{c['name']}", os.path.join(dest, c["name"]))
def upload(rel, path):
    return g(item(rel) + ":/content", "PUT", open(path, "rb").read(), "application/pdf")
def sub(kids, prefix): return next((c["name"] for c in kids or [] if c.get("folder") and c["name"].startswith(prefix)), None)

# ---------------------------------------------------------------- one group
def build(grp, code):
    fp, name = grp["folder_path"], grp["name"]
    kids = children(fp)
    if kids is None: return "failed", f"SharePoint folder '{fp}' not found; fix the folder on the desk.", None
    f06, f07 = sub(kids, "06"), sub(kids, "07")
    calcs = [c["name"] for c in (children(f"{fp}/{f06}") if f06 else []) or [] if re.match(r"DPD Calculator.*\.xlsx$", c["name"], re.I)]
    if not calcs: return "failed", "No group calculator in 06 Decision yet.", None
    if not f07: return "failed", "No '07 Proposal' folder for this group; add one in SharePoint.", None
    wd = os.path.join(WORK, re.sub(r"[^\w.-]+", "_", name)); shutil.rmtree(wd, ignore_errors=True); os.makedirs(wd)
    download(f"{fp}/{f06}/{calcs[0]}", os.path.join(wd, "calc_in.xlsx"))
    for pre, test in (("01", lambda n: "sbc" in n.lower() or "kerix" in n.lower()), ("05", lambda n: True), ("04", lambda n: "kerix" in n.lower())):
        s_ = sub(kids, pre)
        for c in (children(f"{fp}/{s_}") if s_ else []) or []:
            if c.get("file") and c["name"].lower().endswith(".pdf") and test(c["name"]) and c.get("size", 0) < 40e6:
                download(f"{fp}/{s_}/{c['name']}", os.path.join(wd, s_, c["name"]))
    # desk values the runbook allows to fill when the calculator is blank (working copy only; the team's file is not changed)
    import openpyxl
    wb = openpyxl.load_workbook(os.path.join(wd, "calc_in.xlsx")); S, D = wb["Start Here"], wb["Decision"]
    filled = []
    if not S["D14"].value and grp.get("kerix_tier") and str(D["D19"].value).strip() == "Yes": S["D14"] = int(grp["kerix_tier"]); filled.append(f"Kerix Tier {grp['kerix_tier']} from desk")
    if grp.get("broker_comp_pepm") is not None and all(D[f"{c}15"].value in (None, "") for c in "FGHI"):
        for c in "FGHI": D[f"{c}15"] = float(grp["broker_comp_pepm"])
        filled.append("broker commission from desk")
    wb.save(os.path.join(wd, "calc_in.xlsx"))
    # recalculate (LibreOffice), so every formula has a value
    r = subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", os.path.join(wd, "rc"), os.path.join(wd, "calc_in.xlsx")], capture_output=True, text=True, timeout=240)
    calc = os.path.join(wd, "rc", "calc_in.xlsx")
    if not os.path.exists(calc): return "failed", "Could not recalculate the calculator (LibreOffice).", None
    out = os.path.join(wd, "out.pdf")
    r = subprocess.run([sys.executable, "make_proposal.py", calc, out, wd], cwd=code, capture_output=True, text=True, timeout=600,
                       env={**os.environ, "DPD_FONTS": os.path.join(code, "fonts")})
    log = (r.stdout + r.stderr).strip()
    if r.returncode == 2 or "REFUSED" in log:
        m = re.search(r"NOT READY - (.*)", log)
        return "failed", f"Calculator NOT READY: {(m.group(1) if m else log[-200:]).strip()}. Fix it in the calculator (or ask Claude on the desk), then request again.", None
    if r.returncode != 0 or not os.path.exists(out):
        return "failed", "Proposal build error: " + log.splitlines()[-1][:200] if log else "Proposal build error.", None
    S2 = openpyxl.load_workbook(calc, data_only=True)["Start Here"]
    is39 = S2["D17"].value == "Benefit RE / 39N"
    qa = [sys.executable, "verify/prop4.py", out, calc, S2["B9"].value] if is39 else [sys.executable, "verify/prop_any.py", out, calc]
    q = subprocess.run(qa, cwd=code, capture_output=True, text=True, timeout=300)
    qline = (q.stdout.strip().splitlines() or [""])[-1]
    if not qline.endswith("[]"):
        return "failed", f"QA did not pass: {qline[:220]}", None
    carrier = "39N" if is39 else (S2["D17"].value or "Carrier")
    fname = f"DRAFT - {name} - {carrier} Proposal {dt.datetime.now(MT):%Y-%m-%d}.pdf"
    res = upload(f"{fp}/{f07}/{fname}", out)
    when = dt.datetime.now(MT).strftime("%b %-d %-I:%M %p")
    note = f"Draft built {when} MT by the cloud builder, QA passed." + (f" Used {', '.join(filled)}." if filled else "") + " DRAFT, internal only while on HOLD."
    return "ready", note, res.get("webUrl", "") + "?web=1"

def main():
    os.makedirs(WORK, exist_ok=True)
    force = [x.strip() for x in (env("FORCE_GROUPS") or "").split(",") if x.strip()]
    cols = "id,name,folder_path,proposal_status,kerix_tier,broker_comp_pepm"
    groups = sb("GET", f"groups?select={cols}&proposal_status=eq.requested")
    for n in force:
        groups += [x for x in sb("GET", f"groups?select={cols}&name=ilike.*{urllib.parse.quote(n)}*") if x["id"] not in {y['id'] for y in groups}]
    if not groups: print("No proposal requests."); return
    code = os.path.join(WORK, "_code")
    if not os.path.exists(os.path.join(code, "make_proposal.py")):
        download_tree(TOOLKIT, code)
    built = failed = 0
    for grp in groups:
        if not claim(grp["id"], "requested" if grp.get("proposal_status") == "requested" else None):
            print("skip (taken)", grp["name"]); continue
        try: status, note, url = build(grp, code)
        except Exception as e: status, note, url = "failed", f"Builder error: {str(e)[:220]}", None
        finish(grp["id"], status, note, url, "Draft proposal built and saved to 07 Proposal (cloud builder)" if status == "ready" else f"Proposal failed: {note}")
        print(status.upper(), grp["name"], "|", note)
        built += status == "ready"; failed += status != "ready"
    sb("POST", "sync_runs", {"kind": "proposal", "summary": f"cloud builder: {built} built, {failed} failed"}, prefer="return=minimal")

if __name__ == "__main__":
    main()
