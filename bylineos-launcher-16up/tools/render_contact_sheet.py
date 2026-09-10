#!/usr/bin/env python3
"""
Headless, edition-agnostic contact-sheet builder for GitHub Actions.

Discovers the NEWEST edition that has proofs (Drive folder "Editions" ->
"Edition-0NN"), reads its database + proofs via a service account, places each
proof by the saved 'Pages Claimed' (table wins over filename), keeps the latest
proof per slug, rasterises with pdftoppm, renders the self-contained HTML, and
maintains the contact index + current.html from a small manifest.

Env: GDRIVE_SA_KEY_PATH   path to the service-account JSON key
Outputs under bylineos-launcher-16up/contact/: edNNN.html, index.html,
current.html, editions.json, .lastbuild
"""
import os, re, json, glob, shutil, sys, calendar
from openpyxl import load_workbook
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import gen

EDITIONS_PARENT = "13crOw1mdlwRkCWS35tVCGgHQwGhzCXU0"   # Drive "Editions" folder
OUTDIR = "bylineos-launcher-16up/contact"
ISSUE  = 68
WORK   = "work"
BASE_ED, BASE_Y, BASE_M = 87, 2026, 7                   # Edition 87 = July 2026 (monthly)

SCOPES=["https://www.googleapis.com/auth/drive.readonly"]
def _creds():
    # accept the key however the workflow exposes it, in order of preference
    p=os.environ.get("GDRIVE_SA_KEY_PATH")
    if not p:
        rt=os.environ.get("RUNNER_TEMP")
        if rt and os.path.exists(os.path.join(rt,"sa.json")): p=os.path.join(rt,"sa.json")
    if p and os.path.exists(p):
        return service_account.Credentials.from_service_account_file(p, scopes=SCOPES)
    raw=os.environ.get("GDRIVE_SA_KEY")           # raw JSON in the env var itself
    if raw and raw.strip().startswith("{"):
        return service_account.Credentials.from_service_account_info(json.loads(raw), scopes=SCOPES)
    sys.exit("NO SA KEY: set GDRIVE_SA_KEY_PATH, or write the key to $RUNNER_TEMP/sa.json, "
             "or put the raw JSON in GDRIVE_SA_KEY")
creds=_creds()
drive = build("drive", "v3", credentials=creds, cache_discovery=False)

def ls(q, fields="files(id,name,modifiedTime,mimeType)"):
    out=[]; tok=None
    while True:
        r=drive.files().list(q=q, fields="nextPageToken,"+fields, pageSize=200,
                             pageToken=tok, includeItemsFromAllDrives=True,
                             supportsAllDrives=True).execute()
        out+=r.get("files",[]); tok=r.get("nextPageToken")
        if not tok: break
    return out
def dl(fid, path, export=None):
    req=(drive.files().export_media(fileId=fid,mimeType=export) if export
         else drive.files().get_media(fileId=fid, supportsAllDrives=True))
    with open(path,"wb") as fh:
        d=MediaIoBaseDownload(fh,req,chunksize=8*1024*1024); done=False
        while not done: _,done=d.next_chunk()
def month_label(ed):
    idx=(BASE_M-1)+(ed-BASE_ED); y=BASE_Y+idx//12; m=idx%12+1
    return f"Edition {ed} · {calendar.month_name[m]} {y}"

# ---- discover newest edition that actually has proofs ----
ed_folders=[]
for f in ls(f"'{EDITIONS_PARENT}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"):
    m=re.match(r"^Edition-0*(\d+)$", f["name"].strip())
    if m: ed_folders.append((int(m.group(1)), f["id"], f["name"]))
ed_folders.sort(reverse=True)
chosen=None
for ed, fid, fname in ed_folders:
    kids=ls(f"'{fid}' in parents and trashed=false")
    proofs_folder=next((k for k in kids if k["mimeType"]=="application/vnd.google-apps.folder"
                        and k["name"].strip().lower()=="proofs"), None)
    db=next((k for k in kids if k["mimeType"]=="application/vnd.google-apps.spreadsheet"
             and "database" in k["name"].lower()), None)
    if not proofs_folder or not db: continue
    pdfs=[p for p in ls(f"'{proofs_folder['id']}' in parents and mimeType='application/pdf' and trashed=false")
          if "flatplan" not in p["name"].lower()]
    if pdfs:
        chosen=dict(ed=ed, db=db, proofs=proofs_folder["id"], pdfs=pdfs); break
if not chosen:
    print("NOEDITION no edition folder has proofs yet"); sys.exit(0)

ed=chosen["ed"]; edfile=f"ed{ed:03d}.html"; label=month_label(ed)
db_mod=drive.files().get(fileId=chosen["db"]["id"], fields="modifiedTime",
                         supportsAllDrives=True).execute()["modifiedTime"]

# change-gate: skip if same edition + same DB modifiedTime as last build
marker=os.path.join(OUTDIR,".lastbuild"); key=f"{ed}|{db_mod}"
if os.path.exists(marker) and open(marker).read().strip()==key:
    print("NOCHANGE", key); sys.exit(0)

if os.path.exists(WORK): shutil.rmtree(WORK)
os.makedirs(WORK+"/proofs"); os.makedirs(WORK+"/thumbs")
dl(chosen["db"]["id"], WORK+"/db.xlsx",
   export="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ---- DB parse ----
wb=load_workbook(WORK+"/db.xlsx", data_only=True)
ws=wb["Articles"] if "Articles" in wb.sheetnames else wb[wb.sheetnames[0]]
rows=list(ws.iter_rows(values_only=True))
hdr=[str(c).strip() if c is not None else "" for c in rows[0]]
def col(n):
    for i,h in enumerate(hdr):
        if h.lower()==n.lower(): return i
    return -1
C={k:col(k) for k in ["Headline","Section/Rubric","Status","Pages Claimed","Print Slug"]}
def pages(v):
    if v is None: return []
    if isinstance(v,(int,float)): return [int(v)]
    s=str(v).strip()
    m=re.match(r"^(\d+)(?:\.0)?\s*[–\-]\s*(\d+)(?:\.0)?$",s)
    if m: return list(range(int(m.group(1)),int(m.group(2))+1))
    m=re.match(r"^(\d+)(?:\.0)?$",s); return [int(m.group(1))] if m else []
arts=[]
for r in rows[1:]:
    if not any(r): continue
    pg=pages(r[C["Pages Claimed"]])
    if not pg: continue
    arts.append(dict(headline=str(r[C["Headline"]] or "").strip(),
        section=str(r[C["Section/Rubric"]] or "").strip(),
        status=str(r[C["Status"]] or "").strip(),
        pages=pg, slug=str(r[C["Print Slug"]] or "").strip().upper()))
slug2pages={a["slug"]:a["pages"] for a in arts if a["slug"]}

def vkey(t):
    s=t[:-4] if t.lower().endswith(".pdf") else t
    s=re.sub(r"^\s*\d+\s*[-–]\s*\d+\s*","",s)
    s=re.sub(r"^\s*\d+\s*","",s)
    s=re.sub(r"BT\.?\d+","",s,flags=re.I)
    s=re.sub(r"\((?:v)?\d+\)","",s,flags=re.I)
    s=re.sub(r"[_.\-–\s]+"," ",s).strip()
    return s.upper()
def fnpages(t):
    m=re.search(r"(\d+)\s*[-–]\s*(\d+)",t)
    if m: return list(range(int(m.group(1)),int(m.group(2))+1))
    m=re.match(r"^\s*(\d+)",t); return [int(m.group(1))] if m else []
def match(key,name):
    if key and key in slug2pages: return slug2pages[key]
    if key:
        for a in arts:
            if key in a["headline"].upper(): return a["pages"]
    return fnpages(name)

best={}
for f in chosen["pdfs"]:
    k=vkey(f["name"]) or ("_fn_"+f["name"])
    if k not in best or f["modifiedTime"]>best[k]["modifiedTime"]: best[k]=f
items=[]
for f in best.values():
    local=WORK+"/proofs/"+f["name"].replace("/","_"); dl(f["id"],local)
    pg=match(vkey(f["name"]),f["name"])
    if pg: items.append((local,pg))
items.sort(key=lambda x:(len(x[1])==1, x[1][0]))
have=gen.rasterise_placed(items, WORK+"/thumbs", WORK+"/tmp")

meta={}
for a in arts:
    for p in a["pages"]: meta[int(p)]=dict(sec=a["section"],title=a["headline"],status=a["status"])
os.makedirs(OUTDIR, exist_ok=True)
subtitle="Rendered from the saved running order · page numbers from the page table · latest proof per slug · self-contained"
summ=[f"<b>{len(have)}</b>/{ISSUE} pages proofed", f"<b>{len(arts)}</b> articles placed"]
gen.render(f"Edition-{ed:03d}", os.path.join(OUTDIR,edfile), WORK+"/thumbs", meta,
           f"Byline Times — Edition {ed:03d} · contact sheet", subtitle, summ,
           issue_pages=ISSUE, status_legend=True)

# ---- manifest upsert + index + current ----
mpath=os.path.join(OUTDIR,"editions.json")
man=json.load(open(mpath)) if os.path.exists(mpath) else []
man=[e for e in man if e.get("ed")!=ed]
man.append(dict(ed=ed, file=edfile, label=label, pages=len(have)))
man.sort(key=lambda e:e["ed"], reverse=True)
json.dump(man, open(mpath,"w"), indent=2)

def card(e, latest):
    tag=' <span class="tag">LATEST PROOFED</span>' if latest else ''
    return (f'      <a class="ed" href="{e["file"]}" target="_blank" rel="noopener noreferrer">\n'
            f'        <span class="ico">&#128196;</span>\n'
            f'        <span><span class="n">{e["label"]}{tag}</span>'
            f'<small>{e["pages"]} / {ISSUE} pages proofed</small></span>\n      </a>')
cards="\n".join(card(e, i==0) for i,e in enumerate(man))
INDEX=f'''<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex,nofollow,noarchive"><title>BylineOS — Contact Sheets</title>
<style>
:root{{--bg:#111214;--panel:#17181b;--line:#2a2c31;--white:#f5f5f5;--red:#e2231a;--grey:#9aa0a6;--teal:#0d9488}}
*{{box-sizing:border-box;margin:0;padding:0}}html,body{{height:100%}}
body{{background:var(--bg);color:var(--white);font-family:"Helvetica Neue",Arial,sans-serif;min-height:100%;display:flex;align-items:center;justify-content:center;padding:40px 20px}}
.card{{width:100%;max-width:720px}}.head{{display:flex;align-items:baseline;gap:14px;padding:2px 4px 20px;flex-wrap:wrap}}
.logo{{font-family:"Arial Black",Arial,sans-serif;font-weight:900;letter-spacing:-1px;font-size:30px}}.logo .o{{color:var(--red)}}
.head .t{{font-family:"Courier New",monospace;color:var(--grey);letter-spacing:3px;font-size:13px}}
.intro{{color:var(--grey);font-size:13px;line-height:1.6;padding:0 4px 18px;font-family:"Courier New",monospace}}
.btns{{display:grid;gap:14px}}
a.ed{{text-decoration:none;color:var(--white);background:var(--panel);border:1px solid var(--line);border-left:4px solid var(--teal);border-radius:14px;padding:20px 24px;display:flex;align-items:center;gap:18px;transition:.15s ease}}
a.ed:hover{{border-color:var(--teal);background:#1d1f23;transform:translateY(-2px)}}
a.ed .ico{{font-size:24px;width:32px;text-align:center}}
a.ed .n{{font-size:20px;font-weight:700;display:flex;align-items:center;gap:10px;flex-wrap:wrap}}
a.ed small{{display:block;font-weight:400;color:var(--grey);font-size:13px;letter-spacing:.3px;margin-top:4px;font-family:"Courier New",monospace}}
.tag{{font-family:"Courier New",monospace;font-size:11px;font-weight:700;letter-spacing:.5px;color:#04211e;background:var(--teal);border-radius:20px;padding:2px 9px}}
.foot{{color:var(--grey);font-size:12px;margin-top:22px;font-family:"Courier New",monospace;line-height:1.7;padding:0 4px}}.foot a{{color:var(--grey)}}
</style></head>
<body><div class="card">
  <div class="head"><div class="logo">BYLINE<span class="o">OS</span></div><div class="t">CONTACT SHEETS</div></div>
  <p class="intro">Whole-issue flow view &mdash; every page rebuilt from the print proofs, in reading order, spreads across the fold. Self-contained; share the link. Internal pacing view, not for granular proofing. Rebuilt automatically from the edition database.</p>
  <div class="btns">
{cards}
  </div>
  <div class="foot">&ldquo;Latest proofed&rdquo; = the newest edition with print proofs in Drive &mdash; the live edition in BylineOS may be further ahead but not yet laid out.<br><a href="../index.html">&#8592; back to launcher</a></div>
</div></body></html>'''
open(os.path.join(OUTDIR,"index.html"),"w",encoding="utf-8").write(INDEX)

open(os.path.join(OUTDIR,"current.html"),"w",encoding="utf-8").write(
 '<!DOCTYPE html><html lang="en"><head><meta charset="utf-8">'
 '<meta name="robots" content="noindex,nofollow,noarchive">'
 '<title>Current contact sheet — redirecting</title>'
 f'<meta http-equiv="refresh" content="0; url={edfile}">'
 f'<link rel="canonical" href="{edfile}"></head>'
 f'<body><p>Opening the latest contact sheet ({label})… '
 f'<a href="{edfile}">continue</a>.</p></body></html>')

open(marker,"w").write(key)
print(f"BUILT edition {ed} {edfile} {len(have)}/{ISSUE} pages; index has {len(man)} editions")
