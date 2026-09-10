#!/usr/bin/env python3
"""BylineOS contact sheet generator (render + rasterise), lifted from
claude/BylineOS_contactsheet_generator.py — placement by DB page, latest-per-slug."""
import os, subprocess, glob, base64, html, shutil, json, re
from PIL import Image
ISSUE_DEFAULT=68

def _leaves_from_pdf(pdf, tmp, dpi=110):
    if os.path.exists(tmp): shutil.rmtree(tmp)
    os.makedirs(tmp)
    subprocess.run(["pdftoppm","-jpeg","-r",str(dpi),pdf,os.path.join(tmp,"pg")],check=True)
    leaves=[]
    for ip in sorted(glob.glob(tmp+"/pg*.jpg")):
        im=Image.open(ip).convert("RGB"); w,h=im.size
        if w>h*1.2:
            mid=w//2; leaves.append(im.crop((0,0,mid,h))); leaves.append(im.crop((mid,0,w,h)))
        else: leaves.append(im)
    return leaves

def _save(im, outdir, pg, W=340):
    w,h=im.size
    im=im.resize((W,int(h*W/w)), Image.LANCZOS)
    im.save(os.path.join(outdir,f"p{pg:02d}.jpg"),"JPEG",quality=72,optimize=True)

def rasterise_placed(items, outdir, tmp):
    """items: list of (path, [pages]) — leaf i -> pages[i]. Later items win overlaps."""
    if os.path.exists(outdir): shutil.rmtree(outdir)
    os.makedirs(outdir)
    for path,pglist in items:
        if not os.path.exists(path): print("  MISSING",os.path.basename(path)); continue
        leaves=_leaves_from_pdf(path, tmp)
        for i,pg in enumerate(pglist):
            if i<len(leaves): _save(leaves[i], outdir, pg)
    if os.path.exists(tmp): shutil.rmtree(tmp)
    return sorted(int(os.path.basename(x)[1:3]) for x in glob.glob(outdir+"/p*.jpg"))

def _status_style(s):
    s=(s or '').strip().lower()
    return {'signed off':('Signed off','#16a34a'),'approved':('Signed off','#16a34a'),
     'ready':('Ready / laid out','#0d9488'),'ready for design':('Ready / laid out','#0d9488'),
     'laid out':('Ready / laid out','#0d9488'),'designing':('Ready / laid out','#0d9488'),
     'editing':('Editing','#ea580c'),'subediting':('Subediting','#2563eb'),
     'designed':('Designed','#7c3aed')}.get(s,('Planning','#94a3b8'))

def _ranges(l):
    out=[]
    if not l: return out
    a=l[0];prev=l[0]
    for x in l[1:]+[None]:
        if x!=prev+1: out.append(f"{a}" if a==prev else f"{a}–{prev}"); a=x
        prev=x if x else prev
    return out

def render(edition_label, out_html, thumbs, meta, title, subtitle, summarybits,
           issue_pages=ISSUE_DEFAULT, status_legend=True):
    have=sorted(int(os.path.basename(x)[1:3]) for x in glob.glob(thumbs+"/p*.jpg"))
    missing=[p for p in range(1,issue_pages+1) if p not in have]
    def uri(p):
        fp=os.path.join(thumbs,f"p{p:02d}.jpg")
        return "data:image/jpeg;base64,"+base64.b64encode(open(fp,'rb').read()).decode() if os.path.exists(fp) else None
    tiles=[]
    def _cls(p):
        c='pg'
        if issue_pages>=8:
            if p<=2: c+=' cvr cvr-f'
            if p==2: c+=' cvr-f2'
            if p==3: c+=' rowbreak'
            if p>=issue_pages-1: c+=' cvr cvr-b'
            if p==issue_pages-1: c+=' rowbreak'
        return c
    for p in range(1,issue_pages+1):
        u=uri(p); m=meta.get(p)
        if u:
            lbl,col=_status_style(m['status']) if (m and m.get('status')) else ('Proof in','#0d9488')
            sec=html.escape((m['sec'] or '').upper()) if m else ''
            t=html.escape(m['title'] or '') if m else ''
            tiles.append(f'<figure class="{_cls(p)}" style="--sc:{col}" title="{t}"><span class="pgnum">{p}</span><img loading="lazy" src="{u}" alt="p{p}: {t}"><figcaption><span class="rub">{sec}</span><span class="dot" style="background:{col}"></span></figcaption></figure>')
        else:
            sec=f'<span class="mrub">{html.escape((m["sec"] or "").upper())}</span>' if (m and m.get('sec')) else ''
            tiles.append(f'<div class="{_cls(p)} gap"><span class="pgnum">{p}</span><span class="gaplbl">no proof</span>{sec}</div>')
    legend=''
    if status_legend:
        it=[('Signed off','#16a34a'),('Ready / laid out','#0d9488'),('Editing','#ea580c'),('Subediting','#2563eb'),('Designed','#7c3aed'),('Planning','#94a3b8')]
        legend='<div class="legend">'+''.join(f'<span class="lg"><i style="background:{c}"></i>{l}</span>' for l,c in it)+'</div>'
    gaps_txt=", ".join(_ranges(missing)) if missing else "none"
    summ=' '.join(f'<span>{b}</span>' for b in (summarybits+[f'gaps: {gaps_txt}']))
    CSS=r'''
:root{--ink:#0f172a;--mut:#64748b;--line:#e2e8f0;--bg:#f4f6f9}
*{box-sizing:border-box}body{margin:0;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif;color:var(--ink);background:var(--bg);-webkit-font-smoothing:antialiased}
header{padding:18px clamp(14px,4vw,40px) 6px}h1{font-size:clamp(17px,3vw,23px);margin:0 0 2px;letter-spacing:-.01em}
.sub{color:var(--mut);font-size:13px;margin:0 0 10px}.summary{display:flex;flex-wrap:wrap;gap:5px 18px;font-size:13px;margin-bottom:8px}.summary b{font-variant-numeric:tabular-nums}
.legend{display:flex;flex-wrap:wrap;gap:6px 14px;font-size:12px;color:var(--mut);margin:2px 0 4px}.lg{display:inline-flex;align-items:center;gap:5px}.lg i{width:11px;height:11px;border-radius:3px}
.toolbar{position:sticky;top:0;z-index:20;background:rgba(244,246,249,.92);backdrop-filter:blur(6px);border-bottom:1px solid var(--line);padding:9px clamp(14px,4vw,40px);display:flex;align-items:center;gap:10px;flex-wrap:wrap}
.toolbar .lab{font-size:12px;color:var(--mut);font-weight:600}.seg{display:inline-flex;border:1px solid var(--line);border-radius:10px;overflow:hidden;background:#fff}
.seg button{appearance:none;border:0;background:#fff;color:var(--ink);font:inherit;font-size:13px;font-weight:600;padding:7px 14px;cursor:pointer;border-left:1px solid var(--line)}.seg button:first-child{border-left:0}.seg button.on{background:#0d9488;color:#fff}
main{padding:10px clamp(14px,4vw,40px) 60px;max-width:1240px;margin:0 auto}
#grid{position:relative;display:grid;grid-template-columns:repeat(var(--cols,2),1fr);gap:12px}
#grid .spacer{display:none}#grid.cols-2{--cols:2}#grid.cols-4{--cols:4}#grid.cols-6{--cols:6}#grid.cols-8{--cols:8}#grid.cols-16{--cols:16}#grid.cols-2 .spacer{display:block}
#grid.cols-2::before{content:"";position:absolute;top:0;bottom:0;left:50%;border-left:2px dashed #c3cede;opacity:.6;pointer-events:none}
.pg{position:relative;margin:0;background:#fff;border:1px solid var(--line);border-radius:7px;overflow:hidden;box-shadow:0 1px 2px rgba(15,23,42,.05)}
.pg img{display:block;width:100%;height:auto;border-bottom:1px solid var(--line)}
.pgnum{position:absolute;top:5px;right:6px;background:rgba(15,23,42,.72);color:#fff;font-size:10px;font-weight:700;padding:1px 5px;border-radius:5px;font-variant-numeric:tabular-nums;z-index:2}
figcaption{display:flex;align-items:center;gap:6px;padding:4px 7px;border-left:4px solid var(--sc,#cbd5e1)}
.rub{font-size:9.5px;letter-spacing:.04em;font-weight:700;color:var(--mut);text-transform:uppercase;flex:1;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.dot{width:8px;height:8px;border-radius:50%;flex:none}
.pg.gap{aspect-ratio:1/1.32;display:flex;flex-direction:column;align-items:center;justify-content:center;background:repeating-linear-gradient(135deg,#fff,#fff 8px,#f6f8fb 8px,#f6f8fb 16px);border:1px dashed #cbd5e1}
.gaplbl{color:#94a3b8;font-size:11px}.mrub{color:#b6c0cf;font-size:9px;font-weight:700;margin-top:3px;text-transform:uppercase;text-align:center;padding:0 4px}
#grid.cols-6 figcaption,#grid.cols-8 figcaption,#grid.cols-16 figcaption{display:none}#grid.cols-6 .pgnum,#grid.cols-8 .pgnum{font-size:9px;padding:0 4px}
#grid.cols-16{gap:5px}
#grid.cols-16 .pg{border-radius:3px;box-shadow:none;border-bottom:3px solid var(--sc,#cbd5e1)}
#grid.cols-16 .pg.gap{border-bottom:1px dashed #cbd5e1}
#grid.cols-16 .pgnum{top:2px;right:2px;font-size:8px;font-weight:700;padding:0 3px;border-radius:4px;background:rgba(15,23,42,.6)}
#grid.cols-16 .gaplbl,#grid.cols-16 .mrub{display:none}
#grid.cols-16 .rowbreak{grid-column-start:1}
#grid.cols-16 .cvr-f{margin-bottom:16px}#grid.cols-16 .cvr-b{margin-top:16px}
body.wideview main{max-width:none;padding:8px clamp(10px,1.5vw,18px) 10px}
.footer{color:var(--mut);font-size:12px;text-align:center;padding:16px 0 0}
@media(max-width:640px){#grid.cols-2::before{display:none}}
@media print{@page{size:A4 portrait;margin:9mm}.toolbar{display:none}body{background:#fff}header,main{padding-left:0;padding-right:0}#grid{--cols:2 !important}#grid::before{display:none}.pg{break-inside:avoid}}
'''
    JS = """
(function(){
  var g=document.getElementById('grid');
  function setCols(n){g.className='cols-'+n;document.body.classList.toggle('wideview',parseInt(n,10)>=16);}
  document.querySelectorAll('.seg button').forEach(function(b){
    b.addEventListener('click',function(){
      setCols(b.getAttribute('data-cols'));
      document.querySelectorAll('.seg button').forEach(function(x){x.classList.toggle('on',x===b)});
    });
  });
})();
"""
    doc=f'''<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>{html.escape(title)}</title><style>{CSS}</style></head>
<body><header><h1>{html.escape(title)}</h1><p class="sub">{subtitle}</p><div class="summary">{summ}</div>{legend}</header>
<div class="toolbar"><span class="lab">View</span><div class="seg"><button data-cols="2" class="on">Spreads</button><button data-cols="4">4-up</button><button data-cols="6">6-up</button><button data-cols="8">8-up</button><button data-cols="16" title="the whole flatplan on one screen">16-up</button></div><span class="lab" style="opacity:.8">— zoom out to see more pages across</span></div>
<main><div id="grid" class="cols-2"><div class="spacer"></div>
{chr(10).join(tiles)}
</div><p class="footer">Internal flow / pacing view · {edition_label} · low-res thumbnails, not for granular proofing.</p></main>
<script>{JS}</script>
</body></html>'''
    open(out_html,'w',encoding='utf-8').write(doc)
    print("wrote",os.path.basename(out_html),len(doc),"bytes ; pages",len(have),"/",issue_pages,"; gaps",gaps_txt)
    return have, missing
