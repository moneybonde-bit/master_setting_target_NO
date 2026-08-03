/* ==========================================================================
   AIM+ TARGET SETTING — SUMMARY + ANALISIS AP
   Spesifikasi v2.0 · PEARL · Wahana Visi Indonesia
   --------------------------------------------------------------------------
   Halaman ini tidak menghitung target. Ia menampilkan apa yang dikirim AP dan
   menerapkan logika status yang sudah disetujui PEARL (§5.3), tanpa perubahan.

   T1 dan T2 tidak mungkin terjadi di sini: analisis dihitung baris demi baris
   dari tblIndicators memakai Row ID sebagai kunci, bukan dari tabel kedua yang
   panjangnya harus dijaga manual. Setiap baris baru langsung ikut dianalisis.
   ========================================================================== */
"use strict";

const CFG={}, S={ rows:[], asumsi:[], pemetaan:[], apz:{}, cat:[], local:false };
let AP_LIST=[], ZONALS=[], OUTCOMES=[], IND_LIST=[];

const isBlank = v => v===null||v===undefined||v==="";
/* KONVENSI PEARL: N() memperlakukan kosong sebagai 0, dan 0 sebagai belum ada data */
const N = v => { const f=parseFloat(v); return isFinite(f)?f:0; };
function esc(s){return String(s==null?"":s).replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));}

/* ---------- kosakata status (§3) ---------- */
/* Kosakata mengikuti pilihan yang dipakai PEARL di file Excel.
   Logikanya tidak berubah — hanya labelnya. */
const ST_IND=[
 {v:"Sesuai Target/Threshold", ic:"●", cls:"s-baik",   desc:"Endline mencapai threshold DAN delta mencapai target"},
 {v:"Ditinjau",                ic:"◆", cls:"s-tinjau", desc:"Salah satu tercapai, satunya belum"},
 {v:"Perhatian",               ic:"▲", cls:"s-hati",   desc:"Keduanya belum tercapai"},
 {v:"Belum ada data",          ic:"◧", cls:"s-belum",  desc:"Baseline atau endline masih 0"}
];
const ST_THR=[
 {v:">= threshold",       ic:"●", cls:"s-baik"},
 {v:"di bawah threshold", ic:"▲", cls:"s-hati"},
 {v:"Belum ada data",     ic:"◧", cls:"s-belum"}
];
const ST_TGT=[
 {v:">= target",      ic:"●", cls:"s-baik"},
 {v:"dibawah target", ic:"▲", cls:"s-hati"},
 {v:"Belum ada data", ic:"◧", cls:"s-belum"}
];
const ICON={}; ST_IND.concat(ST_THR,ST_TGT).forEach(x=>{ICON[x.v]=x;});
function pill(v){const x=ICON[v]||{ic:"–",cls:"s-belum"};
  return '<span class="pill '+x.cls+'"><span class="ic">'+x.ic+'</span>'+esc(v)+'</span>';}

/* ---------- pemuatan ---------- */
function adopt(){
  const c=window.WVI_CONFIG;
  Object.assign(CFG,{cycle:c.cycle,version:c.version,data_date:c.data_date,owner:c.owner,
    code:c.access_code,target_delta:(window.WVI_ASUMSI||{}).target_delta_default||0.1});
  ZONALS=c.zonal.slice(); OUTCOMES=c.outcomes.slice();
  AP_LIST=c.ap.map(a=>({ap:a.ap,zonal:a.zonal}));
  S.apz={}; AP_LIST.forEach(a=>S.apz[a.ap]=a.zonal);
  S.cat=c.catalogue.map(x=>Object.assign({},x));
  IND_LIST=S.cat.map(x=>x.ind);

  const pack=window.WVI_INDICATORS, col=pack.columns;
  S.rows=pack.rows.map(r=>{const o={}; col.forEach((k,i)=>o[k]=r[i]); return o;});
  S.asumsi=(window.WVI_ASUMSI.rows||[]).map(a=>Object.assign({},a));
  S.pemetaan={aps:window.WVI_PEMETAAN.aps.slice(),
    rows:window.WVI_PEMETAAN.rows.map(p=>({code:p.code,ind:p.ind,v:p.v.slice()}))};
}

/* ---------- indeks bantu ---------- */
let AS_BY_IND={}, PEM_BY_IND={}, CAT_BY_IND={};
function reindex(){
  AS_BY_IND={}; S.asumsi.forEach(a=>{ AS_BY_IND[a.ind]=a; });
  PEM_BY_IND={}; S.pemetaan.rows.forEach(p=>{ PEM_BY_IND[p.ind]=p; });
  CAT_BY_IND={}; S.cat.forEach(c=>{ CAT_BY_IND[c.ind]=c; });
}
const arahOf  = ind => (AS_BY_IND[ind]||{}).arah || "Naik";
const targetOf= ind => { const a=AS_BY_IND[ind];
  return (a&&a.delta!==null&&a.delta!==undefined)?a.delta:CFG.target_delta; };
function berlakuOf(ind,ap){
  const p=PEM_BY_IND[ind]; if(!p) return "Yes";
  const i=S.pemetaan.aps.indexOf(ap); if(i<0) return "Yes";
  const v=p.v[i]; return (v&&String(v).trim().toLowerCase()==="no")?"No":"Yes";
}
const shortOf = ind => (CAT_BY_IND[ind]||{}).short || String(ind||"").slice(0,42);

/* ==========================================================================
   LOGIKA STATUS — persis §5.3, tidak diubah
   ========================================================================== */
function recompute(){
  reindex();
  const seen={}, dupKey={};
  S.rows.forEach(r=>{
    const k=r["Area Program"]+"|"+r.Indicator;
    dupKey[k]=(dupKey[k]||0)+1;
  });
  S.rows.forEach(r=>{
    const ind=r.Indicator, ap=r["Area Program"];
    const base=N(r.Pct_Base), end=N(r.Pct_LOP), thr=N(r.Threshold);
    const arah=arahOf(ind), tgt=targetOf(ind);

    r._arah=arah; r._target=tgt;
    /* EPS: perbandingan >= dan <= pada bilangan pecahan biner tidak stabil di ambang.
       Dua baris (IND-291, IND-296) delta-nya tepat sama dengan Target Delta; tanpa
       toleransi ini, hasilnya berubah hanya karena presisi penyimpanan angka.
       Aturannya "≥" dan "≤", jadi nilai yang sama persis dihitung TERCAPAI. */
    const EPS=1e-9;
    r._thr_status = end===0 ? "Belum ada data"
      : (arah==="Turun" ? (end<=thr+EPS?">= threshold":"di bawah threshold")
                        : (end>=thr-EPS?">= threshold":"di bawah threshold"));
    r._delta = (base===0||end===0) ? null : end-base;
    r._tgt_status = r._delta===null ? "Belum ada data"
      : (arah==="Turun" ? (r._delta<=tgt+EPS?">= target":"dibawah target")
                        : (r._delta>=tgt-EPS?">= target":"dibawah target"));
    r._status = (r._thr_status==="Belum ada data"||r._tgt_status==="Belum ada data") ? "Belum ada data"
      : (r._thr_status===">= threshold"&&r._tgt_status===">= target") ? "Sesuai Target/Threshold"
      : (r._thr_status===">= threshold"||r._tgt_status===">= target") ? "Ditinjau"
      : "Perhatian";
    r._berlaku = berlakuOf(ind,ap);
    r._short = shortOf(ind);
    r._dupe = dupKey[ap+"|"+ind]>1;
    /* delta yang arahnya berlawanan dengan arah indikator (§5.8) */
    r._wrongway = r._delta!==null && ((arah==="Naik"&&r._delta<0)||(arah==="Turun"&&r._delta>0));
    const id=r.Row_ID; r._iddupe = seen[id]?true:false; seen[id]=1;
  });
}

/* ==========================================================================
   PEMERIKSAAN INTEGRITAS — ditampilkan di band header kedua halaman
   ========================================================================== */
function checks(){
  const n=S.rows.length;
  const idDupe=S.rows.filter(r=>r._iddupe).length;
  const comboDupe=uniq(S.rows.filter(r=>r._dupe).map(r=>r["Area Program"]+"|"+r.Indicator)).length;
  const dupeRows=S.rows.filter(r=>r._dupe).length;
  const total=ST_IND.reduce((a,s)=>a+S.rows.filter(r=>r._status===s.v).length,0);
  const oor=S.rows.filter(r=>Math.abs(N(r.Pct_Base))>1||Math.abs(N(r.Pct_LOP))>1).length;
  const nd=S.rows.filter(r=>(N(r.Den_Base)>0&&N(r.Num_Base)>N(r.Den_Base))||
                            (N(r.Den_LOP)>0&&N(r.Num_LOP)>N(r.Den_LOP))).length;
  const unknownAsumsi=S.asumsi.filter(a=>IND_LIST.indexOf(a.ind)<0).length;
  const unknownAP=S.pemetaan.aps.filter(a=>!AP_LIST.some(x=>x.ap===a)).length;
  return {
    n:n,
    rowid: idDupe===0 ? "ROW ID OK" : "ROW ID GANDA ("+idDupe+")",
    sinkron: total===n ? "SINKRON" : "TIDAK SINKRON ("+total+" dari "+n+")",
    dupe: comboDupe===0 ? "TIDAK ADA DUPLIKAT" : comboDupe+" KOMBINASI AP × INDIKATOR GANDA",
    dupeRows:dupeRows, comboDupe:comboDupe, oor:oor,
    range: oor===0 ? "PROPORSI OK" : oor+" PROPORSI DI LUAR 0–100%",
    numden: nd===0 ? "NUM ≤ DEN" : nd+" BARIS NUMERATOR > DENOMINATOR", nd:nd,
    unknownAsumsi:unknownAsumsi, unknownAP:unknownAP,
    ok: idDupe===0 && total===n && comboDupe===0 && oor===0 && unknownAsumsi===0 && unknownAP===0
  };
}

/* ---------- format (§3) ---------- */
const uniq = a => Array.from(new Set(a));
const n0 = v => isBlank(v)?"—":Math.round(Number(v)).toLocaleString("id-ID");
/* 0 dan kosong sama-sama tampil sebagai — , sesuai konvensi PEARL */
const pctD = (v,d) => { const f=N(v); return f===0?"—":(f*100).toFixed(d===undefined?1:d)+"%"; };
const pctT = v => { const f=N(v); return f===0?"TBC":(f*100).toFixed(1)+"%"; };
const ppD = v => v===null||v===undefined?"—":((v>0?"+":"")+(v*100).toFixed(1)+"pp");
const med = a => { const b=a.filter(v=>v!=null&&isFinite(v)).sort((x,y)=>x-y);
  if(!b.length) return null; const m=b.length>>1;
  return b.length%2?b[m]:(b[m-1]+b[m])/2; };
const avg = a => { const b=a.filter(v=>v!=null&&isFinite(v)); 
  return b.length?b.reduce((x,y)=>x+y,0)/b.length:null; };

/* ---------- penyimpanan lokal ---------- */
const LS="wvi_aimplus_v2";
function saveLocal(){
  try{ localStorage.setItem(LS,JSON.stringify({
    saved:new Date().toISOString(), asumsi:S.asumsi, pemetaan:S.pemetaan,
    target_delta:CFG.target_delta, rows:S.rows.map(r=>{
      const o={}; window.WVI_INDICATORS.columns.forEach(k=>o[k]=r[k]); return o;})}));
    S.local=true; chip();
  }catch(e){ S.local=false; }
}
function loadLocal(){ try{ const s=localStorage.getItem(LS); return s?JSON.parse(s):null; }catch(e){ return null; } }
function clearLocal(){ try{ localStorage.removeItem(LS); }catch(e){} }
function localOK(){ try{ localStorage.setItem("__t","1"); localStorage.removeItem("__t"); return true; }catch(e){ return false; } }
/* ==========================================================================
   GRAFIK — SVG sebaris, tanpa library. Judul ditulis di sel, bukan chart title.
   Warna seri persis §3.
   ========================================================================== */
const CLR={
  baik:"#155930", tinjau:"#E8A33C", hati:"#B10831", belum:"#D8D6D1",
  baseline:"#0C7993", threshold:"#D8D6D1", endline:"#FF5515",
  zonal:["#0C7993","#155930","#FF5515","#3F3D4C"], ref:"#3F3D4C"
};
const SER_IND=[["Sesuai Target/Threshold",CLR.baik],["Ditinjau",CLR.tinjau],
               ["Perhatian",CLR.hati],["Belum ada data",CLR.belum]];
const SER_THR=[[">= threshold",CLR.baik],["di bawah threshold",CLR.hati],
               ["Belum ada data",CLR.belum]];
const SER_TGT=[[">= target",CLR.baik],["dibawah target",CLR.hati],["Belum ada data",CLR.belum]];

function legend(series,extra){
  return '<div class="legend">'+series.map(s=>
    '<div><i style="background:'+s[1]+';border-color:'+(s[1]===CLR.belum||s[1]===CLR.threshold?CLR.ref:s[1])+'"></i>'+
    esc(s[0])+'</div>').join('')+(extra||'')+'</div>';
}

/* stacked bar horizontal. pct100=true → 100% stacked. items:[{label,vals:[]}] */
function chartStack(items,series,opt){
  opt=opt||{};
  const labW=opt.labW||150, W=700, BH=opt.bh||17, GAP=opt.gap||9;
  const PITCH=BH+GAP, H=Math.max(28,items.length*PITCH+10), PW=W-labW-52;
  const totals=items.map(it=>it.vals.reduce((a,b)=>a+b,0));
  const scale=opt.pct100?null:Math.max(1,...totals);
  let s='<svg viewBox="0 0 '+W+' '+H+'" role="img">';
  if(!items.length) s+='<text x="'+labW+'" y="18" font-size="11" fill="#8A8894">Tidak ada data pada filter ini</text>';
  items.forEach((it,i)=>{
    const y=i*PITCH+4, tot=totals[i];
    s+='<text x="'+(labW-8)+'" y="'+(y+BH/2+4)+'" text-anchor="end" font-size="10.5" font-weight="600" fill="#111222">'+
       esc(it.label)+'</text>';
    if(!tot){
      s+='<rect x="'+labW+'" y="'+y+'" width="'+PW+'" height="'+BH+'" fill="#F7F6F4" stroke="#EDEBE6"/>'+
         '<text x="'+(labW+6)+'" y="'+(y+BH/2+4)+'" font-size="9" fill="#A9A6B0">tidak ada baris</text>';
      return;
    }
    let x=labW;
    it.vals.forEach((v,k)=>{
      if(!v) return;
      const w=opt.pct100 ? PW*v/tot : PW*v/scale;
      const col=series[k][1];
      s+='<rect x="'+x.toFixed(1)+'" y="'+y+'" width="'+w.toFixed(1)+'" height="'+BH+'" fill="'+col+'"'+
         (col===CLR.belum?' stroke="#C9C6CE" stroke-width=".6"':'')+'/>';
      const lbl = opt.labelCount ? v : (opt.pct100?Math.round(v/tot*100)+"%":v);
      if(w>18) s+='<text x="'+(x+w/2).toFixed(1)+'" y="'+(y+BH/2+3.5)+'" text-anchor="middle" font-size="9" font-weight="700" fill="'+
        (col===CLR.belum||col===CLR.tinjau?"#3F3D4C":"#FFFFFF")+'">'+lbl+'</text>';
      x+=w;
    });
    s+='<text x="'+(labW+PW+8)+'" y="'+(y+BH/2+4)+'" font-size="9.5" fill="#8A8894">'+tot+'</text>';
  });
  return s+'</svg>';
}

/* bar horizontal seri tunggal. items:[{label,v,hl}] */
function chartBar(items,color,opt){
  opt=opt||{};
  const labW=opt.labW||150, W=700, BH=opt.bh||19, PITCH=BH+11;
  const H=Math.max(28,items.length*PITCH+8), PW=W-labW-56;
  const max=Math.max(1,...items.map(i=>i.v));
  let s='<svg viewBox="0 0 '+W+' '+H+'" role="img">';
  items.forEach((it,i)=>{
    const y=i*PITCH+4, w=PW*it.v/max;
    s+='<text x="'+(labW-8)+'" y="'+(y+BH-5)+'" text-anchor="end" font-size="10.5" font-weight="600" fill="#111222">'+esc(it.label)+'</text>'+
       '<rect x="'+labW+'" y="'+y+'" width="'+Math.max(0,w).toFixed(1)+'" height="'+BH+'" fill="'+(it.hl||color)+'"/>'+
       '<text x="'+(labW+w+7).toFixed(1)+'" y="'+(y+BH-5)+'" font-size="10.5" font-weight="700" fill="'+(it.hl||color)+'">'+n0(it.v)+'</text>';
  });
  return s+'</svg>';
}

/* clustered bar 2 seri, skala 0–100%. items:[{label,a,b}] */
function chartPair(items,s1,s2,opt){
  opt=opt||{};
  const labW=opt.labW||210, W=700, BH=9, OV=2, PITCH=BH*2+OV+13;
  const H=Math.max(30,items.length*PITCH+14), PW=W-labW-58;
  const max=Math.max(0.0001,...items.map(i=>Math.max(i.a||0,i.b||0)));
  const sc=Math.min(1,Math.max(0.2,Math.ceil(max*10)/10));
  let s='<svg viewBox="0 0 '+W+' '+H+'" role="img">';
  for(let g=0;g<=4;g++){const x=labW+PW*g/4;
    s+='<line x1="'+x+'" y1="8" x2="'+x+'" y2="'+(H-10)+'" stroke="#EDEBE6"/>'+
       '<text x="'+x+'" y="'+(H-2)+'" text-anchor="middle" font-size="8" fill="#A9A6B0">'+Math.round(sc*100*g/4)+'%</text>';}
  items.forEach((it,i)=>{
    const y=i*PITCH+10;
    s+='<text x="'+(labW-8)+'" y="'+(y+BH+2)+'" text-anchor="end" font-size="9.5" font-weight="600" fill="#111222">'+esc(it.label)+'</text>';
    const wa=PW*Math.min(1,(it.a||0)/sc), wb=PW*Math.min(1,(it.b||0)/sc);
    s+='<rect x="'+labW+'" y="'+y+'" width="'+wa.toFixed(1)+'" height="'+BH+'" fill="'+s1[1]+'"/>'+
       '<rect x="'+labW+'" y="'+(y+BH+OV)+'" width="'+wb.toFixed(1)+'" height="'+BH+'" fill="'+s2[1]+'" stroke="'+CLR.ref+'" stroke-width=".6"/>'+
       '<text x="'+(labW+Math.max(wa,wb)+7).toFixed(1)+'" y="'+(y+BH+2)+'" font-size="8.5" fill="#3F3D4C">'+
       (it.a?(it.a*100).toFixed(1)+"%":"—")+' / '+(it.b?(it.b*100).toFixed(1)+"%":"TBC")+'</text>';
  });
  return s+'</svg>';
}

/* bar menyimpang dari nol, dengan penanda target. items:[{label,v,target}] */
function chartDiverge(items,opt){
  opt=opt||{};
  const labW=opt.labW||210, W=700, BH=13, PITCH=BH+11;
  const H=Math.max(34,items.length*PITCH+22), PW=W-labW-70;
  const vals=items.flatMap(i=>[i.v||0,i.target||0]);
  const m=Math.max(0.05,...vals.map(Math.abs)), sc=Math.ceil(m*20)/20;
  const zero=labW+PW/2, half=PW/2;
  const x=v=>zero+half*Math.max(-1,Math.min(1,(v||0)/sc));
  let s='<svg viewBox="0 0 '+W+' '+H+'" role="img">';
  s+='<line x1="'+zero+'" y1="12" x2="'+zero+'" y2="'+(H-12)+'" stroke="#3F3D4C" stroke-width="1.2"/>'+
     '<text x="'+zero+'" y="8" text-anchor="middle" font-size="8" fill="#3F3D4C">0</text>'+
     '<text x="'+(labW)+'" y="8" text-anchor="start" font-size="8" fill="#A9A6B0">−'+(sc*100).toFixed(0)+'pp</text>'+
     '<text x="'+(labW+PW)+'" y="8" text-anchor="end" font-size="8" fill="#A9A6B0">+'+(sc*100).toFixed(0)+'pp</text>';
  items.forEach((it,i)=>{
    const y=i*PITCH+14, xv=x(it.v), xt=x(it.target);
    s+='<text x="'+(labW-8)+'" y="'+(y+BH-3)+'" text-anchor="end" font-size="9.5" font-weight="600" fill="#111222">'+esc(it.label)+'</text>';
    s+='<rect x="'+Math.min(zero,xv).toFixed(1)+'" y="'+y+'" width="'+Math.abs(xv-zero).toFixed(1)+
       '" height="'+BH+'" fill="'+CLR.endline+'"/>';
    s+='<line x1="'+xt.toFixed(1)+'" y1="'+(y-3)+'" x2="'+xt.toFixed(1)+'" y2="'+(y+BH+3)+
       '" stroke="'+CLR.ref+'" stroke-width="2"/>';
    s+='<text x="'+(labW+PW+8)+'" y="'+(y+BH-3)+'" font-size="8.5" fill="#3F3D4C">'+
       (it.v==null?"—":(it.v>0?"+":"")+(it.v*100).toFixed(1))+' / '+
       (it.target>0?"+":"")+(it.target*100).toFixed(0)+'</text>';
  });
  return s+'</svg>';
}
/* ==========================================================================
   KERANGKA HALAMAN + BAND FILTER
   ========================================================================== */
const F={ zonal:[], ap:[], outcome:[], status:[], ind:[], berlaku:"Yes", period:"Both",
  sort:null, sortDir:1, showAll:false, fbOpen:true };

function card(lab,val,sub,cls){
  return '<div class="card '+(cls||"neutral")+'"><div class="lab">'+lab+'</div>'+
    '<div><div class="val">'+val+'</div><div class="sub">'+(sub||"&nbsp;")+'</div></div></div>';
}

/* ---------- filter ---------- */
function filtered(){
  return S.rows.filter(r=>
    (!F.zonal.length   || F.zonal.indexOf(r.Zonal)>=0) &&
    (!F.ap.length      || F.ap.indexOf(r["Area Program"])>=0) &&
    (!F.outcome.length || F.outcome.indexOf(r.Outcome)>=0) &&
    (!F.status.length  || F.status.indexOf(r._status)>=0) &&
    (F.berlaku==="(Semua)" || r._berlaku===F.berlaku));
}
function chips(field,items,multi){
  const sel=multi?F[field]:[F[field]];
  return items.map(it=>'<button class="chipf'+(sel.indexOf(it.v)>=0?" sel":"")+
    (it.n===0?" nodata":"")+'" data-slice="'+field+'" data-val="'+esc(it.v)+'">'+esc(it.v)+
    (it.n!==undefined?'<span class="n">'+it.n+'</span>':'')+'</button>').join('');
}
function fbRow(label,field,items,multi,hint){
  const active=multi?F[field].length>0:false;
  return '<div class="fb-row"><div class="fb-lab">'+esc(label)+
    (active?'<button class="fb-x" data-clear="'+field+'" title="Bersihkan '+esc(label)+'">×</button>':'')+
    '</div><div class="fb-chips">'+chips(field,items,multi)+
    (hint?'<span class="fb-hint">'+hint+'</span>':'')+'</div></div>';
}
function filterBand(withStatus){
  const all=S.rows;
  const hidden=all.length-all.filter(r=>F.berlaku==="(Semua)"||r._berlaku===F.berlaku).length;
  /* daftar AP mengikuti Zonal yang dipilih — 17 tombol jadi 3–8 */
  const apPool=F.zonal.length?AP_LIST.filter(a=>F.zonal.indexOf(a.zonal)>=0):AP_LIST;
  const nAktif=F.zonal.length+F.ap.length+F.outcome.length+F.status.length+(F.berlaku!=="Yes"?1:0);

  if(!F.fbOpen){
    return '<div class="fb closed"><div class="fb-head">'+
      '<span class="fb-title">Filter</span>'+
      '<span class="fb-sum">'+(nAktif?esc(activeText()):'tanpa filter')+'</span>'+
      '<div class="fb-sp"></div>'+
      '<button class="fb-btn" data-act="fbOpen">Ubah filter ▾</button>'+
      '</div></div>';
  }
  return '<div class="fb"><div class="fb-head">'+
      '<span class="fb-title">Filter</span>'+
      '<span class="fb-sum">'+(nAktif?esc(activeText()):'tanpa filter')+'</span>'+
      '<div class="fb-sp"></div>'+
      (nAktif?'<button class="fb-btn" data-act="fbReset">Bersihkan semua</button>':'')+
      '<button class="fb-btn" data-act="fbClose">Sembunyikan ▴</button>'+
    '</div>'+
    fbRow("Zonal","zonal",ZONALS.map(z=>({v:z,n:all.filter(r=>r.Zonal===z).length})),true)+
    fbRow("Area Program","ap",apPool.map(a=>({v:a.ap,n:all.filter(r=>r["Area Program"]===a.ap).length})),true,
      F.zonal.length?'daftar mengikuti Zonal yang dipilih':'pilih Zonal untuk memendekkan daftar ini')+
    fbRow("Outcome","outcome",OUTCOMES.map(o=>({v:o,n:all.filter(r=>r.Outcome===o).length})),true)+
    (withStatus?fbRow("Status","status",ST_IND.map(s=>({v:s.v,
      n:all.filter(r=>r._status===s.v).length})),true):'')+
    fbRow("Berlaku","berlaku",[{v:"Yes",n:all.filter(r=>r._berlaku==="Yes").length},
      {v:"No",n:all.filter(r=>r._berlaku==="No").length},{v:"(Semua)",n:all.length}],false)+
    (hidden>0?'<div class="fb-note"><b>'+n0(hidden)+' baris</b> disembunyikan oleh Berlaku = '+
      esc(F.berlaku)+'</div>':'')+
  '</div>';
}
function activeText(){
  const b=[];
  if(F.zonal.length) b.push(F.zonal.join(", "));
  if(F.ap.length) b.push(F.ap.length===1?F.ap[0]:F.ap.length+" AP");
  if(F.outcome.length) b.push(F.outcome.join(", "));
  if(F.status.length) b.push(F.status.join(", "));
  if(F.berlaku!=="Yes") b.push("Berlaku "+F.berlaku);
  return b.join(" · ");
}
function activeLine(){
  const s=activeText();
  return s?' <span class="hint">'+esc(s)+'</span>':'';
}
function sortRows(rows,cols){
  if(!F.sort) return rows;
  const k=F.sort;
  return rows.slice().sort((a,b)=>{
    let x=a[k], y=b[k];
    const nx=typeof x==="number", ny=typeof y==="number";
    if(nx&&ny) return (x-y)*F.sortDir;
    return String(x==null?"":x).localeCompare(String(y==null?"":y))*F.sortDir;
  });
}
const th = (label,key,cls) => '<th class="'+(cls||"")+(F.sort===key?" sorted":"")+'" data-sort="'+key+'">'+
  label+(F.sort===key?(F.sortDir>0?' ▲':' ▼'):'')+'</th>';

/* ==========================================================================
   HALAMAN 1 — SUMMARY
   ========================================================================== */
function renderSummary(){
  const rows=filtered();

  const COLS=[["Zonal","Zonal"],["Area Program","Area Program"],["Outcome","Outcome"],["Code","Code"],
    ["Indicator","Indikator"],["Num_Base","Num Base"],["Den_Base","Den Base"],["Pct_Base","% Baseline"],
    ["Num_LOP","Num LOP"],["Den_LOP","Den LOP"],["Pct_LOP","% LOP"],["Delta","Delta"],
    ["Threshold","Threshold"],["Row_ID","Row ID"]];
  const view=sortRows(rows);
  const cap=F.showAll?view.length:Math.min(view.length,150);

  return filterBand(false)+



  '<div class="slabel">Tabel lengkap semua indikator '+
    '<span class="hint">'+n0(view.length)+' baris'+(cap<view.length?' · menampilkan '+cap:'')+
    '</span></div>'+
  '<div class="tscroll"><table class="gt tight"><thead><tr>'+
    COLS.map(c=>th(c[1],c[0],(["Num_Base","Den_Base","Pct_Base","Num_LOP","Den_LOP","Pct_LOP",
      "Delta","Threshold"].indexOf(c[0])>=0)?"r":"")).join('')+
    '</tr></thead><tbody>'+
    view.slice(0,cap).map(r=>{
        return '<tr'+(r._berlaku==="No"?' class="oos"':'')+'>'+
      '<td class="dim">'+esc(r.Zonal)+'</td><td>'+esc(r["Area Program"])+'</td>'+
      '<td class="c dim">'+esc(r.Outcome)+'</td>'+
      '<td class="code'+(isBlank(r.Code)?' miss':'')+'">'+(isBlank(r.Code)?'—':esc(r.Code))+'</td>'+
      '<td class="ind" title="'+esc(r.Indicator)+'">'+esc(r._short)+'</td>'+
      '<td class="r dim">'+n0(r.Num_Base)+'</td><td class="r dim">'+n0(r.Den_Base)+'</td>'+
      '<td class="'+(N(r.Pct_Base)===0?"miss":"r")+'">'+pctD(r.Pct_Base)+'</td>'+
      '<td class="r dim">'+n0(r.Num_LOP)+'</td><td class="r dim">'+n0(r.Den_LOP)+'</td>'+
      '<td class="'+(N(r.Pct_LOP)===0?"belumcell":"r")+'">'+pctD(r.Pct_LOP)+'</td>'+
      '<td class="r '+(N(r.Delta)>0?"up":N(r.Delta)<0?"down":"dim")+'">'+
        (N(r.Delta)===0?"—":(N(r.Delta)>0?"+":"")+(N(r.Delta)*100).toFixed(1)+"pp")+'</td>'+
      '<td class="'+(N(r.Threshold)===0?"miss":"r")+'">'+pctT(r.Threshold)+'</td>'+
      '<td class="code dim'+(r._iddupe?' miss':'')+'">'+esc(r.Row_ID)+'</td></tr>';
    }).join('')+
    '</tbody></table></div>'+
  (cap<view.length?'<div class="morebar"><button class="ghost" data-act="showAll">Tampilkan seluruh '+
    n0(view.length)+' baris</button></div>':'')+
  '<p class="tcap">Sel merah = nilai <b>0</b>, yang menurut konvensi PEARL berarti <b>belum ada data</b>. '+
   'Baris abu-abu miring: <b>Berlaku = No</b>.</p>';
}
/* ==========================================================================
   SUMBER DATA CSV — untuk pipeline Power Automate
   --------------------------------------------------------------------------
   Halaman mencoba data/indicators.csv lebih dulu. Kalau ada dan bisa dibaca,
   itu yang dipakai; kalau tidak, jatuh ke data/indicators.js seperti biasa.
   Jadi pipeline bisa dinyalakan tanpa mengubah apa pun, dan halaman tetap
   jalan kalau flow-nya mati atau kalau file dibuka langsung dari folder.

   Pemisah kolom dideteksi sendiri: Tab, titik koma, atau koma. Excel berbahasa
   Indonesia menulis CSV dengan titik koma dan desimal koma — keduanya terbaca.

   Kolom dipetakan lewat NAMA HEADER, bukan posisi. Jadi kalau urutan kolom di
   Excel berubah, atau ada kolom baru di tengah, data tetap masuk ke tempat
   yang benar.
   ========================================================================== */
const CSV_MAP={
  "zonal":"Zonal",
  "area program":"Area Program","areaprogram":"Area Program","ap":"Area Program","adp":"Area Program",
  "outcome":"Outcome",
  "code":"Code","kode":"Code",
  "indicator deskripsi":"Indicator","indicator":"Indicator","indikator":"Indicator",
  "indicator description":"Indicator",
  "numerator baseline":"Num_Base","num baseline":"Num_Base","numerator base":"Num_Base",
  "denominator baseline":"Den_Base","den baseline":"Den_Base","denominator base":"Den_Base",
  "% baseline":"Pct_Base","pct baseline":"Pct_Base","baseline %":"Pct_Base","baseline":"Pct_Base",
  "numerator lop":"Num_LOP","num lop":"Num_LOP","numerator evaluation":"Num_LOP",
  "denominator lop":"Den_LOP","den lop":"Den_LOP","denominator evaluation":"Den_LOP",
  "% lop":"Pct_LOP","pct lop":"Pct_LOP","lop %":"Pct_LOP","% endline":"Pct_LOP",
  "% endline (lop)":"Pct_LOP","evaluation %":"Pct_LOP",
  "delta":"Delta",
  "threshold":"Threshold",
  "delta (lop-baseline)":"Delta_LOP_Base","delta lop-baseline":"Delta_LOP_Base",
  "delta (lop - baseline)":"Delta_LOP_Base",
  "row id":"Row_ID","rowid":"Row_ID","row_id":"Row_ID",
  "ap decision":"AP_Decision",
  "ap aim+ 2026 >= threshold?":"AP_vs_Threshold","ap aim+ 2026 >= threshold":"AP_vs_Threshold",
  "ap vs threshold":"AP_vs_Threshold",
  /* kolom yang metode PEARL butuhkan */
  "hh targeted":"HH_Targeted","hh target":"HH_Targeted","household targeted":"HH_Targeted",
  "hh_targeted":"HH_Targeted","kk sasaran":"HH_Targeted"
};
const CSV_NUM={Num_Base:1,Den_Base:1,Pct_Base:1,Num_LOP:1,Den_LOP:1,Pct_LOP:1,
  Delta:1,Threshold:1,Delta_LOP_Base:1,HH_Targeted:1,AP_Decision:1};
const CSV_PROP={Pct_Base:1,Pct_LOP:1,Delta:1,Threshold:1,Delta_LOP_Base:1};

function sniffDelim(line){
  let best="\t", n=-1;
  ["\t",";",","].forEach(d=>{
    let c=0,q=false;
    for(let i=0;i<line.length;i++){
      const ch=line[i];
      if(ch==='"'){ q=!q; continue; }
      if(!q&&ch===d) c++;
    }
    if(c>n){ n=c; best=d; }
  });
  return {delim:best, count:n};
}
function parseDelim(text,delim){
  const rows=[]; let row=[], cell="", q=false, i=0;
  const s=text.replace(/\r\n/g,"\n").replace(/\r/g,"\n").replace(/^\uFEFF/,"");
  while(i<s.length){
    const ch=s[i];
    if(q){
      if(ch==='"'){ if(s[i+1]==='"'){cell+='"';i+=2;continue;} q=false;i++;continue; }
      cell+=ch; i++; continue;
    }
    if(ch==='"'&&cell===""){ q=true; i++; continue; }
    if(ch===delim){ row.push(cell); cell=""; i++; continue; }
    if(ch==="\n"){ row.push(cell); rows.push(row); row=[]; cell=""; i++; continue; }
    cell+=ch; i++;
  }
  if(cell!==""||row.length){ row.push(cell); rows.push(row); }
  return rows.filter(r=>r.join("").trim()!=="");
}
/* pemisah desimal ditentukan sekali untuk seluruh file, dari kolom proporsi */
function sniffDec(cells,idx,propCols){
  let dot=0, com=0;
  for(let r=1;r<cells.length;r++){
    propCols.forEach(j=>{
      const v=String(cells[r][j]==null?"":cells[r][j]).replace(/[\s\u00a0%]/g,"");
      const m=v.match(/([.,])(\d{2,})$/);
      if(m) m[1]==="."?dot++:com++;
    });
  }
  return com>dot ? "," : ".";
}
function csvNum(raw,dec){
  let s=String(raw==null?"":raw).replace(/[\s\u00a0]/g,"");
  if(s===""||s==="-"||s==="—") return null;
  if(s.charAt(0)==="#") return null;      /* #VALUE! #REF! #N/A #DIV/0! dari Excel */
  const pct=/%$/.test(s); s=s.replace(/%/g,"");
  const neg=/^\(.*\)$/.test(s)||/^-/.test(s);
  s=s.replace(/^[+-]/,"").replace(/^\(|\)$/g,"");
  if(!/^[\d.,]+$/.test(s)) return null;
  const grp = dec==="," ? "." : ",";
  const d=s.lastIndexOf(dec);
  const out = d<0 ? s.split(grp).join("")
    : s.slice(0,d).split(grp).join("").split(dec).join("")+"."+
      s.slice(d+1).split(grp).join("").split(dec).join("");
  let n=parseFloat(out);
  if(!isFinite(n)) return null;
  if(neg) n=-n;
  return pct ? n/100 : n;
}
function parseCsvData(text){
  const first=text.replace(/^\uFEFF/,"").split(/\r?\n/)[0]||"";
  const sn=sniffDelim(first);
  if(sn.count<3) return {err:"Baris pertama hanya punya "+(sn.count+1)+" kolom. "+
    "Pastikan file CSV punya baris header dan minimal empat kolom."};
  const cells=parseDelim(text,sn.delim);
  if(cells.length<2) return {err:"Tidak ada baris data di bawah header."};
  const head=cells[0].map(h=>String(h||"").replace(/[\s\u00a0]+/g," ").trim().toLowerCase());
  const idx={}; const unknown=[];
  head.forEach((h,j)=>{
    const k=CSV_MAP[h];
    if(k){ if(idx[k]===undefined) idx[k]=j; }
    else if(h) unknown.push(cells[0][j]);
  });
  const need=["Area Program","Indicator"];
  const miss=need.filter(k=>idx[k]===undefined);
  if(miss.length) return {err:"Kolom wajib tidak ditemukan di header: "+miss.join(", ")+
    ". Header yang terbaca: "+head.filter(Boolean).join(" | ")};
  const propCols=Object.keys(CSV_PROP).filter(k=>idx[k]!==undefined).map(k=>idx[k]);
  const dec=sniffDec(cells,idx,propCols);
  const rows=[]; let errCells=0;
  for(let r=1;r<cells.length;r++){
    const c=cells[r], o={};
    Object.keys(idx).forEach(k=>{
      const raw=c[idx[k]];
      if(CSV_NUM[k]) o[k]=csvNum(raw,dec);
      else {
        let s=String(raw==null?"":raw).replace(/[\x00-\x1F\x7F]/g,"").trim();
        if(s.charAt(0)==="#") { s=""; errCells++; }
        o[k]=s||null;
      }
    });
    if(!o["Area Program"]||!o.Indicator) continue;
    if(!o.Row_ID||o.Row_ID==="0") o.Row_ID="CSV-"+String(rows.length+1).padStart(4,"0");
    rows.push(o);
  }
  if(!rows.length) return {err:"Header terbaca tapi tidak ada baris data yang lengkap."};
  return {rows:rows, delim:sn.delim===";"?"titik koma":sn.delim===","?"koma":"Tab",
    dec:dec===","?"koma":"titik", unknown:unknown, errCells:errCells,
    hasHH:idx.HH_Targeted!==undefined};
}
async function tryCsv(){
  for(const p of ["data/indicators.csv","indicators.csv"]){
    try{
      const res=await fetch(p+"?t="+Date.now(),{cache:"no-store"});
      if(!res.ok) continue;
      const txt=await res.text();
      if(!txt.trim()) continue;
      const out=parseCsvData(txt);
      out.path=p;
      return out;
    }catch(e){ /* file:// atau file tidak ada — lanjut ke .js */ }
  }
  return null;
}
/* ==========================================================================
   HALAMAN 1 — NATIONAL SUMMARY DASHBOARD
   --------------------------------------------------------------------------
   Metode: Weighted National (%) = Σ Numerator ÷ Σ Denominator.
   Numerator dijumlahkan dulu, denominator dijumlahkan dulu, baru dibagi.
   TIDAK ADA rata-rata persentase antar AP atau antar zona.

   Konsekuensi yang perlu diingat: sebuah baris hanya bisa masuk hitungan
   kalau denominatornya ada. Baris yang punya % tapi tidak punya denominator
   tidak bisa diberi bobot, jadi ia dikeluarkan — dan jumlahnya dilaporkan.

   Tidak ada total lintas indikator. Menjumlahkan numerator dua indikator yang
   populasinya berbeda menghasilkan angka tanpa makna (di dataset ini hasilnya
   116,9%). Angka nasional hanya sah per indikator.
   ========================================================================== */
const PERIODS=["Baseline","Evaluation","Both"];

/* Metode PEARL, diverifikasi terhadap sheet IND160 (Sigma HH 90.716 -> 20,6%):
     ADP Weight            = HH Targeted ADP / Sigma HH Targeted ADP yang punya rate
     Weighted Contribution = ADP Weight x ADP Rate
     Weighted National     = Sigma Weighted Contribution
   Kalau kolom HH Targeted belum ada di data, angka nasional TIDAK dihitung —
   lebih baik kosong daripada salah bobot. hasHH menandainya. */
const hasHH = () => S.rows.some(r=>N(r.HH_Targeted)>0);

/* Terverifikasi terhadap Master (9): di sheet IND160 Dashboard, kolom HH Targeted
   berisi angka yang SAMA dengan Denominator (LOP) pada seluruh 17 ADP. Karena itu
       Sigma(HH x rate) / Sigma HH  =  Sigma(den x num/den) / Sigma den  =  Sigma num / Sigma den
   Kedua rumus itu identik secara aljabar, dan Total Nasional di sheet Anda
   (18,34% evaluation, 24,72% baseline untuk OIOS 160) sama dengan hasil di bawah. */
function weightedOf(rows){
  let nB=0,dB=0,nE=0,dE=0,skipB=0,skipE=0;
  const apB={}, apE={};
  rows.forEach(r=>{
    const db=N(r.Den_Base), dl=N(r.Den_LOP);
    if(db>0){ nB+=N(r.Num_Base); dB+=db; apB[r["Area Program"]]=1; }
    else if(N(r.Pct_Base)>0) skipB++;
    if(dl>0){ nE+=N(r.Num_LOP); dE+=dl; apE[r["Area Program"]]=1; }
    else if(N(r.Pct_LOP)>0) skipE++;
  });
  const pB=dB>0?nB/dB:null, pE=dE>0?nE/dE:null;
  return {nB:nB,dB:dB,nE:nE,dE:dE,pB:pB,pE:pE,wB:dB,wE:dE,useHH:true,
    apB:Object.keys(apB).length, apE:Object.keys(apE).length,
    skipB:skipB, skipE:skipE,
    delta:(pB!==null&&pE!==null)?(pE-pB):null};
}

function natRows(){
  let rows=S.rows;
  if(F.zonal.length)   rows=rows.filter(r=>F.zonal.indexOf(r.Zonal)>=0);
  if(F.ap.length)      rows=rows.filter(r=>F.ap.indexOf(r["Area Program"])>=0);
  if(F.outcome.length) rows=rows.filter(r=>F.outcome.indexOf(r.Outcome)>=0);
  if(F.ind.length)     rows=rows.filter(r=>F.ind.indexOf(r.Indicator)>=0);
  return rows;
}
const showB = () => F.period!=="Evaluation";
const showE = () => F.period!=="Baseline";

function renderNasional(){
  const pc=v=>v===null?"—":(v*100).toFixed(1)+"%";
  const pp=v=>v===null?"—":(v>0?"+":"")+(v*100).toFixed(1);
  const rows=natRows();
  const per=S.cat.map(c=>{
    const rs=rows.filter(r=>r.Indicator===c.ind);
    return {ind:c.ind, short:c.short, code:c.code, oc:c.oc,
      dir:arahOf(c.ind), w:weightedOf(rs), n:rs.length};
  }).filter(x=>x.n>0);

  const withB=per.filter(x=>x.w.pB!==null), withE=per.filter(x=>x.w.pE!==null);
  const withBoth=per.filter(x=>x.w.delta!==null);
  const BAND=0.001;                                  /* 0,1pp dianggap tetap */
  const naik=withBoth.filter(x=>x.w.delta>BAND).length;
  const turun=withBoth.filter(x=>x.w.delta<-BAND).length;
  const tetap=withBoth.length-naik-turun;
  const skipB=per.reduce((a,x)=>a+x.w.skipB,0), skipE=per.reduce((a,x)=>a+x.w.skipE,0);
  /* weighted di luar 0–100% berarti Σnumerator melebihi Σdenominator — mustahil
     untuk sebuah proporsi, jadi ini masalah input, bukan hasil. */
  const oorInd=per.filter(x=>(x.w.pB!==null&&x.w.pB>1)||(x.w.pE!==null&&x.w.pE>1));
  const oorRows=rows.filter(r=>(N(r.Den_Base)>0&&N(r.Num_Base)>N(r.Den_Base))||
                               (N(r.Den_LOP)>0&&N(r.Num_LOP)>N(r.Den_LOP)));
  const oorAP={}; oorRows.forEach(r=>{oorAP[r["Area Program"]]=(oorAP[r["Area Program"]]||0)+1;});
  const oorTop=Object.keys(oorAP).sort((a,b)=>oorAP[b]-oorAP[a]);


  /* ---------- grafik ---------- */
  const cmp=per.filter(x=>x.w.pB!==null||x.w.pE!==null)
    .map(x=>({label:x.short+(x.dir==="Turun"?" ↓":""),
      a:showB()?x.w.pB:null, b:showE()?x.w.pE:null}));
  const del=withBoth.slice().sort((a,b)=>b.w.delta-a.w.delta)
    .map(x=>({label:x.short+(x.dir==="Turun"?" ↓":""), v:x.w.delta}));

  return '<div class="notice"><b>Weighted National (%) = Σ Numerator ÷ Σ Denominator</b>, dihitung '+
    'per indikator. Setara dengan Σ(ADP Weight × ADP Rate) di sheet Anda, karena kolom '+
    '<b>HH Targeted</b> di Master (9) berisi angka yang sama dengan Denominator (LOP) pada '+
    'seluruh 17 ADP. OIOS 160 di sheet IND160 Dashboard: 18,34% evaluation dan 24,72% baseline — '+
    'sama dengan angka di halaman ini. <b>Tidak ada total lintas indikator</b>, karena populasi '+
    'tiap indikator berbeda.</div>'+
  natFilterBand()+

  '<div class="slabel">Cakupan dataset'+natActiveLine()+'</div>'+
  '<div class="cards">'+
    card("Total<br>AP",uniq(rows.map(r=>r["Area Program"])).length,"","teal")+
    card("Total<br>Zonal",uniq(rows.map(r=>r.Zonal)).length,"","neutral")+
    card("Total<br>Outcome",uniq(rows.map(r=>r.Outcome).filter(Boolean)).length,"","neutral")+
    card("Total<br>Indicators",per.length,"dari "+S.cat.length+" di katalog","accent")+
  '</div>'+

  '<div class="slabel">Ringkasan arah perubahan '+
    '<span class="hint">'+withBoth.length+' indikator punya baseline dan evaluation berbobot</span></div>'+
  '<div class="cards" style="grid-template-columns:repeat(auto-fit,minmax(150px,1fr))">'+
    card("Meningkat",naik,"delta &gt; +0,1pp","ready")+
    card("Tetap",tetap,"|delta| ≤ 0,1pp","belum")+
    card("Menurun",turun,"delta &lt; −0,1pp","critical")+
    card("Belum bisa<br>dibandingkan",per.length-withBoth.length,
      "baseline atau evaluation belum berbobot","review")+
  '</div>'+
  ((skipB+skipE)?'<p class="chartnote">'+n0(skipB+skipE)+' baris punya persentase tetapi tidak punya '+
    'denominator, jadi tidak bisa diberi bobot dan dikeluarkan dari hitungan nasional '+
    '('+skipB+' baseline, '+skipE+' evaluation).</p>':'')+

  (oorInd.length?
  '<div class="warnbox" style="margin-top:18px"><b>'+oorInd.length+' indikator menghasilkan '+
   'weighted di luar 0–100%</b>, artinya Σ numerator melebihi Σ denominator. Itu mustahil untuk '+
   'sebuah proporsi, jadi angkanya <b>tidak boleh dikutip</b> sebelum sumbernya diperbaiki. '+
   'Penyebabnya ada di <b>'+n0(oorRows.length)+' baris</b>'+
   (oorTop.length?', terbanyak di <b>'+esc(oorTop[0])+'</b> ('+oorAP[oorTop[0]]+' baris)'+
     (oorTop[1]?' dan '+esc(oorTop[1])+' ('+oorAP[oorTop[1]]+')':''):'')+'. '+
   'Metode weighted-lah yang memunculkannya — rata-rata persentase akan menyembunyikannya.</div>'+
  '<div class="tscroll" style="margin-top:10px"><table class="gt tight"><thead><tr>'+
    '<th>Indicator</th><th class="r">Σ Num Base</th><th class="r">Σ Den Base</th><th class="r">Baseline</th>'+
    '<th class="r">Σ Num Eval</th><th class="r">Σ Den Eval</th><th class="r">Evaluation</th>'+
    '</tr></thead><tbody>'+oorInd.map(x=>'<tr>'+
      '<td class="ind" title="'+esc(x.ind)+'">'+esc(x.short)+'</td>'+
      '<td class="r dim">'+n0(x.w.dB?x.w.nB:null)+'</td><td class="r dim">'+n0(x.w.dB||null)+'</td>'+
      '<td class="'+(x.w.pB!==null&&x.w.pB>1?"miss":"r")+'">'+pc(x.w.pB)+'</td>'+
      '<td class="r dim">'+n0(x.w.dE?x.w.nE:null)+'</td><td class="r dim">'+n0(x.w.dE||null)+'</td>'+
      '<td class="'+(x.w.pE!==null&&x.w.pE>1?"miss":"r")+'">'+pc(x.w.pE)+'</td></tr>').join('')+
    '</tbody></table></div>':'')+

  '<div class="slabel">Baseline vs Evaluation per indikator '+
    '<span class="hint">Weighted National (%)</span></div>'+
  '<div class="chartbox">'+
    (cmp.length?chartPair(cmp,["Baseline",CLR.baseline],["Evaluation",CLR.endline],{labW:250})
      :'<p class="dim">Tidak ada indikator dengan denominator pada filter ini.</p>')+
    legend([["Baseline",CLR.baseline]].concat(showE()?[["Evaluation",CLR.endline]]:[]))+
    '<p class="chartnote">↓ menandai indikator berarah Turun — pada indikator itu penurunan '+
    'justru perbaikan.</p></div>'+

  (F.period==="Both"?
  '<div class="slabel">Delta per indikator <span class="hint">poin persentase · '+
    'peningkatan terbesar di atas</span></div>'+
  '<div class="chartbox">'+
    (del.length?chartDelta(del,{labW:250}):'<p class="dim">Belum ada indikator yang punya keduanya.</p>')+
    '<div class="legend">'+
      '<div><i style="background:'+CLR.baik+';border-color:'+CLR.baik+'"></i>Meningkat</div>'+
      '<div><i style="background:'+CLR.belum+';border-color:'+CLR.ref+'"></i>Tetap</div>'+
      '<div><i style="background:'+CLR.hati+';border-color:'+CLR.hati+'"></i>Menurun</div>'+
    '</div>'+
    '<p class="chartnote">Warna mengikuti arah angka, bukan arah perbaikan: pada indikator '+
    'bertanda ↓ warna merah berarti angkanya turun, dan itu <b>hasil yang baik</b>.</p></div>':'')+

  '<div class="slabel">National Summary Table <span class="hint">'+per.length+' indikator</span></div>'+
  '<div class="tscroll"><table class="gt tight"><thead><tr>'+
    '<th>Outcome</th><th>Indicator</th>'+
    (showB()?'<th class="r">Base Num</th><th class="r">Base Den</th><th class="r">Baseline (%)</th><th class="r">AP</th>':'')+
    (showE()?'<th class="r">Eval Num</th><th class="r">Eval Den</th><th class="r">Evaluation (%)</th><th class="r">AP</th>':'')+
    (F.period==="Both"?'<th class="r">Delta (pp)</th>':'')+
    '</tr></thead><tbody>'+
    per.map(x=>'<tr><td class="c dim">'+esc(x.oc||"—")+'</td>'+
      '<td class="ind" title="'+esc(x.ind)+'">'+esc(x.short)+
        (x.dir==="Turun"?' <span class="turun">↓</span>':'')+'</td>'+
      (showB()?'<td class="r dim">'+n0(x.w.dB?x.w.nB:null)+'</td><td class="r dim">'+n0(x.w.dB||null)+'</td>'+
        '<td class="'+(x.w.pB===null||x.w.pB>1?"miss":"r")+'"><b>'+pc(x.w.pB)+'</b>'+
        (x.w.pB!==null&&x.w.pB>1?' ▲':'')+'</td>'+
        '<td class="r dim">'+(x.w.apB||"—")+'</td>':'')+
      (showE()?'<td class="r dim">'+n0(x.w.dE?x.w.nE:null)+'</td><td class="r dim">'+n0(x.w.dE||null)+'</td>'+
        '<td class="'+(x.w.pE===null||x.w.pE>1?"miss":"r")+'"><b>'+pc(x.w.pE)+'</b>'+
        (x.w.pE!==null&&x.w.pE>1?' ▲':'')+'</td>'+
        '<td class="r dim">'+(x.w.apE||"—")+'</td>':'')+
      (F.period==="Both"?'<td class="r '+(x.w.delta===null?"dim":x.w.delta>0.001?"up":x.w.delta<-0.001?"down":"dim")+'">'+
        pp(x.w.delta)+'</td>':'')+
      '</tr>').join('')+
    '</tbody></table></div>'+
  '<p class="tcap">Kolom <b>AP</b> adalah jumlah Area Program yang benar-benar menyumbang angka itu. '+
   'Angka nasional yang hanya berasal dari satu atau dua AP tetap ditampilkan, tapi jangan dibaca '+
   'sebagai gambaran nasional.</p>';
}

/* ---------- band filter khusus halaman nasional ---------- */
function natFilterBand(){
  const all=S.rows;
  const apPool=F.zonal.length?AP_LIST.filter(a=>F.zonal.indexOf(a.zonal)>=0):AP_LIST;
  const indPool=S.cat.filter(c=>!F.outcome.length||F.outcome.indexOf(c.oc)>=0);
  const nAktif=F.zonal.length+F.ap.length+F.outcome.length+F.ind.length+(F.period!=="Both"?1:0);

  let h='<div class="fb"><div class="fb-head">'+
    '<span class="fb-title">Filter</span>'+
    '<span class="fb-sum">'+(nAktif?esc(natActiveText()):'seluruh dataset')+'</span>'+
    '<div class="fb-sp"></div>'+
    (nAktif?'<button class="fb-btn" data-act="natReset">Bersihkan semua</button>':'')+
    (F.fbOpen?'<button class="fb-btn" data-act="fbClose">Sembunyikan ▴</button>'
             :'<button class="fb-btn" data-act="fbOpen">Ubah filter ▾</button>')+
    '</div>';
  if(!F.fbOpen) return h+'</div>';
  h+='<div class="fb-row"><div class="fb-lab">Periode</div><div class="fb-chips">'+
      PERIODS.map(p=>'<button class="chipf big'+(F.period===p?" sel":"")+
        '" data-slice="period" data-val="'+p+'">'+
        (F.period===p?"◉ ":"○ ")+p+(p==="Both"?' <span class="n">default</span>':'')+
        '</button>').join('')+
     '</div></div>'+
    fbRow("Zonal","zonal",ZONALS.map(z=>({v:z,n:all.filter(r=>r.Zonal===z).length})),true)+
    fbRow("Area Program","ap",apPool.map(a=>({v:a.ap,n:all.filter(r=>r["Area Program"]===a.ap).length})),true,
      F.zonal.length?'mengikuti Zonal yang dipilih':'')+
    fbRow("Outcome","outcome",OUTCOMES.map(o=>({v:o,n:all.filter(r=>r.Outcome===o).length})),true)+
    '<div class="fb-row"><div class="fb-lab">Indicator'+
      (F.ind.length?'<button class="fb-x" data-clear="ind" title="Bersihkan">×</button>':'')+
      '</div><div class="fb-chips">'+
      indPool.map(c=>'<button class="chipf'+(F.ind.indexOf(c.ind)>=0?" sel":"")+
        '" data-slice="ind" data-val="'+esc(c.ind)+'" title="'+esc(c.ind)+'">'+esc(c.short)+'</button>').join('')+
      (F.outcome.length?'<span class="fb-hint">mengikuti Outcome yang dipilih</span>':'')+
    '</div></div>'+
  '</div>';
  return h;
}
function natActiveText(){
  const b=[];
  if(F.period!=="Both") b.push(F.period);
  if(F.zonal.length) b.push(F.zonal.join(", "));
  if(F.ap.length) b.push(F.ap.length===1?F.ap[0]:F.ap.length+" AP");
  if(F.outcome.length) b.push(F.outcome.join(", "));
  if(F.ind.length) b.push(F.ind.length+" indikator");
  return b.join(" · ");
}
function natActiveLine(){
  const s=natActiveText();
  return s?' <span class="hint">'+esc(s)+'</span>':'';
}

/* ---------- bar delta: hijau naik, abu tetap, merah turun ---------- */
function chartDelta(items,opt){
  opt=opt||{};
  const labW=opt.labW||250, W=700, BH=13, PITCH=BH+9;
  const H=Math.max(34,items.length*PITCH+24), PW=W-labW-76;
  const m=Math.max(0.02,...items.map(i=>Math.abs(i.v||0)));
  const sc=Math.ceil(m*20)/20, zero=labW+PW/2, half=PW/2;
  const x=v=>zero+half*Math.max(-1,Math.min(1,(v||0)/sc));
  let s='<svg viewBox="0 0 '+W+' '+H+'" role="img">';
  s+='<line x1="'+zero+'" y1="14" x2="'+zero+'" y2="'+(H-10)+'" stroke="#3F3D4C" stroke-width="1.2"/>'+
     '<text x="'+zero+'" y="9" text-anchor="middle" font-size="8" fill="#3F3D4C">0</text>'+
     '<text x="'+labW+'" y="9" font-size="8" fill="#A9A6B0">−'+(sc*100).toFixed(0)+'pp</text>'+
     '<text x="'+(labW+PW)+'" y="9" text-anchor="end" font-size="8" fill="#A9A6B0">+'+(sc*100).toFixed(0)+'pp</text>';
  items.forEach((it,i)=>{
    const y=i*PITCH+16, xv=x(it.v);
    const col = it.v>0.001?CLR.baik : it.v<-0.001?CLR.hati : CLR.belum;
    s+='<text x="'+(labW-8)+'" y="'+(y+BH-3)+'" text-anchor="end" font-size="9.5" font-weight="600" fill="#111222">'+
       esc(it.label)+'</text>'+
       '<rect x="'+Math.min(zero,xv).toFixed(1)+'" y="'+y+'" width="'+Math.abs(xv-zero).toFixed(1)+
       '" height="'+BH+'" fill="'+col+'"'+(col===CLR.belum?' stroke="'+CLR.ref+'" stroke-width=".6"':'')+'/>'+
       '<text x="'+(labW+PW+8)+'" y="'+(y+BH-3)+'" font-size="9.5" font-weight="700" fill="'+
       (col===CLR.belum?CLR.ref:col)+'">'+(it.v>0?"+":"")+(it.v*100).toFixed(1)+'pp</text>';
  });
  return s+'</svg>';
}
/* ==========================================================================
   IMPOR DARI EXCEL
   Tempel blok sheet Indicators, halaman memeriksanya, lalu menulis ulang
   data/indicators.js untuk di-commit. Tidak perlu mengedit file dengan tangan.

   Angka ditafsirkan per kolom, bukan seragam: kolom hitungan (Numerator,
   Denominator) memperlakukan pemisah 3 digit sebagai ribuan, kolom proporsi
   memperlakukan pemisah terakhir sebagai desimal. Ini yang membuat blok dari
   Excel berbahasa Indonesia ("0,2349") maupun Inggris ("0.2349") sama-sama
   terbaca benar.
   ========================================================================== */
const IMP_COLS=["Zonal","Area Program","Outcome","Code","Indicator","Num_Base","Den_Base",
  "Pct_Base","Num_LOP","Den_LOP","Pct_LOP","Delta","Threshold","Delta_LOP_Base","Row_ID"];
const IMP_KIND={Num_Base:"count",Den_Base:"count",Num_LOP:"count",Den_LOP:"count",
  Pct_Base:"prop",Pct_LOP:"prop",Delta:"prop",AP_Decision:"prop",Threshold:"prop",
  Delta_LOP_Base:"prop"};

/* Pemisah desimal dideteksi sekali untuk seluruh blok, dari kolom proporsi —
   di sana pemisah yang diikuti 2 digit atau lebih pasti desimal, bukan ribuan.
   Setelah diketahui, setiap angka dibaca dengan aturan yang sama, jadi
   "115.099" tidak lagi ambigu: ia bergantung pada locale blok itu, bukan tebakan
   per sel. Kalau blok tidak memberi petunjuk, dipakai titik. */
let DECSEP=".";
function detectDecSep(lines,start){
  const PROP=[7,10,11,12,13];
  let dot=0, com=0;
  for(let i=start;i<lines.length;i++){
    const c=lines[i].split("\t");
    PROP.forEach(j=>{
      const s=String(c[j]==null?"":c[j]).replace(/[\s\u00a0%]/g,"");
      const m=s.match(/([.,])(\d{2,})$/);
      if(m){ m[1]==="." ? dot++ : com++; }
    });
  }
  DECSEP = com>dot ? "," : ".";
  return {sep:DECSEP, dot:dot, com:com};
}
function parseNum(raw){
  let s=String(raw==null?"":raw).replace(/[\s\u00a0]/g,"");
  if(s===""||s==="-"||s==="—") return null;
  const pct=/%$/.test(s); s=s.replace(/%/g,"");
  const neg=/^\(.*\)$/.test(s)||/^-/.test(s);
  s=s.replace(/^[+-]/,"").replace(/^\(|\)$/g,"");
  if(!/^[\d.,]+$/.test(s)) return null;
  const grp = DECSEP==="," ? "." : ",";
  const dec = s.lastIndexOf(DECSEP);
  let out;
  if(dec<0) out=s.split(grp).join("");
  else out = s.slice(0,dec).split(grp).join("").split(DECSEP).join("")+"."+
             s.slice(dec+1).split(grp).join("").split(DECSEP).join("");
  let n=parseFloat(out);
  if(!isFinite(n)) return null;
  if(neg) n=-n;
  if(pct) n=n/100;
  return n;
}
function clean(v){return String(v==null?"":v).replace(/[\x00-\x1F\x7F]/g,"").trim();}

/* Excel membungkus sel yang berisi Tab, baris baru, atau tanda kutip dengan
   tanda kutip ganda. Kolom Code di data ini memang berisi baris baru
   (mis. "C5G.027623\n"), jadi memecah teks per baris saja akan menggeser kolom.
   Parser di bawah membaca seluruh blok sebagai satu aliran. */
function parseTSV(text){
  const rows=[]; let row=[], cell="", q=false, i=0;
  const s=text.replace(/\r\n/g,"\n").replace(/\r/g,"\n");
  while(i<s.length){
    const ch=s[i];
    if(q){
      if(ch==='"'){ if(s[i+1]==='"'){cell+='"';i+=2;continue;} q=false; i++; continue; }
      cell+=ch; i++; continue;
    }
    if(ch==='"'&&cell===""){ q=true; i++; continue; }
    if(ch==="\t"){ row.push(cell); cell=""; i++; continue; }
    if(ch==="\n"){ row.push(cell); rows.push(row); row=[]; cell=""; i++; continue; }
    cell+=ch; i++;
  }
  if(cell!==""||row.length){ row.push(cell); rows.push(row); }
  return rows.filter(r=>r.join("").trim()!=="");
}

let IMP=null;   /* hasil parse yang menunggu konfirmasi */

function impParse(){
  const raw=document.getElementById("impBox").value;
  const cells=parseTSV(raw);
  const rep=document.getElementById("impReport");
  if(!cells.length){ rep.innerHTML='<p class="vmsg">▲ Belum ada yang ditempel.</p>'; IMP=null; return; }

  let start=0;
  if(/^zonal$/i.test(clean(cells[0][0]))) start=1;   /* baris header ikut tertempel */
  const lines=cells.map(c=>c.join("\t"));
  const loc=detectDecSep(lines,start);

  const rows=[], bad=[];
  for(let i=start;i<cells.length;i++){
    const c=cells[i];
    if(c.length<5){ bad.push({line:i+1, why:"hanya "+c.length+" kolom"}); continue; }
    const r={};
    IMP_COLS.forEach((k,j)=>{
      const v=c[j];
      /* Nilai dimuat apa adanya. Proporsi di luar 0–1 TIDAK dikoreksi otomatis —
         ia dilaporkan, karena mengubahnya diam-diam berarti mengarang data. */
      if(IMP_KIND[k]) r[k]=parseNum(v);
      else r[k]=clean(v)||null;
    });
    if(!r.Row_ID) r.Row_ID="IND-"+String(rows.length+1).padStart(3,"0");
    if(!r["Area Program"]||!r.Indicator){ bad.push({line:i+1, why:"Area Program atau Indikator kosong"}); continue; }
    rows.push(r);
  }
  if(!rows.length){ rep.innerHTML='<p class="vmsg">▲ Tidak ada baris yang bisa dibaca. '+
    'Pastikan blok disalin dari sheet Indicators, 17 kolom, dipisah Tab.</p>'; IMP=null; return; }

  /* ---- pemeriksaan ---- */
  const seen={}, idDupe=[];
  rows.forEach(r=>{ if(seen[r.Row_ID]) idDupe.push(r.Row_ID); seen[r.Row_ID]=1; });
  const unkInd=uniq(rows.map(r=>r.Indicator).filter(x=>IND_LIST.indexOf(x)<0));
  const unkAP=uniq(rows.map(r=>r["Area Program"]).filter(x=>!AP_LIST.some(a=>a.ap===x)));
  const unkOC=uniq(rows.map(r=>r.Outcome).filter(x=>x&&OUTCOMES.indexOf(x)<0));
  const combo={}; rows.forEach(r=>{const k=r["Area Program"]+"|"+r.Indicator; combo[k]=(combo[k]||0)+1;});
  const comboDupe=Object.keys(combo).filter(k=>combo[k]>1).length;
  const oldIds=new Set(S.rows.map(r=>r.Row_ID));
  const added=rows.filter(r=>!oldIds.has(r.Row_ID)).length;
  const gone=S.rows.filter(r=>!seen[r.Row_ID]).length;
  const withBase=rows.filter(r=>N(r.Pct_Base)>0).length;
  const withLop=rows.filter(r=>N(r.Pct_LOP)>0).length;
  const oor=rows.filter(r=>Math.abs(N(r.Pct_Base))>1||Math.abs(N(r.Pct_LOP))>1);
  const shifted=rows.filter(r=>r.Row_ID&&!/^IND-\d+$/i.test(r.Row_ID)).length;

  const blok=(ok,txt)=>'<span class="ichip '+(ok?"ok":"no")+'">'+(ok?"● ":"▲ ")+txt+'</span>';
  rep.innerHTML=
    '<div class="imprep">'+
      '<div class="improw"><b>'+n0(rows.length)+'</b> baris terbaca'+
        (start?' <span class="dim">(baris header diabaikan)</span>':'')+'</div>'+
      '<div class="improw">Dibanding data sekarang ('+n0(S.rows.length)+' baris): '+
        '<b>'+added+'</b> baris baru, <b>'+gone+'</b> baris hilang</div>'+
      '<div class="improw">Ada baseline <b>'+withBase+'</b> · ada endline <b>'+withLop+'</b></div>'+
      '<div class="improw">Pemisah desimal terdeteksi: <b>'+(loc.sep===","?"koma":"titik")+'</b>'+
        ' <span class="dim">('+loc.com+' nilai berkoma, '+loc.dot+' bertitik)</span></div>'+
      '<div class="improw">'+
        blok(idDupe.length===0, idDupe.length?"Row ID ganda: "+idDupe.slice(0,4).join(", "):"Row ID unik")+
        blok(unkInd.length===0, unkInd.length?unkInd.length+" indikator tidak dikenal":"Nama indikator dikenal")+
        blok(unkAP.length===0, unkAP.length?"AP tidak dikenal: "+unkAP.join(", "):"Nama AP dikenal")+
        (unkOC.length?blok(false,"Outcome tidak dikenal: "+unkOC.join(", ")):"")+
        (comboDupe?blok(false,comboDupe+" kombinasi AP × indikator ganda"):blok(true,"Tidak ada duplikat"))+
        (bad.length?blok(false,bad.length+" baris dilewati"):"")+
        (oor.length?blok(false,oor.length+" proporsi di luar 0–100%"):blok(true,"Semua proporsi 0–100%"))+
        (shifted?blok(false,shifted+" Row ID tidak berpola IND-nnn"):"")+
      '</div>'+
      (oor.length?'<div class="improw" style="color:var(--red)">Di luar 0–100%: '+
        oor.slice(0,4).map(r=>esc(r.Row_ID)+" ("+(N(r.Pct_Base)>1?"baseline "+(N(r.Pct_Base)*100).toFixed(0)+"%":
        "endline "+(N(r.Pct_LOP)*100).toFixed(0)+"%")+")").join(" · ")+(oor.length>4?" …":"")+
        '. Dimuat apa adanya, tidak dikoreksi otomatis.</div>':'')+
      (unkInd.length?'<div class="improw dim">Indikator tidak dikenal tidak akan dapat Arah, Target Delta, '+
        'maupun Berlaku dari konfigurasi: '+esc(unkInd.slice(0,3).map(shortOf).join(" · "))+
        (unkInd.length>3?" …":"")+'. Tambahkan dulu di pemetaan.js.</div>':'')+
      (bad.length?'<div class="improw dim">Dilewati: '+bad.slice(0,3).map(b=>"baris "+b.line+" ("+b.why+")").join(" · ")+'</div>':'')+
    '</div>';
  IMP={rows:rows, hard:(idDupe.length>0||unkAP.length>0||bad.length>0)};
  const btn=document.getElementById("impGo");
  btn.disabled=false;
  btn.textContent=IMP.hard?"Muat walaupun ada peringatan":"Muat "+n0(rows.length)+" baris";
}
function impApply(){
  if(!IMP){ return; }
  S.rows=IMP.rows.map(r=>Object.assign({},r));
  recompute(); saveLocal();
  document.getElementById("scrimImport").classList.remove("on");
  document.getElementById("impBox").value=""; document.getElementById("impReport").innerHTML="";
  document.getElementById("impGo").disabled=true;
  const c=checks();
  repaint(); 
  toast('Data dimuat: <b>'+n0(S.rows.length)+' baris</b>. '+esc(c.sinkron)+'. '+
    'Jangan lupa <b>Simpan file data</b> lalu commit <span style="color:#8FE0AE">data/indicators.js</span>.');
  IMP=null;
}
/* ==========================================================================
   HALAMAN 2 — ANALISIS AP
   ========================================================================== */
function renderAnalisis(){
  const rows=filtered(), n=rows.length;
  const cnt=(k,v)=>rows.filter(r=>r[k]===v).length;
  const aps=uniq(rows.map(r=>r["Area Program"]));

  /* G5 · status indikator per AP */
  const g5=aps.map(ap=>{
    const rs=rows.filter(r=>r["Area Program"]===ap);
    return {label:ap, vals:ST_IND.map(s=>rs.filter(r=>r._status===s.v).length),
      score:rs.filter(r=>r._status==="Sesuai Target/Threshold"||r._status==="Ditinjau").length/(rs.length||1)};
  }).sort((a,b)=>b.score-a.score);

  /* endline vs threshold per AP — urutan sama dengan grafik status supaya bisa dibaca berpasangan */
  const g6=g5.map(x=>{
    const rs=rows.filter(r=>r["Area Program"]===x.label);
    return {label:x.label, vals:ST_THR.map(s=>rs.filter(r=>r._thr_status===s.v).length)};
  });
  /* ringkasan yang sama, dilihat per Outcome */
  const ocRows=OUTCOMES.filter(o=>rows.some(r=>r.Outcome===o));
  const gOcStat=ocRows.map(o=>{
    const rs=rows.filter(r=>r.Outcome===o);
    return {label:o, vals:ST_IND.map(s=>rs.filter(r=>r._status===s.v).length)};
  });
  const gOcThr=ocRows.map(o=>{
    const rs=rows.filter(r=>r.Outcome===o);
    return {label:o, vals:ST_THR.map(s=>rs.filter(r=>r._thr_status===s.v).length)};
  });

  /* G7 · delta vs target per indikator */
  const g7=S.cat.map(c=>{
    const rs=rows.filter(r=>r.Indicator===c.ind);
    return {label:c.short, vals:ST_TGT.map(s=>rs.filter(r=>r._tgt_status===s.v).length),
      tot:rs.length};
  }).filter(x=>x.tot>0);

  /* G8 · delta rata-rata vs target delta per indikator */
  const g8=S.cat.map(c=>{
    const rs=rows.filter(r=>r.Indicator===c.ind);
    const d=avg(rs.map(r=>r._delta).filter(v=>v!==null));
    return {label:c.short, v:d, target:targetOf(c.ind), tot:rs.length};
  }).filter(x=>x.tot>0);

  /* urutan dan nama kolom persis seperti tblAnalisis di Master (7) */
  const COLS=[["Zonal","Zonal"],["Area Program","Area Program"],["Outcome","Outcome"],["Code","Code"],
    ["Indicator","Indikator"],["Pct_Base","% Baseline"],["Pct_LOP","% Endline (LOP)"],
    ["Threshold","Threshold"],["_thr_status","Endline vs Threshold"],
    ["_delta","Delta (Endline − Baseline)"],["_target","Target Delta"],
    ["_tgt_status","Delta vs Target"],["_status","Status Indikator"],
    ["_arah","Arah Indikator"],["_berlaku","Berlaku"]];
  const view=sortRows(rows);
  const cap=F.showAll?view.length:Math.min(view.length,150);
  const c=checks();
  const nInd=uniq(rows.map(r=>r.Indicator)).length;
  const indAtas=uniq(rows.filter(r=>r._thr_status===">= threshold").map(r=>r.Indicator)).length;
  const indTarget=uniq(rows.filter(r=>r._tgt_status===">= target").map(r=>r.Indicator)).length;
  const indSesuai=uniq(rows.filter(r=>r._status==="Sesuai Target/Threshold").map(r=>r.Indicator)).length;
  const indTinjau=uniq(rows.filter(r=>r._status==="Ditinjau").map(r=>r.Indicator)).length;
  const indHati=uniq(rows.filter(r=>r._status==="Perhatian").map(r=>r.Indicator)).length;

  return filterBand(true)+

  '<div class="slabel">Ringkasan'+activeLine()+'</div>'+
  '<div class="cards">'+
    card("Area<br>Program",aps.length,"dari "+AP_LIST.length+" terdaftar","teal")+
    card("Zonal",uniq(rows.map(r=>r.Zonal)).length,"","neutral")+
    card("&gt;= threshold",indAtas,"indikator, dari "+nInd,"teal")+
    card("Di bawah<br>threshold",uniq(rows.filter(r=>r._thr_status==="di bawah threshold")
      .map(r=>r.Indicator)).length,"indikator","critical")+
    card("Sesuai Target<br>/ Threshold",cnt("_status","Sesuai Target/Threshold"),indSesuai+" indikator","ready")+
    card("Ditinjau",cnt("_status","Ditinjau"),indTinjau+" indikator","review")+
    card("Perhatian",cnt("_status","Perhatian"),indHati+" indikator","critical")+
    card("&gt;= target",cnt("_tgt_status",">= target"),indTarget+" indikator","accent")+
  '</div>'+
  '<div class="legendrow"><span class="lbl">Legend</span>'+
    ST_IND.map(s=>'<span class="pill '+s.cls+'"><span class="ic">'+s.ic+'</span>'+s.v+'</span> '+
      '<span class="ldesc">'+esc(s.desc)+'</span>').join('')+
    '<span class="ldesc" style="margin-left:auto">Target Delta default <b>'+
      (CFG.target_delta*100).toFixed(0)+'pp</b></span></div>'+

  '<div class="grid2" style="margin-top:22px">'+
    '<div><div class="slabel" style="margin-top:0">Status indikator per Outcome '+
      '<span class="hint">jumlah indikator</span></div>'+
      '<div class="chartbox">'+chartStack(gOcStat,SER_IND,{pct100:true,labelCount:true,labW:70,bh:20,gap:11})+
        legend(SER_IND)+'</div></div>'+
    '<div><div class="slabel" style="margin-top:0">Endline vs Threshold per Outcome '+
      '<span class="hint">jumlah indikator</span></div>'+
      '<div class="chartbox">'+chartStack(gOcThr,SER_THR,{pct100:true,labelCount:true,labW:70,bh:20,gap:11})+
        legend(SER_THR)+'</div></div>'+
  '</div>'+

  '<div class="grid2" style="margin-top:20px">'+
    '<div><div class="slabel" style="margin-top:0">Status indikator per Area Program '+
      '<span class="hint">jumlah indikator</span></div>'+
      '<div class="chartbox">'+chartStack(g5,SER_IND,{pct100:true,labelCount:true,labW:138})+
        legend(SER_IND)+'</div></div>'+
    '<div><div class="slabel" style="margin-top:0">Endline vs Threshold per Area Program '+
      '<span class="hint">jumlah indikator</span></div>'+
      '<div class="chartbox">'+chartStack(g6,SER_THR,{pct100:true,labelCount:true,labW:138})+legend(SER_THR)+
      '<p class="chartnote">Urutan baris sama dengan grafik di sebelah kiri, jadi kedua grafik bisa dibaca berpasangan. '+
      'Panjang bar tetap proporsional supaya antar-AP bisa dibandingkan; angka di dalam segmen adalah jumlah indikator.</p></div></div>'+
  '</div>'+

  '<div class="grid2" style="margin-top:20px">'+
    '<div><div class="slabel" style="margin-top:0">Delta vs Target per indikator</div>'+
      '<div class="chartbox">'+chartStack(g7,SER_TGT,{pct100:true,labW:250,bh:12,gap:6})+legend(SER_TGT)+'</div></div>'+
    '<div><div class="slabel" style="margin-top:0">Delta rata-rata vs Target Delta</div>'+
      '<div class="chartbox">'+chartDiverge(g8,{labW:250})+
        legend([["Delta rata-rata",CLR.endline]],'<div><i style="background:'+CLR.ref+';width:3px"></i>Target Delta</div>')+
        '<p class="chartnote">Delta rata-rata hanya dari baris yang punya baseline dan endline. '+
        'Indikator berarah Turun bernilai negatif — target tercapai bila Delta ≤ Target Delta.</p></div></div>'+
  '</div>'+

  '<div class="slabel">Tabel analisis '+
    '<span class="hint">'+n0(view.length)+' baris'+(cap<view.length?' · menampilkan '+cap:'')+'</span></div>'+
  '<div class="tscroll"><table class="gt tight"><thead><tr>'+
    COLS.map(x=>th(x[1],x[0],(["Pct_Base","Pct_LOP","Threshold","_delta","_target"].indexOf(x[0])>=0)?"r":"")).join('')+
    '</tr></thead><tbody>'+
    view.slice(0,cap).map(r=>'<tr'+(r._berlaku==="No"?' class="oos"':'')+'>'+
      '<td class="dim">'+esc(r.Zonal)+'</td><td>'+esc(r["Area Program"])+'</td>'+
      '<td class="c dim">'+esc(r.Outcome)+'</td>'+
      '<td class="code'+(isBlank(r.Code)?' miss':'')+'">'+(isBlank(r.Code)?'—':esc(r.Code))+'</td>'+
      '<td class="ind" title="'+esc(r.Indicator)+'">'+esc(r._short)+'</td>'+
      '<td class="'+(N(r.Pct_Base)===0?"miss":"r")+'">'+pctD(r.Pct_Base)+'</td>'+
      '<td class="'+(N(r.Pct_LOP)===0?"belumcell":"r")+'">'+pctD(r.Pct_LOP)+'</td>'+
      '<td class="'+(N(r.Threshold)===0?"miss":"r")+'">'+pctT(r.Threshold)+'</td>'+
      '<td>'+pill(r._thr_status)+'</td>'+
      '<td class="r '+(r._wrongway?"wrongway":(r._delta>0?"up":r._delta<0?"down":"dim"))+'"'+
        (r._wrongway?' title="Delta bergerak berlawanan dengan arah indikator"':'')+'>'+ppD(r._delta)+'</td>'+
      '<td class="r dim">'+(r._target>0?"+":"")+(r._target*100).toFixed(0)+'pp</td>'+
      '<td>'+pill(r._tgt_status)+'</td>'+
      '<td>'+pill(r._status)+'</td>'+
      '<td class="c '+(r._arah==="Turun"?"turun":"naik")+'">'+esc(r._arah)+'</td>'+
      '<td class="c '+(r._berlaku==="No"?"nocell":"dim")+'">'+esc(r._berlaku)+'</td></tr>').join('')+
    '</tbody></table></div>'+
  (cap<view.length?'<div class="morebar"><button class="ghost" data-act="showAll">Tampilkan seluruh '+
    n0(view.length)+' baris</button></div>':'')+
  '<p class="tcap">Delta <span class="wrongway">merah tebal</span> bergerak <b>berlawanan</b> '+
   'dengan arah indikator. Arah dan Target Delta diatur di '+
   '<a href="#" data-go="ASUMSI">Asumsi Indikator</a>, Berlaku di '+
   '<a href="#" data-go="PEMETAAN">Pemetaan Indikator</a>.</p>'+

  (c.comboDupe>0?
  '<div class="slabel">Duplikat AP × indikator <span class="hint">tidak tertangkap oleh cek Row ID</span></div>'+
  '<div class="warnbox"><b>'+c.comboDupe+' kombinasi</b> Area Program × indikator muncul lebih dari satu kali, '+
   'mencakup <b>'+n0(c.dupeRows)+' baris</b>. Row ID-nya unik, jadi cek ROW ID tetap lolos — tapi setiap '+
   'hitungan per AP dan setiap rata-rata per indikator menghitungnya dua kali.</div>'+
  '<div class="tscroll"><table class="gt tight"><thead><tr><th>Area Program</th><th>Indikator</th>'+
    '<th class="r">Muncul</th><th>Row ID</th><th class="r">% Baseline</th><th class="r">% Endline</th>'+
    '<th class="r">Threshold</th></tr></thead><tbody>'+
    (()=>{const g={}; S.rows.filter(r=>r._dupe).forEach(r=>{
        const k=r["Area Program"]+"|"+r.Indicator; (g[k]=g[k]||[]).push(r);});
      return Object.keys(g).slice(0,40).map(k=>{const v=g[k];
        return v.map((r,i)=>'<tr'+(i?'':' class="grp"')+'>'+
          (i?'<td></td><td></td><td></td>':'<td><b>'+esc(r["Area Program"])+'</b></td>'+
            '<td class="ind" title="'+esc(r.Indicator)+'">'+esc(r._short)+'</td>'+
            '<td class="r crit" rowspan="'+v.length+'">'+v.length+'×</td>')+
          '<td class="code">'+esc(r.Row_ID)+'</td>'+
          '<td class="'+(N(r.Pct_Base)===0?"miss":"r")+'">'+pctD(r.Pct_Base)+'</td>'+
          '<td class="'+(N(r.Pct_LOP)===0?"belumcell":"r")+'">'+pctD(r.Pct_LOP)+'</td>'+
          '<td class="'+(N(r.Threshold)===0?"miss":"r")+'">'+pctT(r.Threshold)+'</td></tr>').join('');
      }).join('');})()+
    '</tbody></table></div>':'');
}

/* ==========================================================================
   SHEET KONFIGURASI 1 — Asumsi Indikator
   ========================================================================== */
function renderAsumsi(){
  const rows=S.asumsi;
  return '<div class="notice">Indikator yang <b>tidak</b> terdaftar di sini otomatis dianggap '+
    '<b>Arah = Naik</b> dengan <b>Target Delta = '+(CFG.target_delta*100).toFixed(0)+'pp</b>. '+
    'Asumsi default itu tidak disembunyikan: ia tertulis di sini supaya terlihat.</div>'+

  '<div class="fgrid" style="margin-top:16px;max-width:420px">'+
    '<div class="fg"><label>Target Delta default</label>'+
      '<input type="number" step="1" data-cfg="target_delta" value="'+(CFG.target_delta*100).toFixed(0)+'">'+
      '<div class="hlp">Dalam poin persentase. Dipakai untuk semua indikator yang tidak ada di tabel di bawah.</div></div>'+
  '</div>'+

  '<div class="slabel">tblAsumsi <span class="hint">'+rows.length+' indikator diatur khusus</span></div>'+
  '<div class="tscroll"><table class="gt ed"><thead><tr><th>Indikator</th><th class="c">Arah</th>'+
    '<th class="r">Target Delta (pp)</th><th>Cek nama</th><th></th></tr></thead><tbody>'+
    rows.map((a,i)=>{
      const known=IND_LIST.indexOf(a.ind)>=0;
      const used=S.rows.filter(r=>r.Indicator===a.ind).length;
      return '<tr><td><select class="cell wide" data-as="'+i+'" data-k="ind">'+
        IND_LIST.map(x=>'<option'+(x===a.ind?" selected":"")+' value="'+esc(x)+'">'+esc(shortOf(x))+'</option>').join('')+
        (known?'':'<option selected value="'+esc(a.ind)+'">'+esc(a.ind)+'</option>')+'</select></td>'+
        '<td class="c"><select class="cell" data-as="'+i+'" data-k="arah">'+
          ["Naik","Turun"].map(x=>'<option'+(x===a.arah?" selected":"")+'>'+x+'</option>').join('')+'</select></td>'+
        '<td class="r"><input class="cell r" size="6" type="number" step="1" data-as="'+i+
          '" data-k="delta" value="'+(a.delta*100).toFixed(0)+'"></td>'+
        '<td>'+(known?'<span class="ichip ok">● OK · '+used+' baris</span>'
                     :'<span class="ichip no">▲ NAMA TIDAK DIKENAL</span>')+'</td>'+
        '<td><button class="tiny warn" data-act="delAsumsi" data-i="'+i+'">hapus</button></td></tr>';
    }).join('')+'</tbody></table></div>'+
  '<div class="addbar"><span>Tambah asumsi</span>'+
    '<select id="newAsInd" style="max-width:420px">'+
      IND_LIST.filter(x=>!S.asumsi.some(a=>a.ind===x)).map(x=>
        '<option value="'+esc(x)+'">'+esc(shortOf(x))+'</option>').join('')+'</select>'+
    '<select id="newAsArah"><option>Turun</option><option>Naik</option></select>'+
    '<input id="newAsDelta" type="number" step="1" value="-10" size="5" title="poin persentase">'+
    '<button class="ghost" data-act="addAsumsi">Tambahkan</button></div>'+

  '<div class="slabel">Akibatnya pada analisis</div>'+
  '<div class="tscroll"><table class="gt"><thead><tr><th>Arah</th><th class="r">Indikator</th>'+
    '<th class="r">&gt;= target</th><th class="r">dibawah target</th></tr></thead><tbody>'+
    ["Naik","Turun"].map(ar=>{
      const rs=S.rows.filter(r=>r._arah===ar);
      return '<tr><td><b class="'+(ar==="Turun"?"turun":"naik")+'">'+ar+'</b></td>'+
        '<td class="r">'+uniq(rs.map(r=>r.Indicator)).length+'</td>'+
        '<td class="r">'+rs.filter(r=>r._tgt_status===">= target").length+'</td>'+
        '<td class="r">'+rs.filter(r=>r._tgt_status==="dibawah target").length+'</td></tr>';
    }).join('')+'</tbody></table></div>'+
  '<p class="tcap">Mengubah Arah membalik seluruh logika status untuk indikator itu: '+
   'pada arah <b>Turun</b> endline dinilai tercapai bila <b>≤</b> threshold, dan delta tercapai bila <b>≤</b> Target Delta.</p>';
}

/* ==========================================================================
   SHEET KONFIGURASI 2 — Pemetaan Indikator
   ========================================================================== */
function renderPemetaan(){
  const P=S.pemetaan;
  const unknown=P.aps.filter(a=>!AP_LIST.some(x=>x.ap===a));
  const yes=P.rows.reduce((n,p)=>n+p.v.filter(v=>String(v).toLowerCase()!=="no").length,0);
  const tot=P.rows.length*P.aps.length;

  return '<div class="notice">Matriks Yes/No: apakah indikator itu <b>berlaku</b> untuk AP tersebut. '+
    'Kombinasi yang di-set <b>No</b> tetap ada di data, tapi bisa disembunyikan lewat filter '+
    '<b>Berlaku</b> di kedua halaman laporan. Nama kolom AP harus persis sama dengan '+
    '<span class="keyprev">Area Program</span> di data.</div>'+

  '<div class="health" style="margin-top:14px">'+
    '<div class="m">Berlaku<b style="color:var(--green)">'+n0(yes)+'</b></div>'+
    '<div class="m">Tidak berlaku<b style="color:var(--grey)">'+n0(tot-yes)+'</b></div>'+
    '<div class="m">Cek nama AP<b>'+(unknown.length
      ?'<span class="ichip no">▲ '+esc(unknown.join(", "))+'</span>'
      :'<span class="ichip ok">● semua dikenal</span>')+'</b></div>'+
  '</div>'+

  '<div class="tscroll" style="margin-top:12px"><table class="gt ed matrix"><thead><tr>'+
    '<th class="sticky">Indikator</th><th class="c">Code</th>'+
    P.aps.map(a=>'<th class="c vert"><span>'+esc(a)+'</span></th>').join('')+
    '<th class="r">Yes</th></tr></thead><tbody>'+
    P.rows.map((p,i)=>{
      const y=p.v.filter(v=>String(v).toLowerCase()!=="no").length;
      return '<tr><td class="sticky ind" title="'+esc(p.ind)+'">'+esc(shortOf(p.ind))+'</td>'+
        '<td class="c code dim">'+(isBlank(p.code)?"—":esc(p.code))+'</td>'+
        p.v.map((v,k)=>{
          const on=String(v).toLowerCase()!=="no";
          return '<td class="c"><button class="yn '+(on?"y":"n")+'" data-pem="'+i+'" data-k="'+k+
            '" title="'+esc(P.aps[k])+'">'+(on?"Yes":"No")+'</button></td>';
        }).join('')+
        '<td class="r">'+y+'</td></tr>';
    }).join('')+
    '</tbody><tfoot><tr><td class="sticky">Berlaku per AP</td><td></td>'+
      P.aps.map((a,k)=>'<td class="c">'+P.rows.filter(p=>String(p.v[k]).toLowerCase()!=="no").length+'</td>').join('')+
      '<td class="r">'+n0(yes)+'</td></tr></tfoot></table></div>'+
  '<p class="tcap">Klik satu sel untuk membaliknya. Perubahan langsung menghitung ulang kolom '+
   '<b>Berlaku</b> di kedua halaman laporan. Untuk menyimpannya permanen, unduh '+
   '<span class="keyprev">pemetaan.js</span> dari tombol Simpan di atas dan commit ke repository.</p>';
}

/* ==========================================================================
   WIRING
   ========================================================================== */
function wire(id,el){
  el.querySelectorAll("[data-go]").forEach(a=>a.onclick=ev=>{ev.preventDefault();go(a.dataset.go);});
  el.querySelectorAll("[data-sort]").forEach(t=>t.onclick=()=>{
    const k=t.dataset.sort;
    if(F.sort===k) F.sortDir=-F.sortDir; else {F.sort=k;F.sortDir=1;}
    paint(id);
  });
  el.querySelectorAll("[data-slice]").forEach(b=>b.onclick=()=>{
    const f=b.dataset.slice, v=b.dataset.val;
    if(f==="berlaku") F.berlaku=v;
    else if(f==="period") F.period=v;
    else { const i=F[f].indexOf(v); i>=0?F[f].splice(i,1):F[f].push(v); }
    if(f==="outcome"&&F.outcome.length) F.ind=F.ind.filter(x=>F.outcome.indexOf((CAT_BY_IND[x]||{}).oc)>=0);
    if(f==="zonal"&&F.zonal.length) F.ap=F.ap.filter(ap=>F.zonal.indexOf(S.apz[ap])>=0);
    F.showAll=false; paint(id);
  });
  el.querySelectorAll("[data-clear]").forEach(b=>b.onclick=()=>{F[b.dataset.clear]=[];paint(id);});
  el.querySelectorAll("[data-cfg]").forEach(c=>c.onchange=()=>{
    CFG.target_delta=(parseFloat(c.value)||0)/100; recompute(); saveLocal(); paint(id);
    toast('Target Delta default sekarang <b>'+(CFG.target_delta*100).toFixed(0)+'pp</b>. Seluruh analisis dihitung ulang.');
  });
  el.querySelectorAll("[data-as]").forEach(c=>c.onchange=()=>{
    const a=S.asumsi[+c.dataset.as], k=c.dataset.k;
    a[k] = k==="delta" ? (parseFloat(c.value)||0)/100 : c.value;
    recompute(); saveLocal(); paint(id);
  });
  el.querySelectorAll("[data-pem]").forEach(b=>b.onclick=()=>{
    const p=S.pemetaan.rows[+b.dataset.pem], k=+b.dataset.k;
    p.v[k] = String(p.v[k]).toLowerCase()==="no" ? "Yes" : "No";
    recompute(); saveLocal(); paint(id);
  });
  el.querySelectorAll("[data-act]").forEach(b=>b.onclick=()=>act(b,id));
}
function act(b,id){
  const a=b.dataset.act;
  if(a==="showAll"){ F.showAll=true; paint(id); return; }
  if(a==="fbOpen"){ F.fbOpen=true; paint(id); return; }
  if(a==="fbClose"){ F.fbOpen=false; paint(id); return; }
  if(a==="fbReset"){ F.zonal=[];F.ap=[];F.outcome=[];F.status=[];F.berlaku="Yes";F.showAll=false; paint(id); return; }
  if(a==="natReset"){ F.zonal=[];F.ap=[];F.outcome=[];F.ind=[];F.period="Both"; paint(id); return; }
  if(a==="delAsumsi"){
    const g=S.asumsi.splice(+b.dataset.i,1)[0];
    recompute(); saveLocal(); paint(id);
    toast('Asumsi dihapus. <b>'+esc(shortOf(g.ind))+'</b> sekarang dianggap Naik dengan target default.');
    return;
  }
  if(a==="addAsumsi"){
    const ind=document.getElementById("newAsInd").value;
    if(!ind){ toast('▲ Pilih indikatornya dulu.'); return; }
    if(S.asumsi.some(x=>x.ind===ind)){ toast('▲ Indikator itu sudah ada di tabel asumsi.'); return; }
    S.asumsi.push({ind:ind, arah:document.getElementById("newAsArah").value,
      delta:(parseFloat(document.getElementById("newAsDelta").value)||0)/100});
    recompute(); saveLocal(); paint(id);
    toast('Asumsi ditambahkan untuk <b>'+esc(shortOf(ind))+'</b>.');
    return;
  }
  if(a==="dlAsumsi")   return dl("asumsi.js",fileAsumsi(),"text/javascript");
  if(a==="dlPemetaan") return dl("pemetaan.js",filePemetaan(),"text/javascript");
  if(a==="dlIndicators") return dl("indicators.js",fileIndicators(),"text/javascript");
  if(a==="dlCsv")      return dlCsv();
  if(a==="reset"){ clearLocal(); location.reload(); }
}

/* ---------- unduhan ---------- */
function dl(name,text,mime){
  const b=new Blob([text],{type:(mime||"text/plain")+";charset=utf-8;"});
  const a=document.createElement("a");
  a.href=URL.createObjectURL(b); a.download=name; a.click();
  setTimeout(()=>URL.revokeObjectURL(a.href),2000);
}
const jv=v=>v===null||v===undefined||v===""?"null":(typeof v==="number"?String(v):JSON.stringify(v));
const head=(t,note)=>"/* "+"=".repeat(74)+"\n   "+t+"\n   "+"-".repeat(74)+"\n   "+
  note.split("\n").join("\n   ")+"\n   Ditulis dari halaman pada "+new Date().toISOString().slice(0,10)+
  " · "+CFG.version+"\n   "+"=".repeat(74)+" */\n";
function fileAsumsi(){
  return head("tblAsumsi  —  data/asumsi.js",
    "Arah dan Target Delta per indikator. Yang tidak terdaftar dianggap Naik\n"+
    "dengan Target Delta = target_delta_default.")+
    "window.WVI_ASUMSI = {\n  target_delta_default: "+CFG.target_delta+",\n  rows: [\n"+
    S.asumsi.map(a=>'  {ind:'+JSON.stringify(a.ind)+', arah:'+JSON.stringify(a.arah)+
      ', delta:'+a.delta+'}').join(",\n")+"\n  ]\n};\n";
}
function filePemetaan(){
  return head("tblPemetaan  —  data/pemetaan.js",
    "Matriks Yes/No indikator × Area Program.")+
    "window.WVI_PEMETAAN = {\n  aps: "+JSON.stringify(S.pemetaan.aps)+",\n  rows: [\n"+
    S.pemetaan.rows.map(p=>'  {code:'+jv(p.code)+', ind:'+JSON.stringify(p.ind)+
      ', v:'+JSON.stringify(p.v)+'}').join(",\n")+"\n  ]\n};\n";
}
function fileIndicators(){
  const cols=window.WVI_INDICATORS.columns;
  return head("tblIndicators  —  data/indicators.js   ("+S.rows.length+" baris)",
    "17 kolom, urutan tidak diubah. KONVENSI PEARL: nilai 0 pada % Baseline atau\n"+
    "% LOP berarti belum ada data. Row ID wajib unik.\n"+cols.join(", "))+
    "window.WVI_INDICATORS = {\n  columns: "+JSON.stringify(cols)+",\n  rows: [\n"+
    S.rows.map(r=>"  ["+cols.map(k=>jv(r[k])).join(",")+"]").join(",\n")+"\n  ]\n};\n";
}
function dlCsv(){
  const cols=window.WVI_INDICATORS.columns.concat(
    ["_arah","_target","_thr_status","_delta","_tgt_status","_status","_berlaku"]);
  const q=v=>'"'+String(v==null?"":v).replace(/"/g,'""')+'"';
  const csv=[cols.map(q).join(",")].concat(
    filtered().map(r=>cols.map(k=>q(r[k])).join(","))).join("\r\n");
  dl("analisis_ap.csv","\ufeff"+csv,"text/csv");
}

/* ---------- toast ---------- */
let tT=null;
function toast(html){
  const t=document.getElementById("toast");
  t.innerHTML=html; t.classList.add("on");
  clearTimeout(tT); tT=setTimeout(()=>t.classList.remove("on"),4200);
}/* ==========================================================================
   v4 — National Target Dashboard
   Lapisan tampilan saja. Mesin data (recompute, weightedOf, natRows) tidak diubah.
   ========================================================================== */
const CL={brand:"#FF5515",blue:"#005A9C",ok:"#157F4B",warn:"#C77700",bad:"#C62828",
  ln:"#E6E9EF",ln2:"#F0F2F6",ink:"#14161A",ink2:"#565D6D",ink3:"#8B92A1",ink4:"#AEB4C0"};
const pc1=v=>v===null||v===undefined?"—":(v*100).toFixed(1)+"%";
const pc0=v=>v===null||v===undefined?"—":(v*100).toFixed(0)+"%";
const ppv=v=>v===null||v===undefined?"—":(v>0?"+":"")+(v*100).toFixed(1)+"pp";

/* capaian terhadap threshold, sadar arah indikator */
function achOf(w,dir,thr){
  if(w.pE===null||!thr) return null;
  const a = dir===-1 ? (w.pE>0 ? thr/w.pE : null) : w.pE/thr;
  return (a===null||!isFinite(a)) ? null : a;
}
function perInd(rows){
  return S.cat.map(c=>{
    const rs=rows.filter(r=>r.Indicator===c.ind);
    if(!rs.length) return null;
    const w=weightedOf(rs), dir=arahOf(c.ind)==="Turun"?-1:1;
    const thrs=rs.map(r=>N(r.Threshold)).filter(v=>v>0);
    const thr=thrs.length?thrs.sort((a,b)=>a-b)[thrs.length>>1]:null;
    return {ind:c.ind,short:c.short,code:c.code,oc:c.oc,dir:dir,thr:thr,w:w,n:rs.length,
      ach:achOf(w,dir,thr),
      meets:(w.pE!==null&&thr)?(dir===-1?w.pE<=thr:w.pE>=thr):null};
  }).filter(Boolean);
}

/* ---------------- grafik ---------------- */
/* Label SVG dua baris, rata kiri.
   Versi lama memakai text-anchor="end" di x = LW-12: teks memanjang ke KIRI
   melewati batas viewBox lalu terpotong, sementara "..." yang sudah tertanam di
   data/config.js menempel di kanan. Itu yang membuat "OIOS 96" terbaca "OS 96".
   Sekarang teks penuh dipakai, dipatah maksimal dua baris, rata kiri. */
function svgWrap(text,x,yMid,cw,maxLines,size,fill){
  const words=String(text==null?"":text).replace(/\u2026/g,"").trim().split(/\s+/);
  const lines=[]; let cur="", cut=false;
  for(let i=0;i<words.length;i++){
    const t=cur?cur+" "+words[i]:words[i];
    if(t.length<=cw){ cur=t; continue; }
    if(lines.length<maxLines-1){ lines.push(cur); cur=words[i]; }
    else { cut=true; break; }
  }
  if(cur) lines.push(cur);
  if(cut){
    const L=lines.length-1;
    lines[L]=lines[L].slice(0,Math.max(1,cw-1)).replace(/\s+\S*$/,"")+"\u2026";
  }
  const lh=size+1.5, y0=yMid-(lines.length-1)*lh/2+size*0.34;
  let o='<text x="'+x+'" y="'+y0.toFixed(1)+'" font-size="'+size+'" fill="'+fill+'">';
  for(let i=0;i<lines.length;i++){
    o+='<tspan x="'+x+'"'+(i?' dy="'+lh+'"':'')+'>'+esc(lines[i])+'</tspan>';
  }
  return o+'</text>';
}
function chHBar(items){
  if(!items.length) return '<div class="xempty">Tidak ada indikator pada filter ini.</div>';
  const W=680,LW=252,BH=8,GAP=3,P=BH*2+GAP+22,H=items.length*P+22,PW=W-LW-62;
  let s='<svg class="ch" viewBox="0 0 '+W+' '+H+'" role="img"><title>Baseline vs Evaluation</title>';
  for(let g=0;g<=4;g++){const x=LW+PW*g/4;
    s+='<line x1="'+x+'" y1="6" x2="'+x+'" y2="'+(H-16)+'" stroke="'+CL.ln2+'"/>'+
       '<text x="'+x+'" y="'+(H-4)+'" text-anchor="middle" font-size="9.5" fill="'+CL.ink4+'">'+(g*25)+'%</text>';}
  items.forEach((it,i)=>{
    const y=i*P+12;
    const wb=it.a==null?0:PW*Math.min(1,it.a), we=it.b==null?0:PW*Math.min(1,it.b);
    const tip=it.label+' \u00B7 Baseline '+(it.a==null?"belum ada data":pc0(it.a))+
              ' \u00B7 Evaluation '+(it.b==null?"belum ada data":pc0(it.b));
    s+='<g class="chr" style="--i:'+i+'" data-tip="'+esc(tip)+'">'+
       '<rect class="chit" fill="none" pointer-events="all" x="0" y="'+(y-8)+'" width="'+W+'" height="'+(P-2)+'"/>'+
       '<rect class="ctrk" fill="#F4F6F9" x="'+LW+'" y="'+y+'" width="'+PW+'" height="'+(BH*2+GAP)+'" rx="3"/>'+
       svgWrap(it.label,0,y+BH+1,40,2,11.5,CL.ink)+
       '<rect x="'+LW+'" y="'+y+'" width="'+wb.toFixed(1)+'" height="'+BH+'" rx="3" fill="'+CL.blue+'" opacity=".85"/>'+
       '<rect x="'+LW+'" y="'+(y+BH+GAP)+'" width="'+we.toFixed(1)+'" height="'+BH+'" rx="3" fill="'+CL.brand+'"/>'+
       '<text x="'+(LW+Math.max(wb,we)+9).toFixed(1)+'" y="'+(y+BH+2)+'" font-size="10.5" fill="'+CL.ink2+'">'+
       (it.a==null?"\u2014":pc0(it.a))+' \u2192 <tspan font-weight="600" fill="'+CL.ink+'">'+(it.b==null?"\u2014":pc0(it.b))+'</tspan></text>'+
       '</g>';
  });
  return s+'</svg>';
}
function chBullet(items){
  if(!items.length) return '<div class="xempty">Tidak ada indikator dengan threshold pada filter ini.</div>';
  const W=680,LW=252,BH=16,P=BH+22,H=items.length*P+22,PW=W-LW-66;
  let s='<svg class="ch" viewBox="0 0 '+W+' '+H+'" role="img"><title>Baseline vs Threshold, bullet chart</title>';
  for(let g=0;g<=4;g++){const x=LW+PW*g/4;
    s+='<text x="'+x+'" y="'+(H-4)+'" text-anchor="middle" font-size="9.5" fill="'+CL.ink4+'">'+(g*25)+'%</text>';}
  items.forEach((it,i)=>{
    const y=i*P+10;
    const tip=it.label+' \u00B7 Evaluation '+(it.b==null?"belum ada data":pc0(it.b))+
              (it.thr==null?"":' \u00B7 Threshold '+pc0(it.thr)+
               ' \u00B7 '+(it.meets?"tercapai":"belum tercapai"));
    s+='<g class="chr" style="--i:'+i+'" data-tip="'+esc(tip)+'">'+
       '<rect class="chit" fill="none" pointer-events="all" x="0" y="'+(y-6)+'" width="'+W+'" height="'+(P-2)+'"/>';
    s+=svgWrap(it.label,0,y+BH/2,40,2,11.5,CL.ink);
    s+='<rect x="'+LW+'" y="'+y+'" width="'+PW+'" height="'+BH+'" rx="3" fill="#F4F6F9"/>';
    if(it.thr!=null){
      const wt=PW*Math.min(1,it.thr);
      s+='<rect x="'+LW+'" y="'+y+'" width="'+wt.toFixed(1)+'" height="'+BH+'" rx="3" fill="#E8ECF2"/>';
    }
    if(it.b!=null){
      const we=PW*Math.min(1,it.b), col=it.meets?CL.ok:CL.brand;
      s+='<rect x="'+LW+'" y="'+(y+4)+'" width="'+we.toFixed(1)+'" height="'+(BH-8)+'" rx="2" fill="'+col+'"/>';
    }
    if(it.a!=null){
      const wb=LW+PW*Math.min(1,it.a);
      s+='<circle cx="'+wb.toFixed(1)+'" cy="'+(y+BH/2)+'" r="3.2" fill="#fff" stroke="'+CL.blue+'" stroke-width="1.6"/>';
    }
    if(it.thr!=null){
      const xt=LW+PW*Math.min(1,it.thr);
      s+='<line x1="'+xt.toFixed(1)+'" y1="'+(y-2)+'" x2="'+xt.toFixed(1)+'" y2="'+(y+BH+2)+
         '" stroke="'+CL.ink+'" stroke-width="2"/>';
    }
    s+='<text x="'+(LW+PW+10)+'" y="'+(y+BH-3)+'" font-size="10.5" font-weight="600" fill="'+
       (it.meets?CL.ok:it.b==null?CL.ink4:CL.bad)+'">'+(it.b==null?"\u2014":pc0(it.b))+'</text>';
    s+='</g>';
  });
  return s+'</svg>';
}
function chDiverge(items){
  if(!items.length) return '<div class="xempty">Belum ada indikator dengan baseline dan evaluation.</div>';
  const W=680,LW=252,BH=10,P=BH+20,H=items.length*P+26,PW=W-LW-72;
  const m=Math.max(.02,...items.map(i=>Math.abs(i.v||0))), sc=Math.ceil(m*20)/20;
  const z=LW+PW/2, half=PW/2, x=v=>z+half*Math.max(-1,Math.min(1,(v||0)/sc));
  let s='<svg class="ch" viewBox="0 0 '+W+' '+H+'" role="img"><title>Delta per indikator</title>';
  s+='<line x1="'+z+'" y1="14" x2="'+z+'" y2="'+(H-14)+'" stroke="'+CL.ink3+'" stroke-width="1"/>'+
     '<text x="'+z+'" y="9" text-anchor="middle" font-size="9.5" fill="'+CL.ink4+'">0</text>'+
     '<text x="'+LW+'" y="9" font-size="9.5" fill="'+CL.ink4+'">−'+(sc*100).toFixed(0)+'pp</text>'+
     '<text x="'+(LW+PW)+'" y="9" text-anchor="end" font-size="9.5" fill="'+CL.ink4+'">+'+(sc*100).toFixed(0)+'pp</text>';
  items.forEach((it,i)=>{
    const y=i*P+16, xv=x(it.v);
    const col=it.v>.001?CL.ok:it.v<-.001?CL.bad:CL.ink4;
    const tip=it.label+' \u00B7 Delta '+((it.v>0?"+":"")+(it.v*100).toFixed(1))+' poin persentase';
    s+='<g class="chr" style="--i:'+i+'" data-tip="'+esc(tip)+'">'+
       '<rect class="chit" fill="none" pointer-events="all" x="0" y="'+(y-8)+'" width="'+W+'" height="'+(P-2)+'"/>'+
       svgWrap(it.label,0,y+BH/2,40,2,11.5,CL.ink)+
       '<rect x="'+Math.min(z,xv).toFixed(1)+'" y="'+y+'" width="'+Math.abs(xv-z).toFixed(1)+
       '" height="'+BH+'" rx="2" fill="'+col+'"/>'+
       '<text x="'+(LW+PW+10)+'" y="'+(y+BH-1)+'" font-size="10.5" font-weight="600" fill="'+col+'">'+
       ((it.v>0?"+":"")+(it.v*100).toFixed(1))+'</text>'+
       '</g>';
  });
  return s+'</svg>';
}
function chDonut(segs,centerVal,centerLab){
  const tot=segs.reduce((a,s)=>a+s.v,0);
  const W=300,H=200,cx=100,cy=100,R=76,r=50;
  let s='<svg class="ch" viewBox="0 0 '+W+' '+H+'" role="img"><title>'+esc(centerLab||"")+'</title>';
  if(!tot){ s+='<circle cx="'+cx+'" cy="'+cy+'" r="'+((R+r)/2)+'" fill="none" stroke="'+CL.ln2+'" stroke-width="'+(R-r)+'"/>'; }
  let a0=-Math.PI/2;
  segs.forEach(sg=>{
    if(!sg.v) return;
    const a1=a0+2*Math.PI*sg.v/tot, big=(a1-a0)>Math.PI?1:0;
    const p=(rr,a)=>[cx+rr*Math.cos(a),cy+rr*Math.sin(a)];
    const [x1,y1]=p(R,a0),[x2,y2]=p(R,a1),[x3,y3]=p(r,a1),[x4,y4]=p(r,a0);
    s+='<path d="M'+x1.toFixed(1)+' '+y1.toFixed(1)+'A'+R+' '+R+' 0 '+big+' 1 '+x2.toFixed(1)+' '+y2.toFixed(1)+
       'L'+x3.toFixed(1)+' '+y3.toFixed(1)+'A'+r+' '+r+' 0 '+big+' 0 '+x4.toFixed(1)+' '+y4.toFixed(1)+'Z" fill="'+sg.c+'"/>';
    a0=a1;
  });
  s+='<text x="'+cx+'" y="'+(cy-2)+'" text-anchor="middle" font-size="26" font-weight="600" fill="'+CL.ink+'">'+centerVal+'</text>'+
     '<text x="'+cx+'" y="'+(cy+17)+'" text-anchor="middle" font-size="10" fill="'+CL.ink3+'">'+esc(centerLab||"")+'</text>';
  let ly=32;
  segs.forEach(sg=>{
    s+='<rect x="205" y="'+(ly-8)+'" width="9" height="9" rx="2" fill="'+sg.c+'"/>'+
       '<text x="220" y="'+ly+'" font-size="11" fill="'+CL.ink2+'">'+esc(sg.k)+'</text>'+
       '<text x="292" y="'+ly+'" text-anchor="end" font-size="11" font-weight="600" fill="'+CL.ink+'">'+sg.v+'</text>';
    ly+=22;
  });
  return s+'</svg>';
}

/* ---------------- kartu pembungkus ---------------- */
function xcard(id,title,sub,body,foot){
  return '<div class="xc" id="xc_'+id+'"><div class="xc-h"><div class="tt"><h3>'+title+'</h3>'+
    (sub?'<div class="st">'+sub+'</div>':'')+'</div>'+
    '<div class="xmenu" data-menu><button aria-label="Menu" data-mtoggle>⋯</button><div class="xmenu-pop">'+
      '<button data-mact="csv" data-mid="'+id+'"><span>⤓</span> Export CSV</button>'+
      '<button data-mact="print"><span>⎙</span> Cetak halaman</button>'+
    '</div></div></div>'+
    '<div class="xc-b">'+body+'</div>'+(foot?'<div class="xc-f">'+foot+'</div>':'')+'</div>';
}

/* ---------------- filter bar ---------------- */
function ddl(field,label,items,selArr){
  const n=selArr.length;
  return '<div class="ddl'+(n?' act':'')+'" data-ddl="'+field+'"><button data-dtoggle>'+
    '<span>'+label+(n?'':'')+'</span>'+(n?'<span class="cnt">'+n+'</span>':'<span style="color:#AEB4C0">▾</span>')+
    '</button><div class="ddl-pop">'+
    items.map(it=>'<label><input type="checkbox" data-fset="'+field+'" value="'+esc(it.v)+'"'+
      (selArr.indexOf(it.v)>=0?' checked':'')+'><span>'+esc(it.lab||it.v)+'</span>'+
      '<span class="n">'+it.n+'</span></label>').join('')+'</div></div>';
}
function filterBar(){
  const all=S.rows;
  const apPool=F.zonal.length?AP_LIST.filter(a=>F.zonal.indexOf(a.zonal)>=0):AP_LIST;
  const indPool=S.cat.filter(c=>!F.outcome.length||F.outcome.indexOf(c.oc)>=0);
  const q=(F.q||"").toLowerCase();
  const chips=[]
    .concat(F.zonal.map(v=>({f:"zonal",v:v,l:"Zone"})))
    .concat(F.ap.map(v=>({f:"ap",v:v,l:"AP"})))
    .concat(F.outcome.map(v=>({f:"outcome",v:v,l:"Outcome"})))
    .concat(F.ind.map(v=>({f:"ind",v:shortOf(v),raw:v,l:"Indicator"})));
  const nAkt=chips.length+(F.period!=="Both"?1:0)+(q?1:0);
  return '<div class="fbar">'+
    '<div class="fbar-h"><span class="lb">Filter</span>'+
      '<span class="summ">'+(nAkt?esc(natActiveText()||"")+(q?' · "'+esc(F.q)+'"':''):'Seluruh dataset')+'</span>'+
      '<div class="sp"></div>'+
      '<div class="seg">'+["Baseline","Evaluation","Both"].map(p=>
        '<button data-fperiod="'+p+'" class="'+(F.period===p?"on":"")+'">'+p+'</button>').join('')+'</div>'+
      (nAkt?'<button class="tbtn" data-fact="reset"><i>↺</i>Reset filters</button>':'')+
      '<button class="tbtn" data-fact="toggle">'+(F.fbOpen?'Sembunyikan ▴':'Filter lanjutan ▾')+'</button>'+
    '</div>'+
    (F.fbOpen?'<div class="fbar-b">'+
      '<div class="fsearch"><i>⌕</i><input type="search" placeholder="Cari indikator…" value="'+esc(F.q||"")+'" data-fq></div>'+
      ddl("zonal","Zone",ZONALS.map(z=>({v:z,n:all.filter(r=>r.Zonal===z).length})),F.zonal)+
      ddl("ap","Area Program",apPool.map(a=>({v:a.ap,n:all.filter(r=>r["Area Program"]===a.ap).length})),F.ap)+
      ddl("outcome","Outcome",OUTCOMES.map(o=>({v:o,n:all.filter(r=>r.Outcome===o).length})),F.outcome)+
      ddl("ind","Indicator",indPool.map(c=>({v:c.ind,lab:c.short,n:all.filter(r=>r.Indicator===c.ind).length})),F.ind)+
    '</div>':'')+
    (chips.length?'<div class="fchips">'+chips.map(c=>'<span class="fchip"><span class="dim">'+c.l+'</span> <b>'+
      esc(c.v)+'</b><button data-fdel="'+c.f+'" data-fval="'+esc(c.raw||c.v)+'" aria-label="Hapus">×</button></span>').join('')+
      '</div>':'')+
  '</div>';
}

/* ==========================================================================
   HALAMAN 1 — DASHBOARD
   ========================================================================== */
function renderDash(){
  let rows=natRows();
  const q=(F.q||"").toLowerCase();
  if(q) rows=rows.filter(r=>String(r.Indicator).toLowerCase().indexOf(q)>=0||
    String(r.Code||"").toLowerCase().indexOf(q)>=0);
  const per=perInd(rows);
  const withE=per.filter(x=>x.w.pE!==null);
  const withBoth=per.filter(x=>x.w.delta!==null);
  const meets=per.filter(x=>x.meets===true);
  const ach=per.map(x=>x.ach).filter(v=>v!==null);
  const medAch=ach.length?ach.slice().sort((a,b)=>a-b)[ach.length>>1]:null;
  const naik=withBoth.filter(x=>x.w.delta>.001), turun=withBoth.filter(x=>x.w.delta<-.001);

  /* ringkasan capaian */
  const band=[{k:"≥ 90%",c:CL.ok,v:0},{k:"75 – 89%",c:"#7FB069",v:0},
              {k:"50 – 74%",c:CL.warn,v:0},{k:"< 50%",c:CL.bad,v:0},
              {k:"Belum ada data",c:CL.ln,v:0}];
  per.forEach(x=>{
    if(x.ach===null){ band[4].v++; return; }
    const a=x.ach*100;
    if(a>=90) band[0].v++; else if(a>=75) band[1].v++; else if(a>=50) band[2].v++; else band[3].v++;
  });
  /* per outcome */
  const ocSeg=OUTCOMES.filter(o=>per.some(x=>x.oc===o)).map((o,i)=>({
    k:o, v:per.filter(x=>x.oc===o).length,
    c:[CL.blue,"#2E7DB8","#5FA8D3",CL.brand,"#C77700"][i%5]}));
  /* outcome terbaik & terlemah */
  const ocStat=OUTCOMES.map(o=>{
    const g=per.filter(x=>x.oc===o&&x.ach!==null);
    if(!g.length) return null;
    const s=g.map(x=>x.ach).sort((a,b)=>a-b);
    return {oc:o,med:s[s.length>>1],n:g.length,meets:g.filter(x=>x.meets).length};
  }).filter(Boolean).sort((a,b)=>b.med-a.med);
  /* AP */
  const apStat=uniq(rows.map(r=>r["Area Program"])).map(ap=>{
    const rs=rows.filter(r=>r["Area Program"]===ap);
    const ok=rs.filter(r=>r._thr_status===">= threshold").length;
    const m=rs.filter(r=>r._thr_status!=="Belum ada data").length;
    return {ap:ap,ok:ok,m:m,rate:m?ok/m:null};
  }).filter(x=>x.rate!==null).sort((a,b)=>b.rate-a.rate);

  const top=withBoth.slice().sort((a,b)=>b.w.delta-a.w.delta).slice(0,10);
  const bot=withBoth.slice().sort((a,b)=>a.w.delta-b.w.delta).slice(0,10);

  const K=(ic,tone,lab,val,sub)=>'<div class="c3"><div class="kpi click" data-kpi="'+lab+'">'+
    '<div class="kpi-top"><span class="kpi-ic '+tone+'">'+ic+'</span><span class="kpi-lab">'+lab+'</span></div>'+
    '<div class="kpi-val">'+val+'</div><div class="kpi-sub">'+(sub||"&nbsp;")+'</div></div></div>';

  return '<div class="bcrumb">Dashboard <b>·</b> Indonesia National Summary</div>'+
  '<div class="grid" style="margin-bottom:24px">'+
    K("◎","t-brand","Area Program",uniq(rows.map(r=>r["Area Program"])).length,
      "dari "+AP_LIST.length+" terdaftar")+
    K("◍","t-blue","Zones",uniq(rows.map(r=>r.Zonal)).length,ZONALS.join(" · "))+
    K("◈","t-warn","Outcomes",uniq(rows.map(r=>r.Outcome).filter(Boolean)).length,"Goal dan OC 1 – OC 4")+
    K("▦","t-gray","Indicators",per.length,n0(rows.length)+" baris data")+
  '</div>'+

  filterBar()+

  '<div class="grid">'+
    '<div class="c8">'+xcard("cmp","Baseline vs Evaluation",
      "Weighted National — Σ Numerator ÷ Σ Denominator, per indikator",
      chHBar(per.map(x=>({label:(x.ind||x.short)+(x.dir===-1?" \u2193":""),
        a:F.period!=="Evaluation"?x.w.pB:null, b:F.period!=="Baseline"?x.w.pE:null}))),
      '<span style="display:inline-flex;gap:16px"><span><i style="display:inline-block;width:9px;height:9px;'+
      'border-radius:2px;background:'+CL.blue+'"></i> Baseline FY26</span>'+
      '<span><i style="display:inline-block;width:9px;height:9px;border-radius:2px;background:'+CL.brand+
      '"></i> Evaluation FY30</span><span>↓ indikator reduksi — penurunan berarti perbaikan</span></span>')+'</div>'+

    '<div class="c12">'+xcard("bul","Evaluation vs Threshold",
      "Bullet chart — batang gelap = threshold, titik = baseline",
      chBullet(per.filter(x=>x.thr!=null).sort((a,b)=>(a.ach||0)-(b.ach||0))
        .map(x=>({label:(x.ind||x.short)+(x.dir===-1?" \u2193":""),a:x.w.pB,b:x.w.pE,thr:x.thr,meets:x.meets}))),
      "Diurutkan dari capaian terendah. Hijau berarti threshold tercapai.")+'</div>'+

    '<div class="c6">'+xcard("perf","Performance Summary",
      "Capaian evaluation terhadap threshold, disesuaikan arah indikator",
      chDonut(band, String(meets.length), "capai threshold"),
      "Indikator reduksi dihitung terbalik: makin rendah nilainya, makin tinggi capaiannya.")+'</div>'+

    '<div class="c6">'+xcard("ocs","Outcome Summary","Jumlah indikator per outcome",
      chDonut(ocSeg, String(per.length), "indikator"),
      ocStat.length?'Capaian tertinggi <b>'+esc(ocStat[0].oc)+'</b>, terendah <b>'+
        esc(ocStat[ocStat.length-1].oc)+'</b>.':'')+'</div>'+

    '<div class="c12">'+xcard("dlt","Delta per Indicator","Poin persentase, peningkatan terbesar di atas",
      chDiverge(withBoth.slice().sort((a,b)=>b.w.delta-a.w.delta)
        .map(x=>({label:(x.ind||x.short)+(x.dir===-1?" \u2193":""),v:x.w.delta}))),
      (top.length?'Top improvement: <b style="color:'+CL.ok+'">'+esc(top[0].short)+' '+ppv(top[0].w.delta)+
        '</b> &nbsp;·&nbsp; Bottom: <b style="color:'+CL.bad+'">'+esc(bot[0].short)+' '+ppv(bot[0].w.delta)+'</b>':''))+'</div>'+

    '<div class="c6">'+xcard("top","Top Performing Indicators","Peningkatan terbesar dibanding baseline",
      tbl(top,"up"))+'</div>'+
    '<div class="c6">'+xcard("bot","Lowest Performing Indicators","Penurunan terbesar dibanding baseline",
      tbl(bot,"down"))+'</div>'+
  '</div>';
}
function tbl(list,dir){
  if(!list.length) return '<div class="xempty">Belum ada indikator dengan baseline dan evaluation.</div>';
  return '<div class="tscroll"><table class="tb"><thead><tr><th style="width:34px"></th><th>Indicator</th>'+
    '<th class="r">Baseline</th><th class="r">Evaluation</th><th class="r">Delta</th></tr></thead><tbody>'+
    list.map((x,i)=>'<tr><td><span class="rank">'+(i+1)+'</span></td>'+
      '<td class="nm">'+esc(x.short)+(x.dir===-1?' <span class="dim">↓</span>':'')+'</td>'+
      '<td class="r dim">'+pc1(x.w.pB)+'</td><td class="r">'+pc1(x.w.pE)+'</td>'+
      '<td class="r '+(x.w.delta>0?"xup":x.w.delta<0?"xdown":"xflat")+'">'+ppv(x.w.delta)+'</td></tr>').join('')+
    '</tbody></table></div>';
}

/* ==========================================================================
   HALAMAN Reports & Settings
   ========================================================================== */
function renderReports(){
  return '<div class="bcrumb">Reports</div><div class="grid">'+
    '<div class="c6">'+xcard("rp1","Unduh data","Berkas siap dibuka di Excel",
      '<div style="display:flex;flex-direction:column;gap:10px">'+
      '<button class="tbtn" data-fact="csvAll"><i>⤓</i>Data lengkap sesuai filter (CSV)</button>'+
      '<button class="tbtn" data-fact="csvNat"><i>⤓</i>National summary per indikator (CSV)</button>'+
      '<button class="tbtn" data-fact="print"><i>⎙</i>Cetak dashboard</button></div>')+'</div>'+
    '<div class="c6">'+xcard("rp2","Berkas untuk repository","Commit ke folder data/",
      '<div style="display:flex;flex-direction:column;gap:10px">'+
      '<button class="tbtn" data-fact="dlAsumsi"><i>⤒</i>asumsi.js</button>'+
      '<button class="tbtn" data-fact="dlPemetaan"><i>⤒</i>pemetaan.js</button>'+
      '<button class="tbtn" data-fact="dlIndicators"><i>⤒</i>indicators.js</button></div>',
      "Perubahan pada Assumptions dan Mapping baru permanen setelah file-nya di-commit.")+'</div>'+
  '</div>';
}
function renderSettings(){
  const c=checks();
  const chip=(lab,ok,val)=>'<div class="ins-i"><span class="ins-d '+(ok?"t-ok":"t-bad")+'">'+(ok?"✓":"!")+
    '</span><div class="ins-t"><b>'+esc(lab)+'</b><br><span class="dim">'+esc(val)+'</span></div></div>';
  return '<div class="bcrumb">Settings <b>·</b> Admin</div>'+
  '<div class="grid">'+
    '<div class="c6">'+xcard("st1","Integritas data","Pemeriksaan teknis, disembunyikan dari dashboard utama",
      chip("Row ID",c.rowid==="ROW ID OK",c.rowid)+
      chip("Sinkronisasi baris",c.sinkron==="SINKRON",c.sinkron)+
      chip("Duplikat AP × indikator",c.comboDupe===0,c.dupe+" · "+n0(c.dupeRows)+" baris")+
      chip("Proporsi dalam 0–100%",c.oor===0,c.range)+
      chip("Numerator ≤ denominator",c.nd===0,c.numden))+'</div>'+
    '<div class="c6">'+xcard("st2","Sumber data",
      S.src==="csv"?"Otomatis dari pipeline Power Automate":"Berkas statis di repository",
      '<div class="ins">'+
      '<div class="ins-i"><span class="ins-d '+(S.src==="csv"?"t-ok":"t-gray")+'">◈</span><div class="ins-t">'+
        '<b>'+(S.src==="csv"?"data/indicators.csv":"data/indicators.js")+'</b><br>'+
        '<span class="dim">'+n0(S.rows.length)+' baris'+(S.src==="csv"&&S.csv?
        ' · pemisah '+esc(S.csv.delim)+' · desimal '+esc(S.csv.dec):'')+'</span></div></div>'+
      '<div class="ins-i"><span class="ins-d t-gray">◉</span><div class="ins-t"><b>Versi</b><br>'+
        '<span class="dim">'+esc(CFG.version)+' · data per '+esc(CFG.data_date)+'</span></div></div>'+
      (S.local?'<div class="ins-i"><span class="ins-d t-warn">!</span><div class="ins-t">'+
        '<b>Ada perubahan lokal</b><br><span class="dim">Belum di-commit ke repository</span></div></div>':'')+
      '</div>',
      '<div style="display:flex;gap:9px;flex-wrap:wrap"><button class="tbtn" data-fact="import"><i>⤓</i>Impor data</button>'+
      '<button class="tbtn" data-fact="reload"><i>↺</i>Muat ulang data repo</button></div>')+'</div>'+
  '</div>';
}
/* ==========================================================================
   v4 — SHELL: sidebar, topbar, routing
   ========================================================================== */
const PAGES=[
 {id:"dashboard",   ic:"◫", nm:"Dashboard",   t:"National Target Dashboard", s:"Indonesia National Summary", r:renderDash},
 {id:"summary",     ic:"▤", nm:"List Indicator AP", t:"List Indicator AP",   s:"Tabel lengkap seluruh baris per Area Programme", r:()=>legacy(renderSummary)},
 {id:"reports",     ic:"⤓", nm:"Reports",     t:"Reports",                   s:"Unduhan dan cetak",              r:renderReports},
 {id:"settings",    ic:"⚙", nm:"Settings",    t:"Settings",                  s:"Sumber data dan pemeriksaan teknis", r:renderSettings}
];
let PG="dashboard";
const legacy = fn => '<div class="legacy">'+fn()+'</div>';

function shell(){
  document.getElementById("app").innerHTML=
  '<div class="shell">'+
    '<aside class="side" id="side">'+
      '<div class="side-top"><span class="side-logo" id="lgBox">WV</span>'+
        '<span class="side-org"><b>Wahana Visi Indonesia</b><em>AIM+ Target Setting</em></span></div>'+
      '<nav class="nav">'+
        '<div class="nav-lab">Analytics</div>'+
        PAGES.slice(0,5).map(p=>navlink(p)).join('')+
        '<div class="nav-lab">Tools</div>'+
        PAGES.slice(5).map(p=>navlink(p)).join('')+
      '</nav>'+
      '<div class="side-foot"><button data-side><span style="font-style:normal">⇤</span><em class="nav-txt">Perkecil menu</em></button></div>'+
    '</aside>'+
    '<div class="mainw">'+
      '<header class="top">'+
        '<button class="iconbtn" data-side style="display:none" id="burger">☰</button>'+
        '<div class="top-mid"><h1 id="pgT"></h1><div class="subt" id="pgS"></div></div>'+
        '<div class="top-right">'+
          '<div class="upd"><b>Last updated</b><em id="upd"></em></div>'+
          '<button class="iconbtn" data-fact="import" title="Impor data">⤓</button>'+
          '<button class="iconbtn" data-go="settings" title="Settings">⚙</button>'+
          '<span class="avatar" title="'+esc(CFG.owner||"")+'">WV</span>'+
        '</div>'+
      '</header>'+
      '<main class="page" id="pg"></main>'+
    '</div>'+
  '</div>';
  document.getElementById("upd").textContent=CFG.data_date||"";
  logo();
}
const navlink = p => '<a href="#" data-go="'+p.id+'"><i>'+p.ic+'</i><span class="nav-txt">'+p.nm+'</span></a>';

function logo(){
  const b=document.getElementById("lgBox"); if(!b)return;
  const tries=["assets/logo.svg","assets/logo.png","logo.svg","logo.png"]; let i=0;
  const img=document.createElement("img"); img.alt="World Vision";
  const t=()=>{ if(i>=tries.length) return;
    img.onload=()=>{ b.innerHTML=""; b.appendChild(img); };
    img.onerror=()=>{ i++; t(); }; img.src=tries[i]; };
  t();
}
function go(id){
  const p=PAGES.find(x=>x.id===id); if(!p)return;
  PG=id;
  document.getElementById("pgT").textContent=p.t;
  document.getElementById("pgS").textContent=p.s;
  document.querySelectorAll(".nav a").forEach(a=>a.classList.toggle("on",a.dataset.go===id));
  document.getElementById("side").classList.remove("show");
  paint();
  window.scrollTo({top:0,behavior:"instant"});
}
function paint(){
  const p=PAGES.find(x=>x.id===PG); if(!p)return;
  document.getElementById("pg").innerHTML=p.r();
  wireAll();
}
const repaint=paint;

/* ---------------- interaksi chart ----------------
   Tooltip kustom, bukan atribut title= bawaan: title lambat muncul dan tidak
   pernah jalan di perangkat sentuh. Satu elemen tooltip dipakai ulang untuk
   semua chart, dipasang lewat delegasi di elemen svg. */
let CHTIP=null;
function chartInteract(){
  if(!CHTIP){
    CHTIP=document.createElement("div");
    CHTIP.className="chtip"; CHTIP.setAttribute("role","status");
    document.body.appendChild(CHTIP);
  }
  const hide=()=>CHTIP.classList.remove("on");
  document.querySelectorAll("svg.ch").forEach(sv=>{
    if(sv.dataset.wired) return;
    sv.dataset.wired="1";
    const show=e=>{
      const g=e.target.closest?e.target.closest("g.chr"):null;
      if(!g){ hide(); return; }
      CHTIP.textContent=g.getAttribute("data-tip")||"";
      CHTIP.classList.add("on");
      const pad=14, w=CHTIP.offsetWidth, h=CHTIP.offsetHeight;
      let x=e.clientX+pad, y=e.clientY-h-8;
      if(x+w>window.innerWidth-8) x=e.clientX-w-pad;
      if(y<8) y=e.clientY+pad;
      CHTIP.style.left=x+"px"; CHTIP.style.top=y+"px";
    };
    sv.addEventListener("mousemove",show);
    sv.addEventListener("mouseleave",hide);
    sv.addEventListener("touchstart",ev=>{
      const t=ev.touches[0];
      show({target:document.elementFromPoint(t.clientX,t.clientY),
            clientX:t.clientX, clientY:t.clientY});
    },{passive:true});
    sv.classList.add("chgo");
  });
  window.addEventListener("scroll",hide,{passive:true});
}

/* ---------------- wiring ---------------- */
function wireAll(){
  const root=document.getElementById("pg");
  chartInteract();
  document.querySelectorAll("[data-go]").forEach(a=>a.onclick=e=>{e.preventDefault();go(a.dataset.go);});
  document.querySelectorAll("[data-side]").forEach(b=>b.onclick=()=>{
    const s=document.getElementById("side");
    if(window.innerWidth<=900) s.classList.toggle("show"); else s.classList.toggle("mini");
  });
  /* filter */
  root.querySelectorAll("[data-fperiod]").forEach(b=>b.onclick=()=>{F.period=b.dataset.fperiod;paint();});
  root.querySelectorAll("[data-fact]").forEach(b=>b.onclick=()=>fact(b.dataset.fact));
  document.querySelectorAll("[data-fact]").forEach(b=>b.onclick=()=>fact(b.dataset.fact));
  root.querySelectorAll("[data-dtoggle]").forEach(b=>b.onclick=e=>{
    e.stopPropagation();
    const d=b.closest(".ddl"), open=d.classList.contains("open");
    root.querySelectorAll(".ddl").forEach(x=>x.classList.remove("open"));
    if(!open) d.classList.add("open");
  });
  root.querySelectorAll("[data-fset]").forEach(c=>c.onchange=()=>{
    const f=c.dataset.fset, v=c.value, i=F[f].indexOf(v);
    c.checked ? (i<0&&F[f].push(v)) : (i>=0&&F[f].splice(i,1));
    if(f==="zonal") F.ap=F.ap.filter(a=>!F.zonal.length||F.zonal.indexOf(S.apz[a])>=0);
    if(f==="outcome") F.ind=F.ind.filter(x=>!F.outcome.length||F.outcome.indexOf((CAT_BY_IND[x]||{}).oc)>=0);
    paint();
  });
  root.querySelectorAll("[data-fdel]").forEach(b=>b.onclick=()=>{
    const f=b.dataset.fdel, v=b.dataset.fval, i=F[f].indexOf(v);
    if(i>=0) F[f].splice(i,1); paint();
  });
  const q=root.querySelector("[data-fq]");
  if(q) q.oninput=()=>{ clearTimeout(q._t); q._t=setTimeout(()=>{F.q=q.value;paint();},260); };
  /* menu kartu */
  root.querySelectorAll("[data-mtoggle]").forEach(b=>b.onclick=e=>{
    e.stopPropagation();
    const m=b.closest(".xmenu"), open=m.classList.contains("open");
    root.querySelectorAll(".xmenu").forEach(x=>x.classList.remove("open"));
    if(!open) m.classList.add("open");
  });
  root.querySelectorAll("[data-mact]").forEach(b=>b.onclick=()=>{
    if(b.dataset.mact==="print") window.print(); else csvNational();
    root.querySelectorAll(".xmenu").forEach(x=>x.classList.remove("open"));
  });
  document.body.onclick=()=>{
    document.querySelectorAll(".ddl.open,.xmenu.open").forEach(x=>x.classList.remove("open"));
  };
  /* halaman legacy tetap memakai wiring lamanya */
  if(["summary","analysis","mapping","assumptions"].indexOf(PG)>=0) wire(PG,root);
}
function fact(a){
  if(a==="toggle"){ F.fbOpen=!F.fbOpen; paint(); return; }
  if(a==="reset"){ F.zonal=[];F.ap=[];F.outcome=[];F.ind=[];F.status=[];F.period="Both";F.q="";paint(); return; }
  if(a==="print"){ window.print(); return; }
  if(a==="import"){ document.getElementById("scrimImport").classList.add("on"); return; }
  if(a==="reload"){ clearLocal(); location.reload(); return; }
  if(a==="csvNat") return csvNational();
  if(a==="csvAll") return dlCsv();
  if(a==="dlAsumsi")     return dl("asumsi.js",fileAsumsi(),"text/javascript");
  if(a==="dlPemetaan")   return dl("pemetaan.js",filePemetaan(),"text/javascript");
  if(a==="dlIndicators") return dl("indicators.js",fileIndicators(),"text/javascript");
}
function csvNational(){
  const per=perInd(natRows());
  const cols=["Outcome","Code","Indicator","Baseline Numerator","Baseline Denominator","Baseline (%)",
    "Evaluation Numerator","Evaluation Denominator","Evaluation (%)","Delta (pp)","Threshold (%)",
    "Achievement (%)","AP Baseline","AP Evaluation"];
  const q=v=>'"'+String(v==null?"":v).replace(/"/g,'""')+'"';
  const body=per.map(x=>[x.oc,x.code,x.ind,x.w.dB?x.w.nB:"",x.w.dB||"",
    x.w.pB===null?"":(x.w.pB*100).toFixed(4), x.w.dE?x.w.nE:"", x.w.dE||"",
    x.w.pE===null?"":(x.w.pE*100).toFixed(4),
    x.w.delta===null?"":(x.w.delta*100).toFixed(4), x.thr===null?"":(x.thr*100).toFixed(4),
    x.ach===null?"":(x.ach*100).toFixed(1), x.w.apB, x.w.apE].map(q).join(","));
  dl("national_summary.csv","\ufeff"+[cols.map(q).join(",")].concat(body).join("\r\n"),"text/csv");
}

/* ---------------- gate ---------------- */
const GK2="wvi_gate_v4";
function showGate(next){
  const g=document.createElement("div"); g.id="gate";
  g.innerHTML='<div class="box"><div class="lg">WV</div>'+
    '<h1>National Target Dashboard</h1>'+
    '<div class="cy">Wahana Visi Indonesia · '+esc(CFG.cycle)+'</div>'+
    '<label for="gi">Access code</label>'+
    '<input id="gi" type="password" placeholder="••••••••" autocomplete="off" spellcheck="false">'+
    '<button id="gb">Open dashboard</button><div class="er" id="ge"></div>'+
    '<div class="nt">Berkas kerja internal. Baseline dan evaluation masih bergerak; '+
    'belum ada angka final, dan tidak untuk diedarkan di luar WVI.</div></div>';
  document.body.appendChild(g);
  const inp=g.querySelector("#gi"), er=g.querySelector("#ge");
  const sub=()=>{
    if(String(inp.value||"").trim().toLowerCase()===String(CFG.code).toLowerCase()){
      try{sessionStorage.setItem(GK2,"1");}catch(e){}
      g.remove(); next(); return;
    }
    er.textContent="Kode itu tidak membuka dashboard ini."; inp.value=""; inp.focus();
  };
  g.querySelector("#gb").onclick=sub;
  inp.onkeydown=e=>{ if(e.key==="Enter") sub(); };
  setTimeout(()=>inp.focus(),40);
}

/* ---------------- boot ---------------- */
async function boot(){
  if(!window.WVI_CONFIG||!window.WVI_INDICATORS||!window.WVI_ASUMSI||!window.WVI_PEMETAAN){
    document.getElementById("app").innerHTML='<div style="padding:40px;max-width:640px">'+
      '<h2 style="font-size:18px">Berkas data tidak termuat</h2>'+
      '<p style="color:#565D6D">Halaman ini membutuhkan data/config.js, data/indicators.js, '+
      'data/asumsi.js, dan data/pemetaan.js.</p></div>';
    return;
  }
  adopt();
  S.csv=await tryCsv();
  if(S.csv&&S.csv.rows){ S.rows=S.csv.rows.map(r=>Object.assign({},r)); S.src="csv"; }
  else { S.src="js"; if(S.csv&&S.csv.err) S.csvErr=S.csv.err; }
  recompute();
  const sv=loadLocal();
  if(sv&&sv.rows&&sv.rows.length){
    try{
      if(sv.asumsi) S.asumsi=sv.asumsi;
      if(sv.pemetaan) S.pemetaan=sv.pemetaan;
      if(sv.target_delta!=null) CFG.target_delta=sv.target_delta;
      const col=window.WVI_INDICATORS.columns;
      S.rows=sv.rows.map(o=>{const r={};col.forEach(k=>r[k]=o[k]);return r;});
      S.local=true; recompute();
    }catch(e){ S.local=false; }
  }
  F.q=""; F.fbOpen=false;
  let ok=false; try{ ok=sessionStorage.getItem(GK2)==="1"; }catch(e){}
  if(ok) open4(); else showGate(open4);
}
function open4(){
  shell();
  document.querySelectorAll("[data-close]").forEach(b=>b.onclick=()=>b.closest(".scrim").classList.remove("on"));
  document.querySelectorAll(".scrim").forEach(s=>s.onclick=e=>{ if(e.target===s) s.classList.remove("on"); });
  document.addEventListener("keydown",e=>{
    if(e.key==="Escape") document.querySelectorAll(".scrim.on").forEach(s=>s.classList.remove("on"));
  });
  const on=(id,fn)=>{const e=document.getElementById(id); if(e) e.onclick=fn;};
  on("impCheck",impParse); on("impGo",impApply);
  if(window.innerWidth<=900) document.getElementById("burger").style.display="grid";
  go("dashboard");
}
function chip(){ /* indikator sumber data dipindah ke halaman Settings */ }
document.addEventListener("DOMContentLoaded",boot);
