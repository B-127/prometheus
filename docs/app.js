"use strict";
/* Sri Lanka CCPI dashboard — all computation client-side from the workbook.
   Security: no inline scripts (CSP), no eval, all dynamic text inserted via
   textContent (no innerHTML with data), CSV/Excel exports sanitised against
   spreadsheet-formula injection. */

const BASE_ORDER = ["1952=100","2002=100","2006/07=100","2013=100","2021=100"];
const LATEST_BASE = "2021=100";
const AGG = new Set(["All Items","Non Food"]);
const HEAT_MONTHS = 36;

const ymKey = (y,m)=>`${y}-${String(m).padStart(2,"0")}`;
const fmtPct = v => (v>0?"+":"")+v.toFixed(1)+"%";
const toNum = v => { if(v===null||v===undefined||v==="") return null;
  const n=+(""+v).replace(/,/g,""); return isFinite(n)?n:null; };

let STATE={tab1:[],tab2:[],measure:"index",core:"head",ctype:"line",base:"__spliced__",
           cmpView:"index",catView:"index",heat:"yoy",grpView:"index",simView:"index",
           composites:[],updated:"—"};
let cgSeq=1;
const CG_HINT="Tick subgroups above, name the set, and press Add. Composites use official CCPI weights and are session-only \u2014 cleared on refresh.";
const CH={};

/* ---------- splice: identical to the verified Python reference ---------- */
function splice(points){
  const present=[...new Set(points.map(p=>p.base))].sort((a,b)=>BASE_ORDER.indexOf(a)-BASE_ORDER.indexOf(b));
  if(!present.length) return {series:[],scale:{}};
  const latest=present[present.length-1];
  const by={}; present.forEach(b=>by[b]={});
  points.forEach(p=>{ by[p.base][ymKey(p.year,p.month)]=p.value; });
  const scale={[latest]:1};
  for(let i=present.length-1;i>0;i--){
    const newer=present[i],older=present[i-1];
    const common=Object.keys(by[newer]).filter(k=>k in by[older]).sort();
    const step=common.length? by[newer][common[common.length-1]]/by[older][common[common.length-1]] : 1;
    scale[older]=scale[newer]*step;
  }
  const out={};
  points.forEach(p=>{ const k=ymKey(p.year,p.month);
    if(!(k in out)||BASE_ORDER.indexOf(p.base)>BASE_ORDER.indexOf(out[k].base))
      out[k]={base:p.base,val:p.value*scale[p.base],y:p.year,m:p.month}; });
  const series=Object.keys(out).sort().map(k=>({key:k,...out[k]}));
  return {series,scale};
}
function overlapMonths(points){
  const seen={};
  points.forEach(p=>{ const k=ymKey(p.year,p.month); (seen[k]=seen[k]||new Set()).add(p.base); });
  return new Set(Object.keys(seen).filter(k=>seen[k].size>1));
}
// merge sorted month keys into contiguous [start,end] runs (no calendar gaps)
function contiguousRanges(keys){
  const r=[]; let start=null,prev=null;
  keys.forEach(k=>{ if(start===null){start=k;} else if(prevYm(k)!==prev){ r.push([start,prev]); start=k; } prev=k; });
  if(start!==null) r.push([start,prev]); return r;
}
// visible amber bands for multi-base overlap spans (start!==end so they have width)
function overlapBands(){
  const ov=[...overlapMonths(ccpiPoints("CCPI Index"))].sort();
  return contiguousRanges(ov).map(([a,b])=>[{name:"multi-base overlap",xAxis:a},{xAxis:b===a?a:b}]);
}
const rowsOf=(wb,s)=> wb.Sheets[s]? XLSX.utils.sheet_to_json(wb.Sheets[s],{defval:null}) : [];
function ccpiPoints(col){
  return STATE.tab1.filter(r=>toNum(r[col])!=null&&r["Base Year Used"]&&r.Year&&r["Month Number"])
    .map(r=>({year:+r.Year,month:+r["Month Number"],base:r["Base Year Used"],value:toNum(r[col])}));
}
/* group -> spliced series from Tab2 Monthly Index Number */
function groupSeries(g){
  const pts=STATE.tab2.filter(r=>r["Period Type"]==="Monthly"&&r.Group===g&&toNum(r["Index Number"])!=null)
    .map(r=>({year:+r.Year,month:+r["Month No"],base:r["Base Year Used"],value:toNum(r["Index Number"])}));
  return splice(pts).series;
}
function groupList(){
  return [...new Set(STATE.tab2.filter(r=>r["Period Type"]==="Monthly"&&!AGG.has(r.Group)).map(r=>r.Group))].sort();
}
function latestT2(){
  const m=STATE.tab2.filter(r=>r["Period Type"]==="Monthly"&&r["Base Year Used"]===LATEST_BASE);
  if(!m.length) return {ym:null,rows:[]};
  const last=m.map(r=>ymKey(+r.Year,+r["Month No"])).sort().pop();
  return {ym:last,rows:m.filter(r=>ymKey(+r.Year,+r["Month No"])===last)};
}

/* ---------- metric helpers: Index / Y-o-Y / M-o-M / YTD (yearly avg) ----------
   Rates come from the workbook's own published columns (Tab1 decimals ->%, Tab2
   already %), collapsed to the newest base per month for continuity. YTD = the
   12-month / annual-average inflation column. */
const MONTHS=["January","February","March","April","May","June","July","August","September","October","November","December"];
const T1COLS={yoy:{head:"CCPI YoY Inflation",core:"CCPI Core YoY Inflation"},
  mom:{head:"CCPI MoM Change",core:"CCPI Core MoM Change"},
  ytd:{head:"CCPI 12M Moving Avg Inflation",core:"CCPI Core 12M Moving Avg Inflation"}};
const T2COLS={yoy:"Year-on-Year Inflation (%)",mom:"Month-to-Month Inflation (%)",ytd:"Annual Average Inflation (%)"};
const METRIC_UNIT=m=>m==="index"?"":"%";
const METRIC_LABEL=m=>({index:"Index",yoy:"Y-o-Y %",mom:"M-o-M %",ytd:"YTD %"}[m]);
let META={months:[],baseChanges:[],overlap:new Set(),heatMonths:[]};
function buildMeta(){
  const {series}=splice(ccpiPoints("CCPI Index"));
  META.months=series.map(s=>s.key);
  META.baseChanges=[];
  for(let i=1;i<series.length;i++) if(series[i].base!==series[i-1].base) META.baseChanges.push(series[i].key);
  META.overlap=overlapMonths(ccpiPoints("CCPI Index"));
  META.heatMonths=[...new Set(STATE.tab2.filter(r=>r["Period Type"]==="Monthly"&&r["Base Year Used"]===LATEST_BASE)
    .map(r=>ymKey(+r.Year,+r["Month No"])))].sort();
}
function collapseNewest(rows){
  const seen={}; rows.forEach(r=>{ if(!(r.k in seen)||BASE_ORDER.indexOf(r.base)>BASE_ORDER.indexOf(seen[r.k].base)) seen[r.k]=r; });
  return Object.keys(seen).sort().map(k=>({key:k,v:seen[k].v}));
}
function headlineSeries(metric,core,base){
  if(metric==="index"){
    const col=core?"CCPI Core Index":"CCPI Index"; const pts=ccpiPoints(col);
    if(base==="__spliced__") return splice(pts).series.map(s=>({key:s.key,v:s.val}));
    return pts.filter(p=>p.base===base).map(p=>({key:ymKey(p.year,p.month),v:p.value})).sort((a,b)=>a.key<b.key?-1:1);
  }
  const col=T1COLS[metric][core?"core":"head"];
  return collapseNewest(STATE.tab1.filter(r=>toNum(r[col])!=null&&(base==="__spliced__"||r["Base Year Used"]===base))
    .map(r=>({k:ymKey(+r.Year,+r["Month Number"]),v:toNum(r[col])*100,base:r["Base Year Used"]})));
}
function groupMetricSeries(g,metric){
  if(metric==="index") return groupSeries(g).map(s=>({key:s.key,v:s.val}));
  const col=T2COLS[metric];
  return collapseNewest(STATE.tab2.filter(r=>r["Period Type"]==="Monthly"&&r.Group===g&&toNum(r[col])!=null)
    .map(r=>({k:ymKey(+r.Year,+r["Month No"]),v:toNum(r[col]),base:r["Base Year Used"]})));
}
const inRange=(k,from,to)=>(!from||k>=from)&&(!to||k<=to);
function baseChangeMarks(keysSet){ return META.baseChanges.filter(k=>keysSet.has(k)).map(k=>({xAxis:k})); }
const rangeVals=(fromId,toId)=>[document.getElementById(fromId).value,document.getElementById(toId).value];

/* Current date in Sri Lanka (Asia/Colombo), computed at view time. */
function colomboDate(){
  try{
    return new Intl.DateTimeFormat("en-GB",{timeZone:"Asia/Colombo",
      weekday:"short",day:"2-digit",month:"short",year:"numeric"}).format(new Date());
  }catch(e){ return new Date().toDateString(); }
}
/* Published headline anchors for the scenario simulation. Index levels come from
   the spliced continuous CCPI series (authoritative); rates come from the
   workbook's own published columns, so every figure matches the topline KPIs. */
function stepBack(k,n){ let [y,m]=k.split("-").map(Number);
  for(let i=0;i<n;i++){ m--; if(m===0){ m=12; y--; } } return ymKey(y,m); }
function publishedHeadline(){
  const series=splice(ccpiPoints("CCPI Index")).series;
  const idx={}; series.forEach(s=>{ idx[s.key]=s.val; });
  const lastKey=series.length?series[series.length-1].key:null;
  const curIdx=lastKey?idx[lastKey]:null;
  const prevIdx=lastKey?idx[stepBack(lastKey,1)]:null;
  const yearAgoIdx=lastKey?idx[stepBack(lastKey,12)]:null;
  let meanPrev12=null;                                   // mean of months t-23..t-12
  if(lastKey){ let sum=0,ok=true;
    for(let i=12;i<24;i++){ const v=idx[stepBack(lastKey,i)]; if(v==null){ ok=false; break; } sum+=v; }
    if(ok) meanPrev12=sum/12; }
  const t1=STATE.tab1.filter(r=>r["Base Year Used"]===LATEST_BASE&&toNum(r["CCPI Index"])!=null)
    .sort((a,b)=>(+a.Year-+b.Year)||(+a["Month Number"]-+b["Month Number"]));
  const last=t1[t1.length-1]||{};
  return {curIdx,prevIdx,yearAgoIdx,meanPrev12,
    mom:toNum(last["CCPI MoM Change"]),yoy:toNum(last["CCPI YoY Inflation"]),
    ytd:toNum(last["CCPI 12M Moving Avg Inflation"])};
}

/* ------------------------------ KPIs ------------------------------ */
function renderKpis(){
  const t1=STATE.tab1.filter(r=>r["Base Year Used"]===LATEST_BASE&&toNum(r["CCPI Index"])!=null)
    .sort((a,b)=>(+a.Year-+b.Year)||(+a["Month Number"]-+b["Month Number"]));
  const last=t1[t1.length-1]; if(!last) return;
  const set=(id,txt)=>{document.getElementById(id).textContent=txt;};
  set("kIndex",toNum(last["CCPI Index"]).toFixed(1));
  const yoy=toNum(last["CCPI YoY Inflation"]), mom=toNum(last["CCPI MoM Change"]);
  const ma=toNum(last["CCPI 12M Moving Avg Inflation"]), cyoy=toNum(last["CCPI Core YoY Inflation"]);
  const kyoy=document.getElementById("kYoY");
  kyoy.textContent=yoy==null?"—":fmtPct(yoy*100);
  kyoy.style.color=yoy==null?"":(yoy<0?"var(--good)":yoy>0.08?"var(--bad)":"var(--warn)");
  set("kMA","12M avg "+(ma==null?"—":fmtPct(ma*100)));
  set("kMoM",mom==null?"—":fmtPct(mom*100));
  set("kCore",toNum(last["CCPI Core Index"])!=null?toNum(last["CCPI Core Index"]).toFixed(1):"—");
  set("kCoreYoY","core Y-o-Y "+(cyoy==null?"—":fmtPct(cyoy*100)));
  set("fMonth",ymKey(last.Year,last["Month Number"]));
  set("kpiMonth",MONTHS[(+last["Month Number"])-1]+" "+last.Year);
  set("fUpdated",colomboDate());
}

/* ------------------------------ trend ------------------------------ */
function drawTrend(){
  const [from,to]=rangeVals("trendFrom","trendTo");
  const core=STATE.core==="core", metric=STATE.measure;
  const arr=headlineSeries(metric,core,STATE.base).filter(p=>inRange(p.key,from,to));
  const xs=arr.map(p=>p.key), ys=arr.map(p=>+p.v.toFixed(2));
  const keysSet=new Set(xs);
  const marks=(STATE.base==="__spliced__")?baseChangeMarks(keysSet):[];
  let areas=[];
  if(STATE.base==="__spliced__"&&metric==="index")
    areas=contiguousRanges([...META.overlap].filter(k=>keysSet.has(k)).sort()).map(([a,b])=>[{xAxis:a},{xAxis:b}]);
  const isArea=STATE.ctype==="area", isBar=STATE.ctype==="bar", unit=METRIC_UNIT(metric);
  CH.trend.setOption({grid:{left:52,right:16,top:16,bottom:28},
    tooltip:{trigger:"axis",valueFormatter:v=>v==null?"—":v.toFixed(2)+unit},
    xAxis:{type:"category",data:xs,axisLabel:{fontSize:10,color:"#7a869a"}},
    yAxis:{type:"value",scale:true,splitLine:{lineStyle:{color:"#eef2f7"}},axisLabel:{fontSize:10,color:"#7a869a"}},
    series:[{type:isBar?"bar":"line",data:ys,smooth:!isBar,showSymbol:false,sampling:"lttb",
      lineStyle:{width:2.2,color:"#1f6feb"},itemStyle:{color:"#1f6feb"},
      areaStyle:(isArea?{color:"rgba(31,111,235,.10)"}:null),
      markLine:{symbol:"none",silent:true,lineStyle:{color:"#b0b8c4",type:"dashed"},
        label:{show:false},data:marks},
      markArea:{silent:true,itemStyle:{color:"rgba(31,111,235,.14)"},
        label:{show:false},data:areas}}]},true);
  CH.trend.__name=(core?"Core ":"")+"CCPI "+METRIC_LABEL(metric);
  CH.trend.__rows=arr.map(p=>({Month:p.key,[CH.trend.__name]:+p.v.toFixed(2)}));
}

/* ------------------------------ compare ------------------------------ */
function fillCmp(){
  const box=document.getElementById("cmpChk"); if(box.dataset.filled) return;
  const defaults=new Set(["Food and Non-Alcoholic Beverages","Transport","Housing, Water, Electricity, Gas and Other Fuels"]);
  groupList().forEach(g=>{
    const lab=document.createElement("label"); const cb=document.createElement("input");
    cb.type="checkbox"; cb.value=g; cb.checked=defaults.has(g); cb.addEventListener("change",drawCompare);
    lab.appendChild(cb); lab.appendChild(document.createTextNode(g)); box.appendChild(lab);
  });
  box.dataset.filled="1";
}
function drawCompare(){
  const [from,to]=rangeVals("cmpFrom","cmpTo");
  const chosen=[...document.querySelectorAll("#cmpChk input:checked")].map(c=>c.value);
  const metric=STATE.cmpView, unit=METRIC_UNIT(metric);
  const series=[]; const xset=new Set();
  chosen.forEach(g=>{
    const data=groupMetricSeries(g,metric).filter(p=>inRange(p.key,from,to)).map(p=>[p.key,+p.v.toFixed(2)]);
    data.forEach(d=>xset.add(d[0]));
    series.push({name:g,type:"line",showSymbol:false,smooth:true,sampling:"lttb",data});
  });
  (STATE.composites||[]).filter(c=>c.sel).forEach(c=>{
    const data=compositeMetricSeries(c.members,metric).filter(p=>inRange(p.key,from,to)).map(p=>[p.key,+p.v.toFixed(2)]);
    data.forEach(d=>xset.add(d[0]));
    series.push({name:c.name,type:"line",showSymbol:false,smooth:true,sampling:"lttb",lineStyle:{width:3},data});
  });
  const xs=[...xset].sort(), keysSet=new Set(xs);
  if(series.length) series[0].markLine={symbol:"none",silent:true,
    lineStyle:{color:"#b0b8c4",type:"dashed"},label:{show:false},data:baseChangeMarks(keysSet)};
  CH.compare.setOption({grid:{left:52,right:16,top:24,bottom:28},
    tooltip:{trigger:"axis",valueFormatter:v=>v==null?"—":v.toFixed(2)+unit},
    legend:{type:"scroll",top:0,textStyle:{fontSize:10}},
    xAxis:{type:"category",data:xs,axisLabel:{fontSize:10,color:"#7a869a"}},
    yAxis:{type:"value",scale:true,splitLine:{lineStyle:{color:"#eef2f7"}},axisLabel:{fontSize:10,color:"#7a869a"}},
    series},true);
  const wide={}; series.forEach(s=>s.data.forEach(([k,v])=>{(wide[k]=wide[k]||{Month:k})[s.name]=v;}));
  CH.compare.__name="Compare "+METRIC_LABEL(metric);
  CH.compare.__rows=xs.map(k=>wide[k]||{Month:k});
}

/* ---- custom "group of groups" composites (official-weighted, base-spliced) ---- */
// {group:{base:{monthKey:index}}} + weights {group:{base:weight}}, built once
function t2Lookup(){
  if(window.__t2look) return window.__t2look;
  const idx={}, wt={};
  STATE.tab2.forEach(r=>{
    if(r["Period Type"]!=="Monthly") return;
    const g=r.Group, b=r["Base Year Used"]; if(g==null||b==null) return;
    const iv=toNum(r["Index Number"]), wv=toNum(r["Weight (%)"]);
    if(iv!=null){ (idx[g]=idx[g]||{}); (idx[g][b]=idx[g][b]||{})[ymKey(+r.Year,+r["Month No"])]=iv; }
    if(wv!=null){ (wt[g]=wt[g]||{}); if(wt[g][b]==null) wt[g][b]=wv; }
  });
  return window.__t2look={idx,wt};
}
// weighted composite index per base (renormalised over members present that month), then spliced
function compositeIndexSeries(members){
  const {idx,wt}=t2Lookup(); const bases=new Set();
  members.forEach(g=>{ if(idx[g]) Object.keys(idx[g]).forEach(b=>bases.add(b)); });
  const pts=[];
  bases.forEach(b=>{
    const months=new Set();
    members.forEach(g=>{ if(idx[g]&&idx[g][b]) Object.keys(idx[g][b]).forEach(k=>months.add(k)); });
    months.forEach(k=>{
      let num=0,wsum=0;
      members.forEach(g=>{ const iv=idx[g]&&idx[g][b]?idx[g][b][k]:null, wv=wt[g]?wt[g][b]:null;
        if(iv!=null&&wv!=null){ num+=wv*iv; wsum+=wv; } });
      if(wsum>0){ const [y,m]=k.split("-").map(Number); pts.push({year:y,month:m,base:b,value:num/wsum}); }
    });
  });
  return splice(pts).series;
}
// derive Index / Y-o-Y / M-o-M / YTD from the composite index (no published rate exists)
function compositeMetricSeries(members,metric){
  const s=compositeIndexSeries(members);
  if(metric==="index") return s.map(p=>({key:p.key,v:p.val}));
  const map={}; s.forEach(p=>{ map[p.key]=p.val; });
  const out=[];
  s.forEach(p=>{ const k=p.key; let v=null;
    if(metric==="mom"){ const pk=stepBack(k,1); if(map[pk]!=null) v=(map[k]/map[pk]-1)*100; }
    else if(metric==="yoy"){ const pk=stepBack(k,12); if(map[pk]!=null) v=(map[k]/map[pk]-1)*100; }
    else { let a=0,b=0,ok=true;
      for(let i=0;i<12&&ok;i++){ const vv=map[stepBack(k,i)]; if(vv==null)ok=false; else a+=vv; }
      for(let i=12;i<24&&ok;i++){ const vv=map[stepBack(k,i)]; if(vv==null)ok=false; else b+=vv; }
      if(ok&&b>0) v=((a/12)/(b/12)-1)*100; }
    if(v!=null) out.push({key:k,v});
  });
  return out;
}
function fillCmpMembers(){
  const box=document.getElementById("cgMembers"); if(box.dataset.filled) return;
  groupList().forEach(g=>{
    const lab=document.createElement("label");
    const cb=document.createElement("input"); cb.type="checkbox"; cb.value=g;
    cb.addEventListener("change",updateCgCount);
    lab.appendChild(cb); lab.appendChild(document.createTextNode(g)); box.appendChild(lab);
  });
  box.dataset.filled="1"; updateCgCount();
}
function updateCgCount(){
  const n=document.querySelectorAll("#cgMembers input:checked").length;
  const el=document.getElementById("cgCount"); if(el) el.textContent=n+" selected";
}
function renderComposites(){
  const box=document.getElementById("cgList"); box.textContent="";
  if(!STATE.composites.length){ const em=document.createElement("span");
    em.className="hint"; em.style.margin="0"; em.textContent="No custom groups yet."; box.appendChild(em); return; }
  STATE.composites.forEach(c=>{
    const card=document.createElement("div"); card.className="cgcard";
    const cb=document.createElement("input"); cb.type="checkbox"; cb.checked=c.sel; cb.dataset.cid=c.id;
    cb.setAttribute("aria-label","Show "+c.name+" on chart");
    const body=document.createElement("div"); body.className="cgcard-body";
    const nm=document.createElement("div"); nm.className="cgcard-name"; nm.textContent=c.name;
    const nn=document.createElement("span"); nn.className="cgcard-n";
    nn.textContent=" \u00b7 "+c.members.length+(c.members.length===1?" group":" groups");
    nm.appendChild(nn);
    const mem=document.createElement("div"); mem.className="cgcard-members"; mem.textContent=c.members.join(", ");
    body.appendChild(nm); body.appendChild(mem);
    const x=document.createElement("button"); x.type="button"; x.className="cgx"; x.dataset.del=c.id;
    x.textContent="\u00d7"; x.title="Delete "+c.name; x.setAttribute("aria-label","Delete "+c.name);
    card.appendChild(cb); card.appendChild(body); card.appendChild(x); box.appendChild(card);
  });
}
function addComposite(){
  const boxes=[...document.querySelectorAll("#cgMembers input:checked")];
  const members=boxes.map(b=>b.value);
  const nameEl=document.getElementById("cgName"), hint=document.getElementById("cgHint");
  let name=(nameEl.value||"").trim().replace(/\s+/g," ");
  if(!members.length){ hint.textContent="Tick at least one subgroup above to combine."; return; }
  if(!name){ hint.textContent="Give the group a name before adding."; return; }
  const taken=new Set([...STATE.composites.map(c=>c.name.toLowerCase()),...groupList().map(g=>g.toLowerCase())]);
  if(taken.has(name.toLowerCase())){ let n=2; const base=name; while(taken.has((base+" ("+n+")").toLowerCase())) n++; name=base+" ("+n+")"; }
  STATE.composites.push({id:"cg"+(cgSeq++),name,members,sel:true});
  nameEl.value=""; boxes.forEach(b=>{ b.checked=false; }); updateCgCount();
  hint.textContent=CG_HINT; renderComposites(); drawCompare();
}
function deleteComposite(id){ STATE.composites=STATE.composites.filter(c=>c.id!==id); renderComposites(); drawCompare(); }
function resetComposites(){
  STATE.composites=[];
  document.querySelectorAll("#cgMembers input:checked").forEach(b=>{ b.checked=false; }); updateCgCount();
  document.getElementById("cgHint").textContent=CG_HINT; renderComposites(); drawCompare();
}

/* ------------------------------ heatmap ------------------------------ */
function drawHeat(){
  const [from,to]=rangeVals("heatFrom","heatTo");
  const groups=groupList();
  const col=T2COLS[STATE.heat];
  const months=META.heatMonths.filter(k=>inRange(k,from,to));
  const rows=STATE.tab2.filter(r=>r["Period Type"]==="Monthly"&&r["Base Year Used"]===LATEST_BASE);
  const gi=Object.fromEntries(groups.map((g,i)=>[g,i])), mi=Object.fromEntries(months.map((m,i)=>[m,i]));
  const data=[]; let vmax=1;
  rows.forEach(r=>{ const k=ymKey(+r.Year,+r["Month No"]); const v=toNum(r[col]);
    if(r.Group in gi && k in mi && v!=null){ data.push([mi[k],gi[r.Group],+v.toFixed(1)]); vmax=Math.max(vmax,Math.abs(v)); }});
  CH.heat.setOption({grid:{left:8,right:16,top:8,bottom:56,containLabel:true},
    tooltip:{position:"top",formatter:p=>`${groups[p.value[1]]}<br/>${months[p.value[0]]}: ${p.value[2]}%`},
    xAxis:{type:"category",data:months,axisLabel:{fontSize:9,color:"#7a869a",rotate:60}},
    yAxis:{type:"category",data:groups,axisLabel:{fontSize:9,color:"#5b6b80",width:140,overflow:"truncate"}},
    visualMap:{min:-vmax,max:vmax,calculable:true,orient:"horizontal",left:"center",bottom:6,
      inRange:{color:["#146c43","#1a8a55","#3fa46e","#69bb8c","#9ed3b4","#d4ecdd",
        "#f4f6f8",
        "#f7d6cf","#eda192","#df6f5b","#cf4b39","#b5271b","#8f1d14"]},textStyle:{fontSize:10}},
    series:[{type:"heatmap",data,progressive:1000,itemStyle:{borderColor:"#fff",borderWidth:.5}}]},true);
  CH.heat.__name="Heatmap "+METRIC_LABEL(STATE.heat);
  CH.heat.__rows=data.map(d=>({Month:months[d[0]],Group:groups[d[1]],[METRIC_LABEL(STATE.heat)]:d[2]}));
}

/* ------------------------------ drivers ------------------------------ */
function components(){
  const {rows}=latestT2();
  const comp=rows.filter(r=>!AGG.has(r.Group)&&toNum(r["Weight (%)"])!=null&&toNum(r["Index Number"])!=null);
  const wsum=comp.reduce((s,r)=>s+toNum(r["Weight (%)"]),0)||1;
  return {comp,wsum};
}
function drawDrivers(){
  const {comp,wsum}=components();
  const arr=comp.map(r=>({name:r.Group,w:toNum(r["Weight (%)"])/wsum,
      yoy:toNum(r["Year-on-Year Inflation (%)"])})).filter(d=>d.yoy!=null)
    .map(d=>({name:d.name,c:d.w*d.yoy})).sort((a,b)=>a.c-b.c);
  const sum=arr.reduce((s,d)=>s+d.c,0);
  const allrow=latestT2().rows.find(r=>r.Group==="All Items");
  const head=allrow?toNum(allrow["Year-on-Year Inflation (%)"]):null;
  document.getElementById("drvHint").textContent=
    `Approx contribution = (weight ÷ Σweight) × Y-o-Y. Sum ≈ ${sum.toFixed(2)}% vs headline ${head==null?"—":head.toFixed(1)+"%"}.`;
  CH.drivers.setOption({grid:{left:8,right:26,top:8,bottom:8,containLabel:true},
    tooltip:{trigger:"axis",axisPointer:{type:"shadow"},valueFormatter:v=>v+" pp"},
    xAxis:{type:"value",axisLabel:{fontSize:10,color:"#7a869a"},splitLine:{lineStyle:{color:"#eef2f7"}}},
    yAxis:{type:"category",data:arr.map(d=>d.name),axisLabel:{fontSize:10,color:"#5b6b80",width:150,overflow:"truncate"}},
    series:[{type:"bar",data:arr.map(d=>+d.c.toFixed(3)),barWidth:"62%",
      itemStyle:{color:p=>p.value<0?"#1a8a55":"#de4940",borderRadius:[0,4,4,0]}}]},true);
  CH.drivers.__name="Inflation drivers";
  CH.drivers.__rows=arr.map(d=>({Group:d.name,"Contribution (pp)":+d.c.toFixed(3)}));
}

/* ------------------------------ scenario ------------------------------ */
function buildSim(){
  const box=document.getElementById("simRows"); box.textContent="";
  const {comp,wsum}=components();
  window.__head=publishedHeadline();
  window.__sim=comp.map(r=>({name:r.Group,w:toNum(r["Weight (%)"])/wsum,index:toNum(r["Index Number"]),shock:0}));
  window.__sim.forEach((c,i)=>{
    const row=document.createElement("div"); row.className="simrow";
    const lab=document.createElement("span"); lab.textContent=c.name; lab.title=c.name;
    lab.style.overflow="hidden"; lab.style.textOverflow="ellipsis"; lab.style.whiteSpace="nowrap";
    const inp=document.createElement("input"); inp.type="number"; inp.step="1"; inp.value="0";
    inp.setAttribute("aria-label",c.name+" percent change");
    inp.addEventListener("input",()=>{ const v=parseFloat(inp.value); window.__sim[i].shock=isFinite(v)?v:0; runSim(); });
    row.appendChild(lab); row.appendChild(inp); box.appendChild(row);
  });
  runSim();
}
const fmtPP = v => (v>0?"+":"")+v.toFixed(2)+" pp";
function runSim(){
  const sim=window.__sim||[];
  const H=window.__head||publishedHeadline();
  // Weighted reconstruction gives the RELATIVE shock; the level is anchored to the
  // published All Items index so the baseline always equals the topline (no 207.7/207.6 gap).
  const baseRecon=sim.reduce((s,c)=>s+c.w*c.index,0);
  const simRecon=sim.reduce((s,c)=>s+c.w*c.index*(1+c.shock/100),0);
  const ratio=baseRecon?simRecon/baseRecon:1;
  const curIdx=(H.curIdx!=null)?H.curIdx:baseRecon;   // published headline (anchor)
  const simIdx=curIdx*ratio;
  const dIdx=simIdx-curIdx;                            // change in index points
  const view=STATE.simView;
  const baseEl=document.getElementById("simBase");
  const newEl=document.getElementById("simNew");
  const delEl=document.getElementById("simDelta");
  const baseL=document.getElementById("simBaseL");
  let isPct=true, baseVal=0, newVal=0, changeDisp=0, changeTxt="—";
  if(view==="index"){
    isPct=false; baseL.textContent="Baseline index";
    baseVal=curIdx; newVal=simIdx;
    changeDisp=(ratio-1)*100; changeTxt=fmtPct(changeDisp);      // relative % change of the index
  } else if(view==="mom"){
    baseL.textContent="Baseline M-o-M";
    baseVal=(H.mom!=null?H.mom*100:0);                            // published MoM %
    newVal=baseVal+(H.prevIdx?dIdx/H.prevIdx*100:0);             // MoM is linear in current index
    changeDisp=newVal-baseVal; changeTxt=fmtPP(changeDisp);
  } else if(view==="yoy"){
    baseL.textContent="Baseline Y-o-Y";
    baseVal=(H.yoy!=null?H.yoy*100:0);                            // published YoY %
    newVal=baseVal+(H.yearAgoIdx?dIdx/H.yearAgoIdx*100:0);       // YoY is linear in current index
    changeDisp=newVal-baseVal; changeTxt=fmtPP(changeDisp);
  } else { // ytd = 12-month / annual-average inflation
    baseL.textContent="Baseline YTD";
    baseVal=(H.ytd!=null?H.ytd*100:0);                           // published 12M-avg %
    newVal=baseVal+(H.meanPrev12?(dIdx/12)/H.meanPrev12*100:0);  // current month is 1/12 of the trailing mean
    changeDisp=newVal-baseVal; changeTxt=fmtPP(changeDisp);
  }
  baseEl.textContent=isPct?fmtPct(baseVal):baseVal.toFixed(1);
  newEl.textContent =isPct?fmtPct(newVal) :newVal.toFixed(1);
  delEl.textContent=changeTxt;
  delEl.style.color=changeDisp<0?"var(--good)":changeDisp>0?"var(--bad)":"var(--muted)";
}

/* ------------------------------ anomaly ------------------------------ */
function prevYm(k){ let [y,m]=k.split("-").map(Number); m--; if(m===0){y--;m=12;} return ymKey(y,m); }
function drawAnom(){
  const {series}=splice(ccpiPoints("CCPI Index"));
  const idx={}; series.forEach(s=>idx[s.key]=s.val);
  const mom=[]; series.forEach(s=>{ const pk=prevYm(s.key);
    if(pk in idx) mom.push({k:s.key,v:(s.val/idx[pk]-1)*100}); });
  const vals=mom.map(x=>x.v).slice().sort((a,b)=>a-b);
  const med=vals.length?vals[Math.floor(vals.length/2)]:0;
  const mad=(()=>{ const dev=mom.map(x=>Math.abs(x.v-med)).sort((a,b)=>a-b);
    return dev.length?(dev[Math.floor(dev.length/2)]||1e-9):1e-9; })();
  const z=x=>0.6745*(x-med)/mad;
  const flagged=mom.map(x=>({...x,z:z(x.v)})).filter(x=>Math.abs(x.z)>3.5)
    .sort((a,b)=>Math.abs(b.z)-Math.abs(a.z));
  const keys=series.map(s=>s.key), gaps=[];
  for(let i=1;i<keys.length;i++) if(prevYm(keys[i])!==keys[i-1]) gaps.push(keys[i]);
  const tb=document.querySelector("#anom tbody"); tb.textContent="";
  const addRow=(c1,c2,c3,c4)=>{ const tr=document.createElement("tr");
    [c1,c2,c3,c4].forEach((c,j)=>{ const td=document.createElement("td"); td.textContent=c;
      if(j===0)td.style.textAlign="left"; tr.appendChild(td); }); tb.appendChild(tr); };
  flagged.slice(0,40).forEach(x=>addRow(x.k,fmtPct(x.v),x.z.toFixed(1),x.v>0?"spike":"drop"));
  gaps.forEach(g=>addRow(g,"—","—","missing prior month"));
  if(!flagged.length&&!gaps.length) addRow("—","—","—","no anomalies");
}

/* ------------------------------ category + subgroup ------------------------------ */
function drawCats(){
  const {ym,rows}=latestT2();
  document.getElementById("catHint").textContent=`Latest month ${ym||"—"} · aggregates excluded.`;
  const colMap={index:"Index Number",yoy:"Year-on-Year Inflation (%)",mom:"Month-to-Month Inflation (%)",ytd:"Annual Average Inflation (%)"};
  const col=colMap[STATE.catView];
  const data=rows.filter(r=>!AGG.has(r.Group)&&toNum(r[col])!=null)
    .map(r=>({name:r.Group,val:toNum(r[col])})).sort((a,b)=>a.val-b.val);
  CH.cats.setOption({grid:{left:8,right:26,top:8,bottom:8,containLabel:true},
    tooltip:{trigger:"axis",axisPointer:{type:"shadow"}},
    xAxis:{type:"value",axisLabel:{fontSize:10,color:"#7a869a"},splitLine:{lineStyle:{color:"#eef2f7"}}},
    yAxis:{type:"category",data:data.map(d=>d.name),axisLabel:{fontSize:10,color:"#5b6b80",width:150,overflow:"truncate"}},
    series:[{type:"bar",data:data.map(d=>+d.val.toFixed(2)),barWidth:"62%",
      itemStyle:{color:STATE.catView==="index"?"#1f6feb":"#de4940",borderRadius:[0,4,4,0]}}]},true);
  const ch=STATE.catView==="index"?"Index":METRIC_LABEL(STATE.catView);
  CH.cats.__name="Category "+ch;
  CH.cats.__rows=data.map(d=>({Group:d.name,[ch]:+d.val.toFixed(2)}));
}
function fillGroups(){
  const sel=document.getElementById("groupSel"); if(sel.dataset.filled) return;
  groupList().forEach(g=>{const o=document.createElement("option");o.value=g;o.textContent=g;sel.appendChild(o);});
  sel.dataset.filled="1";
  sel.value=groupList().includes("Food and Non-Alcoholic Beverages")?"Food and Non-Alcoholic Beverages":groupList()[0];
}
function drawGroup(){
  const g=document.getElementById("groupSel").value; if(!g) return;
  const [from,to]=rangeVals("grpFrom","grpTo");
  const metric=STATE.grpView, unit=METRIC_UNIT(metric);
  const arr=groupMetricSeries(g,metric).filter(p=>inRange(p.key,from,to));
  const xs=arr.map(p=>p.key), ys=arr.map(p=>+p.v.toFixed(metric==="index"?1:2));
  const keysSet=new Set(xs);
  CH.group.setOption({grid:{left:48,right:16,top:14,bottom:26},
    tooltip:{trigger:"axis",valueFormatter:v=>v==null?"—":v.toFixed(2)+unit},
    xAxis:{type:"category",data:xs,axisLabel:{fontSize:10,color:"#7a869a"}},
    yAxis:{type:"value",scale:true,splitLine:{lineStyle:{color:"#eef2f7"}},axisLabel:{fontSize:10,color:"#7a869a"}},
    series:[{type:"line",data:ys,smooth:true,showSymbol:false,sampling:"lttb",
      lineStyle:{width:2,color:"#1a8a55"},areaStyle:{color:"rgba(26,138,85,.07)"},
      markLine:{symbol:"none",silent:true,lineStyle:{color:"#b0b8c4",type:"dashed"},
        label:{show:false},data:baseChangeMarks(keysSet)}}]},true);
  const gh=g+" "+METRIC_LABEL(metric);
  CH.group.__name="Subgroup "+g;
  CH.group.__rows=arr.map(p=>({Month:p.key,[gh]:+p.v.toFixed(metric==="index"?1:2)}));
}
function drawTable(){
  const ov=overlapMonths(ccpiPoints("CCPI Index"));
  const {series}=splice(ccpiPoints("CCPI Index"));
  const tb=document.querySelector("#tbl tbody"); tb.textContent="";
  series.slice().reverse().forEach(s=>{ const tr=document.createElement("tr");
    [s.key,s.val.toFixed(1),s.base,ov.has(s.key)?"yes":""].forEach((c,j)=>{
      const td=document.createElement("td"); td.textContent=c; if(j===0)td.style.textAlign="left"; tr.appendChild(td); });
    tb.appendChild(tr); });
}

/* ------------------------------ export (injection-safe) ------------------------------ */
function csvCell(v){ let s=(v==null?"":String(v));
  if(/^[=+\-@\t\r]/.test(s)) s="'"+s;                 // neutralise formula injection
  if(/[",\n]/.test(s)) s='"'+s.replace(/"/g,'""')+'"'; return s; }
function download(name,blob){ const u=URL.createObjectURL(blob);
  const a=document.createElement("a"); a.href=u; a.download=name; a.click();
  setTimeout(()=>URL.revokeObjectURL(u),1500); }
function continuousRows(){
  const ov=overlapMonths(ccpiPoints("CCPI Index"));
  return splice(ccpiPoints("CCPI Index")).series.map(s=>({Month:s.key,
    "Spliced Index":+s.val.toFixed(2),"Source Base":s.base,Overlap:ov.has(s.key)?"yes":""}));
}
// numeric cells output raw; only TEXT cells get the formula-injection guard
function csvVal(v){ if(v==null) return ""; if(typeof v==="number") return String(v);
  let s=String(v); if(/^[=+\-@\t\r]/.test(s)) s="'"+s;
  if(/[",\n]/.test(s)) s='"'+s.replace(/"/g,'""')+'"'; return s; }
const slug=s=>String(s||"chart").replace(/[^A-Za-z0-9]+/g,"_").replace(/^_+|_+$/g,"").toLowerCase()||"chart";
const selInst=()=>CH[document.getElementById("expChart").value];
function exportCsv(){
  const inst=selInst(); const rows=(inst&&inst.__rows)||[]; if(!rows.length) return;
  const cols=[...new Set(rows.flatMap(r=>Object.keys(r)))];
  const lines=[cols.map(csvVal).join(",")].concat(rows.map(r=>cols.map(c=>csvVal(r[c])).join(",")));
  download(slug(inst.__name)+".csv",new Blob([lines.join("\n")],{type:"text/csv;charset=utf-8"}));
}
function exportXlsx(){
  const inst=selInst(); const src=(inst&&inst.__rows)||[]; if(!src.length) return;
  const rows=src.map(r=>{const o={};for(const k in r){const v=r[k];
    o[k]=(typeof v==="string"&&/^[=+\-@]/.test(v))?"'"+v:v;}return o;});
  const ws=XLSX.utils.json_to_sheet(rows); const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,"Data"); XLSX.writeFile(wb,slug(inst.__name)+".xlsx");
}
function exportPng(){ const inst=selInst(); if(!inst) return;
  const url=inst.getDataURL({type:"png",pixelRatio:2,backgroundColor:"#fff"});
  const a=document.createElement("a"); a.href=url; a.download=slug(inst.__name)+".png"; a.click(); }

/* which charts are exportable, and their linked on-screen range selects */
const EXPORTS={
  trend:{label:"CCPI trend",rangeIds:["trendFrom","trendTo"],months:()=>META.months},
  compare:{label:"Compare groups",rangeIds:["cmpFrom","cmpTo"],months:()=>META.months},
  group:{label:"Sub-group trend",rangeIds:["grpFrom","grpTo"],months:()=>META.months},
  heat:{label:"Inflation heatmap",rangeIds:["heatFrom","heatTo"],months:()=>META.heatMonths},
  cats:{label:"Category breakdown (latest month)",rangeIds:null},
  drivers:{label:"Inflation drivers (latest month)",rangeIds:null},
};
const EXPORT_DRAW={trend:drawTrend,compare:drawCompare,group:drawGroup,heat:drawHeat,cats:drawCats,drivers:drawDrivers};
function fillOpts(selEl,months,val){ selEl.textContent="";
  months.forEach(k=>{const o=document.createElement("option");o.value=k;o.textContent=k;selEl.appendChild(o);});
  if(val!=null) selEl.value=val; }
function fillExport(){
  const sel=document.getElementById("expChart");
  Object.entries(EXPORTS).forEach(([k,c])=>{const o=document.createElement("option");o.value=k;o.textContent=c.label;sel.appendChild(o);});
  onExpChart();
}
function onExpChart(){
  const cfg=EXPORTS[document.getElementById("expChart").value];
  const f=document.getElementById("expFrom"), t=document.getElementById("expTo");
  if(cfg.rangeIds){
    fillOpts(f,cfg.months(),document.getElementById(cfg.rangeIds[0]).value);
    fillOpts(t,cfg.months(),document.getElementById(cfg.rangeIds[1]).value);
    f.disabled=false; t.disabled=false;
  } else { f.textContent=""; t.textContent=""; f.disabled=true; t.disabled=true; }
}
function onExpRange(){
  const key=document.getElementById("expChart").value, cfg=EXPORTS[key]; if(!cfg.rangeIds) return;
  document.getElementById(cfg.rangeIds[0]).value=document.getElementById("expFrom").value;
  document.getElementById(cfg.rangeIds[1]).value=document.getElementById("expTo").value;
  EXPORT_DRAW[key]();   // redraw on-screen chart to this range so the export matches exactly
}

/* ------------------------------ wiring / boot ------------------------------ */
function seg(e,id,attr,cb){ if(e.target.tagName!=="BUTTON")return;
  document.querySelectorAll(`#${id} button`).forEach(b=>b.classList.remove("on"));
  e.target.classList.add("on"); cb(e.target.dataset[attr]); }
function fillBaseSel(){
  const sel=document.getElementById("baseSel");
  const bases=[...new Set(STATE.tab1.map(r=>r["Base Year Used"]))].sort((a,b)=>BASE_ORDER.indexOf(a)-BASE_ORDER.indexOf(b));
  const opt=document.createElement("option"); opt.value="__spliced__"; opt.textContent="Continuous (spliced)"; sel.appendChild(opt);
  bases.forEach(b=>{const o=document.createElement("option");o.value=b;o.textContent="Only "+b;sel.appendChild(o);});
}
function wire(){
  document.getElementById("segMeasure").onclick=e=>seg(e,"segMeasure","m",v=>{STATE.measure=v;drawTrend();});
  document.getElementById("segCore").onclick=e=>seg(e,"segCore","c",v=>{STATE.core=v;drawTrend();});
  document.getElementById("segType").onclick=e=>seg(e,"segType","t",v=>{STATE.ctype=v;drawTrend();});
  document.getElementById("baseSel").onchange=e=>{STATE.base=e.target.value;drawTrend();};
  document.getElementById("segCmp").onclick=e=>seg(e,"segCmp","v",v=>{STATE.cmpView=v;drawCompare();});
  document.getElementById("cgAdd").onclick=addComposite;
  document.getElementById("cgName").addEventListener("keydown",e=>{ if(e.key==="Enter"){ e.preventDefault(); addComposite(); }});
  document.getElementById("cgReset").onclick=resetComposites;
  document.getElementById("cgList").addEventListener("change",e=>{
    const t=e.target; if(t.matches('input[type=checkbox][data-cid]')){
      const c=STATE.composites.find(x=>x.id===t.dataset.cid); if(c){ c.sel=t.checked; drawCompare(); } }});
  document.getElementById("cgList").addEventListener("click",e=>{
    const b=e.target.closest("button[data-del]"); if(b) deleteComposite(b.dataset.del); });
  document.getElementById("segHeat").onclick=e=>seg(e,"segHeat","h",v=>{STATE.heat=v;drawHeat();});
  document.getElementById("segCat").onclick=e=>seg(e,"segCat","v",v=>{STATE.catView=v;drawCats();});
  document.getElementById("segGrp").onclick=e=>seg(e,"segGrp","g",v=>{STATE.grpView=v;drawGroup();});
  document.getElementById("segSim").onclick=e=>seg(e,"segSim","s",v=>{STATE.simView=v;runSim();});
  document.getElementById("groupSel").onchange=drawGroup;
  ["trendFrom","trendTo"].forEach(id=>document.getElementById(id).onchange=drawTrend);
  ["cmpFrom","cmpTo"].forEach(id=>document.getElementById(id).onchange=drawCompare);
  ["grpFrom","grpTo"].forEach(id=>document.getElementById(id).onchange=drawGroup);
  ["heatFrom","heatTo"].forEach(id=>document.getElementById(id).onchange=drawHeat);
  document.getElementById("simReset").onclick=()=>{document.querySelectorAll("#simRows input").forEach(i=>i.value="0");
    (window.__sim||[]).forEach(c=>c.shock=0); runSim();};
  document.getElementById("expCsv").onclick=exportCsv;
  document.getElementById("expXlsx").onclick=exportXlsx;
  document.getElementById("expPng").onclick=exportPng;
  document.getElementById("expPrint").onclick=()=>window.print();
  document.getElementById("expChart").onchange=onExpChart;
  document.getElementById("expFrom").onchange=onExpRange;
  document.getElementById("expTo").onchange=onExpRange;
  window.addEventListener("resize",()=>Object.values(CH).forEach(c=>c&&c.resize()));
}
function fillRanges(){
  const fill=(fromId,toId,months,defFromIdx)=>{
    const f=document.getElementById(fromId), t=document.getElementById(toId);
    months.forEach(k=>{
      const o1=document.createElement("option"); o1.value=k; o1.textContent=k; f.appendChild(o1);
      const o2=document.createElement("option"); o2.value=k; o2.textContent=k; t.appendChild(o2);
    });
    if(months.length){ f.value=months[Math.max(0,defFromIdx)]; t.value=months[months.length-1]; }
  };
  fill("trendFrom","trendTo",META.months,0);
  fill("cmpFrom","cmpTo",META.months,0);
  fill("grpFrom","grpTo",META.months,0);
  fill("heatFrom","heatTo",META.heatMonths,Math.max(0,META.heatMonths.length-HEAT_MONTHS));
}
function initCharts(){ ["trend","compare","heat","drivers","cats","group"].forEach(id=>CH[id]=echarts.init(document.getElementById(id))); }
function boot(wb){
  STATE.tab1=rowsOf(wb,"CCPI Data"); STATE.tab2=rowsOf(wb,"CCPI Subgroup Breakdown");
  if(!STATE.tab1.length){showBanner("err","Workbook loaded but 'CCPI Data' is empty or renamed.");return;}
  buildMeta();
  initCharts(); wire(); fillBaseSel(); fillCmp(); fillCmpMembers(); fillGroups(); fillRanges();
  document.getElementById("cgHint").textContent=CG_HINT; renderComposites();
  renderKpis(); drawTrend(); drawCompare(); drawHeat(); drawDrivers(); buildSim();
  drawAnom(); drawCats(); drawGroup(); drawTable(); fillExport();
}
function showBanner(kind,msg){const b=document.getElementById("banner");b.className="banner "+kind;b.textContent=msg;b.style.display="block";}

fetch("./data/ccpi_combined_single_tab.xlsx",{cache:"no-store"})
  .then(r=>{ if(!r.ok) throw 0; STATE.updated=(r.headers.get("last-modified")||"—").slice(0,16); return r.arrayBuffer(); })
  .then(buf=>boot(XLSX.read(new Uint8Array(buf),{type:"array"})))
  .catch(()=>{ showBanner("warn","Showing a small SAMPLE — serve this page (GitHub Pages, or `python -m http.server` in /docs) so it can read the workbook.");
    boot(sampleWb()); });

function sampleWb(){
  const t1=[],t2=[]; let idx=150;
  const groups=["Food and Non-Alcoholic Beverages","Transport","Housing, Water, Electricity, Gas and Other Fuels","Health","Education"];
  const w={"Food and Non-Alcoholic Beverages":26,"Transport":13,"Housing, Water, Electricity, Gas and Other Fuels":32,"Health":4,"Education":5};
  for(let i=0;i<40;i++){const y=2021+Math.floor(i/12),m=(i%12)+1;
    const base=(y<2023||(y===2023&&m===1))?"2013=100":"2021=100"; idx=+(idx+1.4+Math.sin(i/3)).toFixed(1);
    t1.push({"Source File":"sample","Base Year Used":base,"Year":y,"Month Number":m,"Month":"Jan","Date":"",
      "CCPI Index":idx,"CCPI Core Index":+(idx*0.95).toFixed(1),"CCPI MoM Change":0.005,"CCPI Core MoM Change":0.004,
      "CCPI YoY Inflation":0.05,"CCPI Core YoY Inflation":0.04,"CCPI 12M Moving Avg Inflation":0.03,"CCPI Core 12M Moving Avg Inflation":0.03});
    groups.forEach((g,j)=>t2.push({"Source File":"sample","Base Year Used":base,"H/H Size":3.8,"Year":y,"Month":"Jan","Month No":m,
      "Period Type":"Monthly","Group":g,"Base Value (Rs. Cts)":1000,"Weight (%)":w[g],"Index Number":+(idx+j*8).toFixed(1),
      "Month-to-Month Inflation (%)":+(Math.sin(i/2+j)).toFixed(1),"Year-on-Year Inflation (%)":+(4+j+Math.sin(i/4)).toFixed(1),"Annual Average Inflation (%)":3}));
  }
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(t1),"CCPI Data");
  XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(t2),"CCPI Subgroup Breakdown");
  return wb;
}
