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
           cmpView:"index",catView:"index",heat:"yoy",grpView:"index",updated:"—"};
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
  set("fUpdated",STATE.updated);
  set("fRows",String(STATE.tab1.length+STATE.tab2.length));
  set("fOverlap",String(overlapMonths(ccpiPoints("CCPI Index")).size));
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
      markArea:{silent:true,itemStyle:{color:"rgba(240,179,74,.22)"},
        label:{show:false},data:areas}}]},true);
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
  const xs=[...xset].sort(), keysSet=new Set(xs);
  if(series.length) series[0].markLine={symbol:"none",silent:true,
    lineStyle:{color:"#b0b8c4",type:"dashed"},label:{show:false},data:baseChangeMarks(keysSet)};
  CH.compare.setOption({grid:{left:52,right:16,top:24,bottom:28},
    tooltip:{trigger:"axis",valueFormatter:v=>v==null?"—":v.toFixed(2)+unit},
    legend:{type:"scroll",top:0,textStyle:{fontSize:10}},
    xAxis:{type:"category",data:xs,axisLabel:{fontSize:10,color:"#7a869a"}},
    yAxis:{type:"value",scale:true,splitLine:{lineStyle:{color:"#eef2f7"}},axisLabel:{fontSize:10,color:"#7a869a"}},
    series},true);
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
      inRange:{color:["#1a8a55","#f4f7fb","#c0392b"]},textStyle:{fontSize:10}},
    series:[{type:"heatmap",data,progressive:1000,itemStyle:{borderColor:"#fff",borderWidth:.5}}]},true);
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
      itemStyle:{color:p=>p.value<0?"#1a8a55":"#c0561a",borderRadius:[0,4,4,0]}}]},true);
}

/* ------------------------------ scenario ------------------------------ */
function buildSim(){
  const box=document.getElementById("simRows"); box.textContent="";
  const {comp,wsum}=components();
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
function runSim(){
  const sim=window.__sim||[];
  const base=sim.reduce((s,c)=>s+c.w*c.index,0);
  const neu=sim.reduce((s,c)=>s+c.w*c.index*(1+c.shock/100),0);
  const d=base?((neu/base-1)*100):0;
  document.getElementById("simBase").textContent=base.toFixed(1);
  document.getElementById("simNew").textContent=neu.toFixed(1);
  const de=document.getElementById("simDelta"); de.textContent=fmtPct(d);
  de.style.color=d<0?"var(--good)":d>0?"var(--bad)":"var(--muted)";
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
      itemStyle:{color:STATE.catView==="index"?"#1f6feb":"#c0561a",borderRadius:[0,4,4,0]}}]},true);
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
function exportCsv(){
  const rows=continuousRows(); const cols=Object.keys(rows[0]||{Month:""});
  const lines=[cols.map(csvCell).join(",")].concat(rows.map(r=>cols.map(c=>csvCell(r[c])).join(",")));
  download("ccpi_continuous.csv",new Blob([lines.join("\n")],{type:"text/csv;charset=utf-8"}));
}
function exportXlsx(){
  const rows=continuousRows().map(r=>{const o={};for(const k in r)o[k]=(typeof r[k]==="string"&&/^[=+\-@]/.test(r[k]))?"'"+r[k]:r[k];return o;});
  const ws=XLSX.utils.json_to_sheet(rows); const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,"CCPI Continuous"); XLSX.writeFile(wb,"ccpi_continuous.xlsx");
}
function exportPng(){ const url=CH.trend.getDataURL({type:"png",pixelRatio:2,backgroundColor:"#fff"});
  const a=document.createElement("a"); a.href=url; a.download="ccpi_trend.png"; a.click(); }

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
  document.getElementById("segHeat").onclick=e=>seg(e,"segHeat","h",v=>{STATE.heat=v;drawHeat();});
  document.getElementById("segCat").onclick=e=>seg(e,"segCat","v",v=>{STATE.catView=v;drawCats();});
  document.getElementById("segGrp").onclick=e=>seg(e,"segGrp","g",v=>{STATE.grpView=v;drawGroup();});
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
  initCharts(); wire(); fillBaseSel(); fillCmp(); fillGroups(); fillRanges();
  renderKpis(); drawTrend(); drawCompare(); drawHeat(); drawDrivers(); buildSim();
  drawAnom(); drawCats(); drawGroup(); drawTable();
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
