"use strict";
/* Sri Lanka CCPI dashboard — Phase 1.
   Compute-once MODEL + pure STATS functions + tabbed, lazy-rendered UI.
   Security: no inline scripts (CSP), no eval, all dynamic text via textContent,
   exports sanitised against spreadsheet-formula injection. */

const BASE_ORDER=["1952=100","2002=100","2006/07=100","2013=100","2021=100"];
const LATEST_BASE="2021=100";
const AGG=new Set(["All Items","Non Food"]);
const HEAT_MONTHS=36, TRIM=0.10, MA_LONG=12, VOL_WIN=12, EWMA_SPAN=6;

const el=id=>document.getElementById(id);
const txt=(id,t)=>{const e=el(id);if(e)e.textContent=t;};
const ymKey=(y,m)=>`${y}-${String(m).padStart(2,"0")}`;
const ymShift=(k,n)=>{let[y,m]=k.split("-").map(Number);let t=y*12+(m-1)-n;return `${Math.floor(t/12)}-${String((t%12)+1).padStart(2,"0")}`;};
const fmtPct=v=>(v>0?"+":"")+v.toFixed(1)+"%";
const toNum=v=>{if(v===null||v===undefined||v==="")return null;const n=+(""+v).replace(/,/g,"");return isFinite(n)?n:null;};
const seriesMap=s=>{const m={};s.forEach(p=>m[p.key]=p.val);return m;};
const sum=(a,f)=>a.reduce((s,x)=>s+f(x),0);

let MODEL=null, FILTER={from:null,to:null,measure:"head"};
let STATE={ctype:"line",base:"__spliced__",catView:"index",cmpView:"index",heat:"yoy"};
const CH={};
function ensure(id){ if(!CH[id]) CH[id]=echarts.init(el(id)); return CH[id]; }

/* ---------------- splice (verified algorithm) ---------------- */
function splice(points){
  const present=[...new Set(points.map(p=>p.base))].sort((a,b)=>BASE_ORDER.indexOf(a)-BASE_ORDER.indexOf(b));
  if(!present.length) return {series:[],scale:{}};
  const latest=present[present.length-1], by={}; present.forEach(b=>by[b]={});
  points.forEach(p=>{by[p.base][ymKey(p.year,p.month)]=p.value;});
  const scale={[latest]:1};
  for(let i=present.length-1;i>0;i--){const nw=present[i],ol=present[i-1];
    const common=Object.keys(by[nw]).filter(k=>k in by[ol]).sort();
    scale[ol]=scale[nw]*(common.length?by[nw][common[common.length-1]]/by[ol][common[common.length-1]]:1);}
  const out={};
  points.forEach(p=>{const k=ymKey(p.year,p.month);
    if(!(k in out)||BASE_ORDER.indexOf(p.base)>BASE_ORDER.indexOf(out[k].base))
      out[k]={base:p.base,val:p.value*scale[p.base],y:p.year,m:p.month};});
  return {series:Object.keys(out).sort().map(k=>({key:k,...out[k]})),scale};
}
function overlapMonths(points){const seen={};
  points.forEach(p=>{const k=ymKey(p.year,p.month);(seen[k]=seen[k]||new Set()).add(p.base);});
  return new Set(Object.keys(seen).filter(k=>seen[k].size>1));}
function contiguousRanges(keys){const r=[];let s=null,p=null;
  keys.forEach(k=>{if(s===null){s=k;}else if(ymShift(k,1)!==p){r.push([s,p]);s=k;}p=k;});
  if(s!==null)r.push([s,p]);return r;}

/* ---------------- build MODEL once ---------------- */
const rowsOf=(wb,s)=> wb.Sheets[s]?XLSX.utils.sheet_to_json(wb.Sheets[s],{defval:null}):[];
function idxPoints(rows,col){
  return rows.filter(r=>toNum(r[col])!=null&&r["Base Year Used"]&&r.Year&&r["Month Number"])
    .map(r=>({year:+r.Year,month:+r["Month Number"],base:r["Base Year Used"],value:toNum(r[col])}));
}
function buildModel(wb){
  const tab1=rowsOf(wb,"CCPI Data"), tab2=rowsOf(wb,"CCPI Subgroup Breakdown");
  const headline=splice(idxPoints(tab1,"CCPI Index")).series;
  const core=splice(idxPoints(tab1,"CCPI Core Index")).series;
  const months=headline.map(s=>s.key);
  const ovPts=idxPoints(tab1,"CCPI Index");
  const overlap=overlapMonths(ovPts);
  // base-change points on the spliced headline
  const baseChanges=[]; for(let i=1;i<headline.length;i++) if(headline[i].base!==headline[i-1].base) baseChanges.push(headline[i].key);
  // reported inflation (decimals) collapsed to newest base per month
  const reported={};
  tab1.forEach(r=>{ if(!r.Year||!r["Month Number"]) return; const k=ymKey(+r.Year,+r["Month Number"]);
    if(!(k in reported)||BASE_ORDER.indexOf(r["Base Year Used"])>BASE_ORDER.indexOf(reported[k].base))
      reported[k]={base:r["Base Year Used"],yoy:toNum(r["CCPI YoY Inflation"]),coreYoy:toNum(r["CCPI Core YoY Inflation"]),
        mom:toNum(r["CCPI MoM Change"]),ma:toNum(r["CCPI 12M Moving Avg Inflation"]),
        index:toNum(r["CCPI Index"]),core:toNum(r["CCPI Core Index"])};});
  // component cross-sections per month (Tab2 Monthly, newest base per month), weights normalised
  const groupsSet=new Set(), byMonthBase={};
  tab2.forEach(r=>{ if(r["Period Type"]!=="Monthly"||AGG.has(r.Group)||!r.Group) return;
    if(toNum(r["Weight (%)"])==null||toNum(r["Index Number"])==null) return;
    const k=ymKey(+r.Year,+r["Month No"]), b=r["Base Year Used"]; groupsSet.add(r.Group);
    (byMonthBase[k]=byMonthBase[k]||{}); (byMonthBase[k][b]=byMonthBase[k][b]||[]).push(
      {group:r.Group,w:toNum(r["Weight (%)"]),yoy:toNum(r["Year-on-Year Inflation (%)"]),
       mom:toNum(r["Month-to-Month Inflation (%)"]),index:toNum(r["Index Number"])});});
  const components={};
  Object.keys(byMonthBase).forEach(k=>{ const bases=Object.keys(byMonthBase[k]).sort((a,b)=>BASE_ORDER.indexOf(a)-BASE_ORDER.indexOf(b));
    const arr=byMonthBase[k][bases[bases.length-1]]; const W=sum(arr,x=>x.w)||1;
    components[k]=arr.map(x=>({...x,w:x.w/W})); });
  const compMonths=Object.keys(components).sort();
  return {tab1,tab2,headline,core,months,overlap,baseChanges,reported,components,compMonths,
    idx:seriesMap(headline),coreIdx:seriesMap(core),groups:[...groupsSet].sort(),
    latest:months[months.length-1],latestComp:compMonths[compMonths.length-1],updated:"—"};
}

/* ---------------- pure STATS ---------------- */
const S={
  yoyFromIndex(series){const m=seriesMap(series);return series.map(p=>{const k=ymShift(p.key,12);
    return (k in m)?{key:p.key,v:(p.val/m[k]-1)*100}:null;}).filter(Boolean);},
  ann(series,k){const m=seriesMap(series);return series.map(p=>{const pk=ymShift(p.key,k);
    return (pk in m)?{key:p.key,v:((p.val/m[pk])**(12/k)-1)*100}:null;}).filter(Boolean);},
  mom(series){const m=seriesMap(series);return series.map(p=>{const pk=ymShift(p.key,1);
    return (pk in m)?{key:p.key,v:(p.val/m[pk]-1)*100}:null;}).filter(Boolean);},
  trailingMean(series,n){const m=seriesMap(series);return series.map(p=>{let s=0,c=0;
    for(let j=0;j<n;j++){const k=ymShift(p.key,j);if(k in m){s+=m[k];c++;}}return c===n?{key:p.key,v:s/n}:null;}).filter(Boolean);},
  ewma(series,alpha){let e=null;return series.map(p=>{e=(e===null)?p.v:alpha*p.v+(1-alpha)*e;return{key:p.key,v:e};});},
  trailingSD(series,n){const m=seriesMap(series);return series.map(p=>{const xs=[];
    for(let j=0;j<n;j++){const k=ymShift(p.key,j);if(k in m)xs.push(m[k]);}
    if(xs.length<n)return null;const mn=xs.reduce((a,b)=>a+b,0)/n;
    return {key:p.key,v:Math.sqrt(xs.reduce((a,b)=>a+(b-mn)**2,0)/n)};}).filter(Boolean);},
  align(a,b,fn){const mb=seriesMap(b);return a.filter(x=>x.key in mb).map(x=>({key:x.key,v:fn(x.v,mb[x.key])}));},
  wmean(items){const W=sum(items,i=>i.w)||1;return sum(items,i=>i.w*i.x)/W;},
  wmedian(items){const a=[...items].sort((p,q)=>p.x-q.x);const W=sum(a,i=>i.w)||1;let c=0;
    for(const i of a){c+=i.w/W;if(c>=0.5)return i.x;}return a.length?a[a.length-1].x:0;},
  wtrim(items,t){const a=[...items].sort((p,q)=>p.x-q.x);const W=sum(a,i=>i.w)||1;let c=0;const keep=[];
    for(const i of a){c+=i.w/W;if(c>t&&c<1-t)keep.push(i);}const k=keep.length?keep:a;
    const KW=sum(k,i=>i.w)||1;return sum(k,i=>i.w*i.x)/KW;},
  wsd(items){const m=S.wmean(items),W=sum(items,i=>i.w)||1;return Math.sqrt(sum(items,i=>i.w*(i.x-m)**2)/W);},
  wskew(items){const m=S.wmean(items),sd=S.wsd(items);if(sd<=0)return 0;const W=sum(items,i=>i.w)||1;
    return sum(items,i=>i.w*((i.x-m)/sd)**3)/W;},
  diffusion(items){const W=sum(items,i=>i.w)||1;return sum(items,i=>i.w*(i.mom>0?1:(i.mom===0?0.5:0)))/W*100;},
};

/* cross-sectional series over months that have components */
function crossSeries(){
  const trim=[],med=[],mean=[],diff=[],disp=[],skew=[];
  MODEL.compMonths.forEach(k=>{const c=MODEL.components[k];
    const yv=c.filter(x=>x.yoy!=null).map(x=>({w:x.w,x:x.yoy}));
    const mv=c.filter(x=>x.mom!=null).map(x=>({w:x.w,mom:x.mom}));
    if(yv.length){mean.push({key:k,v:S.wmean(yv)});med.push({key:k,v:S.wmedian(yv)});
      trim.push({key:k,v:S.wtrim(yv,TRIM)});disp.push({key:k,v:S.wsd(yv)});skew.push({key:k,v:S.wskew(yv)});}
    if(mv.length)diff.push({key:k,v:S.diffusion(mv)});});
  return {trim,med,mean,diff,disp,skew};
}
function carryover(year){
  const idx=MODEL.idx, obs=[]; for(let m=1;m<=12;m++) if(ymKey(year,m) in idx) obs.push(m);
  if(!obs.length) return null; const lastV=idx[ymKey(year,obs[obs.length-1])];
  let projSum=0; for(let m=1;m<=12;m++){const k=ymKey(year,m); projSum+=(k in idx)?idx[k]:lastV;}
  const prev=[]; for(let m=1;m<=12;m++){const k=ymKey(year-1,m); if(k in idx)prev.push(idx[k]);}
  if(!prev.length) return null;
  return ((projSum/12)/(prev.reduce((a,b)=>a+b,0)/prev.length)-1)*100;
}

/* ---------------- filters ---------------- */
const inRange=k=>(!FILTER.from||k>=FILTER.from)&&(!FILTER.to||k<=FILTER.to);
const clip=s=>s.filter(p=>inRange(p.key));
function idxSeries(){return FILTER.measure==="core"?MODEL.core:MODEL.headline;}
function color(v,hi){return v<0?"#1a8a55":(v>hi?"#c0392b":"#c0561a");}
function lineOpt(series,unit,col){const s=clip(series);
  return {grid:{left:52,right:16,top:16,bottom:26},tooltip:{trigger:"axis",valueFormatter:v=>v==null?"—":v.toFixed(2)+(unit||"")},
    xAxis:{type:"category",data:s.map(p=>p.key),axisLabel:{fontSize:10,color:"#7a869a"}},
    yAxis:{type:"value",scale:true,splitLine:{lineStyle:{color:"#eef2f7"}},axisLabel:{fontSize:10,color:"#7a869a"}},
    series:[{type:"line",data:s.map(p=>+p.v.toFixed(2)),smooth:true,showSymbol:false,sampling:"lttb",
      lineStyle:{width:2,color:col||"#1f6feb"},areaStyle:{color:(col||"#1f6feb")+"18"}}]};}
function lastVal(series){const s=clip(series);return s.length?s[s.length-1]:null;}

/* ---------------- OVERVIEW ---------------- */
function renderKpis(){
  const r=MODEL.reported[MODEL.latest]; if(!r)return;
  txt("kIndex",r.index!=null?r.index.toFixed(1):"—");
  const y=el("kYoY"); y.textContent=r.yoy!=null?fmtPct(r.yoy*100):"—";
  y.style.color=r.yoy==null?"":color(r.yoy*100,8);
  txt("kMA","12M avg "+(r.ma!=null?fmtPct(r.ma*100):"—"));
  txt("kMoM",r.mom!=null?fmtPct(r.mom*100):"—");
  txt("kCore",r.core!=null?r.core.toFixed(1):"—");
  txt("kCoreYoY","core Y-o-Y "+(r.coreYoy!=null?fmtPct(r.coreYoy*100):"—"));
}
function overlapBands(){return contiguousRanges([...MODEL.overlap].sort())
  .map(([a,b])=>[{xAxis:a},{xAxis:b}]).filter(([a,b])=>inRange(a[0].xAxis)||inRange(b[1].xAxis));}
function drawTrend(){
  const col=FILTER.measure==="core"?"CCPI Core Index":"CCPI Index";
  let xs=[],ys=[],marks=[],areas=[];
  if(STATE.base==="__spliced__"){
    const s=clip(idxSeries()); xs=s.map(p=>p.key); ys=s.map(p=>+p.val.toFixed(2));
    MODEL.baseChanges.filter(inRange).forEach(k=>marks.push({xAxis:k}));
    areas=overlapBands();
  }else{
    const f=clip(idxPoints(MODEL.tab1,col).filter(p=>p.base===STATE.base)
      .map(p=>({key:ymKey(p.year,p.month),val:p.value}))).sort((a,b)=>a.key<b.key?-1:1);
    xs=f.map(p=>p.key); ys=f.map(p=>+p.val.toFixed(2));
  }
  const isArea=STATE.ctype==="area",isBar=STATE.ctype==="bar";
  ensure("trend").setOption({grid:{left:52,right:16,top:16,bottom:28},tooltip:{trigger:"axis"},
    xAxis:{type:"category",data:xs,axisLabel:{fontSize:10,color:"#7a869a"}},
    yAxis:{type:"value",scale:true,splitLine:{lineStyle:{color:"#eef2f7"}},axisLabel:{fontSize:10,color:"#7a869a"}},
    series:[{type:isBar?"bar":"line",data:ys,smooth:!isBar,showSymbol:false,sampling:"lttb",
      lineStyle:{width:2.2,color:"#1f6feb"},itemStyle:{color:"#1f6feb"},
      areaStyle:isArea?{color:"rgba(31,111,235,.10)"}:null,
      markLine:{symbol:"none",silent:true,lineStyle:{color:"#b0b8c4",type:"dashed"},label:{show:false},data:marks},
      markArea:{silent:true,itemStyle:{color:"rgba(240,179,74,.22)"},label:{show:false},data:areas}}]},true);
}
function drawDrivers(){
  const c=MODEL.components[MODEL.latestComp]||[];
  const arr=c.filter(x=>x.yoy!=null).map(x=>({name:x.group,cbn:x.w*x.yoy})).sort((a,b)=>a.cbn-b.cbn);
  const s=arr.reduce((t,d)=>t+d.cbn,0), head=(MODEL.reported[MODEL.latest]||{}).yoy;
  txt("drvHint",`Latest ${MODEL.latestComp}. Σ contributions ≈ ${s.toFixed(2)}% vs headline ${head==null?"—":(head*100).toFixed(1)+"%"}.`);
  ensure("drivers").setOption({grid:{left:8,right:26,top:8,bottom:8,containLabel:true},
    tooltip:{trigger:"axis",axisPointer:{type:"shadow"},valueFormatter:v=>v+" pp"},
    xAxis:{type:"value",axisLabel:{fontSize:10,color:"#7a869a"},splitLine:{lineStyle:{color:"#eef2f7"}}},
    yAxis:{type:"category",data:arr.map(d=>d.name),axisLabel:{fontSize:10,color:"#5b6b80",width:150,overflow:"truncate"}},
    series:[{type:"bar",data:arr.map(d=>+d.cbn.toFixed(3)),barWidth:"62%",
      itemStyle:{color:p=>p.value<0?"#1a8a55":"#c0561a",borderRadius:[0,4,4,0]}}]},true);
}
function drawCats(){
  const c=MODEL.components[MODEL.latestComp]||[];
  txt("catHint",`Latest month ${MODEL.latestComp} · aggregates excluded.`);
  const key=STATE.catView==="yoy"?"yoy":"index";
  const arr=c.filter(x=>x[key]!=null).map(x=>({name:x.group,v:x[key]})).sort((a,b)=>a.v-b.v);
  ensure("cats").setOption({grid:{left:8,right:26,top:8,bottom:8,containLabel:true},
    tooltip:{trigger:"axis",axisPointer:{type:"shadow"}},
    xAxis:{type:"value",axisLabel:{fontSize:10,color:"#7a869a"},splitLine:{lineStyle:{color:"#eef2f7"}}},
    yAxis:{type:"category",data:arr.map(d=>d.name),axisLabel:{fontSize:10,color:"#5b6b80",width:150,overflow:"truncate"}},
    series:[{type:"bar",data:arr.map(d=>+d.v.toFixed(2)),barWidth:"62%",
      itemStyle:{color:STATE.catView==="yoy"?"#c0561a":"#1f6feb",borderRadius:[0,4,4,0]}}]},true);
}
function drawTable(){
  const s=clip(MODEL.headline);
  const tb=document.querySelector("#tbl tbody"); tb.textContent="";
  s.slice().reverse().forEach(p=>{const tr=document.createElement("tr");
    [p.key,p.val.toFixed(1),p.base,MODEL.overlap.has(p.key)?"yes":""].forEach((v,j)=>{
      const td=document.createElement("td");td.textContent=v;if(j===0)td.style.textAlign="left";tr.appendChild(td);});
    tb.appendChild(tr);});
}

/* ---------------- MOMENTUM & SMOOTHING ---------------- */
function drawMom(){
  const base=idxSeries();
  const yoy=S.yoyFromIndex(base), a3=S.ann(base,3), a6=S.ann(base,6), a1=S.ann(base,1);
  const xkeys=clip(yoy).map(p=>p.key);
  const pick=(ser)=>{const m=seriesMap(ser);return xkeys.map(k=>k in m?+m[k].toFixed(2):null);};
  ensure("chMom").setOption({grid:{left:50,right:16,top:26,bottom:28},tooltip:{trigger:"axis"},
    legend:{top:0,textStyle:{fontSize:10}},
    xAxis:{type:"category",data:xkeys,axisLabel:{fontSize:10,color:"#7a869a"}},
    yAxis:{type:"value",scale:true,splitLine:{lineStyle:{color:"#eef2f7"}},axisLabel:{fontSize:10,color:"#7a869a"}},
    series:[
      {name:"1m ann (noisy)",type:"line",data:pick(a1),showSymbol:false,smooth:true,lineStyle:{width:1,color:"#c9d4e2"}},
      {name:"3m ann",type:"line",data:pick(a3),showSymbol:false,smooth:true,lineStyle:{width:2,color:"#e8893b"}},
      {name:"6m ann",type:"line",data:pick(a6),showSymbol:false,smooth:true,lineStyle:{width:2,color:"#1f6feb"}},
      {name:"12m Y-o-Y",type:"line",data:pick(yoy),showSymbol:false,smooth:true,lineStyle:{width:2.4,color:"#1b263b"}}]},true);
  const l3=lastVal(a3),ly=lastVal(yoy);
  txt("chMom_t",(l3&&ly)?`Latest 3-month pace ${l3.v.toFixed(1)}% vs 12-month ${ly.v.toFixed(1)}% — ${l3.v>ly.v?"running hotter than":"running below"} the annual rate. Not seasonally adjusted.`:"");
}
function drawGap(){
  const base=idxSeries(); const gap=S.align(S.ann(base,3),S.yoyFromIndex(base),(a,b)=>a-b);
  const s=clip(gap);
  ensure("chGap").setOption({grid:{left:50,right:16,top:12,bottom:26},tooltip:{trigger:"axis",valueFormatter:v=>v.toFixed(2)+" pp"},
    xAxis:{type:"category",data:s.map(p=>p.key),axisLabel:{fontSize:10,color:"#7a869a"}},
    yAxis:{type:"value",scale:true,splitLine:{lineStyle:{color:"#eef2f7"}},axisLabel:{fontSize:10,color:"#7a869a"}},
    series:[{type:"bar",data:s.map(p=>+p.v.toFixed(2)),itemStyle:{color:p=>p.value>=0?"#c0561a":"#1a8a55"}}]},true);
  const l=lastVal(gap); txt("chGap_t",l?`Latest gap ${l.v>0?"+":""}${l.v.toFixed(1)}pp — inflation ${l.v>0?"accelerating":"cooling"} relative to its 12-month rate.`:"");
}
function drawSmooth(){
  const yoy=S.yoyFromIndex(idxSeries());
  const ma3=S.trailingMean(yoy,3),ma6=S.trailingMean(yoy,6),ma12=S.trailingMean(yoy,MA_LONG),ew=S.ewma(yoy,2/(EWMA_SPAN+1));
  const xkeys=clip(yoy).map(p=>p.key);
  const pick=ser=>{const m=seriesMap(ser);return xkeys.map(k=>k in m?+m[k].toFixed(2):null);};
  ensure("chSmooth").setOption({grid:{left:48,right:12,top:24,bottom:26},tooltip:{trigger:"axis"},
    legend:{top:0,textStyle:{fontSize:9}},
    xAxis:{type:"category",data:xkeys,axisLabel:{fontSize:9,color:"#7a869a"}},
    yAxis:{type:"value",scale:true,splitLine:{lineStyle:{color:"#eef2f7"}},axisLabel:{fontSize:9,color:"#7a869a"}},
    series:[
      {name:"Y-o-Y",type:"line",data:pick(yoy),showSymbol:false,smooth:true,lineStyle:{width:1,color:"#cbd5e1"}},
      {name:"3m MA",type:"line",data:pick(ma3),showSymbol:false,smooth:true,lineStyle:{width:1.6,color:"#e8893b"}},
      {name:"6m MA",type:"line",data:pick(ma6),showSymbol:false,smooth:true,lineStyle:{width:1.6,color:"#1f6feb"}},
      {name:"12m MA",type:"line",data:pick(ma12),showSymbol:false,smooth:true,lineStyle:{width:2,color:"#1b263b"}},
      {name:"EWMA",type:"line",data:pick(ew),showSymbol:false,smooth:true,lineStyle:{width:1.6,type:"dashed",color:"#1a8a55"}}]},true);
  txt("chSmooth_t","Shorter averages track turns faster; the 12-month average is the smoothest trend.");
}
function drawVol(){
  const vol=S.trailingSD(S.mom(idxSeries()),VOL_WIN);
  ensure("chVol").setOption(lineOpt(vol," pp","#8046c0"));
  const l=lastVal(vol); txt("chVol_t",l?`Rolling 12-month volatility of monthly inflation is ${l.v.toFixed(2)}pp.`:"");
}

/* ---------------- UNDERLYING ---------------- */
function drawUnderlying(){
  const cs=crossSeries(); const r=MODEL.reported[MODEL.latest]||{};
  const lt=k=>{const s=clip(cs[k]);return s.length?s[s.length-1].v:null;};
  const tm=lt("trim"),md=lt("med");
  txt("uHead",r.yoy!=null?fmtPct(r.yoy*100):"—");
  txt("uTrim",tm!=null?fmtPct(tm):"—"); txt("uMed",md!=null?fmtPct(md):"—");
  txt("uGap",(r.yoy!=null&&r.coreYoy!=null)?fmtPct((r.yoy-r.coreYoy)*100):"—");
  const xkeys=clip(cs.mean).map(p=>p.key);
  const pick=ser=>{const m=seriesMap(ser);return xkeys.map(k=>k in m?+m[k].toFixed(2):null);};
  ensure("chUnder").setOption({grid:{left:50,right:16,top:26,bottom:28},tooltip:{trigger:"axis"},
    legend:{top:0,textStyle:{fontSize:10}},
    xAxis:{type:"category",data:xkeys,axisLabel:{fontSize:10,color:"#7a869a"}},
    yAxis:{type:"value",scale:true,splitLine:{lineStyle:{color:"#eef2f7"}},axisLabel:{fontSize:10,color:"#7a869a"}},
    series:[
      {name:"Weighted mean",type:"line",data:pick(cs.mean),showSymbol:false,smooth:true,lineStyle:{width:1,color:"#c9d4e2"}},
      {name:"Trimmed mean",type:"line",data:pick(cs.trim),showSymbol:false,smooth:true,lineStyle:{width:2,color:"#1f6feb"}},
      {name:"Weighted median",type:"line",data:pick(cs.med),showSymbol:false,smooth:true,lineStyle:{width:2,color:"#e8893b"}}]},true);
  txt("chUnder_t","Trimmed mean and median remove one-off category spikes to show the durable trend.");
  // headline minus core over time (reported)
  const hc=MODEL.months.filter(k=>{const r=MODEL.reported[k];return r&&r.yoy!=null&&r.coreYoy!=null;})
    .map(k=>{const r=MODEL.reported[k];return {key:k,v:(r.yoy-r.coreYoy)*100};});
  ensure("chHC").setOption({...lineOpt(hc," pp","#c0561a")});
  const l=lastVal(hc); txt("chHC_t",l?`Headline is ${l.v>=0?"above":"below"} core by ${Math.abs(l.v).toFixed(1)}pp — volatile items are ${l.v>=0?"adding to":"subtracting from"} the headline.`:"");
}

/* ---------------- BREADTH ---------------- */
function drawBreadth(){
  const cs=crossSeries();
  ensure("chDiff").setOption({...lineOpt(cs.diff,"","#1f6feb"),
    yAxis:{type:"value",min:0,max:100,splitLine:{lineStyle:{color:"#eef2f7"}},axisLabel:{fontSize:10,color:"#7a869a"}}});
  const d=lastVal(cs.diff); txt("chDiff_t",d?`${d.v.toFixed(0)}% of the basket (by weight) rose last month — ${d.v>50?"broad-based":"narrow"}.`:"");
  ensure("chDisp").setOption(lineOpt(cs.disp," pp","#8046c0"));
  const dp=lastVal(cs.disp); txt("chDisp_t",dp?`Category inflation rates are spread ±${dp.v.toFixed(1)}pp around the mean.`:"");
  ensure("chSkew").setOption(lineOpt(cs.skew,"","#c0561a"));
  const sk=lastVal(cs.skew); txt("chSkew_t",sk?`Skew ${sk.v.toFixed(2)} — ${sk.v>0.2?"a few large increases pull the average up":sk.v<-0.2?"tilted toward decreases":"fairly symmetric"}.`:"");
}
function drawHeat(){
  const groups=MODEL.groups;
  const col=STATE.heat==="mom"?"mom":"yoy";
  const months=MODEL.compMonths.filter(inRange).slice(-HEAT_MONTHS);
  const gi=Object.fromEntries(groups.map((g,i)=>[g,i])), mi=Object.fromEntries(months.map((m,i)=>[m,i]));
  const data=[]; let vmax=1;
  months.forEach(k=>{(MODEL.components[k]||[]).forEach(x=>{const v=x[col];
    if(x.group in gi && v!=null){data.push([mi[k],gi[x.group],+v.toFixed(1)]);vmax=Math.max(vmax,Math.abs(v));}});});
  ensure("heat").setOption({grid:{left:8,right:16,top:8,bottom:56,containLabel:true},
    tooltip:{position:"top",formatter:p=>`${groups[p.value[1]]}<br/>${months[p.value[0]]}: ${p.value[2]}%`},
    xAxis:{type:"category",data:months,axisLabel:{fontSize:9,color:"#7a869a",rotate:60}},
    yAxis:{type:"category",data:groups,axisLabel:{fontSize:9,color:"#5b6b80",width:140,overflow:"truncate"}},
    visualMap:{min:-vmax,max:vmax,calculable:true,orient:"horizontal",left:"center",bottom:6,
      inRange:{color:["#1a8a55","#f4f7fb","#c0392b"]},textStyle:{fontSize:10}},
    series:[{type:"heatmap",data,progressive:1000,itemStyle:{borderColor:"#fff",borderWidth:.5}}]},true);
}

/* ---------------- BASE EFFECTS & CONTEXT ---------------- */
function drawBase(){
  const mom=seriesMap(S.mom(idxSeries()));
  const keys=MODEL.months.filter(k=>k in mom && ymShift(k,12) in mom).filter(inRange).slice(-24);
  const cur=keys.map(k=>+mom[k].toFixed(2)), base=keys.map(k=>+(-mom[ymShift(k,12)]).toFixed(2));
  const dyoy=keys.map((k,i)=>+(cur[i]+base[i]).toFixed(2));
  ensure("chBase").setOption({grid:{left:50,right:16,top:24,bottom:28},tooltip:{trigger:"axis"},
    legend:{top:0,textStyle:{fontSize:10}},
    xAxis:{type:"category",data:keys,axisLabel:{fontSize:9,color:"#7a869a"}},
    yAxis:{type:"value",splitLine:{lineStyle:{color:"#eef2f7"}},axisLabel:{fontSize:10,color:"#7a869a"}},
    series:[
      {name:"This-month move",type:"bar",stack:"d",data:cur,itemStyle:{color:"#1f6feb"}},
      {name:"Base effect (year-ago drop-out)",type:"bar",stack:"d",data:base,itemStyle:{color:"#e8893b"}},
      {name:"Δ Y-o-Y",type:"line",data:dyoy,showSymbol:false,lineStyle:{width:2,color:"#1b263b"}}]},true);
  txt("chBase_t","Each month's change in the annual rate splits into new price action (blue) and the year-ago month leaving the window (orange).");
}
function drawCarry(){
  const yr=+MODEL.latest.split("-")[0];
  const cur=carryover(yr);
  txt("cCarry",cur==null?"—":fmtPct(cur));
  const r=MODEL.reported[MODEL.latest];
  const yoy=r&&r.yoy!=null?r.yoy*100:null;
  txt("cDouble",(yoy&&yoy>0)?(70/yoy).toFixed(1):"—");
  const years=[]; for(let y=yr-9;y<=yr;y++){const c=carryover(y);if(c!=null)years.push({key:String(y),v:c});}
  ensure("chCarry").setOption({grid:{left:44,right:12,top:12,bottom:24},tooltip:{trigger:"axis",valueFormatter:v=>v.toFixed(2)+"%"},
    xAxis:{type:"category",data:years.map(y=>y.key),axisLabel:{fontSize:10,color:"#7a869a"}},
    yAxis:{type:"value",splitLine:{lineStyle:{color:"#eef2f7"}},axisLabel:{fontSize:10,color:"#7a869a"}},
    series:[{type:"bar",data:years.map(y=>+y.v.toFixed(2)),itemStyle:{color:p=>p.value<0?"#1a8a55":"#1f6feb",borderRadius:[3,3,0,0]}}]},true);
  txt("chCarry_t",cur==null?"":`If prices held flat from ${MODEL.latest}, ${yr} would still average about ${cur.toFixed(1)}% inflation.`);
}
function drawPP(){
  const a=el("ppFrom").value, b=el("ppTo").value, idx=MODEL.idx;
  if(!(a in idx)||!(b in idx)||a>=b){txt("ppCum","—");txt("ppRupee","—");txt("ppAnn","—");txt("pp_t","Pick two months (From earlier than To).");return;}
  const ratio=idx[b]/idx[a], cum=(ratio-1)*100;
  const mons=(+b.split("-")[0]*12+ +b.split("-")[1])-(+a.split("-")[0]*12+ +a.split("-")[1]);
  const ann=(ratio**(12/mons)-1)*100;
  txt("ppCum",fmtPct(cum)); txt("ppRupee","Rs."+(100*ratio).toFixed(0)); txt("ppAnn",fmtPct(ann));
  txt("pp_t",`Rs.100 of goods in ${a} cost about Rs.${(100*ratio).toFixed(0)} in ${b} (${(cum).toFixed(0)}% cumulative, ${ann.toFixed(1)}% per year).`);
}
function buildSim(){
  const box=el("simRows"); box.textContent="";
  const c=(MODEL.components[MODEL.latestComp]||[]).filter(x=>x.index!=null);
  window.__sim=c.map(x=>({name:x.group,w:x.w,index:x.index,shock:0}));
  window.__sim.forEach((cc,i)=>{const row=document.createElement("div");row.className="simrow";
    const lab=document.createElement("span");lab.textContent=cc.name;lab.title=cc.name;
    lab.style.overflow="hidden";lab.style.textOverflow="ellipsis";lab.style.whiteSpace="nowrap";
    const inp=document.createElement("input");inp.type="number";inp.step="1";inp.value="0";
    inp.setAttribute("aria-label",cc.name+" percent change");
    inp.addEventListener("input",()=>{const v=parseFloat(inp.value);window.__sim[i].shock=isFinite(v)?v:0;runSim();});
    row.appendChild(lab);row.appendChild(inp);box.appendChild(row);});
  runSim();
}
function runSim(){const s=window.__sim||[];const base=s.reduce((t,c)=>t+c.w*c.index,0);
  const neu=s.reduce((t,c)=>t+c.w*c.index*(1+c.shock/100),0);const d=base?((neu/base-1)*100):0;
  txt("simBase",base.toFixed(1));txt("simNew",neu.toFixed(1));
  const de=el("simDelta");de.textContent=fmtPct(d);de.style.color=d<0?"var(--good)":d>0?"var(--bad)":"var(--muted)";}

/* ---------------- COMPARE & ALERTS ---------------- */
function groupSeries(g){
  const pts=MODEL.tab2.filter(r=>r["Period Type"]==="Monthly"&&r.Group===g&&toNum(r["Index Number"])!=null)
    .map(r=>({year:+r.Year,month:+r["Month No"],base:r["Base Year Used"],value:toNum(r["Index Number"])}));
  return splice(pts).series;
}
function fillCmp(){const box=el("cmpChk");if(box.dataset.filled)return;
  const def=new Set(["Food and Non-Alcoholic Beverages","Transport","Housing, Water, Electricity, Gas and Other Fuels"]);
  MODEL.groups.forEach(g=>{const lab=document.createElement("label");const cb=document.createElement("input");
    cb.type="checkbox";cb.value=g;cb.checked=def.has(g);cb.addEventListener("change",drawCompare);
    lab.appendChild(cb);lab.appendChild(document.createTextNode(g));box.appendChild(lab);});
  box.dataset.filled="1";}
function drawCompare(){
  const chosen=[...document.querySelectorAll("#cmpChk input:checked")].map(c=>c.value);
  const yoy=STATE.cmpView==="yoy"; const series=[]; const xset=new Set();
  chosen.forEach(g=>{let s=groupSeries(g);let data;
    if(yoy){const m=seriesMap(s);data=clip(s).map(p=>{const pk=ymShift(p.key,12);
      return (pk in m)?[p.key,+((p.val/m[pk]-1)*100).toFixed(2)]:[p.key,null];});}
    else data=clip(s).map(p=>[p.key,+p.val.toFixed(2)]);
    data.forEach(d=>xset.add(d[0])); series.push({name:g,type:"line",showSymbol:false,smooth:true,sampling:"lttb",data});});
  ensure("compare").setOption({grid:{left:52,right:16,top:24,bottom:28},tooltip:{trigger:"axis"},
    legend:{type:"scroll",top:0,textStyle:{fontSize:10}},
    xAxis:{type:"category",data:[...xset].sort(),axisLabel:{fontSize:10,color:"#7a869a"}},
    yAxis:{type:"value",scale:true,splitLine:{lineStyle:{color:"#eef2f7"}},axisLabel:{fontSize:10,color:"#7a869a"}},
    series},true);
}
function fillGroups(){const sel=el("groupSel");if(sel.dataset.filled)return;
  MODEL.groups.forEach(g=>{const o=document.createElement("option");o.value=g;o.textContent=g;sel.appendChild(o);});
  sel.dataset.filled="1";
  sel.value=MODEL.groups.includes("Food and Non-Alcoholic Beverages")?"Food and Non-Alcoholic Beverages":MODEL.groups[0];}
function drawGroup(){const g=el("groupSel").value;if(!g)return;const s=clip(groupSeries(g));
  ensure("group").setOption({grid:{left:48,right:16,top:14,bottom:26},tooltip:{trigger:"axis"},
    xAxis:{type:"category",data:s.map(p=>p.key),axisLabel:{fontSize:10,color:"#7a869a"}},
    yAxis:{type:"value",scale:true,splitLine:{lineStyle:{color:"#eef2f7"}},axisLabel:{fontSize:10,color:"#7a869a"}},
    series:[{type:"line",data:s.map(p=>+p.val.toFixed(1)),smooth:true,showSymbol:false,sampling:"lttb",
      lineStyle:{width:2,color:"#1a8a55"},areaStyle:{color:"rgba(26,138,85,.07)"}}]},true);}
function drawAnom(){
  const mom=S.mom(MODEL.headline); const vals=mom.map(x=>x.v).slice().sort((a,b)=>a-b);
  const med=vals.length?vals[Math.floor(vals.length/2)]:0;
  const dev=mom.map(x=>Math.abs(x.v-med)).sort((a,b)=>a-b); const mad=dev.length?(dev[Math.floor(dev.length/2)]||1e-9):1e-9;
  const z=x=>0.6745*(x-med)/mad;
  const flagged=mom.map(x=>({...x,z:z(x.v)})).filter(x=>Math.abs(x.z)>3.5).sort((a,b)=>Math.abs(b.z)-Math.abs(a.z));
  const keys=MODEL.headline.map(s=>s.key),gaps=[];
  for(let i=1;i<keys.length;i++) if(ymShift(keys[i],1)!==keys[i-1]) gaps.push(keys[i]);
  const tb=document.querySelector("#anom tbody"); tb.textContent="";
  const add=(a,b,c,d)=>{const tr=document.createElement("tr");[a,b,c,d].forEach((v,j)=>{const td=document.createElement("td");
    td.textContent=v;if(j===0)td.style.textAlign="left";tr.appendChild(td);});tb.appendChild(tr);};
  flagged.slice(0,40).forEach(x=>add(x.key,fmtPct(x.v),x.z.toFixed(1),x.v>0?"spike":"drop"));
  gaps.forEach(g=>add(g,"—","—","missing prior month"));
  if(!flagged.length&&!gaps.length) add("—","—","—","no anomalies");
}

/* ---------------- exports ---------------- */
function csvCell(v){let s=(v==null?"":String(v));if(/^[=+\-@\t\r]/.test(s))s="'"+s;
  if(/[",\n]/.test(s))s='"'+s.replace(/"/g,'""')+'"';return s;}
function download(name,blob){const u=URL.createObjectURL(blob);const a=document.createElement("a");
  a.href=u;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(u),1500);}
function contRows(){return clip(MODEL.headline).map(p=>({Month:p.key,"Spliced Index":+p.val.toFixed(2),
  "Source Base":p.base,Overlap:MODEL.overlap.has(p.key)?"yes":""}));}
function exportCsv(){const rows=contRows();const cols=Object.keys(rows[0]||{Month:""});
  const lines=[cols.map(csvCell).join(",")].concat(rows.map(r=>cols.map(c=>csvCell(r[c])).join(",")));
  download("ccpi_continuous.csv",new Blob([lines.join("\n")],{type:"text/csv;charset=utf-8"}));}
function exportXlsx(){const rows=contRows().map(r=>{const o={};for(const k in r)o[k]=(typeof r[k]==="string"&&/^[=+\-@]/.test(r[k]))?"'"+r[k]:r[k];return o;});
  const ws=XLSX.utils.json_to_sheet(rows);const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,"CCPI Continuous");XLSX.writeFile(wb,"ccpi_continuous.xlsx");}
function exportPng(){if(!CH.trend)return;const url=CH.trend.getDataURL({type:"png",pixelRatio:2,backgroundColor:"#fff"});
  const a=document.createElement("a");a.href=url;a.download="ccpi_trend.png";a.click();}

/* ---------------- tabs, filters, boot ---------------- */
const PAGES={
  overview:()=>{renderKpis();drawTrend();drawDrivers();drawCats();drawTable();},
  momentum:()=>{drawMom();drawGap();drawSmooth();drawVol();},
  underlying:()=>{drawUnderlying();},
  breadth:()=>{drawBreadth();drawHeat();},
  base:()=>{drawBase();drawCarry();drawPP();buildSim();},
  compare:()=>{fillCmp();drawCompare();fillGroups();drawGroup();drawAnom();},
};
let CURRENT="overview";
function openTab(name){
  CURRENT=name;
  document.querySelectorAll("#tabs button").forEach(b=>b.classList.toggle("on",b.dataset.tab===name));
  document.querySelectorAll(".tabpage").forEach(p=>p.classList.toggle("on",p.dataset.page===name));
  PAGES[name]();
  setTimeout(()=>Object.values(CH).forEach(c=>c&&c.resize()),0);
}
function refresh(){PAGES[CURRENT]();}
function seg(e,id,cb){if(e.target.tagName!=="BUTTON")return;
  document.querySelectorAll(`#${id} button`).forEach(b=>b.classList.remove("on"));
  e.target.classList.add("on");cb(e.target.dataset);}
function fillMonthSelect(sel,val){MODEL.months.forEach(k=>{const o=document.createElement("option");o.value=k;o.textContent=k;sel.appendChild(o);});sel.value=val;}
function wire(){
  document.getElementById("tabs").addEventListener("click",e=>{if(e.target.dataset.tab)openTab(e.target.dataset.tab);});
  el("segMeasure").onclick=e=>seg(e,"segMeasure",d=>{FILTER.measure=d.m;refresh();});
  el("fromSel").onchange=e=>{FILTER.from=e.target.value;refresh();};
  el("toSel").onchange=e=>{FILTER.to=e.target.value;refresh();};
  el("segType").onclick=e=>seg(e,"segType",d=>{STATE.ctype=d.t;drawTrend();});
  el("baseSel").onchange=e=>{STATE.base=e.target.value;drawTrend();};
  el("segCat").onclick=e=>seg(e,"segCat",d=>{STATE.catView=d.v;drawCats();});
  el("segCmp").onclick=e=>seg(e,"segCmp",d=>{STATE.cmpView=d.v;drawCompare();});
  el("segHeat").onclick=e=>seg(e,"segHeat",d=>{STATE.heat=d.h;drawHeat();});
  el("groupSel").onchange=drawGroup;
  el("ppFrom").onchange=drawPP; el("ppTo").onchange=drawPP;
  el("simReset").onclick=()=>{document.querySelectorAll("#simRows input").forEach(i=>i.value="0");
    (window.__sim||[]).forEach(c=>c.shock=0);runSim();};
  el("expCsv").onclick=exportCsv; el("expXlsx").onclick=exportXlsx; el("expPng").onclick=exportPng;
  el("expPrint").onclick=()=>window.print();
  window.addEventListener("resize",()=>Object.values(CH).forEach(c=>c&&c.resize()));
}
function fillBaseSel(){const sel=el("baseSel");
  const bases=[...new Set(MODEL.tab1.map(r=>r["Base Year Used"]))].sort((a,b)=>BASE_ORDER.indexOf(a)-BASE_ORDER.indexOf(b));
  const o0=document.createElement("option");o0.value="__spliced__";o0.textContent="Continuous (spliced)";sel.appendChild(o0);
  bases.forEach(b=>{const o=document.createElement("option");o.value=b;o.textContent="Only "+b;sel.appendChild(o);});}
function setHeader(){txt("fMonth",MODEL.latest);txt("fUpdated",MODEL.updated);
  txt("fRows",String(MODEL.tab1.length+MODEL.tab2.length));txt("fOverlap",String(MODEL.overlap.size));}
function boot(wb,updated){
  MODEL=buildModel(wb); MODEL.updated=updated||"—";
  if(!MODEL.headline.length){showBanner("err","Workbook loaded but 'CCPI Data' is empty or renamed.");return;}
  const n=MODEL.months.length, defFrom=MODEL.months[Math.max(0,n-181)], last=MODEL.months[n-1];
  FILTER.from=defFrom; FILTER.to=last;
  fillMonthSelect(el("fromSel"),defFrom); fillMonthSelect(el("toSel"),last);
  fillMonthSelect(el("ppFrom"),MODEL.months[0]); fillMonthSelect(el("ppTo"),last);
  fillBaseSel(); setHeader(); wire(); openTab("overview");
}
function showBanner(kind,msg){const b=el("banner");b.className="banner "+kind;b.textContent=msg;b.style.display="block";}

fetch("./data/ccpi_combined_single_tab.xlsx",{cache:"no-store"})
  .then(r=>{if(!r.ok)throw 0;const u=(r.headers.get("last-modified")||"—").slice(0,16);return r.arrayBuffer().then(b=>[b,u]);})
  .then(([buf,u])=>boot(XLSX.read(new Uint8Array(buf),{type:"array"}),u))
  .catch(()=>{showBanner("warn","Showing a small SAMPLE — serve this page (GitHub Pages, or `python -m http.server` in /docs) so it can read the workbook.");boot(sampleWb(),"sample");});

function sampleWb(){
  const t1=[],t2=[];let idx=150;
  const groups=["Food and Non-Alcoholic Beverages","Transport","Housing, Water, Electricity, Gas and Other Fuels","Health","Education"];
  const w={"Food and Non-Alcoholic Beverages":26,"Transport":13,"Housing, Water, Electricity, Gas and Other Fuels":32,"Health":4,"Education":5};
  for(let i=0;i<48;i++){const y=2021+Math.floor(i/12),m=(i%12)+1;
    const base=(y<2023||(y===2023&&m===1))?"2013=100":"2021=100";idx=+(idx+1.4+Math.sin(i/3)).toFixed(1);
    t1.push({"Source File":"sample","Base Year Used":base,"Year":y,"Month Number":m,"Month":"Jan","Date":"",
      "CCPI Index":idx,"CCPI Core Index":+(idx*0.95).toFixed(1),"CCPI MoM Change":0.005,"CCPI Core MoM Change":0.004,
      "CCPI YoY Inflation":0.05,"CCPI Core YoY Inflation":0.04,"CCPI 12M Moving Avg Inflation":0.03,"CCPI Core 12M Moving Avg Inflation":0.03});
    groups.forEach((g,j)=>t2.push({"Source File":"sample","Base Year Used":base,"H/H Size":3.8,"Year":y,"Month":"Jan","Month No":m,
      "Period Type":"Monthly","Group":g,"Base Value (Rs. Cts)":1000,"Weight (%)":w[g],"Index Number":+(idx+j*8).toFixed(1),
      "Month-to-Month Inflation (%)":+(Math.sin(i/2+j)).toFixed(1),"Year-on-Year Inflation (%)":+(4+j+Math.sin(i/4)).toFixed(1),"Annual Average Inflation (%)":3}));}
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(t1),"CCPI Data");
  XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(t2),"CCPI Subgroup Breakdown");
  return wb;
}
