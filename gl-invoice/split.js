// แยกไฟล์ GL Invoice ตามคอลัมน์ -- ย้ายมาจากเครื่องมือเดี่ยว gl-invoice-split (logic เดิมไม่แก้ไข)
// ห่อด้วย IIFE เพื่อไม่ให้ชื่อฟังก์ชัน (เช่น normText) ชนกับ sheetmodel.js/excel.js ของ gl-invoice
(function () {
// ---- Core: split an xlsx by selected columns, editing raw XML only ----
const SPLIT_COLUMNS = ["EMP ID","Alternate ID","NAME","Section","Line Manager","Position",
  "Invoice Sent To","Cost Center 1","Cost Center 2","Cost Center 3","Cost Center 4"];

function xmlDecode(s){
  return s.replace(/&(lt|gt|amp|quot|apos|#\d+|#x[0-9a-fA-F]+);/g,(m,e)=>{
    if(e==="lt")return"<"; if(e==="gt")return">"; if(e==="amp")return"&";
    if(e==="quot")return'"'; if(e==="apos")return"'";
    if(e[1]==="x"||e[1]==="X")return String.fromCodePoint(parseInt(e.slice(2),16));
    return String.fromCodePoint(parseInt(e.slice(1),10));
  });
}
function xmlEncode(s){return String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");}
function colToNum(c){let n=0;for(const ch of c)n=n*26+(ch.charCodeAt(0)-64);return n;}
function attr(a,name){const m=a.match(new RegExp('(?:^|\\s)'+name+'="([^"]*)"'));return m?m[1]:null;}

function parseSharedStrings(xml){
  if(!xml) return [];
  const out=[]; const re=/<si>([\s\S]*?)<\/si>|<si\/>/g; let m;
  while((m=re.exec(xml))){
    const body=(m[1]||"").replace(/<rPh\b[\s\S]*?<\/rPh>/g,"");
    let t=""; const tr=/<t\b[^>]*>([\s\S]*?)<\/t>|<t\b[^>]*\/>/g; let k;
    while((k=tr.exec(body))) t+=xmlDecode(k[1]||"");
    out.push(t);
  }
  return out;
}

function parseRows(sheetData){
  const rows=[]; const re=/<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g; let m;
  while((m=re.exec(sheetData))) rows.push({xml:m[0], r:parseInt(attr(m[1],"r"),10), inner:m[2]||""});
  return rows;
}
function parseCells(rowInner, ss){
  const cells={}; const re=/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g; let m;
  while((m=re.exec(rowInner))){
    const a=m[1], body=m[2]||""; const ref=attr(a,"r"); if(!ref) continue;
    const col=ref.replace(/\d+/g,""); const t=attr(a,"t");
    let v="";
    if(t==="inlineStr"){ const tr=/<t\b[^>]*>([\s\S]*?)<\/t>/g; let k; while((k=tr.exec(body))) v+=xmlDecode(k[1]); }
    else { const vm=body.match(/<v>([\s\S]*?)<\/v>/); if(vm){ v=xmlDecode(vm[1]); if(t==="s") v=ss[parseInt(v,10)]??""; } }
    cells[col]={v:String(v), attrs:a, hasFormula:/<f\b/.test(body)};
  }
  return cells;
}

function resolveTarget(t){ if(t.startsWith("/")) return t.slice(1); return "xl/"+t.replace(/^\.\//,""); }

async function analyzeWorkbook(zip){
  const ss=parseSharedStrings(zip.file("xl/sharedStrings.xml")?await zip.file("xl/sharedStrings.xml").async("string"):"");
  const wb=await zip.file("xl/workbook.xml").async("string");
  const rels=await zip.file("xl/_rels/workbook.xml.rels").async("string");
  const relMap={}; rels.replace(/<Relationship\b([^>]*)\/?>/g,(_,a)=>{relMap[attr(a,"Id")]=attr(a,"Target");});
  const sheets=[]; wb.replace(/<sheet\b([^>]*)\/?>/g,(_,a)=>{sheets.push({name:xmlDecode(attr(a,"name")||""), path:resolveTarget(relMap[attr(a,"r:id")]||"")});});
  for(const sh of sheets){
    const f=zip.file(sh.path); if(!f) continue;
    const xml=await f.async("string");
    const sd=xml.match(/<sheetData>([\s\S]*)<\/sheetData>|<sheetData\/>/); if(!sd||!sd[1]) continue;
    const rows=parseRows(sd[1]);
    let headerIdx=-1, headerMap={};
    for(let i=0;i<rows.length&&i<60;i++){
      const cells=parseCells(rows[i].inner, ss);
      if(Object.values(cells).some(c=>c.v.trim()==="EMP ID")){
        headerIdx=i;
        const norm=x=>x.replace(/\s+/g," ").trim().toLowerCase();
        const byNorm={}; for(const [col,c] of Object.entries(cells)){ const h=norm(c.v); if(h && !(h in byNorm)) byNorm[h]=col; }
        // map canonical names (case/space-insensitive) so "LINE MANAGER" = "Line Manager", "Name" = "NAME"
        for(const h of SPLIT_COLUMNS.concat(["EMP ID","Amount"])) if(byNorm[norm(h)]) headerMap[h]=byNorm[norm(h)];
        break;
      }
    }
    if(headerIdx<0) continue;
    const data=[];
    for(let i=headerIdx+1;i<rows.length;i++){
      const cells=parseCells(rows[i].inner, ss);
      if(!Object.values(cells).some(c=>c.v.trim()!=="")) continue; // fully empty row
      data.push({row:rows[i], cells});
    }
    // Head Count cell above header
    let headCount=null;
    for(let i=0;i<headerIdx;i++){
      const cells=parseCells(rows[i].inner, ss);
      const cols=Object.keys(cells).sort((a,b)=>colToNum(a)-colToNum(b));
      const li=cols.findIndex(c=>/^head\s*count/i.test(cells[c].v.trim()));
      if(li>=0 && cols[li+1]){ headCount={rowIdx:i, col:cols[li+1]}; break; }
      if(li>=0){ const nc=String.fromCharCode(cols[li].charCodeAt(cols[li].length-1)+1); headCount={rowIdx:i,col:nc,missing:true}; break; }
    }
    return {ss, sheet:sh, xml, sdMatch:sd, rows, headerIdx, headerMap, data, headCount,
      hasFormulas:data.some(d=>Object.values(d.cells).some(c=>c.hasFormula))};
  }
  throw new Error("NO_HEADER");
}

function groupData(info, selected){
  const cols=selected.map(h=>info.headerMap[h]);
  const empCol=info.headerMap["EMP ID"], amtCol=info.headerMap["Amount"];
  const groups=new Map();
  for(const d of info.data){
    const vals=cols.map(c=>(d.cells[c]?.v??"").trim()||"(blank)");
    const key=JSON.stringify(vals);
    if(!groups.has(key)) groups.set(key,{values:vals, items:[], emps:new Set(), amount:0});
    const g=groups.get(key); g.items.push(d);
    if(empCol && d.cells[empCol]?.v.trim()) g.emps.add(d.cells[empCol].v.trim());
    if(amtCol){ const n=parseFloat(d.cells[amtCol]?.v); if(!isNaN(n)) g.amount+=n; }
  }
  return [...groups.values()];
}

function safeName(s){ return s.replace(/[\\\/:*?"<>|\r\n\t]+/g,"-").replace(/\s+/g," ").trim().replace(/[. ]+$/,""); }

function renumberRow(xml, oldR, newR){
  let x=xml.replace(/^<row\b([^>]*?)\br="\d+"/, (m,pre)=>`<row${pre}r="${newR}"`);
  x=x.replace(/(<c\b[^>]*?\br=")([A-Z]+)\d+(")/g,(m,a,col,b)=>`${a}${col}${newR}${b}`);
  return x;
}

function setHeadCount(rowXml, col, rowNum, count){
  const ref=col+rowNum;
  const re=new RegExp('<c\\b([^>]*?)\\br="'+ref+'"([^>]*?)(?:\\/>|>([\\s\\S]*?)<\\/c>)');
  const m=rowXml.match(re);
  if(!m) return rowXml; // keep original if no value cell
  let a=' '+(m[1].trim()+' r="'+ref+'"'+m[2]).trim();
  const t=attr(a,"t");
  a=a.replace(/\s+t="[^"]*"/,"");
  let cell;
  if(t==="s"||t==="inlineStr"||t==="str") cell=`<c${a} t="inlineStr"><is><t>${count}</t></is></c>`;
  else cell=`<c${a}><v>${count}</v></c>`;
  return rowXml.replace(m[0], cell);
}

function buildSheetXml(info, group){
  const keep=info.rows.slice(0, info.headerIdx+1).map(r=>({row:r}));
  const all=keep.concat(group.items.map(d=>({row:d.row})));
  const out=all.map((it,i)=>{
    let x=renumberRow(it.row.xml, it.row.r, i+1);
    if(info.headCount && !info.headCount.missing && i===info.headCount.rowIdx)
      x=setHeadCount(x, info.headCount.col, i+1, group.emps.size);
    return x;
  });
  const last=all.length;
  const full=info.xml; const sd=info.sdMatch;
  let nx=full.slice(0, sd.index)+"<sheetData>"+out.join("")+"</sheetData>"+full.slice(sd.index+sd[0].length);
  nx=nx.replace(/(<dimension\b[^>]*\bref="[A-Z]+\d+:[A-Z]+)\d+(")/, (m,a,b)=>a+last+b);
  nx=nx.replace(/(<autoFilter\b[^>]*\bref="[A-Z]+\d+:[A-Z]+)\d+(")/, (m,a,b)=>a+last+b);
  const hdrRow=info.headerIdx+1;
  nx=nx.replace(/<mergeCells\b[^>]*>([\s\S]*?)<\/mergeCells>/, (m,body)=>{
    const keepM=(body.match(/<mergeCell\b[^>]*\/>/g)||[]).filter(mc=>{
      const r=attr(mc.slice(10),"ref")||""; const nums=r.match(/\d+/g)||[]; return nums.every(n=>parseInt(n,10)<=hdrRow);
    });
    return keepM.length?`<mergeCells count="${keepM.length}">${keepM.join("")}</mergeCells>`:"";
  });
  return nx;
}

async function buildFiles(JSZipLib, buffer, info, groups, baseName){
  const src=await JSZipLib.loadAsync(buffer);
  const names=new Set(); const files=[];
  for(const g of groups){
    const zip=new JSZipLib();
    for(const name of Object.keys(src.files)){
      const e=src.files[name]; if(e.dir) continue;
      if(name===info.sheet.path) zip.file(name, buildSheetXml(info,g), {date:e.date});
      else zip.file(name, await e.async("uint8array"), {date:e.date});
    }
    const blob=await zip.generateAsync({type: typeof Blob!=="undefined"?"blob":"nodebuffer", compression:"DEFLATE",
      compressionOptions:{level:6}, mimeType:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
    let fn=safeName(baseName+"_"+g.values.join("_")); let n=fn, k=2;
    while(names.has(n.toLowerCase())) n=`${fn} (${k++})`;
    names.add(n.toLowerCase());
    files.push({filename:n+".xlsx", data:blob, group:g});
  }
  return files;
}

if(typeof module!=="undefined") module.exports={SPLIT_COLUMNS, analyzeWorkbook, groupData, buildFiles};

(function(){
  const $=id=>document.getElementById(id);
  const ICON_ALERT='<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" aria-hidden="true"><circle cx="12" cy="12" r="9"/><path d="M12 8v5"/><path d="M12 16.5v.01"/></svg>';
  let buffer=null, info=null, baseName="", groups=[], cache=null, busy=false;
  let dl=null, dlReady=false;
  if(window.claude && typeof window.claude.use==="function"){
    window.claude.use("downloads").then(d=>{dl=d;dlReady=true;}).catch(()=>{dlReady=true;});
  } else dlReady=true;

  function esc(s){return String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");}
  function showMsg(el,text){ if(!text){el.classList.remove("show");el.innerHTML="";return;} el.innerHTML=ICON_ALERT+"<span>"+esc(text)+"</span>"; el.classList.add("show"); }
  const fmt=n=>n.toLocaleString("en-US",{minimumFractionDigits:2,maximumFractionDigits:2});

  // build checkbox list
  const colsEl=$("cols");
  SPLIT_COLUMNS.forEach((h,i)=>{
    const lab=document.createElement("label"); lab.className="col";
    lab.innerHTML='<input type="checkbox" value="'+esc(h)+'" id="c'+i+'"><span class="col__label">'+esc(h)+'</span>';
    colsEl.appendChild(lab);
  });
  let tickOrder=[];
  colsEl.addEventListener("change",e=>{ const v=e.target.value; tickOrder=tickOrder.filter(x=>x!==v); if(e.target.checked) tickOrder.push(v); refresh(); });

  // file input
  const drop=$("drop"), fileEl=$("file");
  fileEl.addEventListener("change",()=>{ if(fileEl.files[0]) load(fileEl.files[0]); });
  ["dragenter","dragover"].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.add("is-over");}));
  ["dragleave","drop"].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.remove("is-over");}));
  drop.addEventListener("drop",e=>{const f=e.dataTransfer.files[0]; if(f) load(f);});

  async function load(file){
    showMsg($("err"),""); $("fileinfo").classList.remove("show");
    info=null; buffer=null; cache=null; setStep(2,false); setStep(3,false);
    if(!/\.xlsx$/i.test(file.name)){ showMsg($("err"),"ไฟล์นี้ไม่ใช่ .xlsx กรุณาเลือกไฟล์ Excel นามสกุล .xlsx"); return; }
    if(typeof JSZip==="undefined"){ showMsg($("err"),"โหลดตัวอ่านไฟล์ไม่สำเร็จ ตรวจการเชื่อมต่ออินเทอร์เน็ตแล้วรีเฟรชหน้านี้"); return; }
    try{
      buffer=await file.arrayBuffer();
      const zip=await JSZip.loadAsync(buffer);
      info=await analyzeWorkbook(zip);
    }catch(e){
      buffer=null; info=null;
      showMsg($("err"), e&&e.message==="NO_HEADER" ? "ไม่พบหัวตาราง EMP ID ในไฟล์นี้ ตรวจว่าเป็นไฟล์ GL Invoice Report" : "เปิดไฟล์ไม่ได้ ไฟล์อาจเสียหายหรือถูกล็อกรหัสผ่านไว้");
      return;
    }
    baseName=file.name.replace(/\.xlsx$/i,"");
    const emps=new Set(); const ec=info.headerMap["EMP ID"];
    info.data.forEach(d=>{const v=d.cells[ec]?.v.trim(); if(v) emps.add(v);});
    $("fileinfo").innerHTML='<b>'+esc(file.name)+'</b><br><span>ชีต '+esc(info.sheet.name)+', '+info.data.length+' แถวข้อมูล, พนักงาน '+emps.size+' คน'+(info.headCount&&!info.headCount.missing?'':'<br>ไฟล์นี้ไม่มีช่อง Head Count จึงไม่มีการแก้ไขค่าใดเลย')+'</span>';
    $("fileinfo").classList.add("show");
    colsEl.querySelectorAll("input").forEach(inp=>{
      const ok=!!info.headerMap[inp.value]; inp.disabled=!ok; if(!ok) inp.checked=false;
      inp.closest(".col").title=ok?"":"ไม่พบคอลัมน์นี้ในไฟล์";
    });
    setStep(2,true); refresh();
  }

  function setStep(n,on){ $("s"+n).setAttribute("aria-disabled", on?"false":"true"); }
  function selected(){ const on=new Set([...colsEl.querySelectorAll("input:checked:not(:disabled)")].map(i=>i.value)); tickOrder=tickOrder.filter(v=>on.has(v)); return tickOrder.slice(); }

  function refresh(){
    cache=null; const sel=selected();
    const np=$("namepreview");
    if(!info){ np.textContent=""; return; }
    np.innerHTML = sel.length ? 'ตัวอย่างชื่อไฟล์: <code>'+esc(baseName)+'<em>_'+sel.map(esc).join("_")+'</em>.xlsx</code>' : "";
    showMsg($("warn"), info.hasFormulas ? "ไฟล์นี้มีสูตรในแถวข้อมูล หลังแยกไฟล์ควรเปิดตรวจว่าสูตรยังอ้างอิงถูกแถว" : "");
    if(!sel.length){
      groups=[]; setStep(3,false);
      $("summary").textContent="ติ๊กคอลัมน์อย่างน้อย 1 คอลัมน์เพื่อดูรายการไฟล์";
      $("result").innerHTML='<p class="empty">ยังไม่มีรายการ</p>'; buttons(false); return;
    }
    groups=groupData(info, sel);
    setStep(3,true);
    const hasAmt=!!info.headerMap["Amount"];
    $("summary").textContent="จะได้ "+groups.length+" ไฟล์ แยกตาม "+sel.join(" + ");
    let rows="", tot=0, totRows=0;
    groups.forEach((g,i)=>{
      tot+=g.amount; totRows+=g.items.length;
      rows+='<tr><td class="fn">'+esc(baseName+"_"+g.values.join("_"))+'.xlsx</td><td class="num">'+g.items.length+'</td><td class="num">'+g.emps.size+'</td>'+
        (hasAmt?'<td class="num">'+fmt(g.amount)+'</td>':'')+
        '<td class="num"><button class="btn btn--ghost" data-i="'+i+'">ดาวน์โหลด</button></td></tr>';
    });
    $("result").innerHTML='<div class="tablewrap"><table><thead><tr><th>ชื่อไฟล์</th><th class="num">แถว</th><th class="num">Head Count</th>'+(hasAmt?'<th class="num">Amount</th>':'')+'<th class="num"><span class="sr">ไฟล์เดียว</span></th></tr></thead><tbody>'+rows+
      '</tbody><tfoot><tr><td>รวม</td><td class="num">'+totRows+'</td><td></td>'+(hasAmt?'<td class="num">'+fmt(tot)+'</td>':'')+'<td></td></tr></tfoot></table></div>';
    buttons(true);
  }
  function buttons(on){ $("btnZip").disabled=!on||busy; $("btnEach").disabled=!on||busy; document.querySelectorAll("#result [data-i]").forEach(b=>b.disabled=!on||busy); }
  const status=t=>{$("status").textContent=t||"";};

  async function getFiles(){
    if(!cache) cache=await buildFiles(JSZip, buffer, info, groups, baseName);
    return cache;
  }
  const sleep=ms=>new Promise(r=>setTimeout(r,ms));
  async function waitDl(){ let t=0; while(!dlReady && t<10000){ await sleep(100); t+=100; } }

  // returns "saved" | "declined" | "error"
  async function saveFile(filename, blob){
    await waitDl();
    if(dl){
      for(let attempt=0;attempt<3;attempt++){
        try{ await dl.save({filename, data:blob}); return "saved"; }
        catch(e){
          const c=e&&e.code;
          if(c==="rate_limited"){ await sleep(1500); continue; }
          if(c==="declined") return "declined";
          return "error:"+(c||"unknown");
        }
      }
      return "error:rate_limited";
    }
    if(window.claude){ return "error:unavailable"; }
    const url=URL.createObjectURL(blob); const a=document.createElement("a");
    a.href=url; a.download=filename; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(()=>URL.revokeObjectURL(url),4000); await sleep(400); return "saved";
  }
  function errText(r){ return r==="error:unavailable"||r==="error:not_granted" ? "หน้านี้ดาวน์โหลดไฟล์ไม่ได้ในมุมมองนี้" : "ดาวน์โหลดไม่สำเร็จ ("+r.slice(6)+")"; }

  async function run(fn){
    if(busy) return; busy=true; buttons(true);
    try{ await fn(); }catch(e){ status("สร้างไฟล์ไม่สำเร็จ: "+(e&&e.message||e)); }
    busy=false; buttons(groups.length>0);
  }

  $("btnZip").addEventListener("click",()=>run(async()=>{
    status("กำลังสร้างไฟล์…");
    const files=await getFiles();
    const z=new JSZip(); files.forEach(f=>z.file(f.filename,f.data));
    const blob=await z.generateAsync({type:"blob",compression:"STORE"});
    const sel=selected().map(s=>s.replace(/\s+/g,"")).join("-");
    status("รอยืนยันการดาวน์โหลด…");
    const r=await saveFile(baseName+"_split_"+sel+".zip", blob);
    status(r==="saved"?"ดาวน์โหลด ZIP แล้ว ("+files.length+" ไฟล์)": r==="declined"?"ยกเลิกการดาวน์โหลด":errText(r));
  }));

  $("btnEach").addEventListener("click",()=>run(async()=>{
    status("กำลังสร้างไฟล์…");
    const files=await getFiles(); let ok=0;
    for(let i=0;i<files.length;i++){
      status("ไฟล์ "+(i+1)+" จาก "+files.length+": รอยืนยัน…");
      const r=await saveFile(files[i].filename, files[i].data);
      if(r==="saved") ok++;
      else if(r==="declined"){ status("หยุดที่ไฟล์ "+(i+1)+" ดาวน์โหลดแล้ว "+ok+" จาก "+files.length+" ไฟล์"); return; }
      else { status(errText(r)); return; }
    }
    status("ดาวน์โหลดแล้ว "+ok+" จาก "+files.length+" ไฟล์");
  }));

  $("result").addEventListener("click",e=>{
    const b=e.target.closest("[data-i]"); if(!b) return;
    const i=+b.dataset.i;
    run(async()=>{
      status("กำลังสร้างไฟล์…"); const files=await getFiles(); const f=files[i];
      status("รอยืนยันการดาวน์โหลด…");
      const r=await saveFile(f.filename,f.data);
      status(r==="saved"?"ดาวน์โหลด "+f.filename+" แล้ว": r==="declined"?"ยกเลิกการดาวน์โหลด":errText(r));
    });
  });
})();
})();
