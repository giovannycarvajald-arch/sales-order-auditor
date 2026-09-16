import * as pdfjsLib from "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/pdf.min.mjs";

pdfjsLib.GlobalWorkerOptions.workerSrc =
  "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/pdf.worker.min.mjs";

const $ = id => document.getElementById(id);
const state = { pdfFile:null, priceList:new Map(), items:[] };

const RULES = {
  taxRate: 0.0825,
  taxableCustomers: ["XTO", "BTA", "BTA OIL & GAS", "BURLESON"],
  nonTaxableCustomers: ["COTERRA", "DIAMONDBACK"],
  noDeliveryCustomers: ["SUMMIT"],
  afterHoursDelivery: 450,
  endUserDelivery: 350,
  inspectionTrigger: "1014-0142-00",
  inspectionPN: "1004-0009-00",
  inspectionPrice: 450,
  stampCustomers: ["XTO"],
  diamondbackAFE: true
};

$("pdfInput").addEventListener("change", e => {
  state.pdfFile = e.target.files[0] || null;
  $("fileName").textContent = state.pdfFile ? state.pdfFile.name : "Ningún PDF seleccionado";
  $("analyzeBtn").disabled = !state.pdfFile;
});
$("dropzone").addEventListener("dragover", e => { e.preventDefault(); $("dropzone").classList.add("drag"); });
$("dropzone").addEventListener("dragleave", () => $("dropzone").classList.remove("drag"));
$("dropzone").addEventListener("drop", e => {
  e.preventDefault(); $("dropzone").classList.remove("drag");
  state.pdfFile = e.dataTransfer.files[0];
  if(state.pdfFile?.type === "application/pdf"){
    $("fileName").textContent = state.pdfFile.name; $("analyzeBtn").disabled = false;
  }
});
$("analyzeBtn").onclick = analyzePDF;
$("auditBtn").onclick = runAudit;
$("addItemBtn").onclick = () => addItem();
$("xlsxInput").onchange = loadPriceList;
$("saveDraftBtn").onclick = saveDraft;
$("exportHistoryBtn").onclick = exportHistory;
$("clearHistoryBtn").onclick = () => {
  if(confirm("¿Borrar todo el historial local?")) { localStorage.removeItem("soAuditHistory"); renderHistory(); }
};

function money(n){ return Number(n||0).toLocaleString("en-US",{style:"currency",currency:"USD"}); }
function normalize(s){ return String(s||"").toUpperCase().replace(/\s+/g," ").trim(); }
function customerKey(s){
  const x=normalize(s);
  if(x.includes("DIAMONDBACK")) return "DIAMONDBACK";
  if(x.includes("COTERRA")) return "COTERRA";
  if(x.includes("XTO")) return "XTO";
  if(x.includes("BTA")) return "BTA";
  if(x.includes("SUMMIT")) return "SUMMIT";
  if(x.includes("BURLESON")) return "BURLESON";
  if(x.includes("APACHE")) return "APACHE";
  return x;
}
function parseMoney(s){
  if(s==null) return null;
  const x=String(s).replace(/[$,]/g,"").match(/-?\d+(?:\.\d+)?/);
  return x ? Number(x[0]) : null;
}
function setValue(id,v){ $(id).value = v ?? ""; }

async function extractPDFText(file){
  const buf = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({data:buf}).promise;
  const pages=[];
  for(let i=1;i<=pdf.numPages;i++){
    const page=await pdf.getPage(i);
    const content=await page.getTextContent();
    const items=content.items.map(x=>({
      text:x.str || "",
      x:x.transform?.[4] ?? 0,
      y:x.transform?.[5] ?? 0
    })).filter(x=>x.text.trim());

    // Reconstruct visual lines from PDF text positions.
    const rows=[];
    for(const item of items){
      let row=rows.find(r=>Math.abs(r.y-item.y)<3);
      if(!row){ row={y:item.y,items:[]}; rows.push(row); }
      row.items.push(item);
    }
    rows.sort((a,b)=>b.y-a.y);
    const lines=rows.map(r=>r.items.sort((a,b)=>a.x-b.x).map(x=>x.text).join(" ").replace(/\s+/g," ").trim());
    pages.push(lines.join("\n"));
  }
  return pages.join("\n");
}
function firstMatch(text, regex){
  const m=text.match(regex); return m ? (m[1]||m[0]).trim() : "";
}

function parseSO(text){
  const lines=text.split(/\r?\n/).map(x=>x.replace(/\s+/g," ").trim()).filter(Boolean);
  const t=lines.join(" ");

  const so=firstMatch(t, /\bSO[-\s]?(\d{4,})\b/i);

  // Phone formats seen on Odessa SOs include "+ 1 4322907927".
  const phone=firstMatch(t, /(\+\s*1\s*\d{10}|\+?1[\s.-]?\(?\d{3}\)?[\s.-]?\d{3}[\s.-]?\d{4})/);

  const date=firstMatch(t, /\b(0?[1-9]|1[0-2])[\/\-](0?[1-9]|[12]\d|3[01])[\/\-](20\d{2})\b/);
  const time=firstMatch(t, /\b((?:0?[1-9]|1[0-2]):[0-5]\d\s*(?:AM|PM))\b/i);

  // Use the labels and the next meaningful line rather than a greedy
  // expression; this matches the actual Odessa PDF structure better.
  const soldIdx=lines.findIndex(x=>/^Sold To\s*:?\s*$/i.test(x));
  const shipIdx=lines.findIndex(x=>/^Ship To\s*:?\s*$/i.test(x));
  const customer=soldIdx>=0 ? (lines[soldIdx+1]||"") : firstMatch(t,/Sold To\s*:?\s*([A-Z][A-Z0-9 &.,'\/-]{2,80})\s+Ship To/i);
  const shipTo=shipIdx>=0 ? (lines[shipIdx+1]||"") : firstMatch(t,/Ship To\s*:?\s*([A-Z][A-Z0-9 &.,'\/-]{2,100})\s+(?:Delivery|Ticket|Shipping)/i);

  // Ship-to well is normally the line after the Ship To company.
  let shipToWell="";
  if(shipIdx>=0){
    const candidate=lines[shipIdx+2]||"";
    if(candidate && !/Odessa Separator|^\d{3}-\d{3}-\d{4}|SO\d+/i.test(candidate)) shipToWell=candidate.replace(/,$/,"").trim();
  }

  const contactLine=lines.find(x=>/^Contact\s*:/i.test(x))||"";
  const contact=contactLine.replace(/^Contact\s*:\s*/i,"").replace(/\s*-\s*(?:\+\s*1\s*)?\d[\d\s().-]+$/,"").trim();

  setValue("soNumber",so);
  setValue("customer",customer);
  setValue("soldTo",customer);
  setValue("shipTo",shipTo);
  setValue("shipToWell",shipToWell);
  setValue("phone",phone);
  setValue("contact",contact);

  if(date){
    const [m,d,y]=date.split(/[\/\-]/);
    setValue("deliveryDate",`${y}-${String(m).padStart(2,"0")}-${String(d).padStart(2,"0")}`);
  }
  if(time){
    const m=time.match(/(\d{1,2}):(\d{2})\s*(AM|PM)/i);
    let h=Number(m[1]); if(m[3].toUpperCase()==="PM" && h<12) h+=12; if(m[3].toUpperCase()==="AM" && h===12) h=0;
    setValue("deliveryTime",`${String(h).padStart(2,"0")}:${m[2]}`);
  }

  const noteText=t;
  setValue("notes",
    /PLEASE GET STAMPED/i.test(noteText) ? "PLEASE GET STAMPED" :
    /GET THE STAMP!/i.test(noteText) ? "GET THE STAMP!" :
    /GET STAMP!/i.test(noteText) ? "GET STAMP!" : ""
  );

  // Totals in Odessa PDFs can be stacked vertically:
  // Total -> amount -> Tax -> amount -> Subtotal -> amount.
  const totalIdx=lines.findIndex(x=>/^Total\s*$/i.test(x));
  let totalVal=null, taxVal=null, subtotalVal=null;
  if(totalIdx>=0){
    totalVal=parseMoney(lines[totalIdx+1]);
    for(let i=totalIdx+1;i<Math.min(lines.length,totalIdx+12);i++){
      if(/^Tax\s*$/i.test(lines[i])) taxVal=parseMoney(lines[i+1]);
      if(/^Subtotal\s*$/i.test(lines[i])) subtotalVal=parseMoney(lines[i+1]);
    }
  }
  // Fallback for inline formats.
  if(totalVal==null) totalVal=parseMoney(firstMatch(t, /Total\s*\$?\s*([\d,]+\.\d{2})/i));
  if(taxVal==null) taxVal=parseMoney(firstMatch(t, /(?:Sale Tax|Tax)\s*\$?\s*([\d,]+\.\d{2})/i));
  if(subtotalVal==null) subtotalVal=parseMoney(firstMatch(t, /Subtotal\s*\$?\s*([\d,]+\.\d{2})/i));

  setValue("subtotal",subtotalVal);
  setValue("tax",taxVal);
  setValue("total",totalVal);

  $("itemsBody").innerHTML="";
  state.items=[];

  // Exact Odessa item-row pattern:
  // PN + description + Rev + Qty + UnitPrice/EA + Amount.
  // We intentionally use the /EA marker to distinguish prices from
  // dimensions such as 2-7/8" and 1-1/4".
  const itemRows=[];
  for(const line of lines){
    const m=line.match(/^(\d+)\s+(\d{2,4}-\d{4}-\d{2}(?:-\d{2})?)\s+(.+?)\s+(\d+)\s+(\d+)\s+([\d,]+\.\d{2})\/EA\s+([\d,]+\.\d{2})$/);
    if(m){
      itemRows.push({
        pn:m[2],
        desc:m[3].trim(),
        qty:Number(m[5]),
        price:Number(m[6].replace(/,/g,""))
      });
    }
  }

  // If a row wraps over multiple PDF lines, reconstruct it by joining
  // lines between a PN and the next line containing /EA.
  if(itemRows.length===0){
    const pnRe=/^\d+\s+(\d{2,4}-\d{4}-\d{2}(?:-\d{2})?)\s+(.+)/;
    for(let i=0;i<lines.length;i++){
      const m=lines[i].match(pnRe); if(!m) continue;
      let chunk=m[2];
      for(let j=i+1;j<Math.min(lines.length,i+5);j++){
        chunk+=" "+lines[j];
        if(/[\d,]+\.\d{2}\/EA\s+[\d,]+\.\d{2}$/.test(chunk)) break;
      }
      const z=chunk.match(/^(.+?)\s+(\d+)\s+(\d+)\s+([\d,]+\.\d{2})\/EA\s+([\d,]+\.\d{2})$/);
      if(z) itemRows.push({pn:m[1],desc:z[1].trim(),qty:Number(z[3]),price:Number(z[4].replace(/,/g,""))});
    }
  }

  itemRows.forEach(x=>addItem(x.pn,x.desc,x.qty,x.price));

  $("status").textContent = `PDF leído: ${itemRows.length} ítem(s) detectado(s). Revisa los campos y ejecuta la auditoría.`;
  $("status").className="status ok";
}
function addItem(pn="",desc="",qty=1,price=""){
  const tr=document.createElement("tr");
  tr.innerHTML=`<td><input class="pn" value="${escapeHtml(pn)}"></td>
    <td><input class="desc" value="${escapeHtml(desc)}"></td>
    <td><input class="qty" type="number" min="0" step="1" value="${qty}"></td>
    <td><input class="price" type="number" min="0" step="0.01" value="${price}"></td>
    <td class="listPrice">—</td>
    <td><button class="danger remove">✕</button></td>`;
  tr.querySelector(".remove").onclick=()=>tr.remove();
  tr.querySelector(".pn").oninput=updateListPriceRow;
  tr.querySelector(".price").oninput=updateListPriceRow;
  $("itemsBody").appendChild(tr);
  updateListPriceRow({target:tr.querySelector(".pn")});
}
function escapeHtml(s){return String(s).replaceAll("&","&amp;").replaceAll("<","&lt;").replaceAll('"',"&quot;");}
function updateListPriceRow(e){
  const tr=e.target.closest("tr"), pn=normalize(tr.querySelector(".pn").value);
  const p=state.priceList.get(pn);
  tr.querySelector(".listPrice").textContent=p==null?"—":money(p);
}
function loadPriceList(e){
  const file=e.target.files[0]; if(!file)return;
  const reader=new FileReader();
  reader.onload=ev=>{
    try{
      const wb=XLSX.read(ev.target.result,{type:"array"});
      const sheet=wb.Sheets[wb.SheetNames[0]];
      const rows=XLSX.utils.sheet_to_json(sheet,{header:1,defval:""});
      const map=new Map();
      rows.forEach(row=>{
        const pn=row.find(x=>/^\s*\d{2,4}-\d{4}-\d{2}(?:-\d{2})?\s*$/.test(String(x)));
        if(pn){
          const idx=row.indexOf(pn);
          const candidates=row.slice(idx+1).map(parseMoney).filter(x=>x!=null);
          if(candidates.length) map.set(normalize(pn),candidates[0]);
        }
      });
      state.priceList=map;
      localStorage.setItem("soPriceList",JSON.stringify([...map.entries()]));
      document.querySelectorAll(".pn").forEach(x=>updateListPriceRow({target:x}));
      $("priceListStatus").textContent=`Price List cargada: ${map.size} part numbers.`;
      $("priceListStatus").className="status ok";
    }catch(err){ $("priceListStatus").textContent="No se pudo leer el Excel: "+err.message; $("priceListStatus").className="status bad"; }
  };
  reader.readAsArrayBuffer(file);
}
function restorePriceList(){
  try{ state.priceList=new Map(JSON.parse(localStorage.getItem("soPriceList")||"[]")); }
  catch{ state.priceList=new Map(); }
  $("priceListStatus").textContent=state.priceList.size?`Price List local disponible: ${state.priceList.size} part numbers.`:"No hay Price List cargada todavía.";
}
async function analyzePDF(){
  $("status").textContent="Leyendo PDF…"; $("status").className="status";
  try{ parseSO(await extractPDFText(state.pdfFile)); }
  catch(err){ $("status").textContent="Error leyendo PDF: "+err.message; $("status").className="status bad"; }
}

function getData(){
  const items=[...document.querySelectorAll("#itemsBody tr")].map(tr=>({
    pn:tr.querySelector(".pn").value.trim(),
    desc:tr.querySelector(".desc").value.trim(),
    qty:Number(tr.querySelector(".qty").value||0),
    price:Number(tr.querySelector(".price").value||0)
  })).filter(x=>x.pn);
  return {
    so:$("soNumber").value.trim(), customer:$("customer").value.trim(),
    soldTo:$("soldTo").value.trim(), shipTo:$("shipTo").value.trim(),
    soldToWell:$("soldToWell").value.trim(), shipToWell:$("shipToWell").value.trim(),
    contact:$("contact").value.trim(), phone:$("phone").value.trim(),
    date:$("deliveryDate").value, time:$("deliveryTime").value,
    deliveryText:$("deliveryText").value.trim(), notes:$("notes").value.trim(),
    subtotal:Number($("subtotal").value||0), tax:Number($("tax").value||0), total:Number($("total").value||0), items
  };
}

function addResult(list,name,ok,detail,kind=""){
  list.push({name,ok,detail,kind});
}
function deliveryExpected(d){
  const customer=customerKey(d.customer);
  const text=normalize(d.deliveryText);
  if(customer==="SUMMIT" || /CUSTOMER PICKUP|WAREHOUSE PICKUP|PICKUP/.test(text)) return {type:"NONE",price:0,reason:"No Delivery Charge"};
  if(!d.date || !d.time) return {type:"UNKNOWN",price:null,reason:"Fecha/hora no disponible"};
  const dt=new Date(`${d.date}T${d.time}:00`);
  const day=dt.getDay(), mins=dt.getHours()*60+dt.getMinutes();
  if(day===0 || day===6 || mins<480 || mins>1020) return {type:"AFTER HOURS",price:RULES.afterHoursDelivery,reason:"Fuera de horario laboral / fin de semana"};
  return {type:"END USER",price:RULES.endUserDelivery,reason:"Lunes a viernes, 8 AM–5 PM"};
}
function hasPhrase(text, phrases){ const n=normalize(text); return phrases.some(p=>n.includes(p)); }
function runAudit(){
  const d=getData(), r=[], suggestions=[];
  const c=customerKey(d.customer);

  addResult(r,"TAXES / STAMP / DISCOUNT / EXTRA INFO",true,"Revisión inicial ejecutada. Reglas específicas abajo.");
  const sold=customerKey(d.soldTo), ship=customerKey(d.shipTo);
  const companyOK=!sold||!ship||sold===ship;
  addResult(r,"Sold To vs Ship To",companyOK && (!d.soldToWell || !d.shipToWell || normalize(d.soldToWell)===normalize(d.shipToWell)),
    companyOK ? ((d.soldToWell&&d.shipToWell&&normalize(d.soldToWell)!==normalize(d.shipToWell)) ? "ERROR: well diferente entre Sold To y Ship To." : "Misma compañía / well consistente o no especificado.")
      : "ERROR: compañías diferentes entre Sold To y Ship To.");
  addResult(r,"Teléfono del contacto",!!d.phone,"Phone encontrado.");

  const de=deliveryExpected(d);
  const deliveryText=normalize(d.deliveryText);
  const deliveryPriceMatch=deliveryText.includes("AFTER HOURS")?RULES.afterHoursDelivery:deliveryText.includes("END USER")?RULES.endUserDelivery:null;
  let delOK=de.type==="UNKNOWN" ? true : de.type==="NONE" ? !deliveryPriceMatch : deliveryPriceMatch===de.price;
  addResult(r,"Delivery Charge",delOK,de.type==="UNKNOWN"?"No se pudo validar fecha/hora.":`${de.type}${de.price!=null?" — "+money(de.price):" — sin cargo"}. ${de.reason}`);

  const hasTrigger=d.items.some(x=>x.pn==="1014-0142-00");
  const hasInspection=d.items.some(x=>x.pn==="1004-0009-00");
  addResult(r,"Inspection Fee",!hasTrigger||hasInspection,hasTrigger&&!hasInspection?"ERROR: 1014-0142-00 requiere 1004-0009-00 Inspection Fee.":"Correcto.");

  let priceOK=true, priceMissing=[];
  d.items.forEach(x=>{
    const lp=state.priceList.get(normalize(x.pn));
    if(lp!=null && Math.abs(lp-x.price)>0.01){priceOK=false;priceMissing.push(`${x.pn}: SO ${money(x.price)} vs lista ${money(lp)}`);}
  });
  addResult(r,"Precios vs Excel",priceOK,state.priceList.size? (priceMissing.length?priceMissing.join(" | "):"Precios coinciden con la Price List cargada."):"Price List no cargada; no se puede verificar contra Excel.");
  if(!state.priceList.size) suggestions.push("Carga tu OSI LIST PRICE para activar la verificación de precios.");

  let taxOK=true, taxDetail="";
  if(RULES.nonTaxableCustomers.includes(c)){taxOK=Math.abs(d.tax)<0.01; taxDetail=`${c} tratado como no taxable → Tax esperado $0.00.`;}
  else if(RULES.taxableCustomers.includes(c)){const expected=Math.round(d.subtotal*RULES.taxRate*100)/100; taxOK=Math.abs(d.tax-expected)<=0.01; taxDetail=`Tax esperado 8.25%: ${money(expected)}.`;}
  else {taxDetail="Cliente no clasificado en las reglas V1; revisión manual requerida."; taxOK=true;}
  addResult(r,"Cálculo de TAXES",taxOK,taxDetail);

  const stampRequired=RULES.stampCustomers.includes(c);
  const stampOK=!stampRequired || hasPhrase(d.notes,["PLEASE GET STAMPED","GET STAMP!","GET THE STAMP!"]);
  addResult(r,"Frase STAMP",stampOK,stampRequired?"XTO requiere STAMP; frase aceptada si está presente.":"No se requiere STAMP para este cliente.");

  const allPriced=d.items.length>0 && d.items.every(x=>x.price>0);
  addResult(r,"Todos los ítems tienen precio",allPriced,allPriced?"Todos tienen precio.":"ERROR: uno o más ítems no tienen precio.");

  const grs=d.items.some(x=>/GRS|GAS RELEASE SYSTEM/i.test(x.desc));
  const hasReverseTrigger=d.items.some(x=>x.pn==="1015-0006-00"||x.pn==="1015-0008-00");
  if(!grs && hasReverseTrigger) suggestions.push("OBSERVACIÓN GRS: la configuración contiene un PN que puede requerir revisión de GRS/Reverse Flow. Confirmar configuración antes de aprobar.");

  if(c==="DIAMONDBACK" && !hasPhrase(d.notes,["AFE & GL","PLEASE PROVIDE AFE & GL ACCOUNT"]))
    suggestions.push("Diamondback: confirmar que estén presentes AFE & GL Account.");

  const errors=r.filter(x=>!x.ok);
  $("resultBanner").textContent=errors.length?`🔴 NOT APPROVED — ${errors.length} error(es)`:"🟢 APPROVED — 0 errores";
  $("resultBanner").className="result "+(errors.length?"bad":"ok");
  $("results").innerHTML=r.map(x=>`<div class="audit-row"><div class="name">${x.name}</div><div>${escapeHtml(x.detail)}</div><div class="${x.ok?'pass':'fail'}">${x.ok?'✓ OK':'✕ ERROR'}</div></div>`).join("");
  $("suggestions").innerHTML=suggestions.length?`<div class="suggestions"><strong>Correcciones / observaciones</strong><ul>${suggestions.map(x=>`<li>${escapeHtml(x)}</li>`).join("")}</ul></div>`:"";
  saveHistory(d,errors.length,r,suggestions);
}

function saveDraft(){
  localStorage.setItem("soDraft",JSON.stringify(getData()));
  $("status").textContent="Borrador guardado localmente."; $("status").className="status ok";
}
function loadDraft(){
  try{
    const d=JSON.parse(localStorage.getItem("soDraft")||"null"); if(!d)return;
    ["soNumber","customer","soldTo","shipTo","soldToWell","shipToWell","contact","phone","deliveryDate","deliveryTime","deliveryText","notes","subtotal","tax","total"].forEach(k=>setValue(k,d[k]));
    $("itemsBody").innerHTML=""; (d.items||[]).forEach(x=>addItem(x.pn,x.desc,x.qty,x.price));
  }catch{}
}
function saveHistory(d,errorCount,results,suggestions){
  const h=JSON.parse(localStorage.getItem("soAuditHistory")||"[]");
  h.unshift({so:d.so,customer:d.customer,date:new Date().toISOString(),status:errorCount?"NOT APPROVED":"APPROVED",errors:errorCount});
  localStorage.setItem("soAuditHistory",JSON.stringify(h.slice(0,100)));
  renderHistory();
}
function renderHistory(){
  const h=JSON.parse(localStorage.getItem("soAuditHistory")||"[]");
  $("history").innerHTML=h.length?h.map(x=>`<div class="history-item"><strong>${escapeHtml(x.so||"SO")}</strong> — ${escapeHtml(x.customer)} — ${x.status} — ${x.errors} error(es) <span class="small">${new Date(x.date).toLocaleString()}</span></div>`).join(""):"<span class='small'>No hay auditorías guardadas.</span>";
}
function exportHistory(){
  const blob=new Blob([localStorage.getItem("soAuditHistory")||"[]"],{type:"application/json"});
  const a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download="so-audit-history.json";a.click();URL.revokeObjectURL(a.href);
}

restorePriceList(); loadDraft(); renderHistory();
