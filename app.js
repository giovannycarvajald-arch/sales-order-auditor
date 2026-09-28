import * as pdfjsLib from "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/pdf.min.mjs";

pdfjsLib.GlobalWorkerOptions.workerSrc =
  "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/pdf.worker.min.mjs";

const $ = id => document.getElementById(id);
const state = { pdfFile:null, priceList:new Map(), items:[], procedureFile:null, procedureRules:null };

const PN_REGEX = /^\d{2,4}-\d{4}-\d{2}(?:-\d{2})?$/;

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
$("procedureInput").onchange = loadProcedure;
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

  // Odessa item parser V1.6
  // Do NOT require the line to start with "Ln# + PN". PDF.js can position
  // the line number and PN as separate text runs. Instead, locate every PN
  // globally inside the item-table text and parse the first Qty/UnitPrice/EA
  // Amount pattern that follows it.
  const itemRows=[];
  const pnRegex=/\b\d{2,4}-\d{4}-\d{2}(?:-\d{2})?\b/g;
  // In Odessa PDFs the "Return Policy" block is on page 1 BEFORE the
  // actual item rows on page 2. Therefore it cannot be used as item-table end.
  // Start at the first table header, then scan the rest of the document.
  const tableStart=t.search(/Ln#\s+Description\s+Rev\s+Quantity\s+Unit Price/i);
  const from=tableStart>=0 ? tableStart : 0;
  const itemArea=t.slice(from);
  const pns=[...itemArea.matchAll(pnRegex)];
  const seen=new Set();

  pns.forEach((m,idx)=>{
    const pn=m[0];
    if(seen.has(pn)) return;
    seen.add(pn);

    const next=pns[idx+1];
    const segment=itemArea.slice(m.index + pn.length, next ? next.index : itemArea.length)
      .replace(/\s+/g," ").trim();

    // First occurrence after the PN is the item's Qty / Unit Price / Amount.
    // This intentionally does not require it to be at the end of the segment,
    // because "Ship Dates" follows every item in the Odessa PDF.
    const priceMatch=segment.match(/(?:^|\s)(\d+(?:\.\d+)?)\s+([\d,]+\.\d{2})\s*\/EA(?:\s+(\d+(?:\.\d+)?)(?:%)?)?\s+([\d,]+\.\d{2})(?:\s|$)/);
    if(!priceMatch) return;

    const qty=Number(priceMatch[1]);
    const price=Number(priceMatch[2].replace(/,/g,""));
    const discount=priceMatch[3]==null?0:Number(priceMatch[3]);
    const amount=Number(priceMatch[4].replace(/,/g,""));

    let desc=segment.slice(0,priceMatch.index + (priceMatch[0].match(/^\s*/)?.[0].length||0)).trim();

    // Remove Rev immediately before Qty when present.
    const rev=desc.match(/^(.*)\s+(\d+)\s*$/);
    if(rev && Number(rev[2]) !== qty) desc=rev[1].trim();

    itemRows.push({pn,desc,qty,price,amount,discount});
  });

  itemRows.forEach(x=>addItem(x.pn,x.desc,x.qty,x.price,x.discount));
  $("status").textContent = `PDF leído: ${itemRows.length} ítem(s) detectado(s). Revisa los campos y ejecuta la auditoría.`;
  $("status").className = itemRows.length ? "status ok" : "status bad";
}

function addItem(pn="",desc="",qty=1,price="",discount=""){
  const tr=document.createElement("tr");
  tr.innerHTML=`<td><input class="pn" value="${escapeHtml(pn)}"></td>
    <td><input class="desc" value="${escapeHtml(desc)}"></td>
    <td><input class="qty" type="number" min="0" step="1" value="${qty}"></td>
    <td><input class="price" type="number" min="0" step="0.01" value="${price}"></td>
    <td><input class="discount" type="number" min="0" max="100" step="0.01" value="${discount}"></td>
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
  const file=e.target.files && e.target.files[0]; if(!file)return;
  const reader=new FileReader();
  reader.onload=ev=>{
    try{
      const wb=XLSX.read(ev.target.result,{type:"array",raw:true});
      const sheetName=wb.SheetNames.find(n=>normalize(n)==="PN LIST");
      if(!sheetName) throw new Error(`No se encontró la hoja "PN LIST". Hojas: ${wb.SheetNames.join(", ")}`);
      const rows=XLSX.utils.sheet_to_json(wb.Sheets[sheetName],{defval:"",raw:true});
      if(!rows.length) throw new Error("La hoja PN LIST está vacía.");
      const keys=Object.keys(rows[0]);
      const pnKey=keys.find(k=>normalize(k)==="PARTNUMBER");
      const priceKey=keys.find(k=>normalize(k)==="PRICING_UNITPRICE0");
      if(!pnKey || !priceKey) throw new Error("La hoja PN LIST debe contener PartNumber y Pricing_UnitPrice0.");
      const map=new Map();
      rows.forEach(row=>{
        const pn=normalize(row[pnKey]);
        if(!PN_REGEX.test(pn)) return;
        const raw=row[priceKey];
        const price=typeof raw==="number" ? raw : parseMoney(raw);
        if(price!=null && Number.isFinite(price)) map.set(pn,price);
      });
      if(!map.size) throw new Error("No se encontraron precios válidos en Pricing_UnitPrice0.");
      state.priceList=map;
      localStorage.setItem("soPriceList_v2",JSON.stringify([...map.entries()]));
      localStorage.setItem("soPriceListMeta_v2",JSON.stringify({fileName:file.name,sheetName,pnKey,priceKey,count:map.size}));
      localStorage.removeItem("soPriceList");
      document.querySelectorAll(".pn").forEach(x=>updateListPriceRow({target:x}));
      $("priceListStatus").textContent=`🟢 Price List cargada: ${map.size.toLocaleString()} PN | Hoja: ${sheetName} | ${priceKey}`;
      $("priceListStatus").className="status ok";
    }catch(err){ $("priceListStatus").textContent="❌ No se pudo leer el Excel: "+err.message; $("priceListStatus").className="status bad"; }
  };
  reader.readAsArrayBuffer(file);
}

function restorePriceList(){
  try{
    const raw=localStorage.getItem("soPriceList_v2");
    state.priceList=raw?new Map(JSON.parse(raw)):new Map();
    const meta=JSON.parse(localStorage.getItem("soPriceListMeta_v2")||"null");
    $("priceListStatus").textContent=state.priceList.size
      ? `🟢 Price List local disponible: ${state.priceList.size.toLocaleString()} PN${meta?.sheetName?` | ${meta.sheetName}`:""}`
      : "No hay Price List cargada todavía.";
    $("priceListStatus").className=state.priceList.size?"status ok":"status";
  }catch{ state.priceList=new Map(); $("priceListStatus").textContent="No hay Price List cargada todavía."; }
}

function cleanProcedureName(s){
  return normalize(String(s||"").replace(/[•\r\n]+/g," ").replace(/\s+/g," ").replace(/[.;]+$/g,"").trim());
}

function parseProcedureLists(text){
  const n=normalize(text);
  const heading="7. INFORMATION TO CONSIDER WHEN CREATING THE SALES ORDER";
  const start=n.lastIndexOf(heading);
  const source=start>=0 ? n.slice(start) : n;
  const m=source.match(/TAXABLE\s+([\s\S]*?)\s+STAMP\s+([\s\S]*?)\s+DISCOUNT\s+([\s\S]*?)\s+CONFIRMATION CHECK/);
  if(!m) throw new Error("No se encontraron las secciones TAXABLE, STAMP, DISCOUNT y CONFIRMATION CHECK en el procedure.");

  const taxable=m[1].split(",").map(cleanProcedureName).filter(Boolean);
  const stamp=m[2].split(",").map(cleanProcedureName).filter(Boolean);

  const discounts=[];
  const discountText=m[3];
  const pctRe=/\((\d+(?:\.\d+)?)%\)/g;
  let pm, lastEnd=0;
  while((pm=pctRe.exec(discountText))){
    const before=discountText.slice(lastEnd,pm.index);
    const comma=before.lastIndexOf(",");
    const rawName=before.slice(comma+1).trim();
    const name=cleanProcedureName(rawName);
    if(name) discounts.push({name,percent:Number(pm[1])});
    lastEnd=pm.index+pm[0].length;
  }
  // REV20 contains one discount written without parentheses: THRU TUBING 25%.
  const noParen=[...discountText.matchAll(/(?:^|,|\s)(THRU TUBING)\s+(\d+(?:\.\d+)?)%/g)];
  noParen.forEach(x=>discounts.push({name:cleanProcedureName(x[1]),percent:Number(x[2])}));
  // De-duplicate while preserving the first occurrence.
  const seenDiscount=new Set();
  const uniqueDiscounts=discounts.filter(x=>{ const k=`${x.name}|${x.percent}`; if(seenDiscount.has(k)) return false; seenDiscount.add(k); return true; });

  const rev=(source.match(/\bREV\s+(\d+)\b/)||[])[1] || "20";

  // Explicit aliases written by the procedure. REV20 states:
  // "EXXON MOBIL (ALL EXXON ORDERS WILL BE XTO)" in BOTH TAXABLE and STAMP.
  // Store the alias as structured data so the audit does not depend on how
  // PDF.js happened to split the line.
  const aliases=[];
  if(/EXXON MOBIL\s*\(\s*ALL EXXON ORDERS WILL BE XTO\s*\)/i.test(source)){
    aliases.push({from:"EXXON MOBIL",to:"XTO",reason:"ALL EXXON ORDERS WILL BE XTO"});
  }

  return {version:`REV ${rev}`, taxable, stamp, discounts:uniqueDiscounts, aliases, loadedAt:new Date().toISOString()};
}

function procedureCustomerMatch(customer, entry){
  const c=normalize(customer), e=normalize(entry);
  if(!c || !e) return false;

  // Structured aliases extracted from the loaded procedure.
  const aliases=state.procedureRules?.aliases || [];
  for(const a of aliases){
    if(normalize(a.to) && c.includes(normalize(a.to)) && e.includes(normalize(a.from))) return true;
    if(normalize(a.from) && c.includes(normalize(a.from)) && e.includes(normalize(a.from))) return true;
  }

  if(e==="ALL PUMP SHOPS") return /PUMP SHOP/.test(c);
  return c===e || c.includes(e) || e.includes(c);
}

function procedureRuleDisplayName(entry){
  const e=normalize(entry);
  if(e.includes("EXXON MOBIL") && e.includes("ALL EXXON ORDERS WILL BE XTO"))
    return "EXXON MOBIL → XTO";
  return entry;
}

function findProcedureDiscount(customer){
  const p=state.procedureRules;
  if(!p) return null;
  return p.discounts.find(x=>procedureCustomerMatch(customer,x.name)) || null;
}
function procedureHas(customer, list){
  return !!(list||[]).find(x=>procedureCustomerMatch(customer,x));
}

async function loadProcedure(e){
  const file=e.target.files && e.target.files[0];
  if(!file) return;
  try{
    const text=await extractPDFText(file);
    const rules=parseProcedureLists(text);
    state.procedureFile=file;
    state.procedureRules=rules;
    localStorage.setItem("soaProcedureRules_v1",JSON.stringify(rules));
    localStorage.setItem("soaProcedureMeta_v1",JSON.stringify({fileName:file.name,loadedAt:rules.loadedAt,version:rules.version}));
    const total=rules.taxable.length+rules.stamp.length+rules.discounts.length;
    $("procedureStatus").textContent=`🟢 Procedure cargado: ${file.name} | ${rules.version} | ${total} reglas (${rules.taxable.length} taxable, ${rules.stamp.length} stamp, ${rules.discounts.length} discount)`;
    $("procedureStatus").className="status ok";
  }catch(err){
    state.procedureRules=null;
    $("procedureStatus").textContent="❌ No se pudo leer el procedure: "+err.message;
    $("procedureStatus").className="status bad";
  }
}

function restoreProcedure(){
  try{
    const raw=localStorage.getItem("soaProcedureRules_v1");
    if(raw){
      state.procedureRules=JSON.parse(raw);
      const meta=JSON.parse(localStorage.getItem("soaProcedureMeta_v1")||"null");
      const p=state.procedureRules;
      $("procedureStatus").textContent=`🟢 Procedure local disponible${meta?.fileName?` | ${meta.fileName}`:""} | ${p.taxable.length} taxable · ${p.stamp.length} stamp · ${p.discounts.length} discount`;
      $("procedureStatus").className="status ok";
      return;
    }
  }catch{}
  $("procedureStatus").textContent="Carga el Sales Order Procedure para activar las reglas dinámicas.";
  $("procedureStatus").className="status";
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
    price:Number(tr.querySelector(".price").value||0),
    discount:Number(tr.querySelector(".discount")?.value||0)
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
    if(lp==null){ priceOK=false; priceMissing.push(`${x.pn}: PN NOT FOUND IN PRICE LIST`); }
    else if(Math.abs(Number(lp)-Number(x.price))>0.01){priceOK=false;priceMissing.push(`${x.pn}: SO ${money(x.price)} vs lista ${money(lp)}`);}
  });
  const priceListReady=state.priceList.size>0;
  const priceResultOK=priceListReady && priceOK;
  addResult(r,"Precios vs Excel",priceResultOK,priceListReady ? (priceMissing.length?priceMissing.join(" | "):"Precios coinciden con la Price List cargada."):"ERROR: Price List no cargada; no se puede verificar contra Excel.");
  if(!state.priceList.size) suggestions.push("Carga tu OSI LIST PRICE para activar la verificación de precios.");

  const procedureReady=!!state.procedureRules;
  let taxOK=true, taxDetail="";
  const taxableRequired=procedureReady && procedureHas(d.customer,state.procedureRules.taxable);
  const expectedTax=Math.round(d.subtotal*0.0825*100)/100;
  if(taxableRequired){
    taxOK=Math.abs(d.tax-expectedTax)<=0.01;
    const alias=state.procedureRules.aliases?.find(a=>
      normalize(a.to) && customerKey(d.customer)==="XTO" && normalize(a.to)==="XTO"
    );
    const aliasNote=alias ? ` Regla aplicada por alias: ${alias.from} → ${alias.to} (${alias.reason}).` : "";
    taxDetail=`Procedure: ${state.procedureRules.version} reconoce ${d.customer} como TAXABLE → Tax esperado 8.25%: ${money(expectedTax)}.${aliasNote}`;
  }else if(procedureReady){
    taxOK=Math.abs(d.tax)<0.01;
    taxDetail=taxOK?`Procedure: ${d.customer} no aparece en TAXABLE → Tax $0.00.`:`ERROR: ${d.customer} no aparece en TAXABLE pero la SO tiene Tax ${money(d.tax)}.`;
  }else{
    taxOK=false;
    taxDetail="ERROR: carga el Sales Order Procedure vigente antes de aprobar la SO.";
  }
  addResult(r,"Cálculo de TAXES",taxOK,taxDetail);

  const stampRequired=procedureReady ? procedureHas(d.customer,state.procedureRules.stamp) : false;
  const stampOK=procedureReady ? (!stampRequired || hasPhrase(d.notes,["PLEASE GET STAMPED","GET STAMP!","GET THE STAMP!"])) : false;
  addResult(r,"Frase STAMP",stampOK,
    procedureReady
      ? (stampRequired ? `Procedure: ${state.procedureRules.version} requiere STAMP; la SO debe contener una frase STAMP en Notes. Si la regla proviene de EXXON MOBIL, aplica a XTO por la nota “ALL EXXON ORDERS WILL BE XTO”.` : "La empresa no está listada en STAMP del procedure.")
      : "ERROR: carga el Sales Order Procedure vigente antes de aprobar la SO.");

  const discountRule=procedureReady ? findProcedureDiscount(d.customer) : null;
  if(procedureReady && discountRule){
    const discountErrors=[];
    d.items.forEach(x=>{
      if(/DELIVERY CHARGE|INSPECTION FEE/i.test(x.desc)) return;
      if(Math.abs((x.discount||0)-discountRule.percent)>0.01)
        discountErrors.push(`${x.pn}: descuento SO ${x.discount||0}% vs requerido ${discountRule.percent}%`);
    });
    addResult(r,"DISCOUNT",discountErrors.length===0,
      discountErrors.length?discountErrors.join(" | "):`Procedure: ${d.customer} requiere ${discountRule.percent}% y todos los ítems aplicables tienen ese descuento.`);
  }else if(procedureReady){
    addResult(r,"DISCOUNT",true,`La empresa ${d.customer} no aparece en la lista DISCOUNT del procedure.`);
  }else{
    addResult(r,"DISCOUNT",false,"ERROR: carga el Sales Order Procedure vigente antes de aprobar la SO.");
  }

  const allPriced=d.items.length>0 && d.items.every(x=>x.price>0);
  addResult(r,"Todos los ítems tienen precio",allPriced,allPriced?"Todos tienen precio.":"ERROR: uno o más ítems no tienen precio.");

  const grs=d.items.some(x=>/GRS|GAS RELEASE SYSTEM/i.test(x.desc));
  const hasReverseTrigger=d.items.some(x=>x.pn==="1015-0006-00"||x.pn==="1015-0008-00");
  if(!grs && hasReverseTrigger) suggestions.push("OBSERVACIÓN GRS: la configuración contiene un PN que puede requerir revisión de GRS/Reverse Flow. Confirmar configuración antes de aprobar.");

  if(c==="DIAMONDBACK" && !hasPhrase(d.notes,["AFE & GL","PLEASE PROVIDE AFE & GL ACCOUNT"]))
    suggestions.push("Diamondback: confirmar que estén presentes AFE & GL Account.");

  if(!procedureReady) suggestions.push("Carga el Sales Order Procedure REV20 para activar TAXABLE/STAMP/DISCOUNT dinámicos.");

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
    $("itemsBody").innerHTML=""; (d.items||[]).forEach(x=>addItem(x.pn,x.desc,x.qty,x.price,x.discount||0));
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

restorePriceList(); restoreProcedure(); loadDraft(); renderHistory();
