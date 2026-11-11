import * as pdfjsLib from "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/pdf.min.mjs";

pdfjsLib.GlobalWorkerOptions.workerSrc =
  "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/pdf.worker.min.mjs";

const $ = id => document.getElementById(id);
const state = { pdfFile:null, priceList:new Map(), items:[], procedureFile:null, procedureRules:null, pdfLayout:[], shippingMethod:"", salesPerson:"" };

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
  $("fileName").textContent = state.pdfFile ? state.pdfFile.name : "No PDF selected";
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
$("loadDraftBtn").onclick = loadDraft;
$("exportHistoryBtn").onclick = exportHistory;
$("clearHistoryBtn").onclick = () => {
  if(confirm("Clear all local history?")) { localStorage.removeItem("soAuditHistory"); renderHistory(); }
};

function money(n){ return Number(n||0).toLocaleString("en-US",{style:"currency",currency:"USD"}); }
function normalize(s){ return String(s||"").toUpperCase().replace(/\s+/g," ").trim(); }
function customerKey(s){
  const x=normalize(s);
  if(x.includes("DIAMONDBACK")) return "DIAMONDBACK";
  if(x.includes("COTERRA")) return "COTERRA";
  if(/\bXTO(?:\s+ENERGY)?(?:\s+INC\.?)?\b/.test(x)) return "XTO";
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
  const buf=await file.arrayBuffer();
  const pdf=await pdfjsLib.getDocument({data:buf}).promise;
  const pages=[];
  const allTexts=[];
  for(let pageNo=1;pageNo<=pdf.numPages;pageNo++){
    const page=await pdf.getPage(pageNo);
    const content=await page.getTextContent();
    const items=content.items.map(x=>({
      text:x.str||"",
      x:x.transform?.[4]??0,
      y:x.transform?.[5]??0,
      w:x.width??0
    })).filter(x=>x.text.trim());

    const rows=[];
    for(const item of items){
      let row=rows.find(r=>Math.abs(r.y-item.y)<6);
      if(!row){row={y:item.y,items:[]};rows.push(row);}
      row.items.push(item);
    }
    rows.sort((a,b)=>b.y-a.y);
    rows.forEach(r=>r.items.sort((a,b)=>a.x-b.x));
    const lines=rows.map(r=>({
      y:r.y,
      items:r.items,
      text:r.items.map(x=>x.text).join(" ").replace(/\s+/g," ").trim()
    }));
    pages.push({pageNo,items,rows:lines});
    allTexts.push(lines.map(x=>x.text).join("\n"));
  }
  state.pdfLayout=pages;
  return allTexts.join("\n");
}

function findPartyColumn(label){
  const target=normalize(label).replace(/:$/,"");
  for(const page of state.pdfLayout||[]){
    for(const row of page.rows){
      const sorted=[...row.items].sort((a,b)=>a.x-b.x);
      let labelItem=null;
      for(let i=0;i<sorted.length;i++){
        const one=normalize(sorted[i].text).replace(/:$/,"");
        const two=i+1<sorted.length ? normalize(sorted[i].text+" "+sorted[i+1].text).replace(/:$/,"") : "";
        if(one===target || two===target){ labelItem=sorted[i]; break; }
      }
      if(!labelItem) continue;
      const isSold=target==="SOLD TO";
      let otherX=null;
      for(let i=0;i<sorted.length;i++){
        const one=normalize(sorted[i].text).replace(/:$/,"");
        const two=i+1<sorted.length ? normalize(sorted[i].text+" "+sorted[i+1].text).replace(/:$/,"") : "";
        const opposite=isSold?"SHIP TO":"SOLD TO";
        if(one===opposite || two===opposite){ otherX=sorted[i].x; break; }
      }
      if(otherX===null) continue;
      const soldX=isSold?labelItem.x:otherX;
      const shipX=isSold?otherX:labelItem.x;
      return {page,rowY:row.y,soldX,shipX,midpoint:(soldX+shipX)/2,labelX:labelItem.x};
    }
  }
  return null;
}

function columnBlockFromLayout(label){
  const found=findPartyColumn(label);
  if(!found) return [];
  const {page,rowY,midpoint,soldX,shipX}=found;
  const isSold=normalize(label)==="SOLD TO";
  const rows=page.rows
    .filter(r=>r.y<rowY && r.y>rowY-150)
    .map(r=>{
      const texts=r.items.filter(it=>{
        // Split the page into two independent columns using the midpoint
        // between the actual Sold To and Ship To headings. This prevents
        // the left company from leaking into the Ship To result.
        return isSold
          ? it.x >= soldX-10 && it.x < midpoint
          : it.x >= midpoint && it.x < (shipX+330);
      }).sort((a,b)=>a.x-b.x).map(it=>it.text);
      return {y:r.y,text:texts.join(" ").replace(/\\s+/g," ").trim()};
    })
    .filter(r=>r.text);
  return rows;
}

function extractShipWellDirect(){
  // Robust fallback for Odessa PDFs where the well is on the line directly
  // beneath the Ship To company, but the generic party parser misses it.
  for(const page of state.pdfLayout||[]){
    const found=findPartyColumn("Ship To");
    if(!found || found.page!==page) continue;
    const {rowY,shipX}=found;
    const rows=page.rows
      .filter(r=>r.y<rowY && r.y>rowY-90)
      .map(r=>({
        y:r.y,
        text:r.items
          .filter(it=>it.x>=shipX-8 && it.x<610)
          .sort((a,b)=>a.x-b.x)
          .map(it=>it.text).join(" ").replace(/\\s+/g," ").trim()
      }))
      .filter(r=>r.text);
    let companySeen=false;
    for(const row of rows){
      let candidate=cleanPartyLine(row.text);
      if(!candidate || isAddressLine(candidate)) continue;
      if(!companySeen){
        companySeen=true;
        // Sometimes PDF.js combines the company and well into one text row.
        const combined=candidate.match(/^(?:APACHE|XTO ENERGY INC|PIONEER(?: NATURAL RESOURCES)?|OCCIDENTAL PERMIAN)\\s+(.+\\d[^,]*)[,]?$/i);
        if(combined && looksLikeWell(combined[1])) return combined[1].trim();
        continue;
      }
      candidate=candidate.split(",")[0].trim();
      if(looksLikeWell(candidate)) return candidate;
    }
  }
  return "";
}

function cleanPartyLine(x){
  return String(x||"").replace(/\s+/g," ").replace(/,$/,"").trim();
}
function isAddressLine(x){
  const n=normalize(x);
  return !n ||
    /^(?:P\.?\s*O\.?\s+BOX)\s+\d+\b/.test(n) ||
    /^\d+\s+/.test(n) ||
    /\b(?:STREET|ST|ROAD|RD|AVENUE|AVE|HIGHWAY|HWY|DRIVE|DR|LANE|LN|BOULEVARD|BLVD|PARKWAY|PKWY|WAY)\b/.test(n) ||
    /\b(?:TX|OK|LA|CO|ND|NM|PA|KS|WY|UT)\s+\d{5}(?:-\d{4})?\b/.test(n) ||
    /^ODESSA SEPARATOR/.test(n) ||
    /^1001 E\. PEARL/.test(n) ||
    /^DELIVERY TICKET/.test(n) ||
    /^SHIPPING METHOD/.test(n) ||
    /^CUSTOMER P\/?O/.test(n) ||
    /^INVENTORY LOCATION/.test(n) ||
    /^432[-\s]/.test(n) ||
    /^SO\d+/.test(n);
}
function looksLikeWell(x){
  const n=normalize(x);
  if(!n || isAddressLine(n)) return false;
  // Reject fragments of the billing company/address that PDF text extraction
  // can split into a separate token row (for example "INC 810").
  if(/^(?:INC|LLC|CORP|CORPORATION|COMPANY)\s+\d+(?:\s+\d+)?$/.test(n)) return false;
  if(/^(?:XTO|ENERGY|BTA|APACHE|DIAMONDBACK|ODESSA)\s+(?:INC|CORP|LLC)?\s*\d*$/i.test(n)) return false;
  // A well must contain a number and enough identifying text to avoid
  // treating a company suffix or a street-number fragment as its name.
  if(!/\d/.test(n)) return false;
  const alphaTokens=n.split(/\s+/).filter(token=>/[A-Z]/.test(token));
  const numberTokens=n.split(/\s+/).filter(token=>/\d/.test(token));
  // Accept short but valid well names such as "SCRRU 207" (one word + one number).
  // PO BOX address lines are explicitly excluded by isAddressLine above.
  if(alphaTokens.length===0 || (alphaTokens.length===1 && numberTokens.length<1)) return false;
  if(/^(?:\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}|\d{1,2}:\d{2})/.test(n)) return false;
  if(/^(?:DELIVERY|SHIP DATES?|SALES PERSON|CONTACT|CUSTOMER P\/?O|INVENTORY|DELIVERY TICKET|SHIPPING METHOD)/i.test(n)) return false;
  return true;
}

function extractCompanyAndWell(label, fallbackText){
  const block=columnBlockFromLayout(label);
  if(block.length){
    const useful=[];
    for(const r of block){
      const candidate=cleanPartyLine(r.text);
      if(!candidate || isAddressLine(candidate)) continue;
      if(/^(?:DELIVERY TICKET|SHIPPING METHOD|CUSTOMER P\/?O|INVENTORY LOCATION|DELIVERY|SHIP DATES?|SALES PERSON|CONTACT)\b/i.test(candidate)) break;
      useful.push(candidate);
      if(useful.length>=3) break;
    }
    const company=useful[0]||"";
    let well="";
    for(let i=1;i<useful.length;i++){
      const candidate=useful[i];
      if(!looksLikeWell(candidate)) continue;

      // In some Odessa PDFs the well is split across two visual lines:
      // "DAVIDSON" / "HALE 62E 5H,". Rejoin the preceding name fragment.
      // A trailing comma marks the end of the Ship To well line; never append
      // city/address text such as "FORT WORTH" from later lines.
      let wellText=candidate.split(",")[0].trim();
      if(i>1){
        const previous=useful[i-1];
        if(previous && !/\d/.test(previous) && !isAddressLine(previous) &&
           !/^(?:XTO ENERGY INC|BTA OIL & GAS|COTERRA(?: ENERGY)?|DIAMONDBACK(?: ENERGY)?|APACHE(?: CORP| CORPORATION)?)$/i.test(previous)){
          wellText=previous+" "+wellText;
        }
      }
      if(looksLikeWell(wellText)){ well=wellText; break; }
    }
    return {company,well};
  }

  // Text-layer fallback. Look at the label, then the company and the next
  // non-address line. This handles PDFs where the column coordinates are
  // missing or the text is emitted in an unusual order.
  const raw=String(fallbackText||"");
  const lab=label.replace(/\s+/g," ");
  const other=label.toLowerCase()==="sold to" ? "Ship To" : "(?:Delivery Ticket|Odessa Separator Inc|Shipping Method)";
  const flat=raw.replace(/\s+/g," ");
  const rx=new RegExp(lab+"\\s*:?\\s*(.{2,160}?)(?=\\s+"+other+")","i");
  const m=flat.match(rx);
  if(m){
    const chunk=cleanPartyLine(m[1]);
    // Prefer a known company at the start of the block.
    const companyMatch=chunk.match(/^(PIONEER NATURAL RESOURCES|XTO ENERGY INC|BTA OIL & GAS|COTERRA(?: ENERGY)?|DIAMONDBACK(?: ENERGY)?|APACHE(?: CORP| CORPORATION)?|CHEVRON(?: USA)?|EXXON MOBIL)(?=\s|$)/i);
    const company=companyMatch?cleanPartyLine(companyMatch[1]):cleanPartyLine(chunk.split(/\s+(?=\d+\s)/)[0]||chunk);
    const rest=chunk.slice(company.length).trim();
    const parts=rest.split(/\s+(?=[A-Z][A-Z0-9 &'./-]*\d)/).map(x=>String(x||"").trim()).filter(Boolean);
    let well="";
    for(let i=0;i<parts.length;i++){
      // The well line may end with a comma before the city/state address.
      // Treat that comma as a hard boundary so "FORT WORTH" is never appended.
      let candidate=parts[i].split(",")[0].trim();
      if(!looksLikeWell(candidate)) continue;

      // PDF text extraction can split a multi-word well across runs:
      // "DAVIDSON" + "HALE 62E 5H,". Join a preceding alphabetic fragment.
      if(i>0){
        const previous=parts[i-1].split(",")[0].trim();
        const isKnownCompany=/^(?:XTO ENERGY INC|BTA OIL & GAS|COTERRA(?: ENERGY)?|DIAMONDBACK(?: ENERGY)?|APACHE(?: CORP| CORPORATION)?)$/i.test(previous);
        if(previous && !/\d/.test(previous) && !isKnownCompany && !isAddressLine(previous)){
          candidate=previous+" "+candidate;
        }
      }
      if(looksLikeWell(candidate)){ well=candidate; break; }
    }
    return {company,well};
  }
  return {company:"",well:""};
}

function firstMatch(text, regex){
  const m=text.match(regex); return m ? (m[1]||m[0]).trim() : "";
}

function parseSO(text){
  const lines=text.split(/\r?\n/).map(x=>x.replace(/\s+/g," ").trim()).filter(Boolean);
  const t=lines.join(" ");

  const so=firstMatch(t, /\bSO[-\s]?(\d{4,})\b/i);
  const shippingMethodMatch=t.match(/Shipping\s+Method\s*:?\s*(Field\s+Service)\b/i);
  state.shippingMethod=shippingMethodMatch?shippingMethodMatch[1].replace(/\s+/g," ").trim():"";

  // Phone formats seen on Odessa SOs include "+ 1 4322907927".
  const phone=firstMatch(t, /(\+\s*1\s*\d{10}|\+?1[\s.-]?\(?\d{3}\)?[\s.-]?\d{3}[\s.-]?\d{4})/);

  // Delivery date/time can appear in the Notes as:
  //   Delivery: 9/24/2026 4:00 PM
  // but many Odessa SOs also repeat the date under every item as:
  //   Ship Dates:
  //   1 by 9/24/2026
  // Prefer the explicit Delivery line, then fall back to Ship Dates.
  const deliveryLine = t.match(/Delivery\s*:\s*(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4})(?:\s+((?:0?[1-9]|1[0-2]):[0-5]\d\s*(?:AM|PM)))?/i);
  const shipDateLine = t.match(/Ship\s+Dates?\s*:\s*\d+\s+by\s+(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4})/i);
  const date = deliveryLine?.[1] || shipDateLine?.[1] || '';
  const time = deliveryLine?.[2] || '';

  // Sold To / Ship To are two side-by-side columns in the Odessa PDF.
  // The previous parser flattened both columns into one line, which could
  // produce an empty Customer and a Ship To containing the entire address.
  const soldBlock=extractCompanyAndWell("Sold To",t);
  const shipBlock=extractCompanyAndWell("Ship To",t);

  // Company names are extracted from each column's own PDF coordinates.
  // Do not override Ship To with a regex over flattened text: PDF.js may emit
  // both companies on the same visual line, causing the Sold To company to
  // be selected for both fields.
  // Odessa SO structure: the company normally appears in both Sold To and
  // Ship To. The well can appear on the following line in the party block.
  // Customer must come from the company name, never from the address/well.
  let customer=soldBlock.company || shipBlock.company || "";
  let soldTo=soldBlock.company || customer || "";
  let shipTo=shipBlock.company || customer || "";
  let soldToWell=soldBlock.well || "";
  let shipToWell=shipBlock.well || "";
  if(!shipToWell) shipToWell=extractShipWellDirect();

  // Last-resort company detection from the flattened text. This is only a
  // fallback; coordinate-based extraction above is the primary parser.
  if(!customer){
    const companyMatch=t.match(/(?:Sold To|Ship To)\s*:\s*(XTO ENERGY INC|BTA OIL & GAS|COTERRA(?: ENERGY)?|DIAMONDBACK[^\s]*)/i);
    if(companyMatch) customer=companyMatch[1].trim();
    soldTo=soldTo||customer;
    shipTo=shipTo||customer;
  }
  // Odessa SO convention: the well name is in the SHIP TO block,
  // immediately below the customer/company line. Do not infer a well from
  // Sold To because the next line there is normally the billing address.
  soldToWell="";
  if(!shipToWell){
    for(const page of state.pdfLayout||[]){
      for(const row of page.rows){
        const rt=normalize(row.text);
        if(rt!=="SHIP TO:" && rt!=="SHIP TO") continue;

        // PDF coordinates increase downward. Look only BELOW the Ship To
        // header and stay inside the Ship To column.
        const labelItem=row.items.find(it=>normalize(it.text).includes("SHIP TO"));
        const soldItem=row.items.find(it=>normalize(it.text).includes("SOLD TO"));
        // Anchor extraction to the Ship To column itself, not merely
        // to the left edge of the Sold To column.
        const minX=labelItem ? labelItem.x-10 : (soldItem ? soldItem.x+350 : 0);
        const nextRows=page.rows
          .filter(r=>r.y<row.y && r.y>row.y-120)
          .map(r=>({
            y:r.y,
            text:r.items
              .filter(it=>it.x>=minX)
              .sort((a,b)=>a.x-b.x)
              .map(it=>it.text).join(" ").replace(/\s+/g," ").trim()
          }))
          .filter(r=>r.text);

        // First non-address line is the company; the next suitable line is
        // the well. This avoids interpreting "810 HOUSTON STREET" as a well.
        const candidates=nextRows.map(r=>cleanPartyLine(r.text)).filter(x=>x && !isAddressLine(x));
        if(candidates.length>1){
          const well=candidates.slice(1).find(x=>looksLikeWell(x));
          if(well) shipToWell=well;
        }
      }
    }
  }

  const salespersonLine=lines.find(x=>/^Sales\s*person\s*:/i.test(x))||"";
  state.salesPerson=salespersonLine.replace(/^Sales\s*person\s*:\s*/i,"").trim();

  const contactLine=lines.find(x=>/^Contact\s*:/i.test(x))||"";
  const contact=contactLine.replace(/^Contact\s*:\s*/i,"").replace(/\s*-\s*(?:\+\s*1\s*)?\d[\d\s().-]+$/,"").trim();

  setValue("soNumber",so);
  setValue("customer",customer);
  setValue("soldTo",soldTo);
  setValue("shipTo",shipTo);
  setValue("soldToWell",soldToWell);
  setValue("shipToWell",shipToWell);
  setValue("phone",phone);
  setValue("contact",contact);

  if(date){
    const [m,d,y]=date.split(/[\/\-]/);
    setValue("deliveryDate",`${y}-${String(m).padStart(2,"0")}-${String(d).padStart(2,"0")}`);
  }
  if(time){
    const tm=time.match(/(\d{1,2}):(\d{2})\s*(AM|PM)/i);
    if(tm){
      let h=Number(tm[1]); if(tm[3].toUpperCase()==="PM" && h<12) h+=12; if(tm[3].toUpperCase()==="AM" && h===12) h=0;
      setValue("deliveryTime",`${String(h).padStart(2,"0")}:${tm[2]}`);
    }
  }
  // Keep a visible trace of where the delivery date came from. This helps
  // verify the parser when a SO uses Ship Dates instead of a Delivery note.
  if(date){
    const sourceLabel = deliveryLine ? "Delivery note" : "Ship Dates";
    const currentDeliveryText = $("deliveryText").value || "";
    if(!currentDeliveryText) setValue("deliveryText", `${sourceLabel}: ${date}${time ? ` ${time}` : ""}`);
  }

  // Preserve the complete Notes section for documented zero-charge/reinstalled items.
  const notesIdx=lines.findIndex(x=>/^Notes\s*:/i.test(x));
  let extractedNotes="";
  if(notesIdx>=0){
    const noteLines=[];
    for(let i=notesIdx;i<lines.length;i++){
      if(i>notesIdx && /^(?:Subtotal|Tax|Total|Sign|Date)\s*$/i.test(lines[i])) break;
      noteLines.push(lines[i]);
    }
    extractedNotes=noteLines.join(" ").replace(/^Notes\s*:\s*/i,"").trim();
  }
  // STAMP phrases may be highlighted outside the Notes block in Odessa SO PDFs.
  // Keep the full Notes text, while also preserving any recognized stamp instruction
  // found anywhere in the extracted PDF text (same behavior as the prior version).
  const recognizedStampPhrase=[
    "PLEASE GET STAMPED",
    "GET STAMP!",
    "GET THE STAMP!",
    "PLEASE PROVIDE AFE & GL ACCOUNT"
  ].find(phrase=>t.toUpperCase().includes(phrase));
  if(recognizedStampPhrase && !extractedNotes.toUpperCase().includes(recognizedStampPhrase)){
    extractedNotes=(extractedNotes+" "+recognizedStampPhrase).trim();
  }
  setValue("notes",extractedNotes);

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
    // Some SOs leave Unit Price blank and show only Amount $0.00 for a reinstalled tool.
    const zeroAmountMatch=!priceMatch ? segment.match(/(?:^|\s)(\d+(?:\.\d+)?)\s+0\.00(?:\s|$)/) : null;
    if(!priceMatch && !zeroAmountMatch) return;

    const rowMatch=priceMatch||zeroAmountMatch;
    const qty=Number(rowMatch[1]);
    const price=priceMatch ? Number(priceMatch[2].replace(/,/g,"")) : 0;
    const discount=priceMatch && priceMatch[3]!=null ? Number(priceMatch[3]) : 0;
    const amount=priceMatch ? Number(priceMatch[4].replace(/,/g,"")) : 0;

    let desc=segment.slice(0,rowMatch.index + (rowMatch[0].match(/^\s*/)?.[0].length||0)).trim();

    // Remove Rev immediately before Qty when present.
    const rev=desc.match(/^(.*)\s+(\d+)\s*$/);
    if(rev && Number(rev[2]) !== qty) desc=rev[1].trim();

    itemRows.push({pn,desc,qty,price,amount,discount});
  });

  itemRows.forEach(x=>addItem(x.pn,x.desc,x.qty,x.price,x.discount));
  $("status").textContent = `PDF read: ${itemRows.length} item(s) detected. Review the fields and run the audit.`;
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
      if(!sheetName) throw new Error(`Worksheet not found "PN LIST". Available sheets: ${wb.SheetNames.join(", ")}`);
      const rows=XLSX.utils.sheet_to_json(wb.Sheets[sheetName],{defval:"",raw:true});
      if(!rows.length) throw new Error("The PN LIST worksheet is empty.");
      const keys=Object.keys(rows[0]);
      const pnKey=keys.find(k=>normalize(k)==="PARTNUMBER");
      const priceKey=keys.find(k=>normalize(k)==="PRICING_UNITPRICE0");
      if(!pnKey || !priceKey) throw new Error("The PN LIST worksheet must contain PartNumber and Pricing_UnitPrice0.");
      const map=new Map();
      rows.forEach(row=>{
        const pn=normalize(row[pnKey]);
        if(!PN_REGEX.test(pn)) return;
        const raw=row[priceKey];
        const price=typeof raw==="number" ? raw : parseMoney(raw);
        if(price!=null && Number.isFinite(price)) map.set(pn,price);
      });
      if(!map.size) throw new Error("No valid prices were found in Pricing_UnitPrice0.");
      state.priceList=map;
      localStorage.setItem("soPriceList_v2",JSON.stringify([...map.entries()]));
      localStorage.setItem("soPriceListMeta_v2",JSON.stringify({fileName:file.name,sheetName,pnKey,priceKey,count:map.size,loadedAt:new Date().toISOString()}));
      localStorage.removeItem("soPriceList");
      document.querySelectorAll(".pn").forEach(x=>updateListPriceRow({target:x}));
      $("priceListStatus").textContent=`🟢 Price List loaded: ${map.size.toLocaleString()} PN | Hoja: ${sheetName} | ${priceKey}`;
      $("priceListStatus").className="status ok";
      renderConfigurationStatus();
    }catch(err){ $("priceListStatus").textContent="❌ Unable to read the Excel file: "+err.message; $("priceListStatus").className="status bad"; renderConfigurationStatus(); }
  };
  reader.readAsArrayBuffer(file);
}

function renderConfigurationStatus(){
  const priceMeta=JSON.parse(localStorage.getItem("soPriceListMeta_v2")||"null");
  const procedureMeta=JSON.parse(localStorage.getItem("soaProcedureMeta_v2")||"null");
  const priceReady=state.priceList instanceof Map && state.priceList.size>0;
  const procedureReady=!!state.procedureRules;
  const complete=priceReady && procedureReady;
  const details=[];
  details.push('<div class="config-file"><span class="config-dot '+(priceReady?"ready":"missing")+'"></span><div><strong>Price List</strong><small>'+(priceReady?(escapeHtml(priceMeta?.fileName||"Excel loaded")+" · "+(priceMeta?.count||state.priceList.size).toLocaleString()+" PN"+(priceMeta?.sheetName?" · Hoja "+escapeHtml(priceMeta.sheetName):"")):"Not uploaded")+'</small></div></div>');
  details.push('<div class="config-file"><span class="config-dot '+(procedureReady?"ready":"missing")+'"></span><div><strong>Sales Order Procedure</strong><small>'+(procedureReady?(escapeHtml(procedureMeta?.fileName||"Procedure cargado")+" · "+escapeHtml(state.procedureRules.version||"Versión no identificada")):"Not uploaded")+'</small></div></div>');
  const box=$("configurationStatus");
  if(!box) return;
  box.className="config-status "+(complete?"complete":"incomplete");
  box.innerHTML='<div class="config-heading"><div><strong>'+(complete?"✓ SETUP COMPLETE":"⚠ SETUP INCOMPLETE")+'</strong><small>'+(complete?"Both documents are available in this browser. You can now audit Sales Orders.":"Upload the missing documents to enable reliable auditing.")+'</small></div><span class="config-count">'+(Number(priceReady)+Number(procedureReady))+"/2"+'</span></div><div class="config-files">'+details.join("")+'</div><p class="config-footnote">Files remain stored locally in this browser and are not shared with other users.</p>';
}

function restorePriceList(){
  try{
    const raw=localStorage.getItem("soPriceList_v2");
    state.priceList=raw?new Map(JSON.parse(raw)):new Map();
    const meta=JSON.parse(localStorage.getItem("soPriceListMeta_v2")||"null");
    $("priceListStatus").textContent=state.priceList.size
      ? `🟢 Local Price List available: ${state.priceList.size.toLocaleString()} PN${meta?.sheetName?` | ${meta.sheetName}`:""}`
      : "No Price List has been uploaded yet.";
    $("priceListStatus").className=state.priceList.size?"status ok":"status";
    renderConfigurationStatus();
  }catch{ state.priceList=new Map(); $("priceListStatus").textContent="No Price List has been uploaded yet."; renderConfigurationStatus(); }
}

function cleanProcedureName(s){
  return normalize(String(s||"").replace(/[•\r\n]+/g," ").replace(/\s+/g," ").replace(/[.;]+$/g,"").trim());
}

function parseProcedureLists(text,fileName=""){
  // REV021 moved this section to page 17 and renumbered it from section 7
  // to section 2. Locate it by its stable title, not by page/section number.
  const n=normalize(text);
  const heading="INFORMATION TO CONSIDER WHEN CREATING THE SALES ORDER";
  const start=n.lastIndexOf(heading);
  const source=start>=0 ? n.slice(start+heading.length) : n;

  const taxableIdx=source.indexOf("TAXABLE");
  const stampIdx=source.indexOf("STAMP",taxableIdx>=0?taxableIdx+7:0);
  const discountIdx=source.indexOf("DISCOUNT",stampIdx>=0?stampIdx+5:0);
  const confirmationIdx=source.indexOf("CONFIRMATION CHECK",discountIdx>=0?discountIdx+8:0);

  if(taxableIdx<0||stampIdx<0||discountIdx<0||confirmationIdx<0){
    throw new Error("Could not correctly identify the TAXABLE, STAMP, DISCOUNT, and CONFIRMATION CHECK sections. Make sure you uploaded Procedure REV021 or a compatible revision.");
  }

  const taxableText=source.slice(taxableIdx+7,stampIdx).trim();
  const stampText=source.slice(stampIdx+5,discountIdx).trim();
  const discountText=source.slice(discountIdx+8,confirmationIdx).trim();
  const splitCustomers=section=>section.split(",").map(cleanProcedureName).filter(Boolean);
  const taxable=splitCustomers(taxableText);
  const stamp=splitCustomers(stampText);

  // Explicit Procedure relationships / aliases.
  const aliases=[];
  const exxonToXto=/EXXON MOBIL\s*\(\s*ALL EXXON ORDERS WILL BE XTO\s*\)/i.test(source)
    ||/ALL EXXON ORDERS WILL BE XTO/i.test(source);
  if(exxonToXto){
    aliases.push({from:"EXXON MOBIL",to:"XTO ENERGY",reason:"ALL EXXON ORDERS WILL BE XTO"});
    if(!taxable.some(x=>/\bXTO\b/.test(x))) taxable.push("XTO ENERGY");
    if(!stamp.some(x=>/\bXTO\b/.test(x))) stamp.push("XTO ENERGY");
  }

  const hasXtoTaxable=/\bXTO(?:\s+ENERGY)?(?:\s+INC\.?|\s+YELLOWJACKET)?\b/i.test(taxableText);
  const hasXtoStamp=/\bXTO(?:\s+ENERGY)?(?:\s+INC\.?|\s+YELLOWJACKET)?\b/i.test(stampText);
  if(hasXtoTaxable&&!taxable.some(x=>/\bXTO\b/i.test(x))) taxable.push("XTO ENERGY");
  if(hasXtoStamp&&!stamp.some(x=>/\bXTO\b/i.test(x))) stamp.push("XTO ENERGY");

  // Discount entries: REV021 contains both "(25%)" and "THRU TUBING 25%" formats.
  const discounts=[];
  const pctRe=/([^,]+?)\s*\((\d+(?:\.\d+)?)%\)/g;
  let pm;
  while((pm=pctRe.exec(discountText))){
    const name=cleanProcedureName(pm[1]);
    if(name) discounts.push({name,percent:Number(pm[2])});
  }
  const noParen=[...discountText.matchAll(/(?:^|,)\s*(THRU TUBING)\s+(\d+(?:\.\d+)?)%/g)];
  noParen.forEach(x=>discounts.push({name:cleanProcedureName(x[1]),percent:Number(x[2])}));
  const seenDiscount=new Set();
  const uniqueDiscounts=discounts.filter(x=>{
    const k=`${x.name}|${x.percent}`;
    if(seenDiscount.has(k)) return false;
    seenDiscount.add(k);
    return true;
  });

  // Preserve the operational instructions after the discount/tax section too.
  const extraInfoIdx=source.indexOf("EXTRA INFO",confirmationIdx);
  const extraInfo=extraInfoIdx>=0?source.slice(extraInfoIdx):"";
  const revisionMatch=String(fileName).match(/REV\s*0*(\d+)/i)||n.match(/REV\s*0*(\d+)/i);
  const rev=revisionMatch?.[1]?String(Number(revisionMatch[1])):"21";

  return {
    version:`REV ${rev}`,
    taxable:[...new Set(taxable)],
    stamp:[...new Set(stamp)],
    discounts:uniqueDiscounts,
    aliases,
    extraInfo,
    loadedAt:new Date().toISOString()
  };
}

function procedureCustomerMatch(customer, entry){
  const c=normalize(customer), e=normalize(entry);
  if(!c || !e) return false;

  const ck=customerKey(c);

  // Canonical XTO matching: Sales Orders can say XTO, XTO ENERGY,
  // or XTO ENERGY INC., while the Procedure may say XTO ENERGY.
  if(ck==="XTO" && /\bXTO(?:\s+ENERGY)?(?:\s+INC\.?)?\b/.test(e)) return true;

  const aliases=state.procedureRules?.aliases || [];
  for(const a of aliases){
    const from=normalize(a.from), to=normalize(a.to);
    if(ck==="XTO" && to && e.includes(to)) return true;
    if(ck==="XTO" && from && e.includes(from)) return true;
    if(from && c.includes(from) && (e.includes(from) || (to && e.includes(to)))) return true;
  }

  if(e.includes("EXXON MOBIL") && /ALL EXXON ORDERS WILL BE XTO/.test(e)) return ck==="XTO";
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
    const rules=parseProcedureLists(text,file.name);
    if(Number((rules.version.match(/\d+/)||[])[0]||0)<21)
      throw new Error("The app already requires REV021 or later. Upload Procedure REV021 or a newer revision.");
    state.procedureFile=file;
    state.procedureRules=rules;
    localStorage.setItem("soaProcedureRules_v2",JSON.stringify(rules));
    localStorage.setItem("soaProcedureMeta_v2",JSON.stringify({fileName:file.name,loadedAt:rules.loadedAt,version:rules.version}));
    const total=rules.taxable.length+rules.stamp.length+rules.discounts.length;
    const xtoTax=rules.taxable.some(x=>/\bXTO\b/i.test(x));
    const xtoStamp=rules.stamp.some(x=>/\bXTO\b/i.test(x));
    $("procedureStatus").textContent=`🟢 Procedure loaded: ${file.name} | ${rules.version} | ${total} rules (${rules.taxable.length} taxable, ${rules.stamp.length} stamp, ${rules.discounts.length} discount) | XTO TAXABLE: ${xtoTax?"YES":"NO"} | XTO STAMP: ${xtoStamp?"YES":"NO"}`;
    $("procedureStatus").className="status ok";
    renderConfigurationStatus();
  }catch(err){
    state.procedureRules=null;
    $("procedureStatus").textContent="❌ Unable to read the Procedure: "+err.message;
    $("procedureStatus").className="status bad";
    renderConfigurationStatus();
  }
}

function restoreProcedure(){
  try{
    const raw=localStorage.getItem("soaProcedureRules_v2");
    if(raw){
      state.procedureRules=JSON.parse(raw);
      const meta=JSON.parse(localStorage.getItem("soaProcedureMeta_v2")||"null");
      const p=state.procedureRules;
      const revision=Number((String(p.version||"").match(/\d+/)||[])[0]||0);
      if(revision<21){
        state.procedureRules=null;
        $("procedureStatus").textContent=`⚠ The saved Procedure is ${p.version||"an earlier version"}. Upload Procedure REV021 to enable the current rules.`;
        $("procedureStatus").className="status bad";
        return;
      }
      $("procedureStatus").textContent=`🟢 Local Procedure available: ${p.version}${meta?.fileName?` | ${meta.fileName}`:""} | ${p.taxable.length} taxable · ${p.stamp.length} stamp · ${p.discounts.length} discount`;
      $("procedureStatus").className="status ok";
      renderConfigurationStatus();
      return;
    }
  }catch{}
  $("procedureStatus").textContent="Upload the Sales Order Procedure to enable dynamic rules.";
  $("procedureStatus").className="status";
  renderConfigurationStatus();
}

async function analyzePDF(){
  $("status").textContent="Reading PDF…"; $("status").className="status";
  try{ parseSO(await extractPDFText(state.pdfFile)); }
  catch(err){ $("status").textContent="Error reading PDF: "+err.message; $("status").className="status bad"; }
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
    deliveryText:$("deliveryText").value.trim(), shippingMethod:state.shippingMethod||"", salesPerson:state.salesPerson||"", notes:$("notes").value.trim(),
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

  addResult(r,"TAXES / STAMP / DISCOUNT / EXTRA INFO",true,"Initial review completed. Specific checks are listed below.");
  const sold=customerKey(d.soldTo), ship=customerKey(d.shipTo);
  const companyOK=!sold||!ship||sold===ship;
  const wellsOK=!d.soldToWell||!d.shipToWell||normalize(d.soldToWell)===normalize(d.shipToWell);
  const companyDetail=!companyOK
    ? "ERROR: Sold To ("+d.soldTo+") and Ship To ("+d.shipTo+") are different companies. Confirm with the salesperson before approval."
    : "Companies match.";
  addResult(r,"Sold To vs Ship To",companyOK && wellsOK,
    !companyOK ? companyDetail
      : (!wellsOK ? "ERROR: Well differs between Sold To and Ship To." : "Same company / well matches or is not specified."));
  addResult(r,"Contact phone",!!d.phone,"Phone found.");

  const de=deliveryExpected(d);
  const deliveryText=normalize(d.deliveryText);
  const deliveryItems=d.items.filter(x=>/DELIVERY\s+CHARGE|AFTER\s+HOURS\s+DELIVERY|END\s+USER\s+DELIVERY/i.test(x.desc)||x.pn==="1004-0002-00");
  const hasEndUserItem=deliveryItems.some(x=>/END\s+USER\s+DELIVERY/i.test(x.desc)||x.pn==="1004-0002-00");
  const hasAfterHoursItem=deliveryItems.some(x=>/AFTER\s+HOURS\s+DELIVERY/i.test(x.desc));
  // A PICK UP CHARGE line means the customer collects the repaired tool at OSI.
  // Use the charge itself as the reliable signal because Shipping Method can be
  // missed by PDF text extraction even when it is visible in the document.
  const pickupChargeItems=d.items.filter(x=>
    /PICK\\s*UP\\s+CHARGE|PICKUP\\s+CHARGE/i.test(x.desc) || x.pn==="1004-0003-00"
  );
  const fieldServicePickup=pickupChargeItems.length>0;
  const deliveryItemAmount=deliveryItems.reduce((sum,x)=>sum+(Number(x.amount)||Number(x.price||0)*Number(x.qty||1)),0);
  const deliveryPriceMatch=
    deliveryText.includes("AFTER HOURS") ? RULES.afterHoursDelivery :
    deliveryText.includes("END USER") ? RULES.endUserDelivery :
    hasAfterHoursItem ? RULES.afterHoursDelivery :
    hasEndUserItem ? RULES.endUserDelivery : null;
  let delOK=true;
  let deliveryDetail="";
  let deliveryKind="";

  // PICK UP CHARGE never bypasses delivery validation. It only changes the
  // outcome when no outbound Delivery Charge is present: then confirm pickup.
  if(fieldServicePickup && deliveryItems.length===0){
    delOK=true;
    deliveryKind="warning";
    deliveryDetail=(normalize(d.shippingMethod)==="FIELD SERVICE"?"Shipping Method: FIELD SERVICE. ":"")+
      "PICK UP CHARGE found: "+pickupChargeItems.map(x=>x.pn+" "+x.desc+" ("+money(x.price)+" x "+x.qty+")").join(" | ")+
      ". No Delivery Charge was found. Verify whether the customer will pick up the tool at OSI after the repair.";
  }else if(de.type==="UNKNOWN"){
    // Without date/time the schedule and applicable tariff cannot be verified.
    // If a pickup and outbound charge coexist, flag the missing validation.
    delOK=!fieldServicePickup;
    deliveryKind=fieldServicePickup && deliveryItems.length>0?"warning":"";
    deliveryDetail="Date/time could not be validated; the Delivery Charge schedule and rate cannot be confirmed."+
      (fieldServicePickup?" PICK UP CHARGE and Delivery Charge were both found, but the charge must be confirmed against the schedule.":"");
  }else if(de.type==="NONE"){
    delOK=deliveryItems.length===0;
    deliveryDetail=deliveryItems.length
      ?"No Delivery Charge was expected for this case, but the following was found: "+deliveryItems.map(x=>x.pn+" "+x.desc+" ("+money(x.price)+" x "+x.qty+")").join(" | ")+"."
      :"No charge required. "+de.reason;
  }else{
    // Apply the same schedule/type/price checks whether or not PICK UP CHARGE
    // is present. A pickup charge does not make an incorrect delivery fee valid.
    const hasExpectedItem=de.type==="END USER" ? hasEndUserItem : hasAfterHoursItem;
    const expectedPrice=de.price;
    const itemPriceOK=deliveryItems.some(x=>
      (de.type==="END USER"
        ? (normalize(x.desc).includes("END USER DELIVERY") || x.pn==="1004-0002-00")
        : normalize(x.desc).includes("AFTER HOURS DELIVERY")) &&
      Math.abs(Number(x.price)-expectedPrice)<=0.01
    );
    delOK=hasExpectedItem && itemPriceOK;
    deliveryDetail=de.type+" — "+money(de.price)+". "+de.reason+
      (deliveryItems.length
        ?" Found on the Sales Order: "+deliveryItems.map(x=>x.pn+" "+x.desc+" ("+money(x.price)+" x "+x.qty+")").join(" | ")+"."
        :" No matching Delivery Charge item was found on the Sales Order.")+
      (fieldServicePickup?" PICK UP CHARGE was also found; the Delivery Charge was still validated against the applicable schedule and rate.":"");
  }
  addResult(r,"Delivery Charge",delOK,deliveryDetail,deliveryKind);

  const hasTrigger=d.items.some(x=>x.pn==="1014-0142-00");
  // Notes references that require an Inspection Fee for consignment,
  // inventory and refurbishment-related orders.
  const inspectionNoteTrigger=/\\b(FROM CONSIGNMENT|INVENTORY|REFURBISHMENT|REFURBISH|COI|REFURB)\\b/.test(notesNormalized);
  const hasInspection=d.items.some(x=>
    x.pn==="1004-0009-00" || /INSPECTION\\s+FEE/.test(normalize(x.desc))
  );
  const inspectionRequired=hasTrigger||inspectionNoteTrigger;
  const inspectionReasons=[
    hasTrigger?"PN 1014-0142-00":null,
    inspectionNoteTrigger?"referencia en Notes (consignment / inventory / refurbishment / COI)":null
  ].filter(Boolean);
  addResult(r,"Inspection Fee",!inspectionRequired||hasInspection,
    inspectionRequired&&!hasInspection
      ?"ERROR: INSPECTION FEE is required because "+inspectionReasons.join(" and ")+", but the fee was not found on the Sales Order."
      :inspectionRequired
        ?"Inspection Fee found. Rule triggered by "+inspectionReasons.join(" and ")+"."
        :"No condition requiring an Inspection Fee was found.");

  let priceErrors=[], specialPNs=[], reinstalledZeroItems=[];
  const notesNormalized=normalize(d.notes);
  d.items.forEach(x=>{
    const lp=state.priceList.get(normalize(x.pn));
    const pnInNotes=notesNormalized.includes(normalize(x.pn));
    const reinstalledMention=/\b(REINSTALLED|REINSTALL|REUSED|RE-USED|REUTILIZED|REUTILIZADO|REINSTALADO|REINSTALADA)\b/.test(notesNormalized);
    const justifiedZero=Number(x.price)===0 && pnInNotes && reinstalledMention;
    // REV021: ESP Vortex Desander sold to Permian Resources has a fixed
    // special unit price of $3,333, overriding the standard Price List.
    const permianVortexDesander=/PERMIAN RESOURCES/.test(normalize(d.customer)) &&
      /ESP VORTEX DESANDER/.test(normalize(x.desc));
    if(permianVortexDesander){
      if(Math.abs(Number(x.price)-3333)>0.01){
        priceErrors.push(`${x.pn}: ESP Vortex Desander para Permian Resources debe ser $3,333.00; SO ${money(x.price)}.`);
      }
    }else if(justifiedZero && lp!=null){
      reinstalledZeroItems.push({pn:x.pn,listPrice:lp});
    }else if(lp==null){
      // Special/custom part numbers may not be included in the standard OSI Price List.
      specialPNs.push(x.pn);
    }else if(Math.abs(Number(lp)-Number(x.price))>0.01){
      priceErrors.push(`${x.pn}: SO ${money(x.price)} vs lista ${money(lp)}`);
    }
  });
  const priceListReady=state.priceList.size>0;
  const priceResultOK=priceListReady && priceErrors.length===0;
  addResult(r,"Prices vs Excel",priceResultOK,
    priceListReady
      ? (priceErrors.length?priceErrors.join(" | "):"All recognized part numbers match the uploaded Price List.")
      : "ERROR: Price List not uploaded; prices cannot be checked against Excel.");
  if(specialPNs.length){
    const uniqueSpecialPNs=[...new Set(specialPNs)];
    addResult(r,"PN especiales not founds",true,
      `The following part numbers are not listed in the Price List: ${uniqueSpecialPNs.join(", ")}. Confirmar el nombre/descripción and el precio directamente con el vendedor antes de procesar la orden.`,"warning");
  }
  if(reinstalledZeroItems.length){
    addResult(r,"Reused / reinstalled tool",true,
      "Precio $0.00 aceptado porque las notas de la SO identifican el mismo PN como REINSTALLED/REUSED. Verificar que la herramienta fue reinstalada and que no corresponde cobrar una nueva unidad. PN: "+reinstalledZeroItems.map(x=>x.pn+" (precio de lista "+money(x.listPrice)+")").join(" | "),"warning");
  }
  if(!priceListReady) suggestions.push("Upload your OSI LIST PRICE to enable price verification.");

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
    taxDetail=`Procedure: ${state.procedureRules.version} identifies ${d.customer} as TAXABLE → Expected tax at 8.25%: ${money(expectedTax)}.${aliasNote}`;
  }else if(procedureReady){
    taxOK=Math.abs(d.tax)<0.01;
    taxDetail=taxOK?`Procedure: ${d.customer} is not listed as TAXABLE → Tax $0.00.`:`ERROR: ${d.customer} is not listed as TAXABLE, but the Sales Order includes tax ${money(d.tax)}.`;
  }else{
    taxOK=false;
    taxDetail="ERROR: upload the current Sales Order Procedure before approving the Sales Order.";
  }
  addResult(r,"TAX Calculation",taxOK,taxDetail);

  const stampRequired=procedureReady ? procedureHas(d.customer,state.procedureRules.stamp) : false;
  const stampOK=procedureReady ? (!stampRequired || hasPhrase(d.notes,["PLEASE GET STAMPED","GET STAMP!","GET THE STAMP!","PLEASE PROVIDE AFE & GL ACCOUNT"])) : false;
  addResult(r,"Frase STAMP",stampOK,
    procedureReady
      ? (stampRequired ? `Procedure: ${state.procedureRules.version} requires STAMP; the Sales Order Notes must contain a STAMP phrase. If the rule comes from EXXON MOBIL, it applies to XTO based on the note “ALL EXXON ORDERS WILL BE XTO”.` : "The company is not listed under STAMP in the Procedure.")
      : "ERROR: upload the current Sales Order Procedure before approving the Sales Order.");

  const discountRule=procedureReady ? findProcedureDiscount(d.customer) : null;
  if(procedureReady && discountRule){
    const discountErrors=[];
    d.items.forEach(x=>{
      if(/DELIVERY CHARGE|INSPECTION FEE/i.test(x.desc)) return;
      if(Math.abs((x.discount||0)-discountRule.percent)>0.01)
        discountErrors.push(`${x.pn}: Sales Order discount ${x.discount||0}% vs. required ${discountRule.percent}%`);
    });
    addResult(r,"DISCOUNT",discountErrors.length===0,
      discountErrors.length?discountErrors.join(" | "):`Procedure: ${d.customer} requires ${discountRule.percent}% and todos los ítems aplicables tienen ese descuento.`);
  }else if(procedureReady){
    addResult(r,"DISCOUNT",true,`The company ${d.customer} is not listed under DISCOUNT in the Procedure.`);
  }else{
    addResult(r,"DISCOUNT",false,"ERROR: upload the current Sales Order Procedure before approving the Sales Order.");
  }

  const allPriced=d.items.length>0 && d.items.every(x=>x.price>0);
  addResult(r,"All items have a price",allPriced,allPriced?"All items have a price.":"ERROR: one or more items are missing a price.");

  const grs=d.items.some(x=>/GRS|GAS RELEASE SYSTEM/i.test(x.desc));
  const hasReverseTrigger=d.items.some(x=>x.pn==="1015-0006-00"||x.pn==="1015-0008-00");
  if(!grs && hasReverseTrigger) suggestions.push("GRS NOTE: The configuration contains a part number that may require a GRS/Reverse Flow review. Confirm the configuration before approval.");

  if(c==="DIAMONDBACK" && !hasPhrase(d.notes,["AFE & GL","PLEASE PROVIDE AFE & GL ACCOUNT"]))
    suggestions.push("Diamondback: confirm that AFE & GL Account are included.");

  // REV021 — Additional customer/product-specific instructions.
  const hasDesc=(pattern)=>d.items.some(x=>pattern.test(normalize(x.desc)));

  if(c==="DIAMONDBACK"){
    const grsItems=d.items.filter(x=>/\bGRS\b|GAS RELEASE SYSTEM/.test(normalize(x.desc)));
    if(grsItems.length){
      const bad=grsItems.filter(x=>Math.abs((x.discount||0)-25)>0.01);
      addResult(r,"Diamondback — GRS discount",bad.length===0,
        bad.length?bad.map(x=>`${x.pn}: descuento ${x.discount||0}% vs. required 25%.`).join(" | "):
        "REV021: Diamondback GRS equipment must have a 25% discount.");
    }
  }

  if(/VALENCE OPERATING/.test(normalize(d.customer))){
    const signOK=hasPhrase(d.notes,["MUST GET COMPANY MAN SIGN"]);
    const taxExempt=Math.abs(d.tax)<0.01;
    addResult(r,"Valence — firma and tax exempt",signOK&&taxExempt,
      `REV021 requires la nota “MUST GET COMPANY MAN SIGN” and tratamiento Tax Exempt. Nota: ${signOK?"present":"MISSING"}; Tax: ${money(d.tax)}.`);
  }

  if(hasDesc(/SLOTTED GAS SHIELD/)){
    const dipTube=hasDesc(/DIP TUBE/);
    addResult(r,"Slotted Gas Shield — Dip Tube",dipTube,
      dipTube?"Dip Tube incluido.":"REV021: todo Slotted Gas Shield debe incluir un Dip Tube.");
  }

  if(hasDesc(/G[- ]?FORCE PACKERLESS/)){
    const companion=hasDesc(/OAL SPACER SUB|PUP JOINT/);
    addResult(r,"G-Force Packerless — accesorio",companion,
      companion?"OAL Spacer Sub o Pup Joint incluido.":"REV021: agregar un OAL Spacer Sub o un Pup Joint al G-Force Packerless.");
  }

  if(hasDesc(/SAFETY SHEAR JOINT/)){
    const screwCount=/\b\d+\s*(?:SCREWS?|BOLTS?)\b|\b(?:SCREWS?|BOLTS?)\s*[:#-]?\s*\d+\b/i.test(d.notes);
    addResult(r,"Safety Shear Joint — tornillos",true,
      screwCount?"La nota incluye la cantidad de tornillos.":"REV021: indicar en Notes la cantidad de tornillos requerida para el Safety Shear Joint.","warning");
  }

  if(/CONSIGNMENT/.test(notesNormalized)){
    const zeroValue=Math.abs(d.total)<0.01;
    addResult(r,"Consignment — Sales Order value",zeroValue,
      zeroValue?"Consignment order con valor total $0.00.":"REV021: consignment orders must have a Sales Order total of $0.00.");
  }

  if((/OXY|PERMIAN RESOURCES|COTERRA|APACHE/.test(normalize(d.customer))) &&
     hasDesc(/SAND LIFT|DESANDER/)){
    addResult(r,"Consignment — tool verification",true,
      "REV021: verificar el Excel de consignment en sitio and cruzar la información con Field Services; si hay herramientas en consignment, completar la hoja de reinstalación correspondiente.","warning");
  }

  if(/FIREBIRD/.test(normalize(d.customer))){
    addResult(r,"Firebird — entity change",true,
      "REV021 states that, effective 09/23/2026, all Firebird orders must be changed to Continental. Confirm that the order is created under Continental.","warning");
  }

  if(/CRESCENT ENERGY/.test(normalize(d.customer))){
    const crescentErrors=[];
    d.items.filter(x=>/DESANDER/.test(normalize(x.desc))).forEach(x=>{
      const expected=/BORONIZED/.test(normalize(x.desc))?4900:3000;
      if(Math.abs(x.price-expected)>0.01) crescentErrors.push(`${x.pn}: ${money(x.price)} vs. required ${money(expected)}.`);
    });
    if(d.items.some(x=>/DESANDER/.test(normalize(x.desc)))){
      addResult(r,"Crescent Energy — desander pricing",crescentErrors.length===0,
        crescentErrors.length?crescentErrors.join(" | "):"REV021: Desander $3,000 / Boronized Desander $4,900.");
    }
  }

  if(/OXY|OCCIDENTAL PERMIAN/.test(normalize(d.customer))){
    const compact=d.items.filter(x=>/SAND LIFT COMPACT/.test(normalize(x.desc)));
    if(compact.length){
      const bad=compact.filter(x=>Math.abs(x.price-5500)>0.01);
      addResult(r,"OXY / Occidental Permian — Sand Lift Compact",bad.length===0,
        bad.length?bad.map(x=>`${x.pn}: ${money(x.price)} vs. required $5,500.`).join(" | "):
        "REV021: Sand Lift Compact debe tener precio de $5,500.");
    }
  }

  if(c==="XTO" && /NEIL JOHNSON/.test(normalize(d.salesPerson)) &&
     d.items.some(x=>/GRSB|GAS RELEASE SYSTEM BYPASS/.test(normalize(x.desc)))){
    const afterHoursQty=d.items.filter(x=>/AFTER HOURS DELIVERY/.test(normalize(x.desc))).reduce((sum,x)=>sum+x.qty,0);
    addResult(r,"XTO Neil — After Hours para GRSB",afterHoursQty>=2,
      afterHoursQty>=2?"Se encontraron dos o más unidades de After Hours Delivery.":
      `REV021: todas las órdenes XTO de Neil Johnson con GRSB requieren dos cargos After Hours Delivery. Cantidad encontrada: ${afterHoursQty}.`);
  }



  if(hasDesc(/PMM/)){
    const sixScrews=/\b6\s+SCREWS?\b/i.test(d.notes);
    const pressure=/\b4500\s*PSI\b/i.test(d.notes);
    addResult(r,"PMM — configuración",true,
      sixScrews&&pressure?"Notes indica 6 screws and 4,500 psi.":
      "REV021: verificar que PMM tenga 6 screws, set pressure de 4,500 psi and etiqueta con presión and nombre del pozo.","warning");
  }

  if(/DONOVAN SANCHEZ/.test(normalize(d.salesPerson)) &&
     /RILEY PERMIAN|DEVON/.test(normalize(d.customer)) &&
     hasDesc(/COMPACT 350 SAND LIFT/)){
    const bad=d.items.filter(x=>/COMPACT 350 SAND LIFT/.test(normalize(x.desc))&&Math.abs(x.price-5500)>0.01);
    addResult(r,"Donovan — Compact 350 Sand Lift",bad.length===0,
      bad.length?bad.map(x=>`${x.pn}: ${money(x.price)} vs. required $5,500.`).join(" | "):
      "REV021: for Donovan with Riley Permian or Devon, Compact 350 Sand Lift must be priced at $5,500.");
  }

  if(/OXY/.test(normalize(d.customer))){
    addResult(r,"OXY — entidad del pozo",true,
      "REV021: verificar en DrillingEdge si el pozo pertenece a OXY NM u OXY USA; son entidades distintas and tienen tratamiento fiscal diferente.","warning");
  }

  if(/PERMIAN RESOURCES/.test(normalize(d.customer)) &&
     hasDesc(/ESP VORTEX DESANDER/)){
    const vortexItems=d.items.filter(x=>/ESP VORTEX DESANDER/.test(normalize(x.desc)));
    const wrong=vortexItems.filter(x=>Math.abs(Number(x.price)-3333)>0.01);
    addResult(r,"Permian Resources — ESP Vortex Desander",wrong.length===0,
      wrong.length
        ? wrong.map(x=>`${x.pn}: Sales Order price ${money(x.price)}; REV021 special price $3,333.00.`).join(" | ")
        : "REV021: ESP Vortex Desander for Permian Resources validated at the special price of $3,333.00 per unit.");
  }

  const cardFee=d.items.find(x=>x.pn==="1002-0002-00");
  if(/CREDIT CARD|CREDIT CARD PAYMENT|PAYMENT AT THE OFFICE/.test(notesNormalized)||cardFee){
    const feeActual=cardFee?cardFee.price*cardFee.qty:0;
    const baseValue=Math.max(0,d.subtotal-feeActual);
    const feeExpected=Math.round(baseValue*0.0375*100)/100;
    const feeOK=!!cardFee&&Math.abs(feeActual-feeExpected)<=0.02;
    addResult(r,"Credit card processing fee",feeOK,
      feeOK?`REV021: 3.75% fee correctly applied (${money(feeActual)}).`:
      `REV021: se requires PN 1002-0002-00 por el 3.75% del valor de la venta. Esperado: ${money(feeExpected)}; found: ${cardFee?money(feeActual):"not found"}. Also verify that TAX is applied to the fee when the customer is taxable.`);
  }

  if(procedureReady && !procedureHas(d.customer,state.procedureRules.taxable) &&
     !RULES.nonTaxableCustomers.some(x=>procedureCustomerMatch(d.customer,x)) &&
     !procedureHas(d.customer,state.procedureRules.stamp) &&
     !findProcedureDiscount(d.customer)){
    addResult(r,"Tax status — confirmar empresa",true,
      "REV021: the company is not listed under TAXABLE, STAMP, or DISCOUNT. If it is new or uncommon, confirm with the salesperson whether TAX should be applied.","warning");
  }

  if(!procedureReady) suggestions.push("Upload Sales Order Procedure REV021 or a later revision to enable dynamic TAXABLE/STAMP/DISCOUNT rules.");

  const errors=r.filter(x=>!x.ok);
  $("resultBanner").textContent=errors.length?`🔴 NOT APPROVED — ${errors.length} error(s)`:"🟢 APPROVED — 0 errores";
  $("resultBanner").className="result "+(errors.length?"bad":"ok");
  $("results").innerHTML=r.map(x=>`<div class="audit-row"><div class="name">${x.name}</div><div>${escapeHtml(x.detail)}</div><div class="${x.kind==='warning'?'obs':(x.ok?'pass':'fail')}">${x.kind==='warning'?'⚠ REVIEW':(x.ok?'✓ OK':'✕ ERROR')}</div></div>`).join("");
  $("suggestions").innerHTML=suggestions.length?`<div class="suggestions"><strong>Corrections / notes</strong><ul>${suggestions.map(x=>`<li>${escapeHtml(x)}</li>`).join("")}</ul></div>`:"";
  saveHistory(d,errors.length,r,suggestions);
}

function saveDraft(){
  localStorage.setItem("soDraft",JSON.stringify(getData()));
  $("status").textContent="Draft saved locally."; $("status").className="status ok";
}
function loadDraft(){
  try{
    const d=JSON.parse(localStorage.getItem("soDraft")||"null");
    if(!d){
      $("status").textContent="No saved draft was found.";
      $("status").className="status";
      return;
    }
    ["soNumber","customer","soldTo","shipTo","soldToWell","shipToWell","contact","phone","deliveryDate","deliveryTime","deliveryText","notes","subtotal","tax","total"].forEach(k=>setValue(k,d[k]));
    $("itemsBody").innerHTML=""; (d.items||[]).forEach(x=>addItem(x.pn,x.desc,x.qty,x.price,x.discount||0));
    $("status").textContent="Local draft loaded successfully.";
    $("status").className="status ok";
  }catch(err){
    $("status").textContent="Unable to load draft: "+err.message;
    $("status").className="status bad";
  }
}
function saveHistory(d,errorCount,results,suggestions){
  const h=JSON.parse(localStorage.getItem("soAuditHistory")||"[]");
  h.unshift({so:d.so,customer:d.customer,date:new Date().toISOString(),status:errorCount?"NOT APPROVED":"APPROVED",errors:errorCount});
  localStorage.setItem("soAuditHistory",JSON.stringify(h.slice(0,100)));
  renderHistory();
}
function renderHistory(){
  const h=JSON.parse(localStorage.getItem("soAuditHistory")||"[]");
  $("history").innerHTML=h.length?h.map(x=>`<div class="history-item"><strong>${escapeHtml(x.so||"SO")}</strong> — ${escapeHtml(x.customer)} — ${x.status} — ${x.errors} error(s) <span class="small">${new Date(x.date).toLocaleString()}</span></div>`).join(""):"<span class='small'>No audit records saved.</span>";
}
function exportHistory(){
  const blob=new Blob([localStorage.getItem("soAuditHistory")||"[]"],{type:"application/json"});
  const a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download="so-audit-history.json";a.click();URL.revokeObjectURL(a.href);
}

restorePriceList(); restoreProcedure(); renderHistory();
