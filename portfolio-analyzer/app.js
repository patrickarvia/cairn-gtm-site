const fieldDefs = [
  ["account","Account Name",["account name","customer name","company name","account","customer","company"]],
  ["arr","ARR / Revenue",["arr","annual recurring revenue","revenue","account arr","customer arr","contract value","annual value","acv","annual contract value","subscription value"]],
  ["renewalDate","Renewal Date",["renewal date","contract renewal date","next renewal date","renewal","contract end date","end date","expiration date","expiry date"]],
  ["health","Health",["health","health status","account health","customer health","health score","customer health score","account health score"]],
  ["lastActivity","Last Activity Date",["last activity date","last customer activity","last activity","last touch date","last touch","last engagement date","last engagement","last contacted"]],
  ["lastExecMeeting","Last Executive Meeting",["last executive meeting","last executive meeting date","executive meeting date","last exec meeting","last qbr date","last ebr date","last business review"]],
  ["champion","Champion",["champion","customer champion","champion identified","has champion"]],
  ["execSponsor","Executive Sponsor",["executive sponsor","exec sponsor","executive sponsor identified","economic buyer","senior sponsor"]],
  ["stakeholders","Stakeholder Count",["stakeholder count","number of stakeholders","active contact count","contact count","number of contacts","active contacts","stakeholders"]],
  ["adoption","Product Adoption",["product adoption","adoption","usage health","usage","adoption status","product usage","utilization"]],
  ["supportIssues","Open Support Issues",["open support issues","open tickets","support tickets","open cases","open support tickets","active cases","unresolved issues"]],
  ["nps","NPS / CSAT",["nps score","nps","net promoter score","csat","customer satisfaction","customer satisfaction score","satisfaction score"]],
  ["renewalStatus","Renewal Status",["renewal status","renewal stage","renewal forecast","renewal likelihood","renewal disposition"]],
  ["expansionPotential","Expansion Potential",["expansion potential","upsell potential","cross sell potential","growth potential","expansion signal","expansion opportunity"]],
  ["expansionAmount","Expansion Amount",["expansion amount","upsell amount","expansion arr","potential expansion","growth amount","upsell arr"]],
  ["owner","CSM / Owner",["csm","customer success manager","account owner","owner","relationship owner","success manager"]],
  ["segment","Segment",["segment","customer segment","account segment","tier","customer tier"]],
  ["region","Region",["region","customer region","territory","geo","geography"]]
];

const coreFieldKeys = new Set(["account","arr","renewalDate"]);
const numericFieldKeys = new Set(["arr","stakeholders","supportIssues","nps","expansionAmount"]);
const dateFieldKeys = new Set(["renewalDate","lastActivity","lastExecMeeting"]);
const booleanishFieldKeys = new Set(["champion","execSponsor","expansionPotential"]);
const genericAliases = new Set(["account","customer","company","revenue","owner","region","segment","health","usage","renewal"]);

let rawRows=[], headers=[], mapping={}, mappingMeta={}, analyzed=[];
let demoMode=false, pendingWorkbook=null, pendingFileName="", activeFileName="", activeSheetName="";

const $ = id => document.getElementById(id);
const esc = s => String(s==null?"":s).replace(/[&<>"']/g, m => ({"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#039;"}[m]));

function normalize(s){
  return String(s||"")
    .replace(/([a-z0-9])([A-Z])/g,"$1 $2")
    .trim().toLowerCase()
    .replace(/[^a-z0-9]+/g," ")
    .replace(/\s+/g," ")
    .trim();
}
function parseCSV(text){
  const rows=[]; let row=[], cell="", quoted=false;
  for(let i=0;i<text.length;i++){
    const ch=text[i], nx=text[i+1];
    if(quoted){
      if(ch==="\"" && nx==="\""){cell+="\"";i++}
      else if(ch==="\""){quoted=false}
      else cell+=ch;
    }else{
      if(ch==="\"")quoted=true;
      else if(ch===","){row.push(cell);cell=""}
      else if(ch==="\n"){row.push(cell);rows.push(row);row=[];cell=""}
      else if(ch!=="\r")cell+=ch;
    }
  }
  if(cell.length||row.length){row.push(cell);rows.push(row)}
  if(!rows.length)return {headers:[],rows:[]};
  const hs=rows[0].map(v=>String(v||"").trim());
  const data=rows.slice(1).filter(r=>r.some(v=>String(v||"").trim()!=="")).map(r=>{
    const o={};hs.forEach((h,i)=>{if(h)o[h]=String(r[i]==null?"":r[i]).trim()});return o;
  });
  return {headers:hs.filter(Boolean),rows:data};
}
function parseGrid(grid){
  if(!Array.isArray(grid)||!grid.length)return {headers:[],rows:[]};
  const first=grid.findIndex(r=>Array.isArray(r)&&r.some(v=>String(v==null?"":v).trim()!==""));
  if(first<0)return {headers:[],rows:[]};
  const hs=grid[first].map(v=>String(v==null?"":v).trim());
  const data=grid.slice(first+1).filter(r=>Array.isArray(r)&&r.some(v=>String(v==null?"":v).trim()!=="")).map(r=>{
    const o={};hs.forEach((h,i)=>{if(h)o[h]=String(r[i]==null?"":r[i]).trim()});return o;
  });
  return {headers:hs.filter(Boolean),rows:data};
}
function sourceContextHTML(){
  if(demoMode)return '<span class="source-label">Source</span><span class="source-chip">Cairn sample portfolio</span><span class="source-chip sheet-chip">Demo data</span>';
  if(!activeFileName)return "";
  let html='<span class="source-label">Source</span><span class="source-chip" title="'+esc(activeFileName)+'">'+esc(activeFileName)+'</span>';
  if(activeSheetName)html+='<span class="source-chip sheet-chip" title="'+esc(activeSheetName)+'">Sheet: '+esc(activeSheetName)+'</span>';
  return html;
}
function renderSourceContext(){
  $("mappingSource").innerHTML=sourceContextHTML();
  $("resultsSource").innerHTML=sourceContextHTML();
}
function loadParsed(parsed,sourceLabel){
  headers=parsed.headers;rawRows=parsed.rows;
  if(!headers.length||!rawRows.length){
    alert("Cairn could not find usable rows in this "+(sourceLabel||"file")+". Make sure the first populated row contains column headers.");
    return false;
  }
  autoMap();renderMapping();renderSourceContext();
  $("sheetPicker").classList.add("hidden");
  $("mappingSection").classList.remove("hidden");
  $("results").classList.add("hidden");
  return true;
}
function loadText(text){return loadParsed(parseCSV(text),"CSV")}
function parseWorkbookSheet(workbook,sheetName){
  const sheet=workbook.Sheets[sheetName];
  if(!sheet)return {headers:[],rows:[]};
  const grid=XLSX.utils.sheet_to_json(sheet,{header:1,defval:"",raw:false,blankrows:false});
  return parseGrid(grid);
}
function useWorkbookSheet(sheetName){
  if(!pendingWorkbook||!sheetName)return;
  const parsed=parseWorkbookSheet(pendingWorkbook,sheetName);
  activeFileName=pendingFileName||"Excel workbook";
  activeSheetName=sheetName;
  if(loadParsed(parsed,'Excel sheet "'+sheetName+'"')){
    pendingWorkbook=null;pendingFileName="";
    $("mappingSection").scrollIntoView({behavior:"smooth",block:"start"});
  }
}
function openWorkbook(arrayBuffer,fileName){
  if(typeof XLSX==="undefined"){alert("Excel support could not load. Please refresh and try again.");return}
  try{
    const workbook=XLSX.read(arrayBuffer,{type:"array",cellDates:true});
    const names=workbook.SheetNames||[];
    if(!names.length){alert("Cairn could not find any worksheets in that Excel file.");return}
    pendingWorkbook=workbook;pendingFileName=fileName||"Excel workbook";
    if(names.length===1){useWorkbookSheet(names[0]);return}
    $("sheetSelect").innerHTML=names.map(n=>'<option value="'+esc(n)+'">'+esc(n)+'</option>').join("");
    $("sheetPickerNote").textContent=pendingFileName+" contains "+names.length+" sheets. Choose the one that contains your customer portfolio.";
    $("sheetPicker").classList.remove("hidden");
    $("mappingSection").classList.add("hidden");
    $("results").classList.add("hidden");
    $("sheetPicker").scrollIntoView({behavior:"smooth",block:"center"});
  }catch(err){
    console.error(err);
    alert("Cairn could not read that Excel file. Confirm it is a valid .xlsx or .xls workbook.");
  }
}
function handlePortfolioFile(file){
  if(!file)return;
  demoMode=false;activeFileName=file.name||"Customer portfolio";activeSheetName="";
  $("sampleReady").classList.add("hidden");
  $("sheetPicker").classList.add("hidden");
  const name=(file.name||"").toLowerCase();
  if(name.endsWith(".xlsx")||name.endsWith(".xls")){
    const reader=new FileReader();
    reader.onload=()=>openWorkbook(reader.result,file.name);
    reader.onerror=()=>alert("Cairn could not read that Excel file.");
    reader.readAsArrayBuffer(file);return;
  }
  if(name.endsWith(".csv")||file.type==="text/csv"||!name.includes(".")){
    const reader=new FileReader();
    reader.onload=()=>{activeSheetName="";loadText(reader.result)};
    reader.onerror=()=>alert("Cairn could not read that CSV file.");
    reader.readAsText(file);return;
  }
  alert("Please choose a CSV, .xlsx, or .xls customer portfolio export.");
}

function tokens(s){return new Set(normalize(s).split(" ").filter(Boolean))}
function tokenScore(a,b){
  const A=tokens(a),B=tokens(b);if(!A.size||!B.size)return 0;
  let common=0;A.forEach(v=>{if(B.has(v))common++});
  return common/Math.max(A.size,B.size);
}
function sourceValues(header,limit){
  return rawRows.map(r=>String(r[header]==null?"":r[header]).trim()).filter(Boolean).slice(0,limit||18);
}
function valueFit(key,header){
  const vals=sourceValues(header,16);if(!vals.length)return 0;
  if(numericFieldKeys.has(key)){
    const good=vals.filter(v=>Number.isFinite(parseFloat(v.replace(/[$,%\s,]/g,"")))).length/vals.length;
    return good>=.8?6:good>=.5?2:-4;
  }
  if(dateFieldKeys.has(key)){
    const good=vals.filter(v=>!isNaN(new Date(v))).length/vals.length;
    return good>=.8?6:good>=.5?2:-4;
  }
  if(booleanishFieldKeys.has(key)){
    const good=vals.filter(v=>/^(yes|no|y|n|true|false|1|0|high|medium|low|unknown|identified|confirmed|none)$/i.test(v)).length/vals.length;
    return good>=.7?4:0;
  }
  return 0;
}
function matchScore(key,label,syns,header){
  const h=normalize(header);let best=0,reason="";
  [label].concat(syns).forEach(raw=>{
    const s=normalize(raw);if(!s)return;
    let score=0,why="";
    if(h===s){score=genericAliases.has(s)?74:100;why=genericAliases.has(s)?"generic exact match":"exact field match"}
    else if(s.length>=5&&(h.includes(s)||s.includes(h))){
      const coverage=Math.min(h.length,s.length)/Math.max(h.length,s.length);
      score=78+Math.round(10*coverage);why="strong name match";
    }else{
      const overlap=tokenScore(h,s);
      if(overlap>=.75){score=76;why="strong token match"}
      else if(overlap>=.5&&Math.min(tokens(h).size,tokens(s).size)>=2){score=66;why="likely token match"}
    }
    if(score>best){best=score;reason=why}
  });
  best+=valueFit(key,header);
  return {score:Math.max(0,Math.min(100,best)),reason:reason};
}
function confidenceFor(score){
  if(score>=94)return {label:"High",cls:"high"};
  if(score>=78)return {label:"Good",cls:"good"};
  if(score>=60)return {label:"Check",cls:"check"};
  return {label:"Not mapped",cls:"unmapped"};
}
function autoMap(){
  mapping={};mappingMeta={};
  const candidates=[];
  fieldDefs.forEach(def=>{
    const key=def[0],label=def[1],syns=def[2];
    headers.forEach(header=>{
      const m=matchScore(key,label,syns,header);
      if(m.score>=60)candidates.push({key:key,header:header,score:m.score,reason:m.reason});
    });
  });
  candidates.sort((a,b)=>b.score-a.score);
  const usedFields=new Set(),usedHeaders=new Set();
  candidates.forEach(c=>{
    if(usedFields.has(c.key)||usedHeaders.has(c.header))return;
    mapping[c.key]=c.header;mappingMeta[c.key]={score:c.score,reason:c.reason,manual:false};
    usedFields.add(c.key);usedHeaders.add(c.header);
  });
  fieldDefs.forEach(def=>{
    const key=def[0];
    if(!(key in mapping))mapping[key]="";
    if(!mappingMeta[key])mappingMeta[key]={score:0,reason:"No confident match",manual:false};
  });
}
function sampleText(header){
  if(!header)return "No source column selected";
  const vals=[];
  rawRows.some(row=>{
    const v=String(row[header]==null?"":row[header]).trim();
    if(v&&!vals.includes(v))vals.push(v);
    return vals.length===2;
  });
  return vals.length?"Sample: "+vals.map(v=>v.length>29?v.slice(0,28)+"…":v).join(" · "):"Selected column is mostly blank";
}
function setMapping(key,value){
  if(value){
    Object.keys(mapping).forEach(other=>{
      if(other!==key&&mapping[other]===value){
        mapping[other]="";mappingMeta[other]={score:0,reason:"Source column reassigned",manual:false};
      }
    });
  }
  mapping[key]=value;
  mappingMeta[key]=value?{score:100,reason:"Selected manually",manual:true}:{score:0,reason:"Not mapped",manual:true};
  renderMapping();
}
function updateMappingSummary(){
  const mapped=fieldDefs.filter(def=>mapping[def[0]]).length;
  const coreMapped=fieldDefs.filter(def=>coreFieldKeys.has(def[0])&&mapping[def[0]]).length;
  const used=new Set(Object.values(mapping).filter(Boolean));
  const ignored=headers.filter(h=>!used.has(h));
  const review=fieldDefs.filter(def=>mapping[def[0]]&&!mappingMeta[def[0]].manual&&mappingMeta[def[0]].score<78).length;
  $("mappingSummary").innerHTML="<strong>"+mapped+" of "+fieldDefs.length+" fields mapped</strong>"+coreMapped+"/"+coreFieldKeys.size+" core fields ready"+(ignored.length?" · "+ignored.length+" source column"+(ignored.length===1?"":"s")+" unused":"");
  const missing=fieldDefs.filter(def=>coreFieldKeys.has(def[0])&&!mapping[def[0]]).map(def=>def[1]);
  $("mappingActionNote").textContent=missing.length
    ?"Core fields still unmapped: "+missing.join(", ")+". Analysis will be less complete."
    :review
      ?review+" suggested match"+(review===1?"":"es")+" marked Check. Review "+(review===1?"it":"them")+" before analyzing."
      :"Core fields are ready. High-confidence matches do not need manual review.";
}
function renderMapping(){
  const grid=$("mappingGrid");grid.innerHTML="";
  const groups=[
    {title:"Core fields",note:"Minimum fields for a useful portfolio read.",core:true},
    {title:"Customer signals",note:"Optional evidence that strengthens risk and expansion analysis.",core:false}
  ];
  groups.forEach(group=>{
    const title=document.createElement("div");title.className="mapping-group-title";
    title.innerHTML="<strong>"+group.title+"</strong><span>"+group.note+"</span>";grid.appendChild(title);
    fieldDefs.filter(def=>coreFieldKeys.has(def[0])===group.core).forEach(def=>{
      const key=def[0],label=def[1],meta=mappingMeta[key]||{score:0,manual:false};
      const conf=meta.manual&&mapping[key]?{label:"Manual",cls:"manual"}:confidenceFor(meta.score||0);
      const div=document.createElement("div");
      div.className="map-item "+(conf.cls==="check"?"check":conf.cls==="unmapped"?"unmapped":"");
      const opts=['<option value="">Not mapped</option>'].concat(headers.map(h=>'<option value="'+esc(h)+'" '+(mapping[key]===h?"selected":"")+'>'+esc(h)+'</option>'));
      div.innerHTML='<div class="map-item-top"><label for="map-'+key+'">'+esc(label)+'</label><div class="map-badges"><span class="field-tag '+(group.core?"core":"signal")+'">'+(group.core?"Core":"Signal")+'</span><span class="match-badge '+conf.cls+'" title="'+esc(meta.reason||"")+'">'+conf.label+'</span></div></div><select id="map-'+key+'" data-key="'+key+'">'+opts.join("")+'</select><div class="map-sample"><b>'+(mapping[key]?esc(mapping[key]):"No match")+'</b> · '+esc(sampleText(mapping[key]))+'</div>';
      grid.appendChild(div);
    });
  });
  grid.querySelectorAll("select").forEach(sel=>sel.onchange=e=>setMapping(e.target.dataset.key,e.target.value));
  updateMappingSummary();
}

function val(row,key){const col=mapping[key];return col?String(row[col]==null?"":row[col]).trim():""}
function money(v){const n=parseFloat(String(v||"").replace(/[$,\s]/g,""));return Number.isFinite(n)?n:0}
function num(v){const n=parseFloat(String(v||"").replace(/[,%\s]/g,""));return Number.isFinite(n)?n:null}
function date(v){if(!v)return null;const d=new Date(v);return isNaN(d)?null:d}
function calendarDay(d){return new Date(Date.UTC(d.getFullYear(),d.getMonth(),d.getDate()))}
function daysBetween(a,b){return Math.round((b-a)/(1000*60*60*24))}
function fmtMoney(n){return new Intl.NumberFormat("en-US",{style:"currency",currency:"USD",maximumFractionDigits:0}).format(n||0)}
function fmtPct(n){return Math.round(n*100)+"%"}
function yes(v){return /^(yes|y|true|1|identified|confirmed|active|present|high)$/i.test(String(v||"").trim())}
function no(v){return /^(no|n|false|0|not identified|none|missing|inactive)$/i.test(String(v||"").trim())}
function unknown(v){return !String(v||"").trim()||/^(unknown|n\/a|na|tbd|not sure|not known|not documented)$/i.test(String(v||"").trim())}
function healthSignal(v){
  const s=normalize(v);const n=num(v);
  if(n!==null){
    if(n<=5&&n>=0)return n<=2?-2:n===3?0:2;
    if(n<=100&&n>=0)return n<45?-2:n<70?0:2;
  }
  if(/red|critical|risk|poor|unhealthy|churn|bad/.test(s))return -2;
  if(/yellow|amber|watch|mixed|medium|fair|neutral/.test(s))return 0;
  if(/green|healthy|good|strong|excellent|positive/.test(s))return 2;
  return null;
}
function adoptionSignal(v){
  const s=normalize(v);const n=num(v);
  if(n!==null&&n>=0){
    if(n<=5)return n<=2?-2:n===3?0:2;
    if(n<=100)return n<40?-2:n<70?0:2;
  }
  if(/low|poor|declin|inactive|under|weak|red/.test(s))return -2;
  if(/medium|moderate|mixed|average|yellow|amber/.test(s))return 0;
  if(/high|strong|active|growing|healthy|green/.test(s))return 2;
  return null;
}
function renewalSignal(v){
  const s=normalize(v);
  if(/churn|non renew|nonrenew|lost|cancel|at risk|risk/.test(s))return -2;
  if(/undecided|pending|unknown|not started|early|discuss/.test(s))return 0;
  if(/confirmed|committed|renewing|renewed|signed|complete/.test(s))return 2;
  return null;
}
function expansionSignal(v){
  const s=normalize(v);
  if(yes(v)||/high|strong|identified|open|active|qualified/.test(s))return 2;
  if(/medium|possible|potential|explor|early/.test(s))return 1;
  if(no(v)||/none|low|no signal/.test(s))return -1;
  return null;
}
function riskClass(score){return score>=70?"Critical":score>=22?"Watch":"Stable"}

function analyzeRow(row,totalPortfolioArr){
  const today=calendarDay(new Date());
  const account=val(row,"account")||"Unnamed account";
  const arr=money(val(row,"arr"));
  const renewalDateRaw=val(row,"renewalDate");
  const renewal=date(renewalDateRaw);
  const daysToRenewal=renewal?daysBetween(today,calendarDay(renewal)):null;
  const healthRaw=val(row,"health"), health=healthSignal(healthRaw);
  const lastActivityRaw=val(row,"lastActivity"), lastActivity=date(lastActivityRaw);
  const activityAge=lastActivity?daysBetween(calendarDay(lastActivity),today):null;
  const lastExecRaw=val(row,"lastExecMeeting"), lastExec=date(lastExecRaw);
  const execAge=lastExec?daysBetween(calendarDay(lastExec),today):null;
  const championRaw=val(row,"champion"), execRaw=val(row,"execSponsor");
  const stakeholders=num(val(row,"stakeholders"));
  const adoptionRaw=val(row,"adoption"), adoption=adoptionSignal(adoptionRaw);
  const support=num(val(row,"supportIssues"));
  const nps=num(val(row,"nps"));
  const renewalStatusRaw=val(row,"renewalStatus"), renewalStatus=renewalSignal(renewalStatusRaw);
  const expansionRaw=val(row,"expansionPotential"), expansion=expansionSignal(expansionRaw);
  const expansionAmount=money(val(row,"expansionAmount"));
  const owner=val(row,"owner"),segment=val(row,"segment"),region=val(row,"region");

  let riskScore=0;
  const risks=[],positives=[];

  if(daysToRenewal===null){riskScore+=8;risks.push("Renewal date is not documented")}
  else if(daysToRenewal<0){riskScore+=30;risks.push("Renewal date has passed")}
  else if(daysToRenewal<=30){riskScore+=20;risks.push("Renewal is due within 30 days")}
  else if(daysToRenewal<=60){riskScore+=12;risks.push("Renewal is due within 60 days")}
  else if(daysToRenewal<=90){riskScore+=7;risks.push("Renewal is due within 90 days")}
  else positives.push("Renewal is more than 90 days away");

  if(health===-2){riskScore+=25;risks.push("Customer health is negative")}
  else if(health===0){riskScore+=10;risks.push("Customer health is mixed")}
  else if(health===2)positives.push("Customer health is positive");
  else if(daysToRenewal!==null&&daysToRenewal<=90){riskScore+=5;risks.push("Health is not documented ahead of renewal")}

  if(activityAge===null){riskScore+=daysToRenewal!==null&&daysToRenewal<=90?12:5;risks.push("Recent customer activity is not documented")}
  else if(activityAge>90){riskScore+=22;risks.push("Customer activity is more than 90 days old")}
  else if(activityAge>60){riskScore+=16;risks.push("Customer activity is more than 60 days old")}
  else if(activityAge>30){riskScore+=8;risks.push("Customer activity is more than 30 days old")}
  else positives.push("Customer activity is recent");

  if(execAge===null&&daysToRenewal!==null&&daysToRenewal<=120){riskScore+=7;risks.push("Executive engagement is not documented near renewal")}
  else if(execAge!==null&&execAge>120){riskScore+=10;risks.push("Executive engagement is more than 120 days old")}
  else if(execAge!==null&&execAge<=90)positives.push("Executive engagement is recent");

  if(no(championRaw)){riskScore+=12;risks.push("No customer champion is documented")}
  else if(yes(championRaw))positives.push("Customer champion is documented");
  else if(unknown(championRaw)&&daysToRenewal!==null&&daysToRenewal<=90){riskScore+=5;risks.push("Champion status is unknown")}

  if(no(execRaw)){riskScore+=10;risks.push("No executive sponsor is documented")}
  else if(yes(execRaw))positives.push("Executive sponsor is documented");
  else if(unknown(execRaw)&&daysToRenewal!==null&&daysToRenewal<=90){riskScore+=4;risks.push("Executive sponsor status is unknown")}

  if(stakeholders!==null&&stakeholders<2){riskScore+=8;risks.push("Relationship coverage is narrow")}
  else if(stakeholders!==null&&stakeholders>=4)positives.push("Relationship coverage is broad");

  if(adoption===-2){riskScore+=18;risks.push("Product adoption is low")}
  else if(adoption===0){riskScore+=7;risks.push("Product adoption is moderate")}
  else if(adoption===2)positives.push("Product adoption is strong");

  if(support!==null&&support>=5){riskScore+=12;risks.push("Five or more support issues are open")}
  else if(support!==null&&support>=2){riskScore+=5;risks.push("Multiple support issues are open")}
  else if(support!==null&&support===0)positives.push("No open support issues are documented");

  if(nps!==null){
    if(nps<=6){riskScore+=10;risks.push("Customer satisfaction signal is negative")}
    else if(nps>=9)positives.push("Customer satisfaction signal is strong");
  }

  if(renewalStatus===-2){riskScore+=30;risks.push("Renewal status is explicitly at risk")}
  else if(renewalStatus===0&&daysToRenewal!==null&&daysToRenewal<=90){riskScore+=12;risks.push("Renewal decision is not yet confirmed")}
  else if(renewalStatus===2)positives.push("Renewal is confirmed or progressing positively");
  else if(renewalStatus===null&&daysToRenewal!==null&&daysToRenewal<=90){riskScore+=8;risks.push("Renewal status is not documented")}

  const risk=riskClass(riskScore);
  const explicitExpansion=expansion!==null&&expansion>0;
  const evidenceExpansion=(health===2&&adoption===2&&activityAge!==null&&activityAge<=30&&(yes(championRaw)||yes(execRaw)));
  const expansionCandidate=risk!=="Critical"&&(explicitExpansion||evidenceExpansion);
  let expansionReason="";
  if(expansionCandidate){
    if(expansion===2)expansionReason="Expansion potential is explicitly identified";
    else if(expansion===1)expansionReason="Expansion potential is being explored";
    else expansionReason="Strong health, adoption, engagement, and relationship evidence";
  }

  let nextAction="Maintain customer cadence and document the next customer-owned milestone.";
  if((daysToRenewal!==null&&daysToRenewal<=30)&&renewalStatus!==2)nextAction="Confirm the renewal decision, commercial path, and customer owner immediately.";
  else if(renewalStatus===-2||health===-2)nextAction="Open a customer recovery plan with a defined issue, owner, and success condition.";
  else if(activityAge===null||activityAge>60)nextAction="Re-establish customer contact and schedule a dated next meeting.";
  else if(no(championRaw)||unknown(championRaw))nextAction="Identify or revalidate a customer champion and document their role.";
  else if(no(execRaw)||unknown(execRaw))nextAction="Establish executive coverage before the renewal cycle compresses.";
  else if(adoption===-2)nextAction="Build an adoption recovery plan tied to the customer's primary use case.";
  else if(support!==null&&support>=5)nextAction="Resolve or contain the support blockers and confirm customer impact.";
  else if(expansionCandidate)nextAction="Validate the expansion use case, buyer, timing, and commercial value.";

  let diagnosis="";
  if(risk==="Critical")diagnosis="This account has material renewal or relationship exposure: "+(risks.slice(0,2).join("; ")||"multiple customer signals require attention")+".";
  else if(risk==="Watch")diagnosis="This account needs attention before risk compounds: "+(risks.slice(0,2).join("; ")||"evidence is incomplete")+".";
  else diagnosis="Current evidence is broadly stable"+(expansionCandidate?" and the account also shows an expansion signal":"")+".";
  const concentration=totalPortfolioArr>0?arr/totalPortfolioArr:0;

  return {
    account:account,arr:arr,renewalDate:renewalDateRaw,daysToRenewal:daysToRenewal,
    health:healthRaw,lastActivity:lastActivityRaw,lastExecMeeting:lastExecRaw,champion:championRaw,execSponsor:execRaw,
    stakeholders:stakeholders,adoption:adoptionRaw,supportIssues:support,nps:nps,renewalStatus:renewalStatusRaw,
    expansionPotential:expansionRaw,expansionAmount:expansionAmount,owner:owner,segment:segment,region:region,
    riskScore:riskScore,risk:risk,risks:risks,positives:positives,nextAction:nextAction,diagnosis:diagnosis,
    expansionCandidate:expansionCandidate,expansionReason:expansionReason,concentration:concentration
  };
}
function summarize(accounts){
  const sum=list=>list.reduce((n,a)=>n+a.arr,0);
  const total=sum(accounts);
  const riskAccounts=accounts.filter(a=>a.risk!=="Stable");
  const renewal90=accounts.filter(a=>a.daysToRenewal!==null&&a.daysToRenewal>=0&&a.daysToRenewal<=90);
  const expansions=accounts.filter(a=>a.expansionCandidate);
  const sorted=accounts.slice().sort((a,b)=>b.arr-a.arr);
  const top5=sorted.slice(0,5);
  return {
    total:total,riskArr:sum(riskAccounts),riskCount:riskAccounts.length,
    renewalArr:sum(renewal90),renewalCount:renewal90.length,
    expansions:expansions,expansionAmount:expansions.reduce((n,a)=>n+a.expansionAmount,0),
    top5:top5,top5Share:total?sum(top5)/total:0
  };
}
function renewalBuckets(accounts,total){
  const defs=[
    {label:"0–30 days",min:0,max:30},
    {label:"31–60",min:31,max:60},
    {label:"61–90",min:61,max:90},
    {label:"91–180",min:91,max:180},
    {label:"180+ days",min:181,max:Infinity}
  ];
  return defs.map(d=>{
    const list=accounts.filter(a=>a.daysToRenewal!==null&&a.daysToRenewal>=d.min&&a.daysToRenewal<=d.max);
    const arr=list.reduce((n,a)=>n+a.arr,0);
    return {label:d.label,arr:arr,count:list.length,pct:total?arr/total:0};
  });
}
function renderResults(){
  $("results").classList.remove("hidden");
  $("demoResultsLabel").classList.toggle("hidden",!demoMode);
  const s=summarize(analyzed);
  $("kArr").textContent=fmtMoney(s.total);$("kAccounts").textContent=analyzed.length+" accounts";
  $("kRiskArr").textContent=fmtMoney(s.riskArr);$("kRiskAccounts").textContent=s.riskCount+" accounts need attention";
  $("kRenewalArr").textContent=fmtMoney(s.renewalArr);$("kRenewalAccounts").textContent=s.renewalCount+" renewals in window";
  $("kExpansion").textContent=s.expansions.length;$("kExpansionArr").textContent=s.expansionAmount?fmtMoney(s.expansionAmount)+" identified":"candidate accounts";

  const critical=analyzed.filter(a=>a.risk==="Critical").length;
  const watch=analyzed.filter(a=>a.risk==="Watch").length;
  $("portfolioRead").textContent=
    fmtMoney(s.riskArr)+" of "+fmtMoney(s.total)+" is currently risk-exposed across "+s.riskCount+" account"+(s.riskCount===1?"":"s")+
    ". "+fmtMoney(s.renewalArr)+" renews within 90 days. "+critical+" account"+(critical===1?" is":"s are")+" critical and "+watch+" "+(watch===1?"is":"are")+" on watch. "+
    "The five largest accounts represent "+fmtPct(s.top5Share)+" of portfolio ARR. "+s.expansions.length+" account"+(s.expansions.length===1?" shows":"s show")+" a credible expansion signal.";

  const priority=analyzed.filter(a=>a.risk!=="Stable").sort((a,b)=>b.riskScore-a.riskScore||((a.daysToRenewal==null?9999:a.daysToRenewal)-(b.daysToRenewal==null?9999:b.daysToRenewal))||b.arr-a.arr).slice(0,8);
  $("priorityCount").textContent=priority.length+" prioritized";
  $("priorityQueue").innerHTML=priority.map(a=>
    '<div class="queue-item"><div class="queue-top"><div><div class="queue-name">'+esc(a.account)+'</div><div class="queue-meta">'+fmtMoney(a.arr)+(a.daysToRenewal!==null?' · renewal '+(a.daysToRenewal<0?Math.abs(a.daysToRenewal)+' days overdue':a.daysToRenewal+' days'):' · renewal date unknown')+(a.owner?' · '+esc(a.owner):'')+'</div></div><span class="risk-pill '+a.risk+'">'+a.risk+'</span></div><div class="queue-reason">'+esc(a.risks.slice(0,2).join(" · ")||"No major risk signals detected")+'</div><div class="queue-next"><strong>Next:</strong> '+esc(a.nextAction)+'</div></div>'
  ).join("")||'<div class="empty">No accounts currently fall into Critical or Watch under the mapped evidence.</div>';

  const buckets=renewalBuckets(analyzed,s.total);
  const maxBucket=Math.max.apply(null,buckets.map(b=>b.arr).concat([1]));
  $("renewalWindows").innerHTML=buckets.map(b=>
    '<div class="renewal-row"><div>'+b.label+'</div><div class="renewal-track"><div class="renewal-fill" style="width:'+(b.arr/maxBucket*100)+'%"></div></div><div class="renewal-value">'+fmtMoney(b.arr)+'<div class="muted">'+b.count+' acct'+(b.count===1?"":"s")+'</div></div></div>'
  ).join("");

  $("concentration").innerHTML='<div class="concentration-row"><span>Top 5 share of ARR</span><span>'+fmtPct(s.top5Share)+'</span></div>'+
    s.top5.map(a=>'<div class="concentration-row"><span>'+esc(a.account)+'</span><span>'+fmtMoney(a.arr)+' · '+fmtPct(a.concentration)+'</span></div>').join("");

  const expansions=s.expansions.slice().sort((a,b)=>(b.expansionAmount-a.expansionAmount)||b.arr-a.arr);
  $("expansionCount").textContent=expansions.length+" candidates";
  $("expansionList").innerHTML=expansions.slice(0,6).map(a=>
    '<div class="expansion-card"><h4>'+esc(a.account)+'</h4><p>'+esc(a.expansionReason)+'</p><div class="expansion-amount">'+(a.expansionAmount?fmtMoney(a.expansionAmount):"Signal identified")+'</div><p>'+esc(a.nextAction)+'</p></div>'
  ).join("")||'<div class="empty">No accounts currently meet the expansion evidence bar.</div>';

  renderAccounts();
}
function renewalFilterPass(a,filter){
  if(filter==="All")return true;
  if(filter==="Later")return a.daysToRenewal===null||a.daysToRenewal>90;
  const days=parseInt(filter,10);
  return a.daysToRenewal!==null&&a.daysToRenewal>=0&&a.daysToRenewal<=days;
}
function renderAccounts(){
  const risk=$("filterRisk").value,renewal=$("filterRenewal").value;
  const list=analyzed.filter(a=>(risk==="All"||a.risk===risk)&&renewalFilterPass(a,renewal)).slice().sort((a,b)=>{
    const rank={Critical:0,Watch:1,Stable:2};
    return rank[a.risk]-rank[b.risk]||((a.daysToRenewal==null?9999:a.daysToRenewal)-(b.daysToRenewal==null?9999:b.daysToRenewal))||b.arr-a.arr;
  });
  $("accountCount").textContent="Showing "+list.length+" of "+analyzed.length+" accounts";
  $("accountList").innerHTML=list.map(a=>{
    const meta=[fmtMoney(a.arr),a.renewalDate?"renewal "+esc(a.renewalDate):"renewal unknown",a.owner?esc(a.owner):"",a.segment?esc(a.segment):""].filter(Boolean).join(" · ");
    const risks=a.risks.length?'<ul>'+a.risks.map(x=>'<li>'+esc(x)+'</li>').join("")+'</ul>':'<div class="empty">No material risk signals detected.</div>';
    const positives=a.positives.length?'<ul>'+a.positives.map(x=>'<li>'+esc(x)+'</li>').join("")+'</ul>':'<div class="empty">No positive evidence fields were available.</div>';
    return '<article class="account"><div class="account-top"><div><div class="account-name">'+esc(a.account)+'</div><div class="account-meta">'+meta+'</div></div><span class="risk-pill '+a.risk+'">'+a.risk+'</span></div><div class="account-diagnosis">'+esc(a.diagnosis)+'</div><div class="account-cols"><div><h4>Risk signals</h4>'+risks+'</div><div><h4>Supporting evidence</h4>'+positives+'</div></div><div class="next-action"><strong>Next customer action:</strong> '+esc(a.nextAction)+'</div></article>';
  }).join("")||'<div class="empty">No accounts match the selected filters.</div>';
}
function runAnalysis(){
  const portfolioArr=rawRows.reduce((n,row)=>n+money(val(row,"arr")),0);
  analyzed=rawRows.map(row=>analyzeRow(row,portfolioArr));
  renderSourceContext();renderResults();
  setTimeout(()=>$("results").scrollIntoView({behavior:"smooth",block:"start"}),80);
}
function analysisCSV(){
  const fields=["Account","ARR","Renewal Date","Days to Renewal","Risk State","Risk Score","Risk Signals","Supporting Evidence","Next Customer Action","Expansion Candidate","Expansion Amount","Owner","Segment","Region"];
  const rows=analyzed.map(a=>[
    a.account,a.arr,a.renewalDate,a.daysToRenewal==null?"":a.daysToRenewal,a.risk,a.riskScore,
    a.risks.join(" | "),a.positives.join(" | "),a.nextAction,a.expansionCandidate?"Yes":"No",a.expansionAmount,a.owner,a.segment,a.region
  ]);
  const quote=v=>'\"'+String(v==null?"":v).replace(/\"/g,'\"\"')+'\"';
  return [fields].concat(rows).map(r=>r.map(quote).join(",")).join("\n");
}
function downloadAnalysis(){
  const blob=new Blob([analysisCSV()],{type:"text/csv;charset=utf-8"});
  const a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download="cairn_customer_portfolio_analysis.csv";a.click();URL.revokeObjectURL(a.href);
}
function demoDate(offset,reference){
  const d=calendarDay(reference||new Date());d.setUTCDate(d.getUTCDate()+offset);return d.toISOString().slice(0,10);
}

$("fileInput").addEventListener("change",e=>handlePortfolioFile(e.target.files[0]));
const dz=$("dropZone");
["dragenter","dragover"].forEach(ev=>dz.addEventListener(ev,e=>{e.preventDefault();dz.classList.add("drag")}));
["dragleave","drop"].forEach(ev=>dz.addEventListener(ev,e=>{e.preventDefault();dz.classList.remove("drag")}));
dz.addEventListener("drop",e=>handlePortfolioFile(e.dataTransfer.files[0]));
$("useSheetBtn").onclick=()=>useWorkbookSheet($("sheetSelect").value);
$("reset").onclick=()=>location.reload();
$("analyzeBtn").onclick=runAnalysis;
$("filterRisk").onchange=renderAccounts;
$("filterRenewal").onchange=renderAccounts;
$("downloadBtn").onclick=downloadAnalysis;

$("loadSample").onclick=()=>{
  demoMode=true;activeFileName="";activeSheetName="";
  const t=new Date();
  const sample=[
    ["Account Name","ARR","Renewal Date","Health","Last Activity Date","Last Executive Meeting","Champion","Executive Sponsor","Stakeholder Count","Product Adoption","Open Support Issues","NPS","Renewal Status","Expansion Potential","Expansion Amount","CSM / Owner","Segment","Region"],
    ["Summit Manufacturing","420000",demoDate(24,t),"Red",demoDate(-67,t),demoDate(-151,t),"No","No","2","Low","6","5","At Risk","No","","Alex Morgan","Enterprise","North America"],
    ["Northstar Logistics","365000",demoDate(58,t),"Yellow",demoDate(-35,t),demoDate(-82,t),"Yes","No","3","Medium","3","7","Pending","Possible","65000","Jamie Lee","Enterprise","North America"],
    ["BluePeak Foods","310000",demoDate(112,t),"Green",demoDate(-9,t),demoDate(-37,t),"Yes","Yes","6","High","0","10","Confirmed","High","90000","Alex Morgan","Enterprise","North America"],
    ["Vertex Energy","295000",demoDate(41,t),"Yellow",demoDate(-18,t),demoDate(-132,t),"Yes","No","3","Medium","2","8","Pending","No","","Priya Shah","Enterprise","North America"],
    ["Cedar Health","270000",demoDate(15,t),"Green",demoDate(-8,t),demoDate(-44,t),"Yes","Yes","5","High","1","9","Confirmed","High","70000","Jamie Lee","Enterprise","North America"],
    ["Atlas Freight","255000",demoDate(205,t),"Green",demoDate(-14,t),demoDate(-61,t),"Yes","Yes","7","High","0","9","Confirmed","High","80000","Priya Shah","Strategic","North America"],
    ["Redwood Systems","230000",demoDate(76,t),"Red",demoDate(-94,t),"","No","No","1","Low","5","4","At Risk","No","","Alex Morgan","Mid-Market","North America"],
    ["Orion Retail","215000",demoDate(149,t),"Green",demoDate(-12,t),demoDate(-49,t),"Yes","Yes","4","High","1","9","Discussing","Possible","45000","Jamie Lee","Mid-Market","North America"],
    ["Harbor Industrial","195000",demoDate(33,t),"Yellow",demoDate(-63,t),demoDate(-118,t),"Yes","No","2","Medium","4","7","Pending","No","","Priya Shah","Mid-Market","North America"],
    ["Pioneer Mobility","175000",demoDate(287,t),"Green",demoDate(-7,t),demoDate(-31,t),"Yes","Yes","5","High","0","10","Confirmed","High","50000","Alex Morgan","Mid-Market","Europe"],
    ["Lumen Works","150000",demoDate(89,t),"Yellow",demoDate(-28,t),"","Unknown","No","2","Medium","2","8","Unknown","Possible","30000","Jamie Lee","Mid-Market","Europe"],
    ["Stonebridge Labs","120000",demoDate(172,t),"Green",demoDate(-11,t),demoDate(-70,t),"Yes","Yes","4","High","0","9","Confirmed","No","","Priya Shah","Growth","North America"]
  ];
  const csv=sample.map(r=>r.map(v=>'\"'+String(v).replace(/\"/g,'\"\"')+'\"').join(",")).join("\n");
  loadText(csv);
  $("sampleReady").classList.remove("hidden");
  const total=rawRows.reduce((n,row)=>n+money(val(row,"arr")),0);
  $("sampleStats").textContent=rawRows.length+" fictional accounts · "+fmtMoney(total)+" portfolio ARR";
  $("mappingSection").scrollIntoView({behavior:"smooth",block:"start"});
};
