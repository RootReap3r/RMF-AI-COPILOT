import { useState, useCallback, useRef, useEffect, useMemo } from "react";
import * as XLSX from "xlsx";

const C = {
  navy:"#0D1B2A", navyMid:"#1B3A5C", steel:"#2E5F8A",
  accent:"#4A90D9", accentSoft:"#D6E8F7", green:"#1A6B3C",
  greenLight:"#D5F5E3", yellow:"#B07A12", yellowLight:"#FBF1DD",
  red:"#A33327", redLight:"#FBE7E4", orange:"#E67E22",
  orangeLight:"#FEF0D4", gray:"#4A4A4A", grayMid:"#7F8C8D",
  grayLight:"#ECF0F1", white:"#FFFFFF", offwhite:"#F8FAFC",
  purple:"#6C3483", purpleLight:"#E8DAEF",
};

const FULL_BASELINE = [
  "AC-1","AC-2","AC-2(1)","AC-2(2)","AC-2(3)","AC-2(4)","AC-2(5)","AC-2(7)","AC-2(9)","AC-2(10)","AC-2(11)","AC-2(12)","AC-2(13)",
  "AC-3","AC-3(7)","AC-4","AC-4(4)","AC-5","AC-6","AC-6(1)","AC-6(2)","AC-6(5)","AC-6(7)","AC-6(9)","AC-6(10)",
  "AC-7","AC-8","AC-10","AC-11","AC-11(1)","AC-12","AC-12(1)","AC-14","AC-17","AC-17(1)","AC-17(2)","AC-17(3)","AC-17(4)",
  "AC-18","AC-18(1)","AC-19","AC-19(5)","AC-20","AC-20(1)","AC-20(2)","AC-21","AC-22","AC-23",
  "AT-1","AT-2","AT-2(2)","AT-2(3)","AT-3","AT-3(5)","AT-4","AT-6",
  "AU-1","AU-2","AU-3","AU-3(1)","AU-4","AU-4(1)","AU-5","AU-5(1)","AU-5(2)","AU-6","AU-6(1)","AU-6(3)","AU-6(4)","AU-6(5)",
  "AU-7","AU-7(1)","AU-8","AU-9","AU-9(2)","AU-9(3)","AU-9(4)","AU-10","AU-11","AU-12","AU-12(1)","AU-12(3)","AU-14","AU-16",
  "CA-1","CA-2","CA-2(1)","CA-2(2)","CA-3","CA-3(6)","CA-5","CA-6","CA-7","CA-7(1)","CA-7(4)","CA-8","CA-8(1)","CA-9",
  "CM-1","CM-2","CM-2(2)","CM-2(3)","CM-2(7)","CM-3","CM-3(1)","CM-3(2)","CM-3(4)","CM-3(6)","CM-4","CM-4(1)","CM-4(2)",
  "CM-5","CM-5(1)","CM-6","CM-6(1)","CM-6(2)","CM-7","CM-7(1)","CM-7(2)","CM-7(4)","CM-7(5)","CM-8","CM-8(1)","CM-8(2)","CM-8(3)",
  "CM-9","CM-10","CM-11","CM-12","CM-12(1)","CM-14",
  "CP-1","CP-2","CP-2(1)","CP-2(2)","CP-2(3)","CP-2(5)","CP-2(8)","CP-3","CP-3(1)","CP-4","CP-4(1)","CP-4(2)",
  "CP-6","CP-6(1)","CP-6(3)","CP-7","CP-7(1)","CP-7(2)","CP-7(3)","CP-8","CP-8(1)","CP-8(2)",
  "CP-9","CP-9(1)","CP-9(2)","CP-9(3)","CP-9(5)","CP-9(8)","CP-10","CP-10(2)","CP-11","CP-13",
  "IA-1","IA-2","IA-2(1)","IA-2(2)","IA-2(5)","IA-2(8)","IA-2(12)","IA-3","IA-3(1)","IA-4","IA-4(4)",
  "IA-5","IA-5(1)","IA-5(2)","IA-5(6)","IA-6","IA-7","IA-8","IA-8(1)","IA-8(2)","IA-8(4)","IA-9","IA-10","IA-11","IA-12","IA-12(2)","IA-12(3)",
  "IR-1","IR-2","IR-2(1)","IR-2(2)","IR-3","IR-3(2)","IR-4","IR-4(1)","IR-4(4)","IR-4(6)","IR-4(8)","IR-4(11)",
  "IR-5","IR-5(1)","IR-6","IR-6(1)","IR-6(2)","IR-6(3)","IR-7","IR-7(1)","IR-8","IR-8(1)","IR-9","IR-9(1)","IR-9(2)","IR-9(3)","IR-10",
  "MA-1","MA-2","MA-2(2)","MA-3","MA-3(1)","MA-3(2)","MA-3(3)","MA-4","MA-4(3)","MA-5","MA-5(1)","MA-6",
  "MP-1","MP-2","MP-3","MP-4","MP-5","MP-6","MP-6(1)","MP-6(2)","MP-6(3)","MP-7",
  "PE-1","PE-2","PE-2(3)","PE-3","PE-3(1)","PE-4","PE-5","PE-6","PE-6(1)","PE-6(4)","PE-8","PE-8(1)",
  "PE-9","PE-10","PE-11","PE-11(1)","PE-12","PE-13","PE-13(2)","PE-14","PE-14(2)","PE-15","PE-16","PE-17","PE-18","PE-20","PE-22","PE-23",
  "PL-1","PL-2","PL-4","PL-4(1)","PL-7","PL-8","PL-8(1)","PL-9","PL-10","PL-11",
  "PM-1","PM-2","PM-3","PM-4","PM-5","PM-6","PM-7","PM-8","PM-9","PM-10","PM-11","PM-12","PM-13","PM-14","PM-15","PM-16",
  "PM-17","PM-18","PM-19","PM-20","PM-21","PM-22","PM-23","PM-24","PM-25","PM-26","PM-27","PM-28","PM-29","PM-30","PM-31",
  "PS-1","PS-2","PS-3","PS-3(3)","PS-4","PS-4(2)","PS-5","PS-6","PS-7","PS-8","PS-9",
  "PT-1","PT-2","PT-2(1)","PT-2(2)","PT-3","PT-3(2)","PT-4","PT-5","PT-5(2)","PT-6","PT-6(1)","PT-7","PT-8",
  "RA-1","RA-2","RA-3","RA-3(1)","RA-5","RA-5(2)","RA-5(4)","RA-5(5)","RA-5(11)","RA-7","RA-8","RA-9","RA-10",
  "SA-1","SA-2","SA-3","SA-4","SA-4(1)","SA-4(2)","SA-4(5)","SA-4(9)","SA-4(10)","SA-5","SA-8","SA-9","SA-9(1)","SA-9(2)","SA-9(4)","SA-9(5)",
  "SA-10","SA-10(1)","SA-11","SA-11(1)","SA-11(2)","SA-11(8)","SA-15","SA-15(3)","SA-16","SA-17","SA-21","SA-22",
  "SC-1","SC-2","SC-2(1)","SC-3","SC-4","SC-5","SC-5(2)","SC-7","SC-7(3)","SC-7(4)","SC-7(5)","SC-7(7)","SC-7(8)","SC-7(10)","SC-7(11)","SC-7(12)","SC-7(13)","SC-7(18)","SC-7(21)","SC-7(22)",
  "SC-8","SC-8(1)","SC-10","SC-12","SC-12(1)","SC-13","SC-15","SC-17","SC-18","SC-20","SC-21","SC-22","SC-23","SC-23(1)","SC-24","SC-25","SC-26","SC-28","SC-28(1)","SC-29","SC-36","SC-39","SC-45","SC-49","SC-51",
  "SI-1","SI-2","SI-2(2)","SI-2(3)","SI-3","SI-3(1)","SI-3(2)","SI-4","SI-4(1)","SI-4(2)","SI-4(4)","SI-4(5)","SI-4(10)","SI-4(12)","SI-4(14)","SI-4(20)","SI-4(22)","SI-4(23)","SI-4(24)","SI-4(25)",
  "SI-5","SI-5(1)","SI-6","SI-7","SI-7(1)","SI-7(2)","SI-7(5)","SI-7(6)","SI-7(7)","SI-7(15)","SI-8","SI-8(1)","SI-8(2)","SI-10","SI-11","SI-12","SI-16","SI-18","SI-18(1)","SI-19","SI-19(4)","SI-23",
  "SR-1","SR-2","SR-2(1)","SR-3","SR-5","SR-5(1)","SR-5(2)","SR-6","SR-6(1)","SR-8","SR-9","SR-9(1)","SR-10","SR-11","SR-11(1)","SR-11(2)","SR-12",
];
const BASELINE_SET = new Set(FULL_BASELINE);

const FAMILY_NAMES = {
  AC:"Access Control", AT:"Awareness & Training", AU:"Audit & Accountability",
  CA:"Assessment, Authorization & Monitoring", CM:"Configuration Management",
  CP:"Contingency Planning", IA:"Identification & Authentication",
  IR:"Incident Response", MA:"Maintenance", MP:"Media Protection",
  PE:"Physical & Environmental Protection", PL:"Planning",
  PM:"Program Management", PS:"Personnel Security",
  PT:"PII Processing & Transparency", RA:"Risk Assessment",
  SA:"System & Services Acquisition", SC:"System & Comms Protection",
  SI:"System & Info Integrity", SR:"Supply Chain Risk Management",
};
const PHASE_LABELS = ["Gap Analysis","SSP/Policy","Evidence","eMASS Entry","AO Review"];

// Each chunk sent to the LLM for extraction. 60k chars ≈ 15k tokens,
// well within any modern model's context window.
const CHUNK_SIZE = 60000;
const CHUNK_OVERLAP = 3000;

function chunkText(text, size, overlap) {
  if (text.length <= size) return [text];
  const chunks = [];
  let start = 0;
  while (start < text.length) {
    chunks.push(text.slice(start, start + size));
    if (start + size >= text.length) break;
    start += size - overlap;
  }
  return chunks;
}

function sleep(ms){ return new Promise(r=>setTimeout(r,ms)); }

function extractJSON(raw) {
  if (!raw) return null;
  let text = raw.trim();
  text = text.replace(/^```(?:json)?\s*/i, "").replace(/```\s*$/i, "").trim();
  try { return JSON.parse(text); } catch {}
  const firstArr = text.indexOf("[");
  const firstObj = text.indexOf("{");
  let start = -1, closeCh;
  if (firstArr !== -1 && (firstObj === -1 || firstArr < firstObj)) { start = firstArr; closeCh="]"; }
  else if (firstObj !== -1) { start = firstObj; closeCh="}"; }
  if (start === -1) return null;
  const end = text.lastIndexOf(closeCh);
  if (end === -1 || end < start) return null;
  try { return JSON.parse(text.slice(start, end+1)); } catch {}
  return null;
}

const API_BASE = import.meta.env.VITE_API_BASE || "";

// ── Auth token management ─────────────────────────────────────────
function getStoredToken() {
  try { return sessionStorage.getItem("rmf-auth-token") || ""; } catch { return ""; }
}
function storeToken(t) {
  try { sessionStorage.setItem("rmf-auth-token", t); } catch {}
}

function jsonHeaders(token) {
  const h = { "Content-Type": "application/json" };
  if (token) h["X-Api-Key"] = token;
  return h;
}

// ── LLM proxy ─────────────────────────────────────────────────────
async function callClaude(system, user, maxTokens=900, token="") {
  for (let attempt=0; attempt<4; attempt++) {
    try {
      const resp = await fetch(`${API_BASE}/api/chat`, {
        method:"POST", headers: jsonHeaders(token),
        body: JSON.stringify({ system, user, maxTokens }),
      });
      if (resp.status===529 || resp.status===429) { await sleep((attempt+1)*5000); continue; }
      if (resp.status===401) throw new Error("AUTH_REQUIRED");
      if (!resp.ok) {
        const errBody = await resp.json().catch(()=>({}));
        throw new Error(errBody.error || `Backend returned ${resp.status}`);
      }
      const data = await resp.json();
      return data.text ?? "";
    } catch (e) {
      if (e.message==="AUTH_REQUIRED") throw e;
      if(attempt===3) throw new Error(`LLM call failed after 4 attempts: ${e.message}`);
      await sleep((attempt+1)*3000);
    }
  }
  return "";
}

// ── NIST catalog lookup ───────────────────────────────────────────
const catalogCache = {};
async function getCatalogEntry(controlId, token="") {
  if (catalogCache[controlId] !== undefined) return catalogCache[controlId];
  try {
    const h = token ? { "X-Api-Key": token } : {};
    const r = await fetch(`${API_BASE}/api/catalog/${encodeURIComponent(controlId)}`, { headers: h });
    if (!r.ok) { catalogCache[controlId] = null; return null; }
    const data = await r.json();
    catalogCache[controlId] = data;
    return data;
  } catch { catalogCache[controlId] = null; return null; }
}

async function processBatch(items, size, delay, fn, onProgress) {
  const out = [];
  for (let i=0;i<items.length;i+=size) {
    const chunk = items.slice(i,i+size);
    const res = await Promise.all(chunk.map(fn));
    out.push(...res);
    onProgress(Math.min(i+size,items.length), items.length);
    if (i+size<items.length) await sleep(delay);
  }
  return out;
}

// ── Document parsing (server-side) ───────────────────────────────
async function parseSSPFiles(files, token="") {
  const formData = new FormData();
  for (const file of Array.from(files)) formData.append("files", file);
  const headers = token ? { "X-Api-Key": token } : {};
  const resp = await fetch(`${API_BASE}/api/parse-document`, { method:"POST", headers, body: formData });
  if (resp.status===401) throw new Error("AUTH_REQUIRED");
  if (!resp.ok) {
    const err = await resp.json().catch(()=>({}));
    throw new Error(err.error || `Document parsing failed (${resp.status})`);
  }
  const { text } = await resp.json();
  return text;
}

function downloadExcel(rows, filename, sheetName) {
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  XLSX.writeFile(wb, filename);
}

// ── Persistent storage ────────────────────────────────────────────
async function loadAllControls(token="") {
  try {
    const listResp = await fetch(`${API_BASE}/api/storage?prefix=ctrl:`,
      { headers: token ? { "X-Api-Key": token } : {} });
    const { keys } = await listResp.json();
    if (!keys || keys.length===0) return {};
    const out = {};
    for (const key of keys) {
      try {
        const r = await fetch(`${API_BASE}/api/storage/${encodeURIComponent(key)}`,
          { headers: token ? { "X-Api-Key": token } : {} });
        if (!r.ok) continue;
        const { value } = await r.json();
        out[key.replace("ctrl:","")] = JSON.parse(value);
      } catch {}
    }
    return out;
  } catch { return {}; }
}

async function saveControl(controlId, data, token="") {
  try {
    await fetch(`${API_BASE}/api/storage`, {
      method:"POST", headers: jsonHeaders(token),
      body: JSON.stringify({ key:"ctrl:"+controlId, value: JSON.stringify(data) }),
    });
  } catch {}
}

async function clearAllControls(token="") {
  try {
    await fetch(`${API_BASE}/api/storage/clear`, {
      method:"POST", headers: jsonHeaders(token),
      body: JSON.stringify({ prefix:"ctrl:" }),
    });
  } catch {}
}

// ── SCTM workbook ─────────────────────────────────────────────────
const SCTM_HEADERS = ["Control ID","Control Family","Implementation Status","Control Origination",
  "Implementation Statement","Test Method","Gap Flag","Gap Reason","AI Suggestion","Guidance"];

function buildSCTMWorkbook(allControls, documentedIds) {
  const wb = XLSX.utils.book_new();
  for (const fam of Object.keys(FAMILY_NAMES)) {
    const rows = [SCTM_HEADERS];
    const famIds = documentedIds.filter(id=>id.split("-")[0]===fam)
      .sort((a,b)=>a.localeCompare(b, undefined, {numeric:true}));
    for (const id of famIds) {
      const c = allControls[id];
      rows.push([id, FAMILY_NAMES[fam], c.sctmStatus, c.origination,
        c.implementationStatement, c.testMethod,
        c.gap?"GAP":"", c.gapReason||"", c.aiSuggestion||"", c.nistGuidance||""]);
    }
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws["!cols"] = [{wch:10},{wch:22},{wch:16},{wch:16},{wch:50},{wch:14},{wch:8},{wch:30},{wch:50},{wch:50}];
    XLSX.utils.book_append_sheet(wb, ws, fam);
  }
  return wb;
}

// ── eMASS workbook ────────────────────────────────────────────────
const EMASS_HEADERS = ["Control_Acronym","Control_Information","Implementation_Status",
  "Security_Control_Designation","Assessment_Procedures","Inherited",
  "Implementation_Narrative","Responsible_Entities","Planned_Implementation_Date","Test_Results","CCI"];
const EMASS_NARRATIVE_LIMIT = 2000;

function baseControlId(id) {
  const m = id.match(/^([A-Z]{2}-\d+)/);
  return m ? m[1] : id;
}

function buildEmassRows(allControls, documentedIds) {
  const byBase = {};
  for (const id of documentedIds) {
    const base = baseControlId(id);
    (byBase[base] ??= []).push(id);
  }
  const rows = [];
  for (const base of Object.keys(byBase).sort((a,b)=>a.localeCompare(b, undefined, {numeric:true}))) {
    const ids = byBase[base].sort((a,b)=>a.length-b.length);
    const fam = base.split("-")[0];
    const baseEntry = allControls[base];
    let status;
    if (baseEntry) status = baseEntry.sctmStatus;
    else {
      const statuses = ids.map(i=>allControls[i].status);
      status = statuses.every(s=>s==="Gap")?"Planned":statuses.every(s=>s==="Matched")?"Implemented":"Partially Implemented";
    }
    let narrative = baseEntry?.implementationStatement || `${base}: not yet documented in SSP.`;
    for (const eid of ids.filter(i=>i!==base)) {
      const e = allControls[eid];
      const note = e.gap ? `${eid}: gap — ${e.gapReason}` : `${eid}: ${e.implementationStatement}`;
      const candidate = narrative + "\n" + note;
      if (candidate.length <= EMASS_NARRATIVE_LIMIT) narrative = candidate;
      else { narrative = narrative.slice(0, EMASS_NARRATIVE_LIMIT-15) + "... [truncated]"; break; }
    }
    if (narrative.length > EMASS_NARRATIVE_LIMIT) narrative = narrative.slice(0, EMASS_NARRATIVE_LIMIT-15) + "... [truncated]";
    const designation = baseEntry?.emassDesignation ?? (ids.map(i=>allControls[i].origination).includes("Inherited")?"Common":"System-Specific");
    const inherited = ids.some(i=>allControls[i].origination==="Inherited");
    const cciSet = new Set();
    for (const i of ids) (allControls[i].ccis||[]).forEach(c=>cciSet.add(c));
    rows.push({
      Control_Acronym:base, Control_Information:baseEntry?.emassControlInfo||FAMILY_NAMES[fam]||"",
      Implementation_Status:status, Security_Control_Designation:designation,
      Assessment_Procedures:baseEntry?.testMethod||"Test", Inherited:String(inherited),
      Implementation_Narrative:narrative, Responsible_Entities:"System Owner / ISSO",
      Planned_Implementation_Date:"", Test_Results:"", CCI:[...cciSet].sort().join(", "),
    });
  }
  return rows;
}

function buildEmassWorkbook(allControls, documentedIds) {
  const rows = buildEmassRows(allControls, documentedIds);
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(rows, {header:EMASS_HEADERS});
  ws["!cols"] = [{wch:14},{wch:18},{wch:18},{wch:16},{wch:16},{wch:10},{wch:60},{wch:18},{wch:16},{wch:12},{wch:12}];
  XLSX.utils.book_append_sheet(wb, ws, "Test Result Import");
  return wb;
}

// ── UI components ─────────────────────────────────────────────────
function DropZone({label, accept, icon, onFiles, files, color, multi}) {
  const [dragging,setDragging]=useState(false);
  const inputRef=useRef();
  const handleDrop=useCallback(e=>{e.preventDefault();setDragging(false);if(onFiles)onFiles(e.dataTransfer.files);},[onFiles]);
  const fileList=files?Array.from(files):[];
  const has=fileList.length>0;
  return (
    <div onClick={()=>inputRef.current.click()} onDragOver={e=>{e.preventDefault();setDragging(true);}} onDragLeave={()=>setDragging(false)} onDrop={handleDrop}
      style={{border:`2px dashed ${dragging?color:has?color:"#BCC8D4"}`,borderRadius:10,padding:"18px 14px",cursor:"pointer",
        background:dragging?`${color}18`:has?`${color}09`:C.offwhite,textAlign:"center",minHeight:110,
        display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",gap:5}}>
      <input ref={inputRef} type="file" accept={accept} multiple={!!multi} style={{display:"none"}}
        onChange={e=>{if(onFiles&&e.target.files.length)onFiles(e.target.files);}}/>
      <div style={{fontSize:28}}>{icon}</div>
      <div style={{fontWeight:700,color:has?color:C.gray,fontSize:13}}>
        {has?(fileList.length===1?fileList[0].name:`${fileList.length} files selected`):label}
      </div>
      {has&&fileList.length>1&&<div style={{fontSize:10,color:C.grayMid,maxWidth:"100%",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{fileList.map(f=>f.name).join(", ")}</div>}
      <div style={{fontSize:10,color:C.grayMid}}>{has?"click to replace":multi?"drag & drop · multiple files OK":"drag & drop or click"}</div>
    </div>
  );
}

function StepBadge({n,label,state}) {
  const cols={idle:{bg:C.grayLight,fg:C.grayMid,b:"#CCC"},active:{bg:C.accentSoft,fg:C.steel,b:C.accent},done:{bg:C.greenLight,fg:C.green,b:C.green},error:{bg:C.redLight,fg:C.red,b:C.red}};
  const s=cols[state]??cols.idle;
  return (
    <div style={{display:"flex",alignItems:"center",gap:6}}>
      <div style={{width:24,height:24,borderRadius:"50%",background:s.bg,border:`2px solid ${s.b}`,display:"flex",alignItems:"center",justifyContent:"center",fontWeight:700,fontSize:10,color:s.fg,flexShrink:0}}>
        {state==="done"?"✓":state==="error"?"✕":n}
      </div>
      <span style={{fontSize:10,color:s.fg,fontWeight:state==="active"?700:500}}>{label}</span>
    </div>
  );
}

function ProgressBar({value,max,label,color}) {
  const pct=max>0?Math.round((value/max)*100):0;
  return (
    <div style={{marginBottom:6}}>
      <div style={{display:"flex",justifyContent:"space-between",marginBottom:2}}>
        <span style={{fontSize:10,color:C.grayMid}}>{label}</span>
        <span style={{fontSize:10,color,fontWeight:700}}>{value}/{max} ({pct}%)</span>
      </div>
      <div style={{height:5,background:C.grayLight,borderRadius:3,overflow:"hidden"}}>
        <div style={{width:`${pct}%`,height:"100%",background:color,borderRadius:3,transition:"width .4s"}}/>
      </div>
    </div>
  );
}

function statusClass(s){
  if(s==="Complete") return {bg:C.greenLight,fg:C.green};
  if(s==="In Progress") return {bg:C.yellowLight,fg:C.yellow};
  if(s==="Blocked") return {bg:C.redLight,fg:C.red};
  return {bg:C.grayLight,fg:C.grayMid};
}

function ControlCard({control}) {
  const [open,setOpen]=useState(false);
  const sc={Matched:{bg:C.greenLight,fg:C.green,label:"✅ Matched"},Gap:{bg:C.redLight,fg:C.red,label:"🔴 Gap"},Partial:{bg:C.yellowLight,fg:C.yellow,label:"🟡 Partial"},Inherited:{bg:C.accentSoft,fg:C.steel,label:"🔵 Inherited"}};
  const s=sc[control.status]??sc.Gap;
  return (
    <div style={{border:`1px solid ${s.fg}44`,borderLeft:`4px solid ${s.fg}`,borderRadius:7,background:s.bg+"44",marginBottom:5,overflow:"hidden"}}>
      <div onClick={()=>setOpen(!open)} style={{display:"flex",alignItems:"center",gap:9,padding:"7px 11px",cursor:"pointer",justifyContent:"space-between"}}>
        <div style={{display:"flex",alignItems:"center",gap:9}}>
          <span style={{fontFamily:"monospace",fontWeight:700,fontSize:11,background:s.fg+"22",color:s.fg,padding:"1px 6px",borderRadius:4}}>{control.id}</span>
          <span style={{fontSize:10,color:C.gray,fontWeight:600}}>{FAMILY_NAMES[control.family]??control.family}</span>
        </div>
        <span style={{fontSize:10,fontWeight:700,color:s.fg,background:s.bg,border:`1px solid ${s.fg}55`,padding:"2px 7px",borderRadius:20}}>{s.label}</span>
      </div>
      {open && (
        <div style={{padding:"0 11px 11px",display:"flex",flexDirection:"column",gap:7}}>
          {control.statement && <div><div style={{fontSize:9,fontWeight:700,color:C.steel,marginBottom:2,textTransform:"uppercase"}}>SSP Statement</div><div style={{fontSize:11,color:C.gray,background:C.white,border:"1px solid #DDE",borderRadius:5,padding:"7px 9px",lineHeight:1.6}}>{control.statement}</div></div>}
          {control.aiSuggestion && <div><div style={{fontSize:9,fontWeight:700,color:C.orange,marginBottom:2,textTransform:"uppercase"}}>🤖 AI Recommendation</div><div style={{fontSize:11,color:C.gray,background:C.orangeLight,border:`1px solid ${C.orange}44`,borderRadius:5,padding:"7px 9px",lineHeight:1.6}}>{control.aiSuggestion}</div></div>}
          {control.nistGuidance && <div><div style={{fontSize:9,fontWeight:700,color:C.purple,marginBottom:2,textTransform:"uppercase"}}>📖 NIST Guidance</div><div style={{fontSize:11,color:C.gray,background:C.purpleLight,border:`1px solid ${C.purple}44`,borderRadius:5,padding:"7px 9px",lineHeight:1.6}}>{control.nistGuidance}</div></div>}
        </div>
      )}
    </div>
  );
}

// ── POA&M editor row ──────────────────────────────────────────────
function POAMRow({id, control, onFieldChange}) {
  const fam = id.split("-")[0];
  const [expanded, setExpanded] = useState(false);
  const hasAllFields = control.poamDueDate && control.poamPOC && control.poamResources;
  const fieldStyle = {
    width:"100%", boxSizing:"border-box", border:`1px solid ${C.grayLight}`,
    borderRadius:5, padding:"5px 8px", fontSize:11, color:C.navy,
    background:C.white, outline:"none", fontFamily:"inherit",
  };
  return (
    <div style={{border:`1px solid ${C.red}33`,borderLeft:`4px solid ${C.red}`,borderRadius:7,marginBottom:6,overflow:"hidden",background:"#FFF9F9"}}>
      <div style={{display:"flex",alignItems:"center",gap:8,padding:"8px 12px",cursor:"pointer",justifyContent:"space-between"}}
        onClick={()=>setExpanded(!expanded)}>
        <div style={{display:"flex",alignItems:"center",gap:8,flex:1,minWidth:0}}>
          <span style={{fontFamily:"monospace",fontWeight:700,fontSize:11,background:C.redLight,color:C.red,padding:"1px 6px",borderRadius:4,flexShrink:0}}>{id}</span>
          <span style={{fontSize:10,color:C.gray,fontWeight:600,flexShrink:0}}>{FAMILY_NAMES[fam]??fam}</span>
          <span style={{fontSize:10,color:C.grayMid,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
            {control.gapReason || "Gap identified"}
          </span>
        </div>
        <span style={{fontSize:9,fontWeight:700,padding:"2px 7px",borderRadius:10,flexShrink:0,
          background:hasAllFields?C.greenLight:C.yellowLight,
          color:hasAllFields?C.green:C.yellow,
          border:`1px solid ${hasAllFields?C.green:C.yellow}44`}}>
          {hasAllFields?"✓ Complete":"Incomplete"}
        </span>
      </div>
      {expanded && (
        <div style={{padding:"0 12px 12px",display:"flex",flexDirection:"column",gap:8}}>
          {control.gapReason && (
            <div>
              <div style={{fontSize:9,fontWeight:700,color:C.red,textTransform:"uppercase",marginBottom:2}}>Weakness</div>
              <div style={{fontSize:11,color:C.gray,background:C.redLight,border:`1px solid ${C.red}33`,borderRadius:5,padding:"6px 8px",lineHeight:1.5}}>{control.gapReason}</div>
            </div>
          )}
          {control.aiSuggestion && (
            <div>
              <div style={{fontSize:9,fontWeight:700,color:C.orange,textTransform:"uppercase",marginBottom:2}}>AI Recommendation</div>
              <div style={{fontSize:11,color:C.gray,background:C.orangeLight,border:`1px solid ${C.orange}33`,borderRadius:5,padding:"6px 8px",lineHeight:1.5}}>{control.aiSuggestion}</div>
            </div>
          )}
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10}}>
            <div>
              <label style={{display:"block",fontSize:9,fontWeight:700,color:C.steel,textTransform:"uppercase",marginBottom:3}}>Est. Completion Date</label>
              <input type="date" style={fieldStyle} value={control.poamDueDate||""}
                onChange={e=>onFieldChange(id,"poamDueDate",e.target.value)}/>
            </div>
            <div>
              <label style={{display:"block",fontSize:9,fontWeight:700,color:C.steel,textTransform:"uppercase",marginBottom:3}}>Responsible POC</label>
              <input type="text" style={fieldStyle} placeholder="Name / org / email"
                value={control.poamPOC||""}
                onChange={e=>onFieldChange(id,"poamPOC",e.target.value)}/>
            </div>
            <div>
              <label style={{display:"block",fontSize:9,fontWeight:700,color:C.steel,textTransform:"uppercase",marginBottom:3}}>Resource Estimate</label>
              <input type="text" style={fieldStyle} placeholder="e.g. 40 hrs, $5k"
                value={control.poamResources||""}
                onChange={e=>onFieldChange(id,"poamResources",e.target.value)}/>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ── Auth prompt overlay ───────────────────────────────────────────
function AuthPrompt({onSubmit}) {
  const [input,setInput]=useState("");
  return (
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.75)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:9999}}>
      <div style={{background:C.white,borderRadius:12,padding:"28px 32px",width:400,maxWidth:"90vw",boxShadow:"0 8px 32px rgba(0,0,0,0.4)"}}>
        <div style={{fontSize:20,fontWeight:800,color:C.navy,marginBottom:8}}>🔐 API Token Required</div>
        <div style={{fontSize:12,color:C.gray,marginBottom:18,lineHeight:1.6}}>
          Enter the value of <code style={{background:C.grayLight,padding:"1px 5px",borderRadius:3}}>AUTH_TOKEN</code> from your <code style={{background:C.grayLight,padding:"1px 5px",borderRadius:3}}>server/.env</code> file.
        </div>
        <input autoFocus type="password" placeholder="Paste token here…" value={input}
          onChange={e=>setInput(e.target.value)}
          onKeyDown={e=>{ if(e.key==="Enter"&&input.trim()) onSubmit(input.trim()); }}
          style={{width:"100%",boxSizing:"border-box",border:`1px solid ${C.steel}`,borderRadius:7,padding:"9px 12px",fontSize:13,marginBottom:12,outline:"none"}}/>
        <button disabled={!input.trim()} onClick={()=>onSubmit(input.trim())}
          style={{width:"100%",background:input.trim()?`linear-gradient(135deg,${C.steel},${C.navyMid})`:C.grayLight,color:input.trim()?C.white:C.grayMid,border:"none",borderRadius:8,padding:"10px",fontWeight:700,fontSize:13,cursor:input.trim()?"pointer":"not-allowed"}}>
          Connect
        </button>
      </div>
    </div>
  );
}

// ── Main app ──────────────────────────────────────────────────────
export default function App() {
  const [authToken, setAuthToken] = useState(getStoredToken);
  const [showAuthPrompt, setShowAuthPrompt] = useState(false);

  const [sspFiles, setSSPFiles] = useState(null);
  const [stage, setStage] = useState("idle");
  const [stepState, setStepState] = useState({1:"idle",2:"idle",3:"idle"});
  const [progress, setProgress] = useState({s2:{done:0,total:0}});
  const [chunkProgress, setChunkProgress] = useState({done:0,total:0});
  const [log, setLog] = useState([]);
  const [error, setError] = useState(null);

  const [allControls, setAllControls] = useState({});
  const [loadingStorage, setLoadingStorage] = useState(true);
  const [view, setView] = useState("dashboard");
  const [activeTab, setActiveTab] = useState("all");
  const [famFilter, setFamFilter] = useState("ALL");
  const [backendInfo, setBackendInfo] = useState(null);
  const logRef = useRef();
  const debounceTimers = useRef({});

  const addLog = (msg,type="info") => {
    setLog(l=>[...l,{msg,type,ts:Date.now()}]);
    setTimeout(()=>{if(logRef.current)logRef.current.scrollTop=logRef.current.scrollHeight;},40);
  };

  function handleAuthRequired() {
    setShowAuthPrompt(true);
    setStage("idle");
    setStepState({1:"idle",2:"idle",3:"idle"});
  }

  function submitToken(t) {
    storeToken(t);
    setAuthToken(t);
    setShowAuthPrompt(false);
    setLoadingStorage(true);
    loadAllControls(t).then(stored => {
      setAllControls(stored);
      setLoadingStorage(false);
      if (Object.keys(stored).length>0) addLog(`Loaded ${Object.keys(stored).length} previously documented controls.`);
    });
  }

  // Auto-save a POA&M field with 800 ms debounce
  function updatePoamField(controlId, field, value) {
    setAllControls(prev => {
      const updated = { ...prev[controlId], [field]: value };
      const timerKey = `${controlId}-${field}`;
      clearTimeout(debounceTimers.current[timerKey]);
      debounceTimers.current[timerKey] = setTimeout(() => {
        saveControl(controlId, updated, authToken);
      }, 800);
      return { ...prev, [controlId]: updated };
    });
  }

  useEffect(() => {
    (async () => {
      try {
        const r = await fetch(`${API_BASE}/api/health`);
        const info = await r.json();
        setBackendInfo(info);
        if (info.authEnabled && !authToken) setShowAuthPrompt(true);
      } catch { setBackendInfo({ ok:false }); }
    })();
  }, []);

  useEffect(() => {
    (async () => {
      const stored = await loadAllControls(authToken);
      setAllControls(stored);
      setLoadingStorage(false);
      if (Object.keys(stored).length>0) addLog(`Loaded ${Object.keys(stored).length} previously documented controls from storage.`);
    })();
  }, []);

  async function resetProgress() {
    await clearAllControls(authToken);
    setAllControls({});
    addLog("Progress reset — all stored control data cleared.", "warn");
  }

  async function runPipeline() {
    if (!sspFiles || sspFiles.length===0) { setError("Upload an SSP — even a single page covering one control works."); return; }
    setError(null); setStage("running"); setLog([]);
    setStepState({1:"active",2:"idle",3:"idle"});
    setProgress({s2:{done:0,total:0}});
    setChunkProgress({done:0,total:0});

    try {
      // ── STEP 1: Parse document ────────────────────────────────────
      const fileList = Array.from(sspFiles);
      addLog(`Parsing ${fileList.length} file${fileList.length>1?"s":""}: ${fileList.map(f=>f.name).join(", ")}`);
      let sspText;
      try {
        sspText = await parseSSPFiles(sspFiles, authToken);
      } catch(e) {
        if (e.message==="AUTH_REQUIRED") { handleAuthRequired(); return; }
        throw e;
      }
      addLog(`Extracted text: ${sspText.length.toLocaleString()} characters.`);

      // ── STEP 1b: Chunked control extraction ───────────────────────
      const chunks = chunkText(sspText, CHUNK_SIZE, CHUNK_OVERLAP);
      if (chunks.length > 1) addLog(`Document split into ${chunks.length} chunks (${CHUNK_SIZE.toLocaleString()} chars each) for extraction.`);
      setChunkProgress({done:0,total:chunks.length});

      const sys1 = `You are a federal cybersecurity analyst. Extract ALL NIST 800-53 Rev 5 control implementation statements from SSP text.
Return ONLY a valid JSON array, no markdown, no preamble.
Each element: {"id":"AC-2","family":"AC","statement":"...","status":"Implemented|Partially Implemented|Planned|Not Applicable"}
The document may cover just one control or many. Extract whatever is present.
Return [] if none found.`;

      const rawStatements = [];
      for (let i=0; i<chunks.length; i++) {
        if (chunks.length > 1) addLog(`Extracting controls from chunk ${i+1}/${chunks.length}…`);
        let raw;
        try {
          raw = await callClaude(sys1, `Extract all control statements:\n\n${chunks[i]}`, 8000, authToken);
        } catch(e) {
          if (e.message==="AUTH_REQUIRED") { handleAuthRequired(); return; }
          throw e;
        }
        const parsed = extractJSON(raw);
        if (Array.isArray(parsed)) rawStatements.push(...parsed);
        setChunkProgress({done:i+1,total:chunks.length});
      }

      // Deduplicate: same control ID across chunks — keep longer statement
      const stmtMap = {};
      for (const s of rawStatements) {
        if (!s?.id) continue;
        const id = String(s.id).toUpperCase().trim();
        if (!BASELINE_SET.has(id)) continue;
        if (!stmtMap[id] || (s.statement?.length||0) > (stmtMap[id].statement?.length||0))
          stmtMap[id] = { ...s, id };
      }
      const statements = Object.values(stmtMap);
      addLog(`Extracted ${statements.length} unique control statement(s): ${statements.map(s=>s.id).join(", ")||"(none)"}`);

      if (statements.length===0) {
        addLog("No recognizable NIST 800-53 Rev 5 control statements found in this document.", "warn");
        setStage("done"); setStepState({1:"done",2:"done",3:"done"}); return;
      }

      setStepState({1:"done",2:"active",3:"idle"});
      setProgress({s2:{done:0,total:statements.length}});

      // ── STEP 2: SCTM + eMASS + gap per control ───────────────────
      addLog(`Generating SCTM + eMASS + POA&M data for ${statements.length} control(s)…`);
      const sys2 = `You are a federal ISSO building SCTM and eMASS data for NIST SP 800-53 Rev 5 from an SSP statement.
You will be given the OFFICIAL NIST control text and discussion for context — use it to judge whether the
SSP statement actually satisfies the control's requirements.
Return ONLY valid JSON, no markdown. Keys:
  status: "Implemented"|"Partially Implemented"|"Planned"|"Not Applicable"|"Inherited"
  origination: "System Specific"|"Hybrid"|"Inherited"|"Common Control Provider"
  implementation_statement: 2-3 sentence eMASS-ready narrative refining the SSP statement
  test_method: "Test"|"Interview"|"Examine"|"Test, Interview"|"Test, Examine"
  gap: true|false
  gap_reason: string if gap else ""
  ai_suggestion: specific actionable recommendation if gap else ""
  emass_control_information: short control title/category text
  emass_security_designation: "Common"|"System-Specific"|"Hybrid"`;

      const mapOne = async (s) => {
        const controlId = s.id.toUpperCase();
        const fam = controlId.split("-")[0];
        const catalogEntry = await getCatalogEntry(controlId, authToken);
        const catalogContext = catalogEntry
          ? `\n\nOFFICIAL NIST CONTROL TEXT (${controlId} — ${catalogEntry.name}):\n${catalogEntry.text}\n\nOFFICIAL NIST DISCUSSION/GUIDANCE:\n${catalogEntry.discussion}`
          : "";
        let raw;
        try { raw = await callClaude(sys2, `Control: ${controlId}\nSSP Statement: ${s.statement}\nDeclared Status: ${s.status}${catalogContext}`, 800, authToken); }
        catch(e) { if (e.message==="AUTH_REQUIRED") throw e; raw=""; }
        let mapped = extractJSON(raw);
        if (!mapped || typeof mapped!=="object" || Array.isArray(mapped)) {
          mapped = { status:s.status??"Implemented", origination:"System Specific", implementation_statement:s.statement, test_method:"Test", gap:false, gap_reason:"", ai_suggestion:"", emass_control_information:FAMILY_NAMES[fam]??"", emass_security_designation:"System-Specific" };
        }
        const nistGuidance = catalogEntry?.discussion
          ? catalogEntry.discussion.slice(0,600)+(catalogEntry.discussion.length>600?"…":"")
          : `Refer to NIST SP 800-53A assessment procedures for ${controlId}.`;
        const cardStatus = mapped.gap?"Gap":mapped.status==="Inherited"?"Inherited":mapped.status==="Implemented"?"Matched":"Partial";
        return { controlId, fam, statement:s.statement, mapped, nistGuidance, ccis:catalogEntry?.ccis||[], cardStatus };
      };

      let results;
      try {
        results = await processBatch(statements, 5, 900, mapOne,
          (done,total)=>{ setProgress(p=>({...p,s2:{done,total}})); addLog(`[Step 2] Processed ${done}/${total}…`); });
      } catch(e) {
        if (e.message==="AUTH_REQUIRED") { handleAuthRequired(); return; }
        throw e;
      }

      setStepState({1:"done",2:"done",3:"active"});

      // ── STEP 3: Save ──────────────────────────────────────────────
      const updated = {...allControls};
      for (const r of results) {
        const existing = allControls[r.controlId] || {};
        const entry = {
          family:r.fam, statement:r.statement,
          status:r.cardStatus, sctmStatus:r.mapped.status,
          origination:r.mapped.origination,
          implementationStatement:r.mapped.implementation_statement,
          testMethod:r.mapped.test_method,
          gap:r.mapped.gap, gapReason:r.mapped.gap_reason, aiSuggestion:r.mapped.ai_suggestion,
          nistGuidance:r.nistGuidance||"", ccis:r.ccis||[],
          emassControlInfo:r.mapped.emass_control_information,
          emassDesignation:r.mapped.emass_security_designation,
          // Preserve any POA&M milestone fields the user already filled in
          poamDueDate: existing.poamDueDate||"",
          poamPOC: existing.poamPOC||"",
          poamResources: existing.poamResources||"",
          updatedAt:Date.now(),
        };
        updated[r.controlId] = entry;
        await saveControl(r.controlId, entry, authToken);
      }
      addLog(`Saved ${results.length} control(s). Total documented: ${Object.keys(updated).length}/${FULL_BASELINE.length}.`);
      setAllControls(updated);
      setStepState({1:"done",2:"done",3:"done"});
      setStage("done");
      addLog("✅ Done. Dashboard updated with cumulative progress.", "success");

    } catch(err) {
      setError(`Pipeline error: ${err.message}`);
      setStage("error");
      setStepState(s=>{const n={...s};for(const k of Object.keys(n))if(n[k]==="active")n[k]="error";return n;});
    }
  }

  const running = stage==="running";
  const documentedIds = Object.keys(allControls);
  const totalDocumented = documentedIds.length;
  const overallPct = Math.round((totalDocumented/FULL_BASELINE.length)*100);

  const allCards = FULL_BASELINE.map(id => {
    const fam = id.split("-")[0];
    if (allControls[id]) {
      const c = allControls[id];
      return { id, family:fam, status:c.status, statement:c.statement, aiSuggestion:c.aiSuggestion, nistGuidance:c.nistGuidance, documented:true, raw:c };
    }
    return { id, family:fam, status:"Gap", statement:null, aiSuggestion:null, documented:false, raw:null };
  });

  const FAMILIES = ["ALL",...Object.keys(FAMILY_NAMES)];
  const filteredCards = allCards.filter(c=>{
    const famOk = famFilter==="ALL"||c.family===famFilter;
    const tabOk = activeTab==="all"?true:activeTab==="documented"?c.documented:activeTab==="undocumented"?!c.documented:activeTab==="gaps"?(c.documented&&c.status==="Gap"):activeTab==="matched"?c.status==="Matched":activeTab==="partial"?c.status==="Partial":activeTab==="inherited"?c.status==="Inherited":true;
    return famOk&&tabOk;
  });

  const familySummaries = Object.keys(FAMILY_NAMES).map(fam=>{
    const famControls = FULL_BASELINE.filter(id=>id.startsWith(fam+"-"));
    const documented = famControls.filter(id=>allControls[id]);
    const matched = documented.filter(id=>allControls[id].status==="Matched").length;
    const partial = documented.filter(id=>allControls[id].status==="Partial").length;
    const gaps = documented.filter(id=>allControls[id].status==="Gap").length;
    const inherited = documented.filter(id=>allControls[id].status==="Inherited").length;
    const undocumented = famControls.length - documented.length;
    const ph = (cond1, cond2, cond3) => cond1?"Complete":cond2?"In Progress":"Not Started";
    return { id:fam, name:FAMILY_NAMES[fam], total:famControls.length, documented:documented.length,
      matched, partial, gaps, inherited, undocumented,
      phases:[
        ph(false, documented.length>0, false),
        ph(documented.length===famControls.length&&gaps===0, documented.length>0, false),
        ph(documented.length===famControls.length&&gaps===0&&partial===0, documented.length>0, false),
        ph(documented.length===famControls.length, documented.length>0, false),
        "Not Started",
      ]};
  });

  // Memoized exports — only recomputed when allControls changes
  const gapIds = useMemo(() => documentedIds.filter(id=>allControls[id]?.gap), [allControls]);

  const poamRows = useMemo(() => gapIds.map(id=>{
    const c = allControls[id]; const fam = id.split("-")[0];
    return {
      "POA&M ID":`POAM-${id.replace(/[^A-Z0-9]/g,"")}`, "Control ID":id,
      "Control Family":FAMILY_NAMES[fam]??fam, "Weakness Description":c.gapReason||"",
      "Recommended Actions":c.aiSuggestion||"", "Status":"Ongoing",
      "Est Completion Date":c.poamDueDate||"",
      "Responsible POC":c.poamPOC||"",
      "Resource Estimate":c.poamResources||"",
    };
  }), [allControls]);

  const emassRowCount = useMemo(() => buildEmassRows(allControls, documentedIds).length, [allControls]);
  const poamComplete = useMemo(() => gapIds.filter(id=>allControls[id]?.poamDueDate&&allControls[id]?.poamPOC&&allControls[id]?.poamResources).length, [allControls]);

  function downloadSCTM() {
    const wb = buildSCTMWorkbook(allControls, documentedIds);
    XLSX.writeFile(wb, "SCTM_Rev5.xlsx");
    addLog(`SCTM exported — ${documentedIds.length} controls across 20 family tabs.`);
  }
  function downloadeMASSFile() {
    const wb = buildEmassWorkbook(allControls, documentedIds);
    XLSX.writeFile(wb, "eMASS_Upload_Rev5.xlsx");
    addLog(`eMASS export — ${emassRowCount} base-control rows.`);
  }
  function downloadTracker() {
    const rows = familySummaries.map(f=>({
      "Family ID":f.id,"Family Name":f.name,"Total Controls":f.total,"Documented":f.documented,"Undocumented":f.undocumented,
      "Matched":f.matched,"Partial":f.partial,"Gaps":f.gaps,"Inherited":f.inherited,
      "Phase 1 - Gap Analysis":f.phases[0],"Phase 2 - SSP/Policy":f.phases[1],
      "Phase 3 - Evidence":f.phases[2],"Phase 4 - eMASS Entry":f.phases[3],"Phase 5 - AO Review":f.phases[4],
    }));
    downloadExcel(rows,"Migration_Tracker.xlsx","Tracker");
  }

  return (
    <div style={{fontFamily:"'Inter','Arial',sans-serif",background:C.navy,minHeight:"100vh",paddingBottom:40}}>
      {showAuthPrompt && <AuthPrompt onSubmit={submitToken}/>}

      {/* Header */}
      <div style={{background:`linear-gradient(135deg,${C.navy},${C.navyMid})`,padding:"20px 26px 16px",borderBottom:`3px solid ${C.steel}`}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:12}}>
          <div style={{display:"flex",alignItems:"center",gap:11}}>
            <div style={{background:C.accent,borderRadius:8,padding:"5px 9px",fontSize:19}}>🛡️</div>
            <div>
              <div style={{color:C.white,fontWeight:800,fontSize:16,letterSpacing:"-0.02em"}}>NIST 800-53 Rev 4 → Rev 5 · Incremental AI Pipeline</div>
              <div style={{color:"#8AAFC8",fontSize:11,marginTop:1}}>Upload SSP content one control (or one family) at a time — progress accumulates automatically</div>
            </div>
          </div>
          <div style={{display:"flex",alignItems:"center",gap:12}}>
            {authToken && (
              <button onClick={()=>{storeToken("");setAuthToken("");setShowAuthPrompt(true);}}
                style={{background:"none",border:`1px solid ${C.steel}`,color:"#8AAFC8",borderRadius:6,padding:"3px 9px",fontSize:10,cursor:"pointer"}}>
                🔐 Re-auth
              </button>
            )}
            <div style={{textAlign:"right"}}>
              <div style={{color:C.white,fontWeight:800,fontSize:22}}>{totalDocumented} / {FULL_BASELINE.length}</div>
              <div style={{color:"#8AAFC8",fontSize:10,fontWeight:600}}>controls documented ({overallPct}%)</div>
            </div>
          </div>
        </div>
      </div>

      {/* Backend status bar */}
      {backendInfo && (
        <div style={{background:backendInfo.ok?(backendInfo.apiKeyConfigured?"#0F2A1A":"#3A2A0A"):"#3A0F0A",borderBottom:"1px solid #1E3A52",padding:"6px 26px",fontSize:10,fontFamily:"monospace",color:"#9FB8CC",display:"flex",gap:14,flexWrap:"wrap",alignItems:"center"}}>
          {backendInfo.ok ? (
            <>
              <span>🔌 backend: connected</span>
              <span>provider: <strong style={{color:"#fff"}}>{backendInfo.provider}</strong></span>
              <span>model: <strong style={{color:"#fff"}}>{backendInfo.model}</strong></span>
              <span>endpoint: {backendInfo.baseUrl}</span>
              {backendInfo.authEnabled && <span style={{color:"#27AE60",fontWeight:700}}>🔐 auth: on</span>}
              {!backendInfo.apiKeyConfigured && <span style={{color:"#F39C12",fontWeight:700}}>⚠ LLM_API_KEY not set — pipeline runs will fail</span>}
            </>
          ) : (
            <span style={{color:"#E74C3C",fontWeight:700}}>⚠ Cannot reach backend at {API_BASE||"(same origin)"}/api/health — is the server running?</span>
          )}
        </div>
      )}

      <div style={{padding:"20px 26px",maxWidth:1100,margin:"0 auto"}}>

        {/* Progress */}
        <div style={{background:C.white,borderRadius:12,border:"1px solid #DDE6EF",padding:"14px 18px",marginBottom:14}}>
          <ProgressBar value={totalDocumented} max={FULL_BASELINE.length} label="Overall baseline coverage (cumulative across all uploads)" color={C.green}/>
          {totalDocumented>0 && (
            <button onClick={resetProgress} style={{marginTop:6,background:"none",border:`1px solid ${C.red}55`,color:C.red,borderRadius:6,padding:"4px 10px",fontSize:10,fontWeight:600,cursor:"pointer"}}>
              ↺ Reset all progress
            </button>
          )}
        </div>

        {/* Upload */}
        <div style={{background:C.white,borderRadius:12,border:"1px solid #DDE6EF",padding:"16px 18px",marginBottom:14}}>
          <div style={{fontWeight:700,color:C.navy,fontSize:13,marginBottom:4}}>Upload SSP Content</div>
          <div style={{fontSize:11,color:C.grayMid,marginBottom:12}}>
            Upload as little or as much as you have. Large documents are automatically chunked for extraction — no page limit.
            Accepts <strong>.docx</strong>, <strong>.pdf</strong>, and <strong>.txt</strong>.
          </div>
          <DropZone label="SSP document(s) — .docx / .pdf / .txt, any size" accept=".docx,.pdf,.txt" icon="📄" multi onFiles={setSSPFiles} files={sspFiles} color={C.steel}/>
        </div>

        {/* Format note */}
        <div style={{background:C.white,borderRadius:12,border:"1px solid #DDE6EF",padding:"14px 18px",marginBottom:14}}>
          <div style={{fontSize:11,color:C.grayMid,lineHeight:1.6}}>
            <strong style={{color:C.navy}}>Standard output formats:</strong> SCTM is a 20-tab workbook (one tab per control family). eMASS export is one row per base control (enhancements rolled up, ≤2,000 chars). Guidance text comes directly from the official NIST SP 800-53 Rev 5 catalog.
          </div>
        </div>

        {/* Pipeline run */}
        <div style={{background:C.white,borderRadius:12,border:"1px solid #DDE6EF",padding:"16px 18px",marginBottom:14}}>
          <div style={{display:"flex",gap:14,flexWrap:"wrap",alignItems:"center",marginBottom:10}}>
            <StepBadge n={1} label="Parse + Extract from SSP" state={stepState[1]}/>
            <span style={{color:"#CCC"}}>▶</span>
            <StepBadge n={2} label="SCTM + eMASS + Gap Check" state={stepState[2]}/>
            <span style={{color:"#CCC"}}>▶</span>
            <StepBadge n={3} label="Save to Tracker" state={stepState[3]}/>
          </div>
          {running && chunkProgress.total>1 && (
            <ProgressBar value={chunkProgress.done} max={chunkProgress.total} label="Extraction chunks" color={C.steel}/>
          )}
          {running && progress.s2.total>0 && (
            <ProgressBar value={progress.s2.done} max={progress.s2.total} label="Processing controls" color={C.accent}/>
          )}
          <button onClick={runPipeline} disabled={running||!sspFiles} style={{
            background:running?C.grayMid:!sspFiles?C.grayLight:`linear-gradient(135deg,${C.steel},${C.navyMid})`,
            color:!sspFiles?C.grayMid:C.white,border:"none",borderRadius:8,padding:"10px 24px",fontWeight:700,fontSize:13,
            cursor:running||!sspFiles?"not-allowed":"pointer",
          }}>{running?"⏳ Processing…":"▶ Process This Upload"}</button>
          {!sspFiles && <span style={{marginLeft:10,fontSize:11,color:C.grayMid}}>Upload SSP content to begin</span>}
        </div>

        {error && <div style={{background:C.redLight,border:`1px solid ${C.red}66`,borderRadius:8,padding:"10px 14px",color:C.red,fontSize:12,marginBottom:12}}>⚠️ {error}</div>}

        {log.length>0 && (
          <div ref={logRef} style={{background:"#0B1520",borderRadius:10,border:"1px solid #1E3A52",padding:"11px 13px",marginBottom:14,maxHeight:130,overflowY:"auto",fontFamily:"monospace",fontSize:10}}>
            {log.map((e,i)=>(<div key={i} style={{color:e.type==="warn"?"#F39C12":e.type==="success"?"#27AE60":e.type==="error"?"#E74C3C":"#8AAFC8",marginBottom:2,lineHeight:1.5}}><span style={{color:"#3A5A72"}}>[{new Date(e.ts).toLocaleTimeString()}]</span> {e.msg}</div>))}
          </div>
        )}

        {!loadingStorage && (
          <div>
            {/* Navigation */}
            <div style={{display:"flex",gap:8,marginBottom:14}}>
              <button onClick={()=>setView("dashboard")} style={{background:view==="dashboard"?C.steel:C.grayLight,color:view==="dashboard"?C.white:C.gray,border:"none",borderRadius:20,padding:"6px 16px",fontWeight:700,fontSize:12,cursor:"pointer"}}>📊 Tracker Dashboard</button>
              <button onClick={()=>setView("controls")} style={{background:view==="controls"?C.steel:C.grayLight,color:view==="controls"?C.white:C.gray,border:"none",borderRadius:20,padding:"6px 16px",fontWeight:700,fontSize:12,cursor:"pointer"}}>🔍 Control-Level Detail</button>
              <button onClick={()=>setView("poam")} style={{
                background:view==="poam"?C.red:C.grayLight,
                color:view==="poam"?C.white:gapIds.length>0?C.red:C.gray,
                border:gapIds.length>0&&view!=="poam"?`1px solid ${C.red}55`:"none",
                borderRadius:20,padding:"6px 16px",fontWeight:700,fontSize:12,cursor:"pointer"
              }}>🔴 POA&M Manager {gapIds.length>0?`(${poamComplete}/${gapIds.length})`:""}</button>
            </div>

            {/* Dashboard view */}
            {view==="dashboard" && (
              <div>
                <div style={{background:C.white,borderRadius:12,border:"1px solid #DDE6EF",padding:"20px 22px",marginBottom:14,display:"grid",gridTemplateColumns:"140px 1fr",gap:24,alignItems:"center"}}>
                  <div style={{width:120,height:120,borderRadius:"50%",border:`3px solid ${C.navy}`,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center"}}>
                    <div style={{fontSize:32,fontWeight:800,color:C.navy}}>{overallPct}%</div>
                    <div style={{fontSize:9,fontWeight:700,color:C.steel,textTransform:"uppercase",letterSpacing:"0.08em"}}>Documented</div>
                  </div>
                  <div>
                    <div style={{fontSize:16,fontWeight:700,color:C.navy,marginBottom:8}}>
                      {totalDocumented===0?"No controls documented yet — upload SSP content to begin":overallPct===100?"All 323 baseline controls documented":`${totalDocumented} of ${FULL_BASELINE.length} controls documented so far`}
                    </div>
                    <div style={{display:"flex",gap:24,flexWrap:"wrap"}}>
                      {[
                        {label:"Documented",val:totalDocumented,fg:C.steel},
                        {label:"✅ Matched",val:documentedIds.filter(id=>allControls[id].status==="Matched").length,fg:C.green},
                        {label:"🟡 Partial",val:documentedIds.filter(id=>allControls[id].status==="Partial").length,fg:C.yellow},
                        {label:"🔴 Gaps",val:gapIds.length,fg:C.red},
                        {label:"Remaining",val:FULL_BASELINE.length-totalDocumented,fg:C.grayMid},
                      ].map(k=>(<div key={k.label}><div style={{fontSize:22,fontWeight:800,color:k.fg}}>{k.val}</div><div style={{fontSize:10,color:C.grayMid,fontWeight:600}}>{k.label}</div></div>))}
                    </div>
                  </div>
                </div>

                <div style={{background:C.white,borderRadius:12,border:"1px solid #DDE6EF",padding:"15px 18px",marginBottom:14}}>
                  <div style={{fontWeight:700,color:C.navy,fontSize:13,marginBottom:10}}>Download Cumulative Outputs</div>
                  <div style={{fontSize:11,color:C.grayMid,marginBottom:10}}>Reflects all {totalDocumented} controls documented so far.</div>
                  <div style={{display:"flex",gap:9,flexWrap:"wrap"}}>
                    <button disabled={totalDocumented===0} onClick={downloadSCTM} style={{background:totalDocumented?`linear-gradient(135deg,${C.green},#145C30)`:C.grayLight,color:totalDocumented?C.white:C.grayMid,border:"none",borderRadius:7,padding:"9px 16px",fontWeight:700,fontSize:12,cursor:totalDocumented?"pointer":"not-allowed"}}>📊 SCTM ({totalDocumented} controls)</button>
                    <button disabled={totalDocumented===0} onClick={downloadeMASSFile} style={{background:totalDocumented?`linear-gradient(135deg,${C.purple},#4A235A)`:C.grayLight,color:totalDocumented?C.white:C.grayMid,border:"none",borderRadius:7,padding:"9px 16px",fontWeight:700,fontSize:12,cursor:totalDocumented?"pointer":"not-allowed"}}>🗄️ eMASS ({emassRowCount} rows)</button>
                    <button disabled={gapIds.length===0} onClick={()=>downloadExcel(poamRows,"POAM_Cumulative.xlsx","POA&M")} style={{background:gapIds.length?`linear-gradient(135deg,${C.red},#9B2D23)`:C.grayLight,color:gapIds.length?C.white:C.grayMid,border:"none",borderRadius:7,padding:"9px 16px",fontWeight:700,fontSize:12,cursor:gapIds.length?"pointer":"not-allowed"}}>🔴 POA&M ({gapIds.length} items{poamComplete>0?`, ${poamComplete} complete`:""})</button>
                    <button onClick={downloadTracker} style={{background:`linear-gradient(135deg,${C.steel},${C.navyMid})`,color:C.white,border:"none",borderRadius:7,padding:"9px 16px",fontWeight:700,fontSize:12,cursor:"pointer"}}>📋 Tracker (20 families)</button>
                  </div>
                </div>

                <div style={{background:C.white,borderRadius:12,border:"1px solid #DDE6EF",padding:"15px 18px"}}>
                  <div style={{fontWeight:700,color:C.navy,fontSize:13,marginBottom:4}}>Family Status</div>
                  <div style={{fontSize:11,color:C.grayMid,marginBottom:12}}>Click a family to view its controls.</div>
                  <div style={{display:"grid",gridTemplateColumns:"60px 1.4fr 1fr repeat(5,0.7fr)",gap:6,fontSize:10,fontWeight:700,color:C.steel,textTransform:"uppercase",letterSpacing:"0.04em",borderBottom:`1px solid ${C.grayLight}`,paddingBottom:6,marginBottom:6}}>
                    <div>ID</div><div>Family</div><div>Coverage</div>
                    {PHASE_LABELS.map((p,i)=><div key={i} style={{textAlign:"center"}}>P{i+1}</div>)}
                  </div>
                  {familySummaries.map(f=>(
                    <div key={f.id} onClick={()=>{setView("controls");setFamFilter(f.id);setActiveTab("all");}}
                      style={{display:"grid",gridTemplateColumns:"60px 1.4fr 1fr repeat(5,0.7fr)",gap:6,alignItems:"center",padding:"7px 0",borderBottom:`1px solid ${C.grayLight}`,cursor:"pointer"}}>
                      <div style={{fontWeight:800,fontSize:12,color:C.navy}}>{f.id}</div>
                      <div style={{fontSize:11,color:C.gray}}>{f.name}</div>
                      <div style={{fontSize:10,color:C.gray}}>{f.documented}/{f.total}
                        {f.documented>0&&<span style={{marginLeft:6,fontSize:9,color:C.grayMid}}>({f.matched}✅ {f.partial}🟡 {f.gaps}🔴{f.inherited?` ${f.inherited}🔵`:""})</span>}
                      </div>
                      {f.phases.map((ph,i)=>{const sc=statusClass(ph);return(<div key={i} style={{textAlign:"center"}}><span style={{fontSize:9,fontWeight:700,color:sc.fg,background:sc.bg,border:`1px solid ${sc.fg}44`,padding:"2px 5px",borderRadius:10}}>{ph==="Complete"?"✓":ph==="In Progress"?"…":ph==="Blocked"?"✕":"–"}</span></div>);})}
                    </div>
                  ))}
                  <div style={{marginTop:10,fontSize:10,color:C.grayMid,display:"flex",gap:16,flexWrap:"wrap"}}>
                    <span><strong>P1</strong> Gap Analysis</span><span><strong>P2</strong> SSP/Policy</span>
                    <span><strong>P3</strong> Evidence</span><span><strong>P4</strong> eMASS Entry</span><span><strong>P5</strong> AO Review (manual)</span>
                  </div>
                </div>
              </div>
            )}

            {/* Controls view */}
            {view==="controls" && (
              <div style={{background:C.white,borderRadius:12,border:"1px solid #DDE6EF",padding:"15px 18px"}}>
                <div style={{display:"flex",justifyContent:"space-between",flexWrap:"wrap",gap:8,marginBottom:10}}>
                  <div style={{display:"flex",gap:5,flexWrap:"wrap"}}>
                    {[
                      {key:"all",label:`All (${allCards.length})`,color:C.steel},
                      {key:"documented",label:`📝 Documented (${documentedIds.length})`,color:C.navyMid},
                      {key:"undocumented",label:`⬜ Not Yet (${FULL_BASELINE.length-documentedIds.length})`,color:C.grayMid},
                      {key:"matched",label:`✅ Matched (${documentedIds.filter(id=>allControls[id].status==="Matched").length})`,color:C.green},
                      {key:"partial",label:`🟡 Partial (${documentedIds.filter(id=>allControls[id].status==="Partial").length})`,color:C.yellow},
                      {key:"gaps",label:`🔴 Gaps (${gapIds.length})`,color:C.red},
                    ].map(t=>(<button key={t.key} onClick={()=>setActiveTab(t.key)} style={{background:activeTab===t.key?t.color:C.grayLight,color:activeTab===t.key?C.white:C.gray,border:"none",borderRadius:20,padding:"4px 11px",fontWeight:600,fontSize:11,cursor:"pointer"}}>{t.label}</button>))}
                  </div>
                  <select value={famFilter} onChange={e=>setFamFilter(e.target.value)} style={{fontSize:11,border:`1px solid ${C.grayLight}`,borderRadius:5,padding:"4px 9px",color:C.navy,background:C.white}}>
                    {FAMILIES.map(f=><option key={f} value={f}>{f==="ALL"?"All Families":`${f} — ${FAMILY_NAMES[f]}`}</option>)}
                  </select>
                </div>
                <div style={{fontSize:11,color:C.grayMid,marginBottom:7}}>Showing {filteredCards.length} of {allCards.length} baseline controls</div>
                <div style={{maxHeight:520,overflowY:"auto"}}>
                  {filteredCards.map(c=>c.documented?<ControlCard key={c.id} control={c}/>:(
                    <div key={c.id} style={{border:`1px dashed ${C.grayLight}`,borderRadius:7,padding:"7px 11px",marginBottom:5,display:"flex",justifyContent:"space-between",alignItems:"center",background:C.offwhite}}>
                      <span style={{fontFamily:"monospace",fontWeight:700,fontSize:11,color:C.grayMid}}>{c.id}</span>
                      <span style={{fontSize:10,color:C.grayMid}}>{FAMILY_NAMES[c.family]}</span>
                      <span style={{fontSize:10,color:C.grayMid,fontStyle:"italic"}}>not yet uploaded</span>
                    </div>
                  ))}
                  {filteredCards.length===0 && <div style={{textAlign:"center",color:C.grayMid,padding:"24px 0",fontSize:13}}>No controls match this filter.</div>}
                </div>
              </div>
            )}

            {/* POA&M Manager view */}
            {view==="poam" && (
              <div style={{background:C.white,borderRadius:12,border:"1px solid #DDE6EF",padding:"15px 18px"}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10,flexWrap:"wrap",gap:8}}>
                  <div>
                    <div style={{fontWeight:700,color:C.navy,fontSize:14}}>POA&M Manager</div>
                    <div style={{fontSize:11,color:C.grayMid,marginTop:2}}>
                      {gapIds.length===0?"No gaps identified yet — run the pipeline to detect gaps."
                        :`${gapIds.length} gap${gapIds.length>1?"s":""} — ${poamComplete} of ${gapIds.length} have all milestone fields filled. Fields auto-save as you type.`}
                    </div>
                  </div>
                  {gapIds.length>0 && (
                    <button onClick={()=>downloadExcel(poamRows,"POAM_Cumulative.xlsx","POA&M")}
                      style={{background:`linear-gradient(135deg,${C.red},#9B2D23)`,color:C.white,border:"none",borderRadius:7,padding:"9px 16px",fontWeight:700,fontSize:12,cursor:"pointer"}}>
                      🔴 Export POA&M ({poamComplete}/{gapIds.length} complete)
                    </button>
                  )}
                </div>
                {gapIds.length>0 && poamComplete<gapIds.length && (
                  <div style={{background:C.yellowLight,border:`1px solid ${C.yellow}55`,borderRadius:7,padding:"8px 12px",marginBottom:12,fontSize:11,color:C.yellow}}>
                    ⚠ {gapIds.length-poamComplete} POA&M item{gapIds.length-poamComplete>1?"s are":" is"} missing milestone fields (Due Date, POC, or Resource Estimate). Fill them in below before exporting for AO review.
                  </div>
                )}
                {gapIds.length===0 ? (
                  <div style={{textAlign:"center",color:C.grayMid,padding:"40px 0",fontSize:13}}>No gaps detected. Upload and process SSP content to identify gaps.</div>
                ) : (
                  <div style={{maxHeight:560,overflowY:"auto"}}>
                    {gapIds.map(id=>(
                      <POAMRow key={id} id={id} control={allControls[id]} onFieldChange={updatePoamField}/>
                    ))}
                  </div>
                )}
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
}
