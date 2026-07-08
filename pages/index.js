import React, { useState, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import { createClient } from "@supabase/supabase-js";

const SUPABASE_URL = process.env.NEXT_PUBLIC_SUPABASE_URL;
const SUPABASE_KEY = process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY;
const sb = SUPABASE_URL && SUPABASE_KEY ? createClient(SUPABASE_URL, SUPABASE_KEY) : null;

const RED = "#C00000";

const SYSTEM_PROMPT = `You are an expert education recruitment assistant for rTriibe, a UK-founded agency placing teachers in UAE international schools.
Extract structured candidate information from the provided CV/resume.
Return ONLY a valid JSON object with exactly these keys (use "" for any field not found):
full_name, nationality, email, phone, location, degree_subject, degree_university,
teaching_qualification, qts_holder, additional_certs, years_experience, teaching_level,
subject_specialism, current_last_school, current_last_role, curriculum,
available_from, notice_period, role_type, notes.
Values: qts_holder "Yes"/"No"/"In Progress"; teaching_level "EYFS"/"Primary"/"Secondary"/"All Through"/"SEN"/"LSA"/"Leadership"; curriculum "British"/"IB"/"American"/"CBSE"/"MOE"/"French"/"Multiple"/"Unknown"; role_type "Permanent"/"Supply"/"Both"/"Tutor"/"LSA".
Return ONLY the JSON object. No markdown, no backticks, no other text.`;

const MATCH_PROMPT = `You are an expert education recruitment consultant at rTriibe, placing teachers in UAE international schools.
Evaluate how well each candidate matches the vacancy. Score 0-100 and sort highest first.
Labels: Strong Match (90-100), Good Match (70-89), Partial Match (50-69), Weak Match (0-49).
Consider: subject/specialism fit, teaching level, curriculum, QTS for British schools, years of experience, availability, and notes/flags.
Return ONLY a JSON array:
[{"id":"_id value","name":"full name","score":85,"label":"Good Match","reason":"2 concise sentences on match strength and key concerns."}]
No markdown, no backticks, no other text.`;

const DETAIL_FIELDS = [
  ["full_name","Full Name"],["nationality","Nationality"],["email","Email"],["phone","Phone"],["location","Location"],
  ["degree_subject","Degree Subject"],["degree_university","Degree University"],
  ["teaching_qualification","Teaching Qualification"],["qts_holder","QTS Holder"],
  ["additional_certs","Additional Certifications"],["years_experience","Years of Experience"],
  ["teaching_level","Teaching Level"],["subject_specialism","Subject / Specialism"],
  ["current_last_school","Current / Last School"],["current_last_role","Current / Last Role"],
  ["curriculum","Curriculum"],["available_from","Available From"],
  ["notice_period","Notice Period"],["role_type","Role Type"],
];

const CATEGORIES = [
  { label:"IDENTITY", icon:"🪪", color:"#37474F", cols:[
    { key:"candidate_id", label:"ID (rTR)", width:88, editable:true },
    { key:"full_name", label:"Full Name", width:145 },
    { key:"nationality", label:"Nationality", width:100 },
  ]},
  { label:"CONTACT", icon:"📞", color:"#1565C0", cols:[
    { key:"email", label:"Email", width:170 },
    { key:"phone", label:"Phone", width:105 },
    { key:"location", label:"Location", width:110 },
  ]},
  { label:"QUALIFICATIONS", icon:"🎓", color:"#4A148C", cols:[
    { key:"degree_subject", label:"Degree Subject", width:120 },
    { key:"degree_university", label:"University", width:130 },
    { key:"teaching_qualification", label:"Teaching Qual", width:105 },
    { key:"qts_holder", label:"QTS", width:68 },
    { key:"additional_certs", label:"Add. Certs", width:110 },
  ]},
  { label:"EXPERIENCE", icon:"📚", color:"#1B5E20", cols:[
    { key:"years_experience", label:"Yrs", width:45 },
    { key:"teaching_level", label:"Level", width:95 },
    { key:"subject_specialism", label:"Specialism", width:130 },
    { key:"current_last_school", label:"Last School", width:145 },
    { key:"current_last_role", label:"Last Role", width:125 },
  ]},
  { label:"CURRICULUM", icon:"🌍", color:"#006064", cols:[
    { key:"curriculum", label:"Curriculum", width:95 },
  ]},
  { label:"AVAILABILITY", icon:"📅", color:"#BF360C", cols:[
    { key:"available_from", label:"Available From", width:100 },
    { key:"notice_period", label:"Notice", width:80 },
    { key:"role_type", label:"Role Type", width:95 },
  ]},
  { label:"NOTES", icon:"📝", color:"#880E4F", cols:[
    { key:"notes", label:"Notes / Flags", width:190 },
  ]},
];

const ALL_COLS = CATEGORIES.flatMap(cat => cat.cols);
const TABLE_W  = ALL_COLS.reduce((s,c) => s+c.width, 0) + 38;

const SUBJECT_SYNONYMS = {
  "english":["english","language arts","esl","eal","elt","ell","literacy"],
  "maths":["math","maths","mathematics","numeracy"],
  "math":["math","maths","mathematics","numeracy"],
  "mathematics":["math","maths","mathematics","numeracy"],
  "science":["science","biology","chemistry","physics","stem"],
  "biology":["biology","science"],
  "chemistry":["chemistry","science"],
  "physics":["physics","science"],
  "pe":["pe","physical education","sport","sports"],
  "physical":["pe","physical education","sport","sports"],
  "ict":["ict","computing","computer","technology","it","digital"],
  "computing":["ict","computing","computer","technology"],
  "arabic":["arabic"],
  "french":["french"],
  "spanish":["spanish"],
  "geography":["geography"],
  "history":["history","humanities","social studies"],
  "humanities":["history","humanities","geography"],
  "art":["art","arts","design","creative"],
  "music":["music"],
  "drama":["drama","theatre","theater","performing arts"],
  "business":["business","economics","commerce"],
  "economics":["economics","business"],
  "religious":["religious","re","islam","islamic","divinity"],
  "islamic":["islamic","islam","religious","re"],
  "lsa":["lsa","learning support","sen","send","special needs","inclusion"],
  "send":["send","sen","special needs","lsa","learning support"],
  "sen":["send","sen","special needs","lsa","learning support"],
  "eyfs":["eyfs","early years","foundation","nursery","reception","kindergarten"],
  "esl":["esl","eal","elt","ell","english"],
  "eal":["eal","esl","elt","ell","english"],
};

const preFilterBySubject = (vacancyText, allCandidates) => {
  const vacLower = vacancyText.toLowerCase();
  const matched = allCandidates.filter(c => {
    const spec = (c.subject_specialism||"").toLowerCase();
    if (!spec) return false;
    const words = spec.split(/[\s,\/&\-+]+/).filter(w => w.length >= 2);
    return words.some(word => {
      const synonyms = SUBJECT_SYNONYMS[word] || [word];
      return synonyms.some(syn => {
        try { return new RegExp("\\b"+syn.replace(/[.*+?^${}()|[\]\\]/g,"\\$&")+"\\b","i").test(vacLower); }
        catch { return vacLower.includes(syn); }
      });
    });
  });
  if (matched.length === 0) return { pool: allCandidates, info: null };
  return { pool: matched, info: `Subject filter applied — checking ${matched.length} of ${allCandidates.length} candidates` };
};

const btn = (bg, color, extra={}) => ({ border:"none", borderRadius:6, padding:"8px 18px", fontWeight:"bold", fontSize:12, cursor:"pointer", fontFamily:"Arial", background:bg, color, ...extra });
const initials = (name) => (name||"?").split(" ").map(n=>n[0]).slice(0,2).join("");
const scoreStyle = (score) => {
  if (score>=90) return { bg:"#e8f5e9", color:"#2e7d32", border:"#a5d6a7" };
  if (score>=70) return { bg:"#fff8e1", color:"#f57f17", border:"#ffe082" };
  if (score>=50) return { bg:"#fff3e0", color:"#e65100", border:"#ffcc80" };
  return { bg:"#ffebee", color:"#c62828", border:"#ef9a9a" };
};

const QTSBadge = ({ val }) => (
  <span style={{ padding:"1px 7px", borderRadius:8, fontSize:10, fontWeight:"bold",
    background:val==="Yes"?"#e8f5e9":val==="In Progress"?"#fff8e1":"#f5f5f5",
    color:val==="Yes"?"#2e7d32":val==="In Progress"?"#e65100":"#999" }}>{val||"—"}</span>
);

const Cell = ({ col, c, onUpdate }) => {
  const val = c[col.key]||"";
  if (col.editable) return (
    <input value={val} onChange={e=>{e.stopPropagation();onUpdate(c._id,col.key,e.target.value);}} onClick={e=>e.stopPropagation()} placeholder="rTR..."
      style={{ width:"100%", border:"1px solid #ddd", borderRadius:4, padding:"2px 5px", fontSize:10, fontFamily:"Arial", outline:"none", boxSizing:"border-box" }} />
  );
  if (col.key==="qts_holder") return <QTSBadge val={val}/>;
  if (col.key==="notes") return <span style={{ color:val?"#b71c1c":"#ddd", fontSize:10 }} title={val}>{val?(val.length>50?val.slice(0,50)+"…":val):"—"}</span>;
  return <span style={{ color:val?"#333":"#ccc" }} title={val}>{val||"—"}</span>;
};

const iStyle = { flex:1, border:"1px solid #e0e0e0", borderRadius:5, padding:"4px 9px", fontSize:12, fontFamily:"Arial", outline:"none", background:"#fafafa", boxSizing:"border-box" };
const sStyle = { flex:1, border:"1px solid #e0e0e0", borderRadius:5, padding:"4px 9px", fontSize:12, fontFamily:"Arial", outline:"none", background:"#fafafa", cursor:"pointer" };

export default function App() {
  const [view, setView]             = useState("extract");
  const [candidates, setCandidates] = useState([]);
  const [selectedCandidate, setSelectedCandidate] = useState(null);
  const [dbReady, setDbReady]       = useState(false);
  const [mode, setMode]             = useState("paste");
  const [cvText, setCvText]         = useState("");
  const [pdfBase64, setPdfBase64]   = useState(null);
  const [fileName, setFileName]     = useState("");
  const [loading, setLoading]       = useState(false);
  const [result, setResult]         = useState(null);
  const [error, setError]           = useState("");
  const [justAdded, setJustAdded]   = useState(false);
  const [confirmClear, setConfirmClear] = useState(false);
  const [bulkFiles, setBulkFiles]   = useState([]);
  const [bulkCurrent, setBulkCurrent] = useState(0);
  const [bulkResults, setBulkResults] = useState([]);
  const [bulkDone, setBulkDone]     = useState(false);
  const [bulkPaused, setBulkPaused] = useState(false);
  const [bulkStats, setBulkStats]   = useState({ added:0, duplicates:0, failed:0 });
  const [bulkStartTime, setBulkStartTime] = useState(0);
  const [failedFiles, setFailedFiles] = useState([]);
  const shouldPauseRef = useRef(false);
  const addedNamesRef  = useRef(new Set());
  const candidatesRef  = useRef([]);
  const [searchName, setSearchName] = useState("");
  const [filterLevel, setFilterLevel] = useState("");
  const [filterCurriculum, setFilterCurriculum] = useState("");
  const [filterQTS, setFilterQTS]   = useState("");
  const [filterRoleType, setFilterRoleType] = useState("");
  const [hoveredRow, setHoveredRow] = useState(null);
  const [vacancies, setVacancies]   = useState([]);
  const [vacancyText, setVacancyText] = useState("");
  const [matchResults, setMatchResults] = useState([]);
  const [matchLoading, setMatchLoading] = useState(false);
  const [matchError, setMatchError] = useState("");
  const [matchFilterInfo, setMatchFilterInfo] = useState("");
  const [selectedVacancyId, setSelectedVacancyId] = useState(null);
  const fileRef = useRef();

  // Keep candidatesRef in sync
  useEffect(() => { candidatesRef.current = candidates; }, [candidates]);

  // Load from Supabase on mount
  useEffect(() => {
    const load = async () => {
      if (!sb) { setDbReady(true); return; }
      try {
        const [{ data: cd }, { data: vd }] = await Promise.all([
          sb.from("candidates").select("data").order("created_at"),
          sb.from("vacancies").select("data").order("created_at", { ascending:false })
        ]);
        if (cd) setCandidates(cd.map(r=>r.data));
        if (vd) setVacancies(vd.map(r=>r.data));
      } catch(e) { console.error("Load error:", e); }
      setDbReady(true);
    };
    load();
  }, []);

  // Supabase helpers
  const dbUpsert = (table, id, data) => {
    if (!sb) return;
    sb.from(table).upsert({ custom_id:String(id), data }, { onConflict:"custom_id" }).then(({error})=>{ if(error) console.error("Upsert error:",error); });
  };
  const dbDelete = (table, id) => {
    if (!sb) return;
    sb.from(table).delete().eq("custom_id", String(id));
  };
  const dbDeleteAll = (table) => {
    if (!sb) return;
    sb.from(table).delete().not("custom_id","is",null);
  };

  // Candidate operations
  const addCandidate = (r) => {
    const entry = { _id:Date.now(), candidate_id:"", ...r };
    setCandidates(prev => [...prev, entry]);
    dbUpsert("candidates", entry._id, entry);
    return entry;
  };
  const addCandidateSafe = (r) => {
    const entry = { _id:Date.now()+Math.random(), candidate_id:"", ...r };
    setCandidates(prev => [...prev, entry]);
    dbUpsert("candidates", entry._id, entry);
  };
  const updateField = (id, key, val) => {
    setCandidates(prev => {
      const updated = prev.map(c => c._id===id ? {...c,[key]:val} : c);
      const changed = updated.find(c => c._id===id);
      if (changed) dbUpsert("candidates", id, changed);
      return updated;
    });
    if (selectedCandidate?._id===id) setSelectedCandidate(prev=>({...prev,[key]:val}));
  };
  const removeCandidate = (id) => {
    setCandidates(prev=>prev.filter(c=>c._id!==id));
    setSelectedCandidate(null);
    dbDelete("candidates", id);
  };
  const clearAll = () => {
    setCandidates([]); setConfirmClear(false); setSelectedCandidate(null);
    dbDeleteAll("candidates");
  };

  // Vacancy operations
  const saveVacancy = () => {
    if (!vacancyText.trim()) return;
    const title = vacancyText.trim().slice(0,65)+(vacancyText.length>65?"...":"");
    if (selectedVacancyId) {
      const base = vacancies.find(v=>v._id===selectedVacancyId);
      if (!base) return;
      const updated = {...base, description:vacancyText.trim(), matchResults, updatedAt:new Date().toLocaleDateString()};
      setVacancies(prev=>prev.map(v=>v._id===selectedVacancyId?updated:v));
      dbUpsert("vacancies", selectedVacancyId, updated);
    } else {
      const entry = { _id:Date.now(), title, description:vacancyText.trim(), savedAt:new Date().toLocaleDateString(), matchResults };
      setVacancies(prev=>[entry,...prev]);
      setSelectedVacancyId(entry._id);
      dbUpsert("vacancies", entry._id, entry);
    }
  };
  const deleteVacancy = (id) => {
    setVacancies(prev=>prev.filter(v=>v._id!==id));
    if (selectedVacancyId===id) { setSelectedVacancyId(null); setMatchResults([]); }
    dbDelete("vacancies", id);
  };
  const loadVacancy = (v) => { setVacancyText(v.description); setSelectedVacancyId(v._id); setMatchResults(v.matchResults||[]); setMatchError(""); setMatchFilterInfo(""); };

  // File handler
  const handleFile = (e) => {
    const files = Array.from(e.target.files).slice(0,50);
    if (!files.length) return;
    setError(""); setResult(null); setJustAdded(false);
    setBulkResults([]); setBulkDone(false); setBulkCurrent(0);
    if (files.length===1) {
      setFileName(files[0].name); setBulkFiles([]);
      const reader = new FileReader();
      reader.onload = ()=>setPdfBase64(reader.result.split(",")[1]);
      reader.readAsDataURL(files[0]);
    } else { setBulkFiles(files); setPdfBase64(null); setFileName(""); }
  };

  // API helper
  const callAPI = async (body, maxTok=1000) => {
    const res = await fetch("/api/extract", { method:"POST", headers:{ "Content-Type":"application/json" }, body:JSON.stringify({...body, max_tokens:maxTok}) });
    const data = await res.json();
    if (!res.ok||data.error) throw new Error(data.error?.message||JSON.stringify(data));
    return (data.content||[]).map(b=>b.text||"").join("").replace(/```json|```/g,"").trim();
  };

  // Extract single
  const extract = async () => {
    const hasInput = mode==="paste"?cvText.trim():pdfBase64;
    if (!hasInput) { setError(mode==="paste"?"Please paste a CV first.":"Please upload a PDF first."); return; }
    setLoading(true); setError(""); setResult(null); setJustAdded(false);
    try {
      const content = mode==="pdf"
        ? [{ type:"document", source:{ type:"base64", media_type:"application/pdf", data:pdfBase64 }},{ type:"text", text:"Extract all candidate information." }]
        : cvText.trim();
      const raw = await callAPI({ model:"claude-sonnet-4-6", system:SYSTEM_PROMPT, messages:[{ role:"user", content }] });
      const parsed = JSON.parse(raw);
      const entry = addCandidate(parsed);
      setResult({...parsed, _id:entry._id}); setJustAdded(true);
    } catch(err) { setError("Error: "+(err.message||String(err))); }
    setLoading(false);
  };

  // Duplicate checker
  const isDuplicate = (parsed) => {
    const name  = (parsed.full_name||"").toLowerCase().trim();
    const email = (parsed.email||"").toLowerCase().trim();
    if (name  && addedNamesRef.current.has(name))  return true;
    if (email && addedNamesRef.current.has(email)) return true;
    return candidatesRef.current.some(c => {
      if (email && (c.email||"").toLowerCase().trim()===email) return true;
      if (name  && (c.full_name||"").toLowerCase().trim()===name)  return true;
      return false;
    });
  };

  const getETA = (current, total, startTime) => {
    if (!current||!startTime) return "";
    const elapsed = Date.now()-startTime;
    const remaining = (total-current)*(elapsed/current);
    const mins = Math.floor(remaining/60000);
    const secs = Math.floor((remaining%60000)/1000);
    if (mins>0) return `~${mins}m ${secs}s left`;
    return `~${secs}s left`;
  };

  // Core bulk extraction engine
  const extractBulkFiles = async (files) => {
    setLoading(true); setBulkCurrent(0); setBulkResults([]); setBulkDone(false); setBulkPaused(false);
    setBulkStats({ added:0, duplicates:0, failed:0 });
    shouldPauseRef.current = false;
    addedNamesRef.current  = new Set();
    const startTime = Date.now(); setBulkStartTime(startTime);
    const localResults = [];

    for (let i=0; i<files.length; i++) {
      // Pause gate
      while (shouldPauseRef.current) { await new Promise(r=>setTimeout(r,400)); }

      const file = files[i]; setBulkCurrent(i+1);
      try {
        const base64 = await new Promise((res,rej)=>{ const r=new FileReader(); r.onload=()=>res(r.result.split(",")[1]); r.onerror=()=>rej(new Error("Read failed")); r.readAsDataURL(file); });
        const content = [{ type:"document", source:{ type:"base64", media_type:"application/pdf", data:base64 }},{ type:"text", text:"Extract all candidate information." }];
        const raw    = await callAPI({ model:"claude-sonnet-4-6", system:SYSTEM_PROMPT, messages:[{ role:"user", content }] });
        const parsed = JSON.parse(raw);

        if (isDuplicate(parsed)) {
          localResults.push({ file, name:parsed.full_name||file.name.replace(/\.pdf$/i,""), status:"duplicate" });
          setBulkResults([...localResults]);
          setBulkStats(p=>({...p, duplicates:p.duplicates+1}));
        } else {
          addCandidateSafe(parsed);
          const n=(parsed.full_name||"").toLowerCase().trim();
          const e=(parsed.email||"").toLowerCase().trim();
          if(n) addedNamesRef.current.add(n);
          if(e) addedNamesRef.current.add(e);
          localResults.push({ file, name:parsed.full_name||file.name.replace(/\.pdf$/i,""), status:"success" });
          setBulkResults([...localResults]);
          setBulkStats(p=>({...p, added:p.added+1}));
        }
      } catch {
        localResults.push({ file, name:file.name.replace(/\.pdf$/i,""), status:"error" });
        setBulkResults([...localResults]);
        setBulkStats(p=>({...p, failed:p.failed+1}));
      }

      // Rate limit — 1.5s between requests
      if (i<files.length-1) await new Promise(r=>setTimeout(r,1500));
    }

    setFailedFiles(localResults.filter(r=>r.status==="error").map(r=>r.file));
    setBulkDone(true); setLoading(false);
  };

  const extractBulk  = ()  => extractBulkFiles(bulkFiles);
  const retryFailed  = ()  => { const f=[...failedFiles]; setFailedFiles([]); setBulkResults([]); setBulkDone(false); extractBulkFiles(f); };
  const pauseBulk    = ()  => { shouldPauseRef.current=true;  setBulkPaused(true);  };
  const resumeBulk   = ()  => { shouldPauseRef.current=false; setBulkPaused(false); };

  // Match vacancy
  const matchVacancy = async () => {
    if (!vacancyText.trim()) { setMatchError("Please enter a vacancy description."); return; }
    if (!candidates.length) { setMatchError("No candidates in database."); return; }
    setMatchLoading(true); setMatchError(""); setMatchResults([]); setMatchFilterInfo("");
    try {
      const { pool, info } = preFilterBySubject(vacancyText, candidates);
      if (info) setMatchFilterInfo(info);
      const summary = pool.map(c=>({ id:c._id, name:c.full_name||"Unknown", level:c.teaching_level, specialism:c.subject_specialism, curriculum:c.curriculum, qts:c.qts_holder, experience:c.years_experience, qualification:c.teaching_qualification, available:c.available_from, role_type:c.role_type, notes:c.notes }));
      const userMsg = `VACANCY:\n${vacancyText.trim()}\n\nCANDIDATES:\n${JSON.stringify(summary,null,2)}`;
      const raw = await callAPI({ model:"claude-sonnet-4-6", system:MATCH_PROMPT, messages:[{ role:"user", content:userMsg }] }, 3000);
      setMatchResults(JSON.parse(raw));
    } catch(err) { setMatchError("Matching failed: "+(err.message||String(err))); }
    setMatchLoading(false);
  };

  const resetExtract = () => {
    setCvText(""); setPdfBase64(null); setFileName(""); setResult(null); setError(""); setJustAdded(false);
    setBulkFiles([]); setBulkResults([]); setBulkDone(false); setBulkCurrent(0);
    setBulkPaused(false); setBulkStats({ added:0, duplicates:0, failed:0 }); setFailedFiles([]);
    shouldPauseRef.current = false;
    if (fileRef.current) fileRef.current.value="";
  };

  const downloadExcel = () => {
    const headers = ["Candidate ID","Full Name","Nationality","Email","Phone","Location","Degree Subject","Degree University","Teaching Qual","QTS Holder","Additional Certs","Years Exp","Teaching Level","Subject / Specialism","Current/Last School","Current/Last Role","Curriculum","Available From","Notice Period","Role Type","CV Built","CV Code","Source","Status","Placed At","Notes"];
    const rows = candidates.map(c=>[c.candidate_id||"",c.full_name||"",c.nationality||"",c.email||"",c.phone||"",c.location||"",c.degree_subject||"",c.degree_university||"",c.teaching_qualification||"",c.qts_holder||"",c.additional_certs||"",c.years_experience||"",c.teaching_level||"",c.subject_specialism||"",c.current_last_school||"",c.current_last_role||"",c.curriculum||"",c.available_from||"",c.notice_period||"",c.role_type||"","Yes","","","New","",c.notes||""]);
    const ws = XLSX.utils.aoa_to_sheet([headers,...rows]);
    ws["!cols"]=[{wch:15},{wch:22},{wch:14},{wch:28},{wch:16},{wch:16},{wch:18},{wch:22},{wch:18},{wch:11},{wch:20},{wch:10},{wch:16},{wch:22},{wch:24},{wch:22},{wch:18},{wch:15},{wch:14},{wch:14},{wch:10},{wch:12},{wch:15},{wch:13},{wch:22},{wch:45}];
    const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,"Candidates"); XLSX.writeFile(wb,"rTriibe_Candidate_Database.xlsx");
  };

  const isBulkMode   = bulkFiles.length>1;
  const progress     = bulkFiles.length>0?Math.round((bulkCurrent/bulkFiles.length)*100):0;
  const successCount = bulkResults.filter(r=>r.status==="success").length;
  const errorCount   = bulkResults.filter(r=>r.status==="error").length;

  const filteredCandidates = candidates.filter(c=>{
    if (searchName&&!(c.full_name||"").toLowerCase().includes(searchName.toLowerCase())) return false;
    if (filterLevel&&c.teaching_level!==filterLevel) return false;
    if (filterCurriculum&&c.curriculum!==filterCurriculum) return false;
    if (filterQTS&&c.qts_holder!==filterQTS) return false;
    if (filterRoleType&&c.role_type!==filterRoleType) return false;
    return true;
  });

  const hasFilters = searchName||filterLevel||filterCurriculum||filterQTS||filterRoleType;
  const levels     = [...new Set(candidates.map(c=>c.teaching_level).filter(Boolean))].sort();
  const curricula  = [...new Set(candidates.map(c=>c.curriculum).filter(Boolean))].sort();
  const roleTypes  = [...new Set(candidates.map(c=>c.role_type).filter(Boolean))].sort();
  const selStyle   = { fontSize:12, padding:"5px 10px", border:"1px solid #ddd", borderRadius:6, fontFamily:"Arial", background:"#fff", outline:"none", cursor:"pointer" };

  const renderDetailField = (key, label) => {
    const val = selectedCandidate[key]||"";
    const isText   = ["available_from","notice_period"].includes(key);
    const isSelect = key==="role_type";
    return (
      <div key={key} style={{ borderBottom:"1px solid #f0f0f0", padding:"9px 0", display:"flex", gap:10, alignItems:"center" }}>
        <div style={{ width:165, fontSize:11, color:"#999", flexShrink:0 }}>{label}</div>
        {isText ? (
          <input value={val} onChange={e=>updateField(selectedCandidate._id,key,e.target.value)} placeholder="e.g. August 2025, Immediately..." style={iStyle}/>
        ) : isSelect ? (
          <select value={val} onChange={e=>updateField(selectedCandidate._id,key,e.target.value)} style={sStyle}>
            <option value="">Select...</option>
            <option value="Permanent">Permanent</option>
            <option value="Supply">Supply</option>
            <option value="Both">Both</option>
            <option value="Tutor">Tutor</option>
            <option value="LSA">LSA</option>
          </select>
        ) : (
          <div style={{ fontSize:12, color:val?"#1a1a1a":"#ccc", fontWeight:val?"500":"normal" }}>{val||"—"}</div>
        )}
      </div>
    );
  };

  return (
    <div style={{ fontFamily:"Arial,sans-serif", fontSize:13, minHeight:"100vh", background:"#f0f0f0" }}>

      <div style={{ background:RED, padding:"14px 24px" }}>
        <div style={{ color:"#fff", fontWeight:"bold", fontSize:18 }}>rTriibe</div>
        <div style={{ color:"rgba(255,255,255,0.72)", fontSize:11, marginTop:2 }}>CV Extractor · Candidate Database · Vacancy Matcher</div>
      </div>

      <div style={{ display:"flex", borderBottom:"2px solid #e0e0e0", background:"#fff" }}>
        {[["extract","📄 Extract CV"],["database",`📋 Database (${candidates.length})`],["match","🎯 Match Vacancy"]].map(([key,label])=>(
          <button key={key} onClick={()=>{setView(key);setSelectedCandidate(null);}} style={{ padding:"12px 24px", border:"none", background:"none", cursor:"pointer", fontFamily:"Arial", fontSize:13, fontWeight:view===key?"bold":"normal", color:view===key?RED:"#666", borderBottom:view===key?`3px solid ${RED}`:"3px solid transparent", marginBottom:-2 }}>{label}</button>
        ))}
      </div>

      {/* ── EXTRACT ── */}
      {view==="extract" && (
        <div style={{ margin:16, background:"#fff", borderRadius:8, padding:"20px 24px", boxShadow:"0 1px 4px rgba(0,0,0,0.08)" }}>
          {!result ? (
            <>
              <div style={{ display:"flex", gap:8, marginBottom:16 }}>
                {[["paste","Paste CV text"],["pdf","Upload PDF"]].map(([m,label])=>(
                  <button key={m} onClick={()=>{setMode(m);resetExtract();}} style={{ fontSize:12, padding:"5px 14px", borderRadius:6, border:`1px solid ${mode===m?RED:"#ccc"}`, background:mode===m?RED:"#fff", color:mode===m?"#fff":"#555", fontWeight:mode===m?"bold":"normal", cursor:"pointer", fontFamily:"Arial" }}>{label}</button>
                ))}
              </div>
              {mode==="paste"&&<textarea value={cvText} onChange={e=>setCvText(e.target.value)} placeholder="Paste the full CV or any candidate info here..." style={{ width:"100%", height:210, border:"1px solid #ddd", borderRadius:6, padding:12, fontSize:12, fontFamily:"Arial", resize:"vertical", boxSizing:"border-box", outline:"none" }}/>}
              {mode==="pdf"&&(
                <>
                  {!pdfBase64&&!isBulkMode&&(
                    <div onClick={()=>fileRef.current.click()} style={{ border:"2px dashed #ccc", borderRadius:6, padding:"36px 24px", textAlign:"center", cursor:"pointer", background:"#fafafa" }}>
                      <div style={{ fontSize:14, color:"#555", marginBottom:6 }}>Click to upload PDF files</div>
                      <div style={{ fontSize:12, color:"#aaa" }}>Select 1 CV or up to 50 at once</div>
                    </div>
                  )}
                  {pdfBase64&&!isBulkMode&&(
                    <div style={{ border:"1px solid #ddd", borderRadius:6, padding:"14px 16px", background:"#fafafa", display:"flex", alignItems:"center", justifyContent:"space-between" }}>
                      <div style={{ fontSize:13, color:"#555" }}>✓ {fileName}</div>
                      <button onClick={()=>fileRef.current.click()} style={btn("#fff","#555",{border:"1px solid #ccc",padding:"5px 12px",fontWeight:"normal"})}>Change</button>
                    </div>
                  )}
                  {isBulkMode&&(
                    <div>
                      {/* Header */}
                      <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:12 }}>
                        <div style={{ fontSize:13, fontWeight:"bold", color:"#333" }}>
                          {bulkFiles.length} CVs selected
                          <span style={{ fontSize:11, color:"#aaa", fontWeight:"normal", marginLeft:8 }}>(max 50)</span>
                        </div>
                        {!loading&&!bulkPaused&&<button onClick={()=>fileRef.current.click()} style={btn("#fff","#555",{border:"1px solid #ccc",padding:"5px 12px",fontWeight:"normal"})}>Change files</button>}
                      </div>

                      {/* Active progress */}
                      {(loading||bulkPaused)&&(
                        <div style={{ marginBottom:12, padding:"12px 14px", background:"#f9f9f9", borderRadius:8, border:"1px solid #eee" }}>
                          <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:8 }}>
                            <span style={{ fontSize:12, fontWeight:"bold", color:"#333" }}>
                              {bulkPaused?"⏸ Paused":"⟳ Processing"} {bulkCurrent} of {bulkFiles.length}
                            </span>
                            <span style={{ fontSize:11, color:"#888" }}>{getETA(bulkCurrent, bulkFiles.length, bulkStartTime)}</span>
                          </div>
                          <div style={{ background:"#e0e0e0", borderRadius:6, height:8, overflow:"hidden", marginBottom:10 }}>
                            <div style={{ background:bulkPaused?"#f57f17":RED, width:`${progress}%`, height:"100%", borderRadius:6, transition:"width 0.4s" }}/>
                          </div>
                          <div style={{ display:"flex", gap:16, fontSize:11, marginBottom:10 }}>
                            <span style={{ color:"#2e7d32", fontWeight:"bold" }}>✓ Added: {bulkStats.added}</span>
                            <span style={{ color:"#f57f17", fontWeight:"bold" }}>⚠ Duplicate: {bulkStats.duplicates}</span>
                            <span style={{ color:"#c00", fontWeight:"bold" }}>✗ Failed: {bulkStats.failed}</span>
                          </div>
                          <div style={{ display:"flex", gap:8 }}>
                            {!bulkPaused
                              ? <button onClick={pauseBulk} style={btn("#fff8e1","#f57f17",{border:"1px solid #ffe082",padding:"5px 14px",fontWeight:"normal",fontSize:12})}>⏸ Pause</button>
                              : <button onClick={resumeBulk} style={btn("#e8f5e9","#2e7d32",{border:"1px solid #a5d6a7",padding:"5px 14px",fontWeight:"normal",fontSize:12})}>▶ Resume</button>
                            }
                          </div>
                        </div>
                      )}

                      {/* Done summary */}
                      {bulkDone&&(
                        <div style={{ marginBottom:12, padding:"12px 14px", borderRadius:8, background:bulkStats.failed===0?"#e8f5e9":"#fff8e1", border:`1px solid ${bulkStats.failed===0?"#a5d6a7":"#ffe082"}` }}>
                          <div style={{ fontWeight:"bold", fontSize:13, color:bulkStats.failed===0?"#2e7d32":"#f57f17", marginBottom:6 }}>
                            {bulkStats.failed===0?"✓ All done!":"⚠ Completed with some failures"}
                          </div>
                          <div style={{ display:"flex", gap:16, fontSize:12 }}>
                            <span style={{ color:"#2e7d32" }}>✓ Added: <b>{bulkStats.added}</b></span>
                            <span style={{ color:"#f57f17" }}>⚠ Duplicate: <b>{bulkStats.duplicates}</b></span>
                            <span style={{ color:"#c00" }}>✗ Failed: <b>{bulkStats.failed}</b></span>
                          </div>
                          {failedFiles.length>0&&(
                            <button onClick={retryFailed} style={btn("#c00","#fff",{marginTop:10,padding:"6px 14px",fontSize:12})}>
                              🔁 Retry {failedFiles.length} failed CV{failedFiles.length!==1?"s":""}
                            </button>
                          )}
                        </div>
                      )}

                      {/* Scrollable log */}
                      {bulkResults.length>0&&(
                        <div style={{ maxHeight:220, overflowY:"auto", border:"1px solid #eee", borderRadius:6 }}>
                          {bulkResults.map((r,i)=>(
                            <div key={i} style={{ display:"flex", alignItems:"center", padding:"6px 12px", borderBottom:"1px solid #f5f5f5" }}>
                              <div style={{ width:20, marginRight:10, fontSize:13, textAlign:"center" }}>
                                {r.status==="success"&&<span style={{ color:"#2e7d32" }}>✓</span>}
                                {r.status==="duplicate"&&<span style={{ color:"#f57f17" }}>⚠</span>}
                                {r.status==="error"&&<span style={{ color:"#c00" }}>✗</span>}
                              </div>
                              <div style={{ flex:1, fontSize:11, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap",
                                color:r.status==="success"?"#333":r.status==="duplicate"?"#b06000":"#c00" }}>
                                {r.name}
                              </div>
                              <div style={{ fontSize:10, color:"#aaa", marginLeft:8, flexShrink:0 }}>
                                {r.status==="duplicate"?"already in system":r.status==="error"?"failed":""}
                              </div>
                            </div>
                          ))}
                          {/* Pending files */}
                          {loading&&bulkFiles.slice(bulkCurrent).map((file,i)=>(
                            <div key={`p${i}`} style={{ display:"flex", alignItems:"center", padding:"6px 12px", borderBottom:"1px solid #f5f5f5" }}>
                              <div style={{ width:20, marginRight:10, fontSize:13, textAlign:"center" }}>
                                {i===0?<span style={{ color:RED }}>⟳</span>:<span style={{ color:"#ddd" }}>○</span>}
                              </div>
                              <div style={{ flex:1, fontSize:11, color:"#bbb", overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>
                                {file.name.replace(/\.pdf$/i,"")}
                              </div>
                            </div>
                          ))}
                        </div>
                      )}
                    </div>
                  )}
                  <input ref={fileRef} type="file" accept=".pdf" multiple onChange={handleFile} style={{ display:"none" }}/>
                </>
              )}
              {error&&<div style={{ marginTop:10, fontSize:12, color:RED, background:"#fff5f5", border:"1px solid #fcc", padding:"8px 12px", borderRadius:6 }}>{error}</div>}
              {!bulkDone&&!bulkPaused&&(
                <button onClick={isBulkMode?extractBulk:extract} disabled={loading} style={btn(loading?"#aaa":RED,"#fff",{marginTop:14,fontSize:13,padding:"9px 22px",cursor:loading?"not-allowed":"pointer"})}>
                  {loading?(isBulkMode?`Processing ${bulkCurrent} of ${bulkFiles.length}...`:"Extracting..."):(isBulkMode?`Extract All ${bulkFiles.length} CVs`:"Extract candidate data")}
                </button>
              )}
              {bulkDone&&(
                <div style={{ display:"flex", gap:10, marginTop:14, flexWrap:"wrap" }}>
                  <button onClick={()=>setView("database")} style={btn(RED,"#fff",{fontSize:13,padding:"9px 22px"})}>View database ({candidates.length}) →</button>
                  <button onClick={resetExtract} style={btn("#fff","#555",{border:"1px solid #ccc",fontSize:13,padding:"9px 22px"})}>Extract more</button>
                </div>
              )}
            </>
          ) : (
            <>
              {justAdded&&(
                <div style={{ background:"#e8f5e9", border:"1px solid #a5d6a7", borderRadius:6, padding:"10px 16px", marginBottom:16, display:"flex", alignItems:"center", justifyContent:"space-between" }}>
                  <span style={{ fontSize:13, color:"#2e7d32", fontWeight:"bold" }}>✓ Added to database — {candidates.length} total</span>
                  <button onClick={()=>setView("database")} style={btn("#fff",RED,{border:`1px solid ${RED}`,padding:"5px 14px",fontWeight:"normal"})}>View database →</button>
                </div>
              )}
              <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:20, flexWrap:"wrap", gap:10 }}>
                <div style={{ display:"flex", alignItems:"center", gap:10 }}>
                  <div style={{ width:38, height:38, borderRadius:"50%", background:RED, display:"flex", alignItems:"center", justifyContent:"center", color:"#fff", fontWeight:"bold", fontSize:13 }}>{initials(result.full_name)}</div>
                  <div>
                    <div style={{ fontWeight:"bold", fontSize:15, color:"#1a1a1a" }}>{result.full_name||"—"}</div>
                    <div style={{ fontSize:12, color:"#888" }}>{result.subject_specialism||result.teaching_level||""}</div>
                  </div>
                </div>
                <button onClick={resetExtract} style={btn("#fff","#555",{border:"1px solid #ccc"})}>Extract another</button>
              </div>
              <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:"0 32px" }}>
                {DETAIL_FIELDS.map(([key,label])=>(
                  <div key={key} style={{ borderBottom:"1px solid #f0f0f0", padding:"8px 0", display:"flex", gap:10 }}>
                    <div style={{ width:160, fontSize:11, color:"#999", flexShrink:0 }}>{label}</div>
                    <div style={{ fontSize:12, color:result[key]?"#1a1a1a":"#ccc", fontWeight:result[key]?"500":"normal" }}>{result[key]||"—"}</div>
                  </div>
                ))}
              </div>
              {result.notes&&<div style={{ marginTop:16, padding:"10px 14px", borderLeft:`3px solid ${RED}`, background:"#fff5f5" }}><div style={{ fontSize:10, color:"#999", marginBottom:4, textTransform:"uppercase", letterSpacing:"0.5px" }}>Notes / flags</div><div style={{ fontSize:12, color:"#a00" }}>{result.notes}</div></div>}
            </>
          )}
        </div>
      )}

      {/* ── DATABASE ── */}
      {view==="database" && (
        <div style={{ padding:16 }}>
          {!dbReady ? (
            <div style={{ background:"#fff", borderRadius:8, padding:"60px 24px", textAlign:"center", boxShadow:"0 1px 4px rgba(0,0,0,0.08)" }}>
              <div style={{ fontSize:14, color:"#888" }}>Loading database...</div>
            </div>
          ) : candidates.length===0 ? (
            <div style={{ background:"#fff", borderRadius:8, padding:"60px 24px", textAlign:"center", boxShadow:"0 1px 4px rgba(0,0,0,0.08)" }}>
              <div style={{ fontSize:40, marginBottom:12 }}>📋</div>
              <div style={{ fontSize:14, color:"#888", marginBottom:4 }}>No candidates yet</div>
              <button onClick={()=>setView("extract")} style={btn(RED,"#fff",{marginTop:16,fontSize:13,padding:"9px 20px"})}>Go to extractor</button>
            </div>
          ) : selectedCandidate ? (
            <div style={{ background:"#fff", borderRadius:8, padding:"20px 24px", boxShadow:"0 1px 4px rgba(0,0,0,0.08)" }}>
              <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:20, flexWrap:"wrap", gap:10 }}>
                <button onClick={()=>setSelectedCandidate(null)} style={btn("#fff","#555",{border:"1px solid #ddd",padding:"6px 14px",fontWeight:"normal",fontSize:12})}>← Back</button>
                <button onClick={()=>removeCandidate(selectedCandidate._id)} style={btn("#fff","#c00",{border:"1px solid #fcc",padding:"6px 14px",fontWeight:"normal",fontSize:12})}>Remove candidate</button>
              </div>
              <div style={{ display:"flex", alignItems:"center", gap:12, marginBottom:20 }}>
                <div style={{ width:48, height:48, borderRadius:"50%", background:RED, display:"flex", alignItems:"center", justifyContent:"center", color:"#fff", fontWeight:"bold", fontSize:16 }}>{initials(selectedCandidate.full_name)}</div>
                <div>
                  <div style={{ fontWeight:"bold", fontSize:17, color:"#1a1a1a" }}>{selectedCandidate.full_name||"—"}</div>
                  <div style={{ fontSize:12, color:"#888", marginTop:2 }}>{selectedCandidate.subject_specialism||selectedCandidate.teaching_level||""}</div>
                </div>
              </div>
              <div style={{ marginBottom:16, padding:"10px 14px", background:"#f9f9f9", borderRadius:6, display:"flex", alignItems:"center", gap:12 }}>
                <div style={{ fontSize:11, color:"#999", width:130, flexShrink:0 }}>Candidate ID (rTR)</div>
                <input value={selectedCandidate.candidate_id||""} onChange={e=>updateField(selectedCandidate._id,"candidate_id",e.target.value)} placeholder="e.g. rTRJD01" style={iStyle}/>
              </div>
              <div style={{ background:"#fffbf0", border:"1px solid #ffe082", borderRadius:6, padding:"8px 14px", marginBottom:16, fontSize:11, color:"#b07d00" }}>✏️ Availability fields are editable below</div>
              <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:"0 32px" }}>
                {DETAIL_FIELDS.map(([key,label])=>renderDetailField(key,label))}
              </div>
              {selectedCandidate.notes&&<div style={{ marginTop:16, padding:"12px 16px", borderLeft:`3px solid ${RED}`, background:"#fff5f5" }}><div style={{ fontSize:10, color:"#999", marginBottom:6, textTransform:"uppercase", letterSpacing:"0.5px" }}>Notes / flags</div><div style={{ fontSize:12, color:"#a00", lineHeight:1.6 }}>{selectedCandidate.notes}</div></div>}
            </div>
          ) : (
            <div style={{ background:"#fff", borderRadius:8, padding:"20px 24px", boxShadow:"0 1px 4px rgba(0,0,0,0.08)" }}>
              <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:14, flexWrap:"wrap", gap:10 }}>
                <div style={{ fontSize:13, color:"#555" }}>{hasFilters?`${filteredCandidates.length} of ${candidates.length} candidates`:`${candidates.length} candidates`} · click a row to view & edit</div>
                <div style={{ display:"flex", gap:8, alignItems:"center", flexWrap:"wrap" }}>
                  <button onClick={downloadExcel} style={btn("#1a6e3c","#fff",{padding:"6px 16px"})}>↓ Download Excel</button>
                  {!confirmClear
                    ? <button onClick={()=>setConfirmClear(true)} style={btn("#fff","#c00",{border:"1px solid #fcc",padding:"6px 12px"})}>Clear all</button>
                    : <><span style={{ fontSize:12, color:"#c00" }}>Sure?</span>
                        <button onClick={clearAll} style={btn("#c00","#fff",{padding:"5px 10px"})}>Yes</button>
                        <button onClick={()=>setConfirmClear(false)} style={btn("#fff","#555",{border:"1px solid #ccc",padding:"5px 10px"})}>Cancel</button></>
                  }
                </div>
              </div>
              <div style={{ display:"flex", gap:8, marginBottom:14, flexWrap:"wrap", alignItems:"center" }}>
                <input value={searchName} onChange={e=>setSearchName(e.target.value)} placeholder="🔍 Search by name..." style={{ ...selStyle, padding:"6px 12px", width:190 }}/>
                <select value={filterLevel} onChange={e=>setFilterLevel(e.target.value)} style={selStyle}><option value="">All Levels</option>{levels.map(l=><option key={l} value={l}>{l}</option>)}</select>
                <select value={filterCurriculum} onChange={e=>setFilterCurriculum(e.target.value)} style={selStyle}><option value="">All Curricula</option>{curricula.map(c=><option key={c} value={c}>{c}</option>)}</select>
                <select value={filterQTS} onChange={e=>setFilterQTS(e.target.value)} style={selStyle}><option value="">QTS — All</option><option value="Yes">Yes</option><option value="No">No</option><option value="In Progress">In Progress</option></select>
                <select value={filterRoleType} onChange={e=>setFilterRoleType(e.target.value)} style={selStyle}><option value="">All Role Types</option>{roleTypes.map(r=><option key={r} value={r}>{r}</option>)}</select>
                {hasFilters&&<button onClick={()=>{setSearchName("");setFilterLevel("");setFilterCurriculum("");setFilterQTS("");setFilterRoleType("");}} style={btn("#fff","#c00",{border:"1px solid #fcc",padding:"5px 12px",fontWeight:"normal",fontSize:12})}>✕ Clear</button>}
              </div>
              <div style={{ overflowX:"auto", border:"1px solid #e0e0e0", borderRadius:6 }}>
                <table style={{ borderCollapse:"collapse", fontSize:11, width:TABLE_W, tableLayout:"fixed" }}>
                  <colgroup>{ALL_COLS.map(col=><col key={col.key} width={col.width}/>)}<col width={38}/></colgroup>
                  <thead>
                    <tr>{CATEGORIES.map(cat=><th key={cat.label} colSpan={cat.cols.length} style={{ background:cat.color, color:"#fff", padding:"7px 10px", textAlign:"left", fontWeight:"bold", fontSize:10, borderRight:"2px solid rgba(255,255,255,0.3)", whiteSpace:"nowrap" }}>{cat.icon} {cat.label}</th>)}<th style={{ background:"#263238" }}></th></tr>
                    <tr style={{ background:"#f5f5f5" }}>{ALL_COLS.map(col=><th key={col.key} style={{ padding:"6px 8px", textAlign:"left", fontSize:10, fontWeight:"bold", color:"#555", borderBottom:"2px solid #ddd", overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{col.label}</th>)}<th style={{ borderBottom:"2px solid #ddd" }}></th></tr>
                  </thead>
                  <tbody>
                    {filteredCandidates.length===0 ? (
                      <tr><td colSpan={ALL_COLS.length+1} style={{ padding:"24px", textAlign:"center", color:"#aaa", fontSize:13 }}>No candidates match your filters</td></tr>
                    ) : filteredCandidates.map((c,i)=>(
                      <tr key={c._id} onClick={()=>setSelectedCandidate(c)}
                        style={{ background:hoveredRow===c._id?"#fff3f3":i%2===0?"#fafafa":"#fff", borderBottom:"1px solid #eee", cursor:"pointer" }}
                        onMouseEnter={()=>setHoveredRow(c._id)} onMouseLeave={()=>setHoveredRow(null)}>
                        {ALL_COLS.map(col=>(
                          <td key={col.key} style={{ padding:"5px 8px", overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>
                            <Cell col={col} c={c} onUpdate={updateField}/>
                          </td>
                        ))}
                        <td style={{ padding:"4px 6px", textAlign:"center" }}>
                          <button onClick={e=>{e.stopPropagation();removeCandidate(c._id);}} title="Remove" style={{ background:"none", border:"none", color:"#ddd", cursor:"pointer", fontSize:15, lineHeight:1, padding:"0 2px" }} onMouseEnter={e=>e.target.style.color="#c00"} onMouseLeave={e=>e.target.style.color="#ddd"}>✕</button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <div style={{ marginTop:10, fontSize:11, color:"#aaa" }}>Scroll right to see all fields. Click any row to view & edit.</div>
            </div>
          )}
        </div>
      )}

      {/* ── MATCH ── */}
      {view==="match" && (
        <div style={{ padding:16, display:"grid", gridTemplateColumns:"240px 1fr", gap:16, alignItems:"start" }}>
          <div style={{ background:"#fff", borderRadius:8, padding:16, boxShadow:"0 1px 4px rgba(0,0,0,0.08)" }}>
            <div style={{ fontWeight:"bold", fontSize:13, color:"#333", marginBottom:12 }}>Saved Vacancies</div>
            {vacancies.length===0 ? (
              <div style={{ fontSize:12, color:"#aaa", textAlign:"center", padding:"20px 0" }}>No saved vacancies yet</div>
            ) : (
              <div style={{ display:"flex", flexDirection:"column", gap:6 }}>
                {vacancies.map(v=>(
                  <div key={v._id} onClick={()=>loadVacancy(v)} style={{ padding:"8px 10px", borderRadius:6, cursor:"pointer", fontSize:11, background:selectedVacancyId===v._id?"#fff3f3":"#f9f9f9", border:`1px solid ${selectedVacancyId===v._id?"#f5c5c5":"#eee"}`, display:"flex", justifyContent:"space-between", alignItems:"flex-start", gap:6 }}>
                    <div style={{ flex:1, minWidth:0 }}>
                      <div style={{ color:"#333", fontWeight:selectedVacancyId===v._id?"bold":"normal", overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{v.title}</div>
                      <div style={{ display:"flex", gap:8, marginTop:3, alignItems:"center" }}>
                        <span style={{ color:"#aaa", fontSize:10 }}>{v.savedAt||v.updatedAt}</span>
                        {v.matchResults&&v.matchResults.length>0&&<span style={{ background:"#e8f5e9", color:"#2e7d32", fontSize:10, padding:"1px 6px", borderRadius:8, fontWeight:"bold" }}>{v.matchResults.length} matched</span>}
                      </div>
                    </div>
                    <button onClick={e=>{e.stopPropagation();deleteVacancy(v._id);}} style={{ background:"none", border:"none", color:"#ddd", cursor:"pointer", fontSize:14, padding:0 }} onMouseEnter={e=>e.target.style.color="#c00"} onMouseLeave={e=>e.target.style.color="#ddd"}>✕</button>
                  </div>
                ))}
              </div>
            )}
          </div>

          <div style={{ background:"#fff", borderRadius:8, padding:"20px 24px", boxShadow:"0 1px 4px rgba(0,0,0,0.08)" }}>
            {candidates.length===0 ? (
              <div style={{ textAlign:"center", padding:"40px 0" }}>
                <div style={{ fontSize:32, marginBottom:10 }}>📋</div>
                <div style={{ fontSize:14, color:"#888", marginBottom:4 }}>No candidates in database</div>
                <button onClick={()=>setView("extract")} style={btn(RED,"#fff",{marginTop:16,fontSize:13,padding:"9px 20px"})}>Go to extractor</button>
              </div>
            ) : (
              <>
                <div style={{ marginBottom:16 }}>
                  <div style={{ fontWeight:"bold", fontSize:14, color:"#333", marginBottom:8 }}>Vacancy Description</div>
                  <textarea value={vacancyText} onChange={e=>setVacancyText(e.target.value)}
                    placeholder="Paste the full vacancy — job title, school, subject, level, curriculum, requirements..."
                    style={{ width:"100%", height:150, border:"1px solid #ddd", borderRadius:6, padding:12, fontSize:12, fontFamily:"Arial", resize:"vertical", boxSizing:"border-box", outline:"none" }}/>
                </div>
                <div style={{ display:"flex", gap:10, marginBottom:16, flexWrap:"wrap", alignItems:"center" }}>
                  {selectedVacancyId ? (
                    <button onClick={async()=>{await matchVacancy();}} disabled={matchLoading} style={btn(matchLoading?"#aaa":"#1565C0","#fff",{fontSize:13,padding:"9px 22px",cursor:matchLoading?"not-allowed":"pointer"})}>
                      {matchLoading?"Rematching...":"🔄 Rematch & Update"}
                    </button>
                  ) : (
                    <button onClick={matchVacancy} disabled={matchLoading} style={btn(matchLoading?"#aaa":RED,"#fff",{fontSize:13,padding:"9px 22px",cursor:matchLoading?"not-allowed":"pointer"})}>
                      {matchLoading?"Matching...":`🎯 Match with ${candidates.length} candidates`}
                    </button>
                  )}
                  {!selectedVacancyId&&matchResults.length>0&&(
                    <button onClick={saveVacancy} style={btn(RED,"#fff",{fontSize:12,padding:"9px 16px"})}>💾 Save vacancy with results</button>
                  )}
                  {selectedVacancyId&&matchResults.length>0&&(
                    <button onClick={saveVacancy} style={btn("#fff","#555",{border:"1px solid #ccc",fontSize:12,padding:"9px 16px",fontWeight:"normal"})}>💾 Update saved results</button>
                  )}
                  {!selectedVacancyId&&!matchResults.length&&vacancyText&&(
                    <button onClick={saveVacancy} style={btn("#fff","#555",{border:"1px solid #ccc",fontSize:12,padding:"9px 16px",fontWeight:"normal"})}>Save vacancy</button>
                  )}
                  {vacancyText&&<button onClick={()=>{setVacancyText("");setMatchResults([]);setSelectedVacancyId(null);setMatchFilterInfo("");}} style={btn("#fff","#aaa",{border:"1px solid #eee",fontSize:12,padding:"9px 14px",fontWeight:"normal"})}>New vacancy</button>}
                </div>
                {matchError&&<div style={{ marginBottom:14, fontSize:12, color:RED, background:"#fff5f5", border:"1px solid #fcc", padding:"8px 12px", borderRadius:6 }}>{matchError}</div>}
                {matchFilterInfo&&!matchLoading&&<div style={{ marginBottom:14, fontSize:12, color:"#1565C0", background:"#e3f2fd", border:"1px solid #90caf9", padding:"8px 12px", borderRadius:6 }}>🎯 {matchFilterInfo}</div>}
                {matchResults.length>0&&(
                  <div>
                    <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:12, flexWrap:"wrap", gap:8 }}>
                      <div style={{ fontWeight:"bold", fontSize:13, color:"#333" }}>Results — {matchResults.length} candidates ranked</div>
                      {selectedVacancyId&&<span style={{ fontSize:11, color:"#888", fontStyle:"italic" }}>Saved · click Rematch to refresh</span>}
                    </div>
                    <div style={{ display:"flex", flexDirection:"column", gap:10 }}>
                      {matchResults.map((r,i)=>{
                        const ss=scoreStyle(r.score);
                        const cand=candidates.find(c=>String(c._id)===String(r.id));
                        return (
                          <div key={i} style={{ border:`1px solid ${ss.border}`, borderRadius:8, padding:"12px 16px", background:ss.bg }}>
                            <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:6, flexWrap:"wrap", gap:8 }}>
                              <div style={{ display:"flex", alignItems:"center", gap:10 }}>
                                <div style={{ width:32, height:32, borderRadius:"50%", background:ss.color, display:"flex", alignItems:"center", justifyContent:"center", color:"#fff", fontWeight:"bold", fontSize:12 }}>{i+1}</div>
                                <div>
                                  <div style={{ fontWeight:"bold", fontSize:13, color:"#1a1a1a" }}>{r.name}</div>
                                  {cand&&<div style={{ fontSize:11, color:"#777" }}>{cand.subject_specialism||""}{cand.teaching_level?` · ${cand.teaching_level}`:""}</div>}
                                </div>
                              </div>
                              <div style={{ display:"flex", alignItems:"center", gap:8 }}>
                                <span style={{ fontSize:24, fontWeight:"bold", color:ss.color }}>{r.score}</span>
                                <span style={{ fontSize:11, fontWeight:"bold", color:ss.color, padding:"3px 10px", borderRadius:12, border:`1px solid ${ss.border}`, background:"#fff" }}>{r.label}</span>
                              </div>
                            </div>
                            <div style={{ fontSize:12, color:"#555", lineHeight:1.5, paddingLeft:42 }}>{r.reason}</div>
                            {cand&&<button onClick={()=>{setView("database");setSelectedCandidate(cand);}} style={{ marginTop:8, marginLeft:42, background:"none", border:"none", color:ss.color, fontSize:11, cursor:"pointer", fontFamily:"Arial", padding:0, textDecoration:"underline" }}>View full profile →</button>}
                          </div>
                        );
                      })}
                    </div>
                  </div>
                )}
              </>
            )}
          </div>
        </div>
      )}
    </div>
  );
}
