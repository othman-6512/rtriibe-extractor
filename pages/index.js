import React, { useState, useRef, useEffect } from "react";
import * as XLSX from "xlsx";

const RED = "#C00000";
const STORAGE_KEY = "rtriibe_candidates_v1";

const SYSTEM_PROMPT = `You are an expert education recruitment assistant for rTriibe, a UK-founded agency placing teachers in UAE international schools.
Extract structured candidate information from the provided CV/resume.
Return ONLY a valid JSON object with exactly these keys (use "" for any field not found):
full_name, nationality, email, phone, location, degree_subject, degree_university,
teaching_qualification, qts_holder, additional_certs, years_experience, teaching_level,
subject_specialism, current_last_school, current_last_role, curriculum,
available_from, notice_period, role_type, notes.
Values: qts_holder is "Yes"/"No"/"In Progress"; teaching_level is "EYFS"/"Primary"/"Secondary"/"All Through"/"SEN"/"LSA"/"Leadership"; curriculum is "British"/"IB"/"American"/"CBSE"/"MOE"/"French"/"Multiple"/"Unknown"; role_type is "Permanent"/"Supply"/"Both"/"Tutor"/"LSA"; notes covers visa/work permit, gaps, PGCEi vs QTS, strengths, concerns.
Return ONLY the JSON object. No markdown, no backticks, no other text.`;

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
    { key:"email", label:"Email", width:180 },
    { key:"phone", label:"Phone", width:110 },
    { key:"location", label:"Location", width:120 },
  ]},
  { label:"QUALIFICATIONS", icon:"🎓", color:"#4A148C", cols:[
    { key:"degree_subject", label:"Degree Subject", width:130 },
    { key:"degree_university", label:"University", width:145 },
    { key:"teaching_qualification", label:"Teaching Qual", width:110 },
    { key:"qts_holder", label:"QTS", width:72 },
    { key:"additional_certs", label:"Add. Certs", width:120 },
  ]},
  { label:"EXPERIENCE", icon:"📚", color:"#1B5E20", cols:[
    { key:"years_experience", label:"Yrs", width:50 },
    { key:"teaching_level", label:"Level", width:100 },
    { key:"subject_specialism", label:"Specialism", width:130 },
    { key:"current_last_school", label:"Last School", width:155 },
    { key:"current_last_role", label:"Last Role", width:135 },
  ]},
  { label:"CURRICULUM", icon:"🌍", color:"#006064", cols:[
    { key:"curriculum", label:"Curriculum", width:100 },
  ]},
  { label:"AVAILABILITY", icon:"📅", color:"#BF360C", cols:[
    { key:"available_from", label:"Available From", width:105 },
    { key:"notice_period", label:"Notice Period", width:95 },
    { key:"role_type", label:"Role Type", width:100 },
  ]},
  { label:"NOTES", icon:"📝", color:"#880E4F", cols:[
    { key:"notes", label:"Notes / Flags", width:210 },
  ]},
];

const ALL_COLS = CATEGORIES.flatMap(cat => cat.cols);
const TABLE_W  = ALL_COLS.reduce((s, c) => s + c.width, 0) + 38;

const btn = (bg, color, extra={}) => ({
  border:"none", borderRadius:6, padding:"8px 18px", fontWeight:"bold",
  fontSize:12, cursor:"pointer", fontFamily:"Arial", background:bg, color, ...extra,
});

const initials = (name) => (name||"?").split(" ").map(n=>n[0]).slice(0,2).join("");

const QTSBadge = ({ val }) => (
  <span style={{
    padding:"1px 7px", borderRadius:8, fontSize:10, fontWeight:"bold",
    background:val==="Yes"?"#e8f5e9":val==="In Progress"?"#fff8e1":"#f5f5f5",
    color:val==="Yes"?"#2e7d32":val==="In Progress"?"#e65100":"#999",
  }}>{val||"—"}</span>
);

const Cell = ({ col, c, onUpdate }) => {
  const val = c[col.key] || "";
  if (col.editable) return (
    <input value={val} onChange={e => { e.stopPropagation(); onUpdate(c._id, col.key, e.target.value); }}
      onClick={e => e.stopPropagation()}
      placeholder="rTR..."
      style={{ width:"100%", border:"1px solid #ddd", borderRadius:4, padding:"2px 5px", fontSize:10, fontFamily:"Arial", outline:"none", boxSizing:"border-box" }} />
  );
  if (col.key==="qts_holder") return <QTSBadge val={val} />;
  if (col.key==="notes") return <span style={{ color:val?"#b71c1c":"#ddd", fontSize:10 }} title={val}>{val?(val.length>60?val.slice(0,60)+"…":val):"—"}</span>;
  return <span style={{ color:val?"#333":"#ccc" }} title={val}>{val||"—"}</span>;
};

export default function App() {
  const [view, setView]             = useState("extract");
  const [candidates, setCandidates] = useState([]);
  const [selectedCandidate, setSelectedCandidate] = useState(null);
  const [mode, setMode]             = useState("paste");
  const [cvText, setCvText]         = useState("");
  const [pdfBase64, setPdfBase64]   = useState(null);
  const [fileName, setFileName]     = useState("");
  const [loading, setLoading]       = useState(false);
  const [result, setResult]         = useState(null);
  const [error, setError]           = useState("");
  const [debug, setDebug]           = useState("");
  const [justAdded, setJustAdded]   = useState(false);
  const [confirmClear, setConfirmClear] = useState(false);
  const [bulkFiles, setBulkFiles]   = useState([]);
  const [bulkCurrent, setBulkCurrent] = useState(0);
  const [bulkResults, setBulkResults] = useState([]);
  const [bulkDone, setBulkDone]     = useState(false);
  const fileRef = useRef();

  useEffect(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY);
      if (saved) setCandidates(JSON.parse(saved));
    } catch {}
  }, []);

  const persistList = (list) => { try { localStorage.setItem(STORAGE_KEY, JSON.stringify(list)); } catch {} };
  const persist         = (list)         => { setCandidates(list); persistList(list); };
  const removeCandidate = (id)           => { persist(candidates.filter(c => c._id !== id)); setSelectedCandidate(null); };
  const updateField     = (id, key, val) => {
    const updated = candidates.map(c => c._id===id ? {...c,[key]:val} : c);
    persist(updated);
    if (selectedCandidate?._id === id) setSelectedCandidate(prev => ({...prev, [key]:val}));
  };
  const clearAll = () => { persist([]); setConfirmClear(false); setSelectedCandidate(null); };

  const addCandidateSafe = (r) => {
    const entry = { _id: Date.now() + Math.random(), candidate_id: "", ...r };
    setCandidates(prev => {
      const newList = [...prev, entry];
      persistList(newList);
      return newList;
    });
  };

  const addCandidate = (r) => persist([...candidates, { _id: Date.now(), candidate_id: "", ...r }]);

  const handleFile = (e) => {
    const files = Array.from(e.target.files).slice(0, 50);
    if (files.length === 0) return;
    setError(""); setDebug(""); setResult(null); setJustAdded(false);
    setBulkResults([]); setBulkDone(false); setBulkCurrent(0);
    if (files.length === 1) {
      setFileName(files[0].name); setBulkFiles([]);
      const reader = new FileReader();
      reader.onload = () => setPdfBase64(reader.result.split(",")[1]);
      reader.readAsDataURL(files[0]);
    } else {
      setBulkFiles(files); setPdfBase64(null); setFileName("");
    }
  };

  const extract = async () => {
    const hasInput = mode==="paste" ? cvText.trim() : pdfBase64;
    if (!hasInput) { setError(mode==="paste" ? "Please paste a CV first." : "Please upload a PDF first."); return; }
    setLoading(true); setError(""); setDebug(""); setResult(null); setJustAdded(false);
    try {
      const content = mode==="pdf"
        ? [{ type:"document", source:{ type:"base64", media_type:"application/pdf", data:pdfBase64 }},{ type:"text", text:"Extract all candidate information." }]
        : cvText.trim();
      const res = await fetch("/api/extract", {
        method:"POST", headers:{ "Content-Type":"application/json" },
        body:JSON.stringify({ model:"claude-sonnet-4-6", max_tokens:1000, system:SYSTEM_PROMPT, messages:[{ role:"user", content }]}),
      });
      const data = await res.json();
      if (!res.ok || data.error) { setError("API error — see debug."); setDebug(data.error?.message||JSON.stringify(data)); setLoading(false); return; }
      const raw   = (data.content||[]).map(b=>b.text||"").join("").trim();
      const clean = raw.replace(/```json|```/g,"").trim();
      try {
        const parsed = JSON.parse(clean);
        setResult(parsed); addCandidate(parsed); setJustAdded(true);
      } catch { setError("Could not parse the response."); setDebug(raw.slice(0,400)); }
    } catch(err) { setError("Network error."); setDebug(err.message||String(err)); }
    setLoading(false);
  };

  const extractBulk = async () => {
    setLoading(true); setBulkCurrent(0); setBulkResults([]); setBulkDone(false);
    for (let i = 0; i < bulkFiles.length; i++) {
      const file = bulkFiles[i];
      setBulkCurrent(i + 1);
      try {
        const base64 = await new Promise((resolve, reject) => {
          const reader = new FileReader();
          reader.onload = () => resolve(reader.result.split(",")[1]);
          reader.onerror = () => reject(new Error("Failed to read file"));
          reader.readAsDataURL(file);
        });
        const content = [
          { type:"document", source:{ type:"base64", media_type:"application/pdf", data:base64 }},
          { type:"text", text:"Extract all candidate information." }
        ];
        const res = await fetch("/api/extract", {
          method:"POST", headers:{ "Content-Type":"application/json" },
          body:JSON.stringify({ model:"claude-sonnet-4-6", max_tokens:1000, system:SYSTEM_PROMPT, messages:[{ role:"user", content }]}),
        });
        const data = await res.json();
        if (!res.ok || data.error) throw new Error(data.error?.message || "API error");
        const raw   = (data.content||[]).map(b=>b.text||"").join("").trim();
        const clean = raw.replace(/```json|```/g,"").trim();
        const parsed = JSON.parse(clean);
        addCandidateSafe(parsed);
        setBulkResults(prev => [...prev, { name:file.name, status:"success", fullName:parsed.full_name||"" }]);
      } catch(err) {
        setBulkResults(prev => [...prev, { name:file.name, status:"error", error:err.message }]);
      }
    }
    setBulkDone(true);
    setLoading(false);
  };

  const resetExtract = () => {
    setCvText(""); setPdfBase64(null); setFileName(""); setResult(null);
    setError(""); setDebug(""); setJustAdded(false);
    setBulkFiles([]); setBulkResults([]); setBulkDone(false); setBulkCurrent(0);
    if (fileRef.current) fileRef.current.value = "";
  };

  const downloadExcel = () => {
    const headers = ["Candidate ID","Full Name","Nationality","Email","Phone","Location",
      "Degree Subject","Degree University","Teaching Qual","QTS Holder","Additional Certs",
      "Years Exp","Teaching Level","Subject / Specialism","Current/Last School","Current/Last Role",
      "Curriculum","Available From","Notice Period","Role Type","CV Built","CV Code","Source","Status","Placed At","Notes"];
    const rows = candidates.map(c => [
      c.candidate_id||"",c.full_name||"",c.nationality||"",c.email||"",c.phone||"",c.location||"",
      c.degree_subject||"",c.degree_university||"",c.teaching_qualification||"",c.qts_holder||"",
      c.additional_certs||"",c.years_experience||"",c.teaching_level||"",c.subject_specialism||"",
      c.current_last_school||"",c.current_last_role||"",c.curriculum||"",
      c.available_from||"",c.notice_period||"",c.role_type||"",
      "Yes","","","New","",c.notes||"",
    ]);
    const ws = XLSX.utils.aoa_to_sheet([headers,...rows]);
    ws["!cols"] = [{wch:15},{wch:22},{wch:14},{wch:28},{wch:16},{wch:16},{wch:18},{wch:22},
      {wch:18},{wch:11},{wch:20},{wch:10},{wch:16},{wch:22},{wch:24},{wch:22},
      {wch:18},{wch:15},{wch:14},{wch:14},{wch:10},{wch:12},{wch:15},{wch:13},{wch:22},{wch:45}];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Candidates");
    XLSX.writeFile(wb, "rTriibe_Candidate_Database.xlsx");
  };

  const isBulkMode   = bulkFiles.length > 1;
  const progress     = bulkFiles.length > 0 ? Math.round((bulkCurrent / bulkFiles.length) * 100) : 0;
  const successCount = bulkResults.filter(r => r.status==="success").length;
  const errorCount   = bulkResults.filter(r => r.status==="error").length;

  return (
    <div style={{ fontFamily:"Arial,sans-serif", maxWidth:960, margin:"0 auto", fontSize:13 }}>

      <div style={{ background:RED, padding:"16px 24px" }}>
        <div style={{ color:"#fff", fontWeight:"bold", fontSize:18 }}>rTriibe</div>
        <div style={{ color:"rgba(255,255,255,0.72)", fontSize:11, marginTop:2 }}>CV Extractor · Candidate Database</div>
      </div>

      <div style={{ display:"flex", borderBottom:"2px solid #eee", background:"#fff" }}>
        {[["extract","Extract CV"],["database",`Database (${candidates.length})`]].map(([key,label]) => (
          <button key={key} onClick={() => { setView(key); setSelectedCandidate(null); }} style={{
            padding:"12px 24px", border:"none", background:"none", cursor:"pointer",
            fontFamily:"Arial", fontSize:13, fontWeight:view===key?"bold":"normal",
            color:view===key?RED:"#666",
            borderBottom:view===key?`3px solid ${RED}`:"3px solid transparent", marginBottom:-2,
          }}>{label}</button>
        ))}
      </div>

      {/* ── EXTRACT VIEW ── */}
      {view==="extract" && (
        <div style={{ background:"#fff", padding:"20px 24px" }}>
          {!result ? (
            <>
              <div style={{ display:"flex", gap:8, marginBottom:16 }}>
                {[["paste","Paste CV text"],["pdf","Upload PDF"]].map(([m,label]) => (
                  <button key={m} onClick={() => { setMode(m); resetExtract(); }} style={{
                    fontSize:12, padding:"5px 14px", borderRadius:6,
                    border:`1px solid ${mode===m?RED:"#ccc"}`,
                    background:mode===m?RED:"#fff", color:mode===m?"#fff":"#555",
                    fontWeight:mode===m?"bold":"normal", cursor:"pointer", fontFamily:"Arial",
                  }}>{label}</button>
                ))}
              </div>

              {mode==="paste" && (
                <textarea value={cvText} onChange={e=>setCvText(e.target.value)}
                  placeholder="Paste the full CV or any candidate info here..."
                  style={{ width:"100%", height:210, border:"1px solid #ddd", borderRadius:6, padding:12, fontSize:12, fontFamily:"Arial", resize:"vertical", boxSizing:"border-box", outline:"none" }} />
              )}

              {mode==="pdf" && (
                <>
                  {!pdfBase64 && bulkFiles.length===0 && (
                    <div onClick={() => fileRef.current.click()}
                      style={{ border:"2px dashed #ccc", borderRadius:6, padding:"36px 24px", textAlign:"center", cursor:"pointer", background:"#fafafa" }}>
                      <div style={{ fontSize:14, color:"#555", marginBottom:6 }}>Click to upload PDF files</div>
                      <div style={{ fontSize:12, color:"#aaa" }}>Select 1 CV or up to 50 at once for bulk extraction</div>
                    </div>
                  )}
                  {pdfBase64 && bulkFiles.length===0 && (
                    <div style={{ border:"1px solid #ddd", borderRadius:6, padding:"14px 16px", background:"#fafafa", display:"flex", alignItems:"center", justifyContent:"space-between" }}>
                      <div style={{ fontSize:13, color:"#555" }}>✓ {fileName}</div>
                      <button onClick={() => fileRef.current.click()} style={btn("#fff","#555",{ border:"1px solid #ccc", padding:"5px 12px", fontWeight:"normal" })}>Change</button>
                    </div>
                  )}
                  {isBulkMode && (
                    <div>
                      <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:12 }}>
                        <div style={{ fontSize:13, fontWeight:"bold", color:"#333" }}>
                          {bulkFiles.length} CVs selected
                          <span style={{ fontSize:11, color:"#aaa", fontWeight:"normal", marginLeft:8 }}>(max 50)</span>
                        </div>
                        {!loading && <button onClick={() => fileRef.current.click()} style={btn("#fff","#555",{ border:"1px solid #ccc", padding:"5px 12px", fontWeight:"normal" })}>Change files</button>}
                      </div>
                      {loading && (
                        <div style={{ marginBottom:12 }}>
                          <div style={{ display:"flex", justifyContent:"space-between", fontSize:12, color:"#555", marginBottom:6 }}>
                            <span>Extracting CV {bulkCurrent} of {bulkFiles.length}...</span>
                            <span style={{ fontWeight:"bold" }}>{progress}%</span>
                          </div>
                          <div style={{ background:"#eee", borderRadius:6, height:10, overflow:"hidden" }}>
                            <div style={{ background:RED, width:`${progress}%`, height:"100%", borderRadius:6, transition:"width 0.4s" }} />
                          </div>
                        </div>
                      )}
                      {bulkDone && (
                        <div style={{ marginBottom:12, padding:"10px 14px", borderRadius:6,
                          background:errorCount===0?"#e8f5e9":"#fff8e1",
                          border:`1px solid ${errorCount===0?"#a5d6a7":"#ffe082"}` }}>
                          <div style={{ fontWeight:"bold", fontSize:13, color:errorCount===0?"#2e7d32":"#f57f17" }}>
                            {errorCount===0 ? `✓ All ${successCount} CVs extracted successfully!` : `✓ ${successCount} extracted · ${errorCount} failed`}
                          </div>
                          <div style={{ fontSize:11, color:"#888", marginTop:2 }}>All added to your database automatically</div>
                        </div>
                      )}
                      <div style={{ maxHeight:220, overflowY:"auto", border:"1px solid #eee", borderRadius:6 }}>
                        {bulkFiles.map((file, i) => {
                          const res = bulkResults[i];
                          const isCurrent = loading && bulkCurrent === i + 1;
                          return (
                            <div key={i} style={{ display:"flex", alignItems:"center", padding:"7px 12px", borderBottom:"1px solid #f5f5f5", background:isCurrent?"#fff8f8":"transparent" }}>
                              <div style={{ width:22, marginRight:10, fontSize:14, textAlign:"center" }}>
                                {!res && !isCurrent && <span style={{ color:"#ddd" }}>○</span>}
                                {isCurrent && <span style={{ color:RED }}>⟳</span>}
                                {res?.status==="success" && <span style={{ color:"#2e7d32" }}>✓</span>}
                                {res?.status==="error" && <span style={{ color:"#c00" }}>✗</span>}
                              </div>
                              <div style={{ flex:1, fontSize:11, color:res?.status==="success"?"#333":"#888", overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>
                                {res?.status==="success" && res.fullName ? res.fullName : file.name.replace(".pdf","").replace(".PDF","")}
                              </div>
                              {res?.status==="error" && <div style={{ fontSize:10, color:"#c00", marginLeft:8, flexShrink:0 }}>Failed</div>}
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  )}
                  <input ref={fileRef} type="file" accept=".pdf" multiple onChange={handleFile} style={{ display:"none" }} />
                </>
              )}

              {error && <div style={{ marginTop:10, fontSize:12, color:RED, background:"#fff5f5", border:"1px solid #fcc", padding:"8px 12px", borderRadius:6 }}>{error}</div>}
              {debug && <div style={{ marginTop:6, fontSize:11, color:"#666", fontFamily:"monospace", background:"#f5f5f5", padding:8, borderRadius:6, wordBreak:"break-all" }}>Debug: {debug}</div>}

              {!bulkDone && (
                <button onClick={isBulkMode ? extractBulk : extract} disabled={loading}
                  style={btn(loading?"#aaa":RED,"#fff",{ marginTop:14, fontSize:13, padding:"9px 22px", cursor:loading?"not-allowed":"pointer" })}>
                  {loading
                    ? (isBulkMode ? `Extracting ${bulkCurrent} of ${bulkFiles.length}...` : "Extracting...")
                    : (isBulkMode ? `Extract All ${bulkFiles.length} CVs` : "Extract candidate data")}
                </button>
              )}
              {bulkDone && (
                <div style={{ display:"flex", gap:10, marginTop:14 }}>
                  <button onClick={() => setView("database")} style={btn(RED,"#fff",{ fontSize:13, padding:"9px 22px" })}>View database ({candidates.length}) →</button>
                  <button onClick={resetExtract} style={btn("#fff","#555",{ border:"1px solid #ccc", fontSize:13, padding:"9px 22px" })}>Extract more</button>
                </div>
              )}
            </>
          ) : (
            <>
              {justAdded && (
                <div style={{ background:"#e8f5e9", border:"1px solid #a5d6a7", borderRadius:6, padding:"10px 16px", marginBottom:16, display:"flex", alignItems:"center", justifyContent:"space-between" }}>
                  <span style={{ fontSize:13, color:"#2e7d32", fontWeight:"bold" }}>✓ Added to database — {candidates.length} candidate{candidates.length!==1?"s":""} total</span>
                  <button onClick={() => setView("database")} style={btn("#fff",RED,{ border:`1px solid ${RED}`, padding:"5px 14px", fontWeight:"normal" })}>View database →</button>
                </div>
              )}
              <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:20, flexWrap:"wrap", gap:10 }}>
                <div style={{ display:"flex", alignItems:"center", gap:10 }}>
                  <div style={{ width:38, height:38, borderRadius:"50%", background:RED, display:"flex", alignItems:"center", justifyContent:"center", color:"#fff", fontWeight:"bold", fontSize:13, flexShrink:0 }}>{initials(result.full_name)}</div>
                  <div>
                    <div style={{ fontWeight:"bold", fontSize:15, color:"#1a1a1a" }}>{result.full_name||"—"}</div>
                    <div style={{ fontSize:12, color:"#888" }}>{result.subject_specialism||result.teaching_level||""}</div>
                  </div>
                </div>
                <button onClick={resetExtract} style={btn("#fff","#555",{ border:"1px solid #ccc" })}>Extract another</button>
              </div>
              <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:"0 32px" }}>
                {DETAIL_FIELDS.map(([key,label]) => (
                  <div key={key} style={{ borderBottom:"1px solid #f0f0f0", padding:"8px 0", display:"flex", gap:10 }}>
                    <div style={{ width:160, fontSize:11, color:"#999", flexShrink:0 }}>{label}</div>
                    <div style={{ fontSize:12, color:result[key]?"#1a1a1a":"#ccc", fontWeight:result[key]?"500":"normal" }}>{result[key]||"—"}</div>
                  </div>
                ))}
              </div>
              {result.notes && (
                <div style={{ marginTop:16, padding:"10px 14px", borderLeft:`3px solid ${RED}`, background:"#fff5f5" }}>
                  <div style={{ fontSize:10, color:"#999", marginBottom:4, textTransform:"uppercase", letterSpacing:"0.5px" }}>Notes / flags</div>
                  <div style={{ fontSize:12, color:"#a00" }}>{result.notes}</div>
                </div>
              )}
            </>
          )}
        </div>
      )}

      {/* ── DATABASE VIEW ── */}
      {view==="database" && (
        <div style={{ background:"#fff", padding:"20px 24px" }}>
          {candidates.length===0 ? (
            <div style={{ textAlign:"center", padding:"60px 24px", color:"#aaa" }}>
              <div style={{ fontSize:40, marginBottom:12 }}>📋</div>
              <div style={{ fontSize:14, marginBottom:4, color:"#888" }}>No candidates yet</div>
              <div style={{ fontSize:12 }}>Extract your first CV to get started</div>
              <button onClick={() => setView("extract")} style={btn(RED,"#fff",{ marginTop:16, fontSize:13, padding:"9px 20px" })}>Go to extractor</button>
            </div>
          ) : selectedCandidate ? (
            /* ── CANDIDATE DETAIL VIEW ── */
            <div>
              <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:20, flexWrap:"wrap", gap:10 }}>
                <button onClick={() => setSelectedCandidate(null)}
                  style={btn("#fff","#555",{ border:"1px solid #ddd", padding:"6px 14px", fontWeight:"normal", fontSize:12 })}>
                  ← Back to database
                </button>
                <button onClick={() => removeCandidate(selectedCandidate._id)}
                  style={btn("#fff","#c00",{ border:"1px solid #fcc", padding:"6px 14px", fontWeight:"normal", fontSize:12 })}>
                  Remove candidate
                </button>
              </div>

              <div style={{ display:"flex", alignItems:"center", gap:12, marginBottom:20 }}>
                <div style={{ width:48, height:48, borderRadius:"50%", background:RED, display:"flex", alignItems:"center", justifyContent:"center", color:"#fff", fontWeight:"bold", fontSize:16, flexShrink:0 }}>
                  {initials(selectedCandidate.full_name)}
                </div>
                <div>
                  <div style={{ fontWeight:"bold", fontSize:17, color:"#1a1a1a" }}>{selectedCandidate.full_name||"—"}</div>
                  <div style={{ fontSize:12, color:"#888", marginTop:2 }}>{selectedCandidate.subject_specialism||selectedCandidate.teaching_level||""}</div>
                </div>
              </div>

              <div style={{ marginBottom:16, padding:"10px 14px", background:"#f9f9f9", borderRadius:6, display:"flex", alignItems:"center", gap:12 }}>
                <div style={{ fontSize:11, color:"#999", width:130, flexShrink:0 }}>Candidate ID (rTR)</div>
                <input
                  value={selectedCandidate.candidate_id||""}
                  onChange={e => updateField(selectedCandidate._id, "candidate_id", e.target.value)}
                  placeholder="e.g. rTRJD01"
                  style={{ flex:1, border:"1px solid #ddd", borderRadius:5, padding:"5px 10px", fontSize:12, fontFamily:"Arial", outline:"none" }}
                />
              </div>

              <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:"0 32px" }}>
                {DETAIL_FIELDS.map(([key,label]) => (
                  <div key={key} style={{ borderBottom:"1px solid #f0f0f0", padding:"9px 0", display:"flex", gap:10 }}>
                    <div style={{ width:160, fontSize:11, color:"#999", flexShrink:0, paddingTop:1 }}>{label}</div>
                    <div style={{ fontSize:12, color:selectedCandidate[key]?"#1a1a1a":"#ccc", fontWeight:selectedCandidate[key]?"500":"normal" }}>
                      {selectedCandidate[key]||"—"}
                    </div>
                  </div>
                ))}
              </div>

              {selectedCandidate.notes && (
                <div style={{ marginTop:16, padding:"12px 16px", borderLeft:`3px solid ${RED}`, background:"#fff5f5" }}>
                  <div style={{ fontSize:10, color:"#999", marginBottom:6, textTransform:"uppercase", letterSpacing:"0.5px" }}>Notes / flags</div>
                  <div style={{ fontSize:12, color:"#a00", lineHeight:1.6 }}>{selectedCandidate.notes}</div>
                </div>
              )}
            </div>
          ) : (
            /* ── TABLE VIEW ── */
            <>
              <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:14, flexWrap:"wrap", gap:10 }}>
                <div style={{ fontSize:13, color:"#555" }}>
                  {candidates.length} candidate{candidates.length!==1?"s":""} · click any row to view details
                </div>
                <div style={{ display:"flex", gap:8, alignItems:"center" }}>
                  <button onClick={downloadExcel} style={btn("#1a6e3c","#fff")}>↓ Download Excel</button>
                  {!confirmClear
                    ? <button onClick={() => setConfirmClear(true)} style={btn("#fff","#c00",{ border:"1px solid #fcc" })}>Clear all</button>
                    : <div style={{ display:"flex", gap:6, alignItems:"center" }}>
                        <span style={{ fontSize:12, color:"#c00" }}>Sure?</span>
                        <button onClick={clearAll} style={btn("#c00","#fff",{ padding:"6px 12px" })}>Yes</button>
                        <button onClick={() => setConfirmClear(false)} style={btn("#fff","#555",{ border:"1px solid #ccc", padding:"6px 12px" })}>Cancel</button>
                      </div>
                  }
                </div>
              </div>
              <div style={{ overflowX:"auto", border:"1px solid #e0e0e0", borderRadius:6 }}>
                <table style={{ borderCollapse:"collapse", fontSize:11, width:TABLE_W, tableLayout:"fixed" }}>
                  <colgroup>
                    {ALL_COLS.map(col => <col key={col.key} width={col.width} />)}
                    <col width={38} />
                  </colgroup>
                  <thead>
                    <tr>
                      {CATEGORIES.map(cat => (
                        <th key={cat.label} colSpan={cat.cols.length} style={{
                          background:cat.color, color:"#fff", padding:"7px 10px",
                          textAlign:"left", fontWeight:"bold", fontSize:10, letterSpacing:"0.4px",
                          borderRight:"2px solid rgba(255,255,255,0.3)", whiteSpace:"nowrap",
                        }}>{cat.icon} {cat.label}</th>
                      ))}
                      <th style={{ background:"#263238" }}></th>
                    </tr>
                    <tr style={{ background:"#f5f5f5" }}>
                      {ALL_COLS.map(col => (
                        <th key={col.key} style={{
                          padding:"6px 8px", textAlign:"left", fontSize:10, fontWeight:"bold",
                          color:"#555", borderBottom:"2px solid #ddd", overflow:"hidden",
                          textOverflow:"ellipsis", whiteSpace:"nowrap",
                        }}>{col.label}</th>
                      ))}
                      <th style={{ borderBottom:"2px solid #ddd" }}></th>
                    </tr>
                  </thead>
                  <tbody>
                    {candidates.map((c,i) => (
                      <tr key={c._id}
                        onClick={() => setSelectedCandidate(c)}
                        style={{ background:i%2===0?"#fafafa":"#fff", borderBottom:"1px solid #eee", cursor:"pointer" }}
                        onMouseEnter={e => e.currentTarget.style.background="#fff3f3"}
                        onMouseLeave={e => e.currentTarget.style.background=i%2===0?"#fafafa":"#fff"}>
                        {ALL_COLS.map(col => (
                          <td key={col.key} style={{ padding:"5px 8px", overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>
                            <Cell col={col} c={c} onUpdate={updateField} />
                          </td>
                        ))}
                        <td style={{ padding:"4px 6px", textAlign:"center" }}>
                          <button onClick={e => { e.stopPropagation(); removeCandidate(c._id); }} title="Remove"
                            style={{ background:"none", border:"none", color:"#ddd", cursor:"pointer", fontSize:15, lineHeight:1, padding:"0 2px" }}
                            onMouseEnter={e=>e.target.style.color="#c00"}
                            onMouseLeave={e=>e.target.style.color="#ddd"}>✕</button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <div style={{ marginTop:10, fontSize:11, color:"#aaa" }}>
                Click any row to view full candidate profile. ID column is editable inline.
              </div>
            </>
          )}
        </div>
      )}
    </div>
  );
}
