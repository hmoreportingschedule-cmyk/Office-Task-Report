import React, { useEffect, useMemo, useState } from "react";
import { createRoot } from "react-dom/client";
import {
  Bell, CalendarDays, CheckCircle2, Clock3, FileText, LayoutDashboard, ClipboardCheck,
  LogOut, Menu, ShieldCheck, Users, ClipboardList, Search, Pencil,
  UserCircle2, AlertCircle, Check, X, Plus, RefreshCw, Upload, Download, FileSpreadsheet, Camera, KeyRound, StickyNote
} from "lucide-react";
import { AreaChart, Area, CartesianGrid, XAxis, YAxis, Tooltip, ResponsiveContainer, BarChart, Bar } from "recharts";
import "./styles.css";

// Office Task Report V.74
// IMPORTANT: Browser -> Google Apps Script POST can hang/fail because the Apps
// Script Web App redirects to googleusercontent.com and browser CORS handling
// can block the response. V.30 sends requests through the same-origin Vercel
// serverless proxy instead.
const API_URL = "/api/office-task";

let XLSX_MODULE = null;
async function getXLSX(){
  // Keep the declaration outside the conditional statement so every Vite/esbuild
  // transform target parses this helper consistently.
  if (XLSX_MODULE) return XLSX_MODULE;
  const mod = await import("xlsx");
  XLSX_MODULE = mod.default || mod;
  return XLSX_MODULE;
}

async function api(action, payload = {}) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 60000);
  try {
    const res = await fetch(API_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ action, ...payload }),
      signal: controller.signal,
      cache: "no-store"
    });

    const text = await res.text();
    let data;
    try { data = JSON.parse(text); }
    catch {
      throw new Error(`Backend non-JSON response (${res.status}). Apps Script deployment/access check karein.`);
    }

    if (!res.ok || !data.ok) {
      const extra = data.detail ? ` | ${String(data.detail).slice(0, 300)}` : "";
      throw new Error((data.message || `Backend request failed (${res.status}).`) + extra);
    }
    return data;
  } catch (e) {
    if (e && e.name === "AbortError") {
      throw new Error("Sign in timeout. Vercel proxy ya Apps Script backend response check karein.");
    }
    throw e;
  } finally {
    clearTimeout(timeout);
  }
}

function App() {
  const [session, setSession] = useState(() => {
    try { return JSON.parse(localStorage.getItem("otr_session") || "null"); } catch { return null; }
  });
  const [view, setView] = useState("dashboard");
  const [sidebar, setSidebar] = useState(false);
  const [loading, setLoading] = useState(false);
  const [toast, setToast] = useState(null);

  const notify = (type, message) => {
    setToast({ type, message });
    setTimeout(() => setToast(null), 3500);
  };

  const logout = () => {
    localStorage.removeItem("otr_session");
    setSession(null);
  };
  useEffect(()=>{const open=()=>setView("notifications");window.addEventListener("open-notifications",open);return()=>window.removeEventListener("open-notifications",open)},[]);

  if (!session) {
    return <Login onLogin={(s) => {
      localStorage.setItem("otr_session", JSON.stringify(s));
      setSession(s);
    }} notify={notify} />;
  }

  return (
    <div className="app-shell">
      <Sidebar open={sidebar} setOpen={setSidebar} session={session} view={view} setView={setView} logout={logout} />
      <main className="main">
        <Topbar session={session} onMenu={() => setSidebar(v => !v)} onLogout={logout} />
        <div className="content">
          {view === "dashboard" && <Dashboard session={session} notify={notify} />}
          {view === "attendance" && <Attendance session={session} notify={notify} />}
          {view === "tasks" && <Tasks session={session} notify={notify} />}
          {view === "templates" && <Templates session={session} notify={notify} />}
          {view === "employees" && <Employees session={session} notify={notify} />}
          {view === "approvals" && <Approvals session={session} notify={notify} />}
          {view === "reports" && <Reports session={session} notify={notify} />}
          {view === "notifications" && <Notifications session={session} notify={notify} />}
          {view === "requests" && <RequestsCenter session={session} notify={notify} />}
          {view === "profile" && <Profile session={session} notify={notify} />}
          {view === "notes" && <Notes session={session} notify={notify} />}
        </div>
      </main>
      {toast && <div className={`toast ${toast.type}`}>{toast.type === "success" ? <Check size={17}/> : <AlertCircle size={17}/>} {toast.message}</div>}
    </div>
  );
}

function Login({ onLogin, notify }) {
  const [user, setUser] = useState("");
  const [password, setPassword] = useState("");
  const [busy, setBusy] = useState(false);

  async function submit(e) {
    e.preventDefault();
    setBusy(true);
    try {
      // Login is intentionally a single request. The previous health-before-login
      // call made low-bandwidth/slow Apps Script deployments feel stuck.
      const r = await api("login", { username: user.trim(), password });
      onLogin(r.session);
    } catch (e) {
      notify("error", e.message);
    } finally { setBusy(false); }
  }

  return (
    <div className="login-page">
      <div className="login-orb orb-a"/><div className="login-orb orb-b"/>
      <form className="login-card" onSubmit={submit}>
        <div className="brand-mark"><ClipboardList size={27}/></div>
        <div className="eyebrow">OFFICE MANAGEMENT</div>
        <h1>Office Task Report</h1>
        <p className="muted">Attendance, tasks, approvals & employee progress.</p>
        <label>Username<input value={user} onChange={e=>setUser(e.target.value)} autoComplete="username" required /></label>
        <label>Password<input value={password} onChange={e=>setPassword(e.target.value)} type="password" autoComplete="current-password" required /></label>
        <button className="primary full" disabled={busy}>{busy ? "Signing in…" : "Sign In"}</button>
        <div className="login-note"><ShieldCheck size={15}/> Secure role-based access</div>
      </form>
    </div>
  );
}

function Sidebar({ open, setOpen, session, view, setView, logout }) {
  const items = [
    ["dashboard", "Dashboard", LayoutDashboard],
    ["attendance", "Attendance", Clock3],
    ["tasks", "Check Task", ClipboardList],
    ["templates", "Task Templates", FileText],
    ["employees", "Users", Users],
    ["approvals", "Approvals", CheckCircle2],
    ["reports", "Progress Reports", CalendarDays],
    ["notifications", "Notifications", Bell],
    ["requests", "Requests & Advanced", ClipboardCheck],
    ["profile", "My Profile", UserCircle2],
    ["notes", "Notes", StickyNote],
  ];
  const canAdmin = ["MASTER_ADMIN","ADMIN","HOD"].includes(session.role);
  return <aside className={`sidebar ${open ? "open" : ""}`}>
    <div className="side-brand"><div className="brand-mark small"><ClipboardList size={21}/></div><div><b>Office Task</b><span>Report V.74</span></div></div>
    <div className="side-user"><div className="avatar">{(session.name || "U").slice(0,1).toUpperCase()}</div><div><b>{session.name}</b><span>{session.role.replaceAll("_"," ")}</span></div></div>
    <nav>
      {items.map(([id,label,Icon]) => {
        if ((id==="employees" || id==="templates" || id==="approvals") && !canAdmin) return null;
        return <button key={id} className={view===id ? "nav active" : "nav"} onClick={()=>{setView(id);setOpen(false)}}><Icon size={18}/><span>{label}</span></button>
      })}
    </nav>
    <button className="nav logout" onClick={logout}><LogOut size={18}/><span>Logout</span></button>
  </aside>
}

function Topbar({session,onMenu,onLogout}) {
  const [now,setNow]=useState(new Date()); const [count,setCount]=useState(0);
  useEffect(()=>{const t=setInterval(()=>setNow(new Date()),1000); return()=>clearInterval(t)},[]);
  useEffect(()=>{let mounted=true; const load=async()=>{try{const r=await api("notifications",{session}); if(mounted)setCount((r.rows||[]).filter(x=>String(x.read||"N")!=="Y").length)}catch(e){}}; load(); const t=setInterval(load,5000); return()=>{mounted=false;clearInterval(t)}},[session.username]);
  return <header className="topbar">
    <button className="icon-btn mobile-menu" onClick={onMenu}><Menu/></button>
    <div><div className="top-title">Good day, {session.name.split(" ")[0]}</div><div className="top-date">{now.toLocaleDateString("en-IN",{weekday:"long",day:"2-digit",month:"short",year:"numeric"})} · {now.toLocaleTimeString("en-IN")}</div></div>
    <div className="top-actions"><div className="live-pill"><span/> Live</div><button className="icon-btn notif-btn" onClick={()=>window.dispatchEvent(new CustomEvent("open-notifications"))}><Bell size={19}/>{count>0&&<span className="notif-badge">{count>99?"99+":count}</span>}</button><button className="avatar sm" onClick={onLogout}>{session.name.slice(0,1).toUpperCase()}</button></div>
  </header>
}

function Dashboard({session,notify}) {
  const [data,setData]=useState(null);
  const [busy,setBusy]=useState(true);
  const load=async()=>{setBusy(true);try{setData(await api("dashboard",{session}));}catch(e){notify("error",e.message)}finally{setBusy(false)}};
  useEffect(()=>{load()},[]);
  if(busy && !data) return <Loader text="Loading dashboard…"/>;
  const k=data?.kpis || {};
  return <div>
    <PageHead title="Dashboard" subtitle="Your office activity at a glance." action={<button className="secondary" onClick={load}><RefreshCw size={16}/> Sync</button>}/>
    <div className="hero"><div><div className="eyebrow light">TODAY'S WORKSPACE</div><h2>Stay on top of your day.</h2><p>Attendance, tasks, reminders and approvals in one place.</p></div><div className="hero-date"><CalendarDays/> {new Date().toLocaleDateString("en-IN",{day:"2-digit",month:"short",year:"numeric"})}</div></div>
    <div className="kpi-grid">
      <Kpi icon={<Users/>} label="Present Today" value={k.present ?? 0}/>
      <Kpi icon={<ClipboardList/>} label="My Tasks" value={k.myTasks ?? 0}/>
      <Kpi icon={<CheckCircle2/>} label="Completed" value={k.completed ?? 0}/>
      <Kpi icon={<Clock3/>} label="Pending" value={k.pending ?? 0}/>
    </div>
    <div className="two-col">
      <section className="panel"><PanelTitle title="Task Progress" /><div className="chart"><ResponsiveContainer width="100%" height={260}><AreaChart data={data?.trend || []}><CartesianGrid strokeDasharray="3 3" vertical={false}/><XAxis dataKey="label"/><YAxis/><Tooltip/><Area type="monotone" dataKey="completed" stroke="#0f766e" fill="#ccfbf1" strokeWidth={2}/></AreaChart></ResponsiveContainer></div></section>
      <section className="panel"><PanelTitle title="Today’s Activity"/><ActivityList items={data?.activities || []}/></section>
    </div>
    <section className="panel"><PanelTitle title="Upcoming Reminders"/><ReminderList items={data?.reminders || []}/></section>
  </div>
}

function formatAttendanceDate(v){
  if(!v) return "-";
  const s=String(v).trim();
  let d=null;
  if(/^\d{4}-\d{2}-\d{2}/.test(s)){ const m=s.match(/^(\d{4})-(\d{2})-(\d{2})/); d=new Date(Number(m[1]),Number(m[2])-1,Number(m[3])); }
  else { const t=new Date(s); if(!Number.isNaN(t.getTime())) d=t; }
  if(!d || Number.isNaN(d.getTime())) return s;
  return `${String(d.getDate()).padStart(2,"0")}-${d.toLocaleString("en-IN",{month:"short"})}-${d.getFullYear()}`;
}
function formatAttendanceStatus(v){
  const s=String(v??"").trim();
  if(!s) return "-";
  const u=s.toUpperCase();
  if(u==="PRESENT") return "Present";
  if(u==="HOLIDAY") return "Holiday";
  if(u==="WEEKLY OFF"||u==="WEEKOFF"||u==="WEEK OFF") return "Weekly Off";
  if(u==="LEAVE") return "Leave";
  if(u==="ABSENT") return "Absent";
  return s.charAt(0).toUpperCase()+s.slice(1).toLowerCase();
}
function formatApprovalStatus(v){
  const u=String(v??"Not Approve").trim().toUpperCase();
  return ["APPROVE","APPROVED","APPROVAL","YES","Y","OK","TRUE","1"].includes(u)?"Approve":"Not Approve";
}

function formatTaskPeriod(r){ const from=r.assignmentFrom&&r.assignmentMonth&&r.assignmentYear?`${String(r.assignmentYear)}-${String(r.assignmentMonth).padStart(2,"0")}-${String(r.assignmentFrom).padStart(2,"0")}`:r.date; const to=r.assignmentFrom&&r.assignmentMonth&&r.assignmentYear?`${String(r.assignmentYear)}-${String(r.assignmentMonth).padStart(2,"0")}-${String(r.assignmentTo||r.assignmentFrom).padStart(2,"0")}`:r.due; return `${formatAttendanceDate(from)} To ${formatAttendanceDate(to)}`; }
function formatAttendanceTime(v){
  if(v===null || v===undefined || v==="") return "-";
  const s=String(v).trim();
  let h=null,m=null;
  let match=s.match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);
  if(match){ h=Number(match[1]); m=Number(match[2]); }
  else {
    match=s.match(/(?:1899-12-30\s+)?(\d{1,2}):(\d{2})(?::\d{2})?/);
    if(match){ h=Number(match[1]); m=Number(match[2]); }
  }
  if(h===null || h>23 || m>59) return s;
  const ap=h>=12?"PM":"AM"; h=h%12||12;
  return `${String(h).padStart(2,"0")}:${String(m).padStart(2,"0")} ${ap}`;
}

function Attendance({session,notify}) {
  const [data,setData]=useState(null); const [busy,setBusy]=useState(false);
  const nowForPeriod=new Date();
  const [reportMonth,setReportMonth]=useState(String(nowForPeriod.getMonth()+1).padStart(2,'0'));
  const [reportYear,setReportYear]=useState(String(nowForPeriod.getFullYear()));
  const [importBusy,setImportBusy]=useState(false); const [importResult,setImportResult]=useState(null); const [importPreview,setImportPreview]=useState(null);
  const [attendanceDate,setAttendanceDate]=useState("");
  const [inTime,setInTime]=useState(""); const [outTime,setOutTime]=useState("");
  const emptyBreak=()=>({type:"",namazType:"",start:"",end:"",reason:""});
  const normalizeBreaksForForm=(selected)=>[1,2,3].map(i=>({type:i===3?"Lunch":"Namaz",namazType:selected?.[`break${i}NamazType`]||"",start:(selected?.[`break${i}Start`]||"").slice(0,5),end:(selected?.[`break${i}End`]||"").slice(0,5),reason:""}));
  const [breaks,setBreaks]=useState([emptyBreak(),emptyBreak(),emptyBreak()]);
  const [ramadanFreeze,setRamadanFreeze]=useState(false);
  const fileRef=React.useRef(null); const loadSeq=React.useRef(0);
  const isAdmin=["MASTER_ADMIN","ADMIN","HOD"].includes(session.role);
  const load=async()=>{
    const seq=++loadSeq.current;
    setBusy(true);
    try{
      const syncKey=`otr_attendance_${session.username}_${reportYear}_${reportMonth}`; const previousSync=localStorage.getItem(syncKey)||"";
      const r=await api("attendance",{session,month:reportMonth,year:reportYear,since:previousSync});
      if(seq!==loadSeq.current) return;
      let mergedRows=r.rows||[];
      if(r.incremental){
        try{
          const cached=JSON.parse(localStorage.getItem(`${syncKey}_rows`)||"[]");
          const byDate=new Map(cached.map(x=>[String(x.date),x]));
          (r.rows||[]).forEach(x=>byDate.set(String(x.date),x));
          mergedRows=Array.from(byDate.values()).sort((a,b)=>String(a.date).localeCompare(String(b.date))).reverse();
        }catch(_e){}
      }
      const viewData={...r,rows:mergedRows};
      setData(viewData);
      setRamadanFreeze(Boolean(r.break3Frozen));
      try{localStorage.setItem(`${syncKey}_rows`,JSON.stringify(mergedRows));}catch(_e){}
      if(r.syncAt) localStorage.setItem(syncKey,r.syncAt);
      const firstEditable=r.editableDates?.[0]?.date || "";
      setAttendanceDate(firstEditable);
      const selected=(r.rows||[]).find(x=>x.date===firstEditable);
      setInTime((selected?.in||"").slice(0,5));
      setOutTime((selected?.out||"").slice(0,5));
      setBreaks(normalizeBreaksForForm(selected));
    }catch(e){if(seq===loadSeq.current) notify("error",e.message)}finally{if(seq===loadSeq.current)setBusy(false)}
  };
  const selectDate=(date)=>{
    setAttendanceDate(date);
    const row=(data?.rows||[]).find(x=>x.date===date);
    setInTime((row?.in||"").slice(0,5));
    setOutTime((row?.out||"").slice(0,5));
    setBreaks(normalizeBreaksForForm(row));
  };

  useEffect(()=>{load(); const t=setInterval(()=>load(),10000); return()=>clearInterval(t)},[reportMonth,reportYear,session.username]);

  const saveManualAttendance=async()=>{
    if(!attendanceDate){notify("error","Attendance date select karein.");return;}
    if(!inTime && !outTime){notify("error","IN Time ya OUT Time enter karein.");return;}
    try{
      setBusy(true);
      const saveBreaks=breaks.map((b,i)=>({...b,type:i===2?"Lunch":"Namaz",namazType:i===0||i===1?b.namazType:""}));
      await api("saveAttendance",{session,date:attendanceDate,inTime,outTime,breaks:saveBreaks});
      notify("success","Attendance aur break details save ho gayi.");
      await load();
    }catch(e){
      notify("error", e?.message || "Attendance save nahi ho saki. Please dobara try karein.");
    }finally{setBusy(false);}
  };

  const downloadFormat=async(type)=>{
    const headers=["Employee Id","Date","In Time","Out Time","Status","Approve/Not Approve"];
    const sample=[[session.code||"12345",attendanceDate||new Date().toISOString().slice(0,10),"09:30 AM","07:00 PM","Present","Approve"]];
    if(type==="csv"){
      const csv=[headers.join(","),...sample.map(r=>r.map(v=>`"${String(v).replaceAll('"','""')}"`).join(","))].join("\n");
      const blob=new Blob([csv],{type:"text/csv;charset=utf-8;"}); const url=URL.createObjectURL(blob); const a=document.createElement("a");
      a.href=url;a.download="Employee_Attendance_Import_Format.csv";a.click();URL.revokeObjectURL(url);return;
    }
    const XLSX=await getXLSX(); const ws=XLSX.utils.aoa_to_sheet([headers,...sample]); ws["!cols"]=[{wch:16},{wch:14},{wch:12},{wch:12},{wch:18},{wch:16},{wch:14}];
    const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,"Attendance"); XLSX.writeFile(wb,"Employee_Attendance_Import_Format.xlsx");
  };

  const importFile=async(e)=>{
    const file=e.target.files?.[0]; if(!file)return;
    setImportResult(null); setImportBusy(true);
    try{
      const XLSX=await getXLSX(); const buf=await file.arrayBuffer(); const wb=XLSX.read(buf,{type:"array",cellDates:true});
      const sheet=wb.Sheets[wb.SheetNames[0]]; const rows=XLSX.utils.sheet_to_json(sheet,{defval:"",raw:false});
      if(!rows.length)throw new Error("Excel/CSV file mein koi record nahi mila.");
      const preview=await api("previewAttendanceImport",{session,rows}); setImportPreview(preview);
      if(preview.errors?.length){throw new Error(`Import preview mein ${preview.errors.length} error(s) hain. Pehle file correct karein.`)}
      const result=await api("importAttendance",{session,rows}); setImportResult(result); setImportPreview(null);
      notify("success",`Attendance import complete: ${result.added||0} added, ${result.updated||0} pending fields filled, ${result.skipped||0} already complete. Employee sync done.`); await load();
    }catch(err){notify("error",err.message);setImportResult({errors:[err.message]});}
    finally{setImportBusy(false);if(fileRef.current)fileRef.current.value="";}
  };

  const selectedRule=data?.editableDates?.find(x=>x.date===attendanceDate);
  const oldestEditable=data?.editableDates?.[data?.editableDates?.length-1]?.date||"";
  return <div>
    <PageHead title="Attendance" subtitle="IN / OUT time manually add karein. Aaj aur previous dates allowed hain; future date allowed nahi hai."
      action={<div className="top-actions-inline"><button className="secondary" onClick={load}><RefreshCw size={16}/> Refresh</button></div>}/>

    {isAdmin && <section className="panel attendance-import-panel">
      <PanelTitle title="Employee Attendance Import" action={<span className="muted">Excel / CSV</span>}/>
      <div className="import-toolbar">
        <div className="import-help"><FileSpreadsheet size={20}/><div><b>Bulk Attendance Import</b><small>Existing attendance overwrite nahi hogi. Sirf blank/pending fields fill honge.</small></div></div>
        <div className="import-actions">
          <button className="secondary" type="button" onClick={()=>downloadFormat("xlsx")}><Download size={15}/> Excel Format</button>
          <button className="secondary" type="button" onClick={()=>downloadFormat("csv")}><Download size={15}/> CSV Format</button>
          <label className="primary upload-btn"><Upload size={15}/> {importBusy?"Importing…":"Import Excel / CSV"}<input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" onChange={importFile} disabled={importBusy}/></label>
        </div>
      </div>
      {importResult && <div className="import-summary"><b>Import Result:</b> Added {importResult.added||0} · Updated Pending {importResult.updated||0} · Skipped {importResult.skipped||0}{importResult.errors?.length?` · Errors ${importResult.errors.length}`:""}</div>}
      {importPreview && <div className="import-summary"><b>Import Preview:</b> Valid {importPreview.valid||0} · Approve {importPreview.approved||0} · Not Approve {importPreview.notApproved||0} · Preview rows {importPreview.preview?.length||0}</div>}
      {importResult?.errors?.length>0 && <div className="import-errors">{importResult.errors.slice(0,8).map((x,i)=><div key={i}>{x}</div>)}</div>}
    </section>}

    <div className="attendance-two-part">
      <section className="panel manual-attendance-panel">
        <PanelTitle title="Manual Attendance" action={<span className="muted">No future date</span>}/>
        <div className="form-grid compact-form" style={{gridTemplateColumns:"repeat(3,minmax(0,1fr))",gap:"14px"}}>
          <label>Attendance Date
            <input type="date" value={attendanceDate} min={oldestEditable} max={new Date().toISOString().slice(0,10)} onChange={e=>selectDate(e.target.value)} disabled={busy} required/>
          </label>
          <label>IN Time
            <input type="time" value={inTime} onChange={e=>setInTime(e.target.value)} disabled={!!selectedRule?.nonWorking}/>
          </label>
          <label>OUT Time
            <input type="time" value={outTime} onChange={e=>setOutTime(e.target.value)} disabled={!!selectedRule?.nonWorking}/>
          </label>
          <div className="assign-note full-span"><Clock3 size={15}/> Office Time: <b>{data?.settings?.officeInTime||"--:--"}</b> to <b>{data?.settings?.officeOutTime||"--:--"}</b> · Weekoff: <b>{data?.settings?.weekoffLabel||"Default"}</b></div>
          <div className="break-section full-span">
            <div className="break-section-head"><div><b>Break Time (Only Namaz & Lunch)</b><small>Maximum 3 breaks per working day.</small></div></div>
            <div className="break-card" style={{display:"grid",gridTemplateColumns:"180px repeat(3,minmax(0,1fr))",gap:"12px",alignItems:"end"}}><div className="break-title">Break 1: For Namaz</div><label>Namaz Name<select value={breaks[0]?.namazType||""} onChange={e=>setBreaks(bs=>bs.map((x,j)=>j===0?{...x,namazType:e.target.value,type:"Namaz"}:x))}><option value="">Select</option><option value="Zohar">Zohar</option><option value="Juma">Juma</option></select></label><label>Start Time<input type="time" value={breaks[0]?.start||""} onChange={e=>setBreaks(bs=>bs.map((x,j)=>j===0?{...x,start:e.target.value,type:"Namaz"}:x))}/></label><label>End Time<input type="time" value={breaks[0]?.end||""} onChange={e=>setBreaks(bs=>bs.map((x,j)=>j===0?{...x,end:e.target.value,type:"Namaz"}:x))}/></label></div>
            <div className="break-card" style={{display:"grid",gridTemplateColumns:"180px repeat(3,minmax(0,1fr))",gap:"12px",alignItems:"end"}}><div className="break-title">Break 2: For Namaz</div><label>Namaz Name<select value={breaks[1]?.namazType||""} onChange={e=>setBreaks(bs=>bs.map((x,j)=>j===1?{...x,namazType:e.target.value,type:"Namaz"}:x))}><option value="">Select</option><option value="Asr">Asr</option><option value="Magrib">Magrib</option></select></label><label>Start Time<input type="time" value={breaks[1]?.start||""} onChange={e=>setBreaks(bs=>bs.map((x,j)=>j===1?{...x,start:e.target.value,type:"Namaz"}:x))}/></label><label>End Time<input type="time" value={breaks[1]?.end||""} onChange={e=>setBreaks(bs=>bs.map((x,j)=>j===1?{...x,end:e.target.value,type:"Namaz"}:x))}/></label></div>
            <div className={`break-card ${ramadanFreeze?"break-frozen":""}`} style={{display:"grid",gridTemplateColumns:"180px repeat(2,minmax(0,1fr))",gap:"12px",alignItems:"end"}}><div className="break-title">Break 3: For Lunch Time</div><label>Start Time<input type="time" value={breaks[2]?.start||""} onChange={e=>setBreaks(bs=>bs.map((x,j)=>j===2?{...x,start:e.target.value,type:"Lunch"}:x))} disabled={ramadanFreeze}/></label><label>End Time<input type="time" value={breaks[2]?.end||""} onChange={e=>setBreaks(bs=>bs.map((x,j)=>j===2?{...x,end:e.target.value,type:"Lunch"}:x))} disabled={ramadanFreeze}/></label>{ramadanFreeze&&<small className="muted">Ramadan mein Lunch Break Admin ne freeze kiya hai.</small>}</div>
          </div>
          {['MASTER_ADMIN','ADMIN'].includes(session.role) && <div className="assign-note full-span ramadan-freeze-control"><label className="inline-check"><input type="checkbox" checked={ramadanFreeze} onChange={async e=>{const next=e.target.checked; try{await api("setRamadanBreakFreeze",{session,frozen:next});setRamadanFreeze(next);notify("success",next?"Ramadan Lunch Break freeze kar diya gaya.":"Ramadan Lunch Break unfreeze kar diya gaya.")}catch(err){notify("error",err.message)}}}/> Ramadan Roza: Freeze Lunch Break</label></div>}
          {selectedRule?.nonWorking && <div className="assign-note full-span">{selectedRule.reason || "Leave / Weekoff / Weekoff Adjustment"}. Attendance select karne ki zarurat nahi hai.</div>}
          <button className="primary full-span" onClick={saveManualAttendance} disabled={busy || !!selectedRule?.nonWorking}><CheckCircle2 size={16}/> Save Attendance</button>
        </div>
      </section>
      <section className="panel recent-attendance-panel"><PanelTitle title="Attendance Record" action={<div className="table-actions"><select value={reportMonth} onChange={e=>setReportMonth(e.target.value)} aria-label="Attendance Month">{Array.from({length:12},(_,i)=>{const v=String(i+1).padStart(2,'0');return <option key={v} value={v}>{new Date(2000,i,1).toLocaleString('en-IN',{month:'long'})}</option>})}</select><select value={reportYear} onChange={e=>setReportYear(e.target.value)} aria-label="Attendance Year">{Array.from({length:13},(_,i)=>{const y=String(new Date().getFullYear()-5+i);return <option key={y} value={y}>{y}</option>})}</select><button className="secondary" onClick={load}><RefreshCw size={14}/> Refresh</button></div>}/>{busy&&!data?<Loader/>:<SimpleTable columns={["Date","In Time","Out Time","Office Minutes","Total Break Minutes","Extra Break Minutes","Status","Approval"]} rows={(data?.rows||[]).map(r=>[formatAttendanceDate(r.date),formatAttendanceTime(r.in),formatAttendanceTime(r.out),r.officeMinutes||0,r.breakMinutes||0,r.extraBreakMinutes||0,formatAttendanceStatus(r.status),formatApprovalStatus(r.approveStatus)])}/>}</section>
      <section className="panel"><PanelTitle title="Break Time Report" action={<span className="muted">Only over maximum time</span>}/>{(data?.breakAlerts||[]).length?<SimpleTable columns={["Date","Break","Allowed Max","Actual","Extra Time"]} rows={(data.breakAlerts||[]).map(r=>[formatAttendanceDate(r.date),r.type,`${r.allowedMinutes} min`,`${r.actualMinutes} min`,`${r.excessMinutes} min`])}/>:<div className="empty">No break over-time records for selected month.</div>}</section>
    </div>
  </div>
}

function Notes({session,notify}) {
  const now=new Date();
  const [rows,setRows]=useState([]), [text,setText]=useState(""), [date,setDate]=useState(now.toISOString().slice(0,10)), [time,setTime]=useState(now.toTimeString().slice(0,5)), [busy,setBusy]=useState(false);
  const load=async()=>{try{const r=await api("notes",{session});setRows(r.rows||[])}catch(e){notify("error",e.message)}};
  useEffect(()=>{load()},[]);
  const save=async(e)=>{e.preventDefault();if(!text.trim()){notify("error","Note text required hai.");return;}setBusy(true);try{
    const note={text:text.trim(),date,time};
    let saved=false,lastError=null;
    for(const action of ["saveNote","createNote","addNote","saveNotes"]){
      try{await api(action,{session,note});saved=true;break}catch(err){lastError=err;if(!/Unknown action/i.test(String(err.message||err)))throw err;}
    }
    if(!saved) throw lastError||new Error("Notes save nahi ho saka.");
    notify("success","Note save ho gaya.");setText("");await load();
  }catch(e){notify("error",e.message)}finally{setBusy(false)}};
  return <div><PageHead title="Notes" subtitle="Important baatein, follow-up aur yaad rakhne wali details yahan save karein." action={<button className="secondary" onClick={load}><RefreshCw size={15}/> Refresh</button>}/>
    <div className="two-col notes-layout"><section className="panel"><PanelTitle title="Add Note"/><form className="form-grid" onSubmit={save}><label className="full-span">Text<textarea value={text} onChange={e=>setText(e.target.value)} placeholder="Note likhein..." rows="6" required/></label><label>Date<input type="date" value={date} onChange={e=>setDate(e.target.value)} required/></label><label>Time<input type="time" value={time} onChange={e=>setTime(e.target.value)} required/></label><button className="primary full-span" disabled={busy}><StickyNote size={16}/> {busy?"Saving…":"Save Note"}</button></form></section>
    <section className="panel"><PanelTitle title="My Notes"/>{rows.length?<div className="notes-list">{rows.map(r=><div className="note-item" key={r.id}><div className="note-meta"><span>{formatAttendanceDate(r.date)}</span><span>{formatAttendanceTime(r.time)}</span></div><div className="note-text">{r.text}</div></div>)}</div>:<div className="empty">No notes found.</div>}</section></div>
  </div>;
}

function Profile({session,notify}) {
  const [data,setData]=useState(null);
  const [busy,setBusy]=useState(true);
  const [pwd,setPwd]=useState({currentPassword:"",newPassword:"",confirm:""});
  const [contact,setContact]=useState({phone:"",email:""});
  const [photoBusy,setPhotoBusy]=useState(false);
  const load=async()=>{setBusy(true);try{const r=await api("profile",{session});setData(r.profile);setContact({phone:r.profile?.phone||"",email:r.profile?.email||""})}catch(e){notify("error",e.message)}finally{setBusy(false)}};
  useEffect(()=>{load()},[]);
  const changePassword=async(e)=>{e.preventDefault();if(pwd.newPassword!==pwd.confirm){notify("error","New password aur confirm password same hona chahiye.");return}try{await api("changePassword",{session,currentPassword:pwd.currentPassword,newPassword:pwd.newPassword});notify("success","Password successfully change ho gaya.");setPwd({currentPassword:"",newPassword:"",confirm:""})}catch(e){notify("error",e.message)}};
  const saveContact=async(e)=>{e.preventDefault();try{await api("profileUpdate",{session,profile:contact});notify("success","Mobile Number aur Email Id update ho gaya.");load()}catch(e){notify("error",e.message)}};
  const uploadOwnPhoto=async(e)=>{const f=e.target.files?.[0]; if(!f)return; if(!f.type.startsWith("image/")){notify("error","Sirf image file upload karein.");e.target.value="";return} if(f.size>5*1024*1024){notify("error","Photo 5 MB se chhoti honi chahiye.");e.target.value="";return} try{setPhotoBusy(true);const photo=await compressProfilePhoto(f);const r=await api("profileUpdate",{session,profile:{photoDataUrl:photo.photoDataUrl,photoFileName:photo.photoFileName,photoMimeType:photo.photoMimeType}});setData(d=>({...d,photoUrl:r.photoUrl||d.photoUrl,photoDataUrl:r.photoDataUrl||photo.photoDataUrl,photoPath:r.photoPath||d.photoPath}));notify("success","Profile photo update ho gayi.")}catch(err){notify("error",err.message)}finally{setPhotoBusy(false);e.target.value=""}};
  if(busy&&!data)return <Loader text="Loading profile…"/>;
  return <div><PageHead title="My Profile" subtitle="Apni profile details dekhein aur allowed information update karein." action={<button className="secondary" onClick={load}><RefreshCw size={15}/> Refresh</button>}/>
    <div className="two-col">
      <section className="panel profile-card"><PanelTitle title="Employee Profile"/><div className="profile-head"><div className="profile-photo">{(data?.photoDataUrl||data?.photoUrl)?<img src={data.photoDataUrl||data.photoUrl} alt="Profile"/>:<UserCircle2 size={72}/>}<label className="photo-upload" title="Update Profile Photo"><Camera size={14}/><input type="file" accept="image/*" onChange={uploadOwnPhoto} disabled={photoBusy}/></label></div><div><h2>{data?.name||session.name}</h2><p>{data?.designation||session.role}</p><span className="tag">Employee ID: {data?.employeeId||session.employeeId||session.code||"-"}</span>{data?.photoPath&&<small className="profile-note">Photo: Employees Photo</small>}</div></div>
      <div className="profile-grid"><div><span>Office City</span><b>{data?.officeCity||"-"}</b></div><div><span>Office Address</span><b>{data?.officeAddress||"-"}</b></div><div><span>Office Time</span><b>{data?.officeInTime&&data?.officeOutTime?`${formatTime(data.officeInTime)} To ${formatTime(data.officeOutTime)}`:"-"}</b></div><div><span>Weekoff</span><b>{data?.weekoffLabel||"-"}</b></div><div><span>Department</span><b>{data?.department||"-"}</b></div><div><span>Designation</span><b>{data?.designation||"-"}</b></div><div><span>Employee ID</span><b>{data?.employeeId||"-"}</b></div><div><span>Username</span><b>{data?.username||"-"}</b></div></div></section>
      <section className="panel"><PanelTitle title="My Contact"/><form className="form-grid" onSubmit={saveContact}><label>Mobile Number<input value={contact.phone} onChange={e=>setContact({...contact,phone:e.target.value})} placeholder="Mobile Number"/></label><label>Email Id<input type="email" value={contact.email} onChange={e=>setContact({...contact,email:e.target.value})} placeholder="Email Id"/></label><div className="assign-note full-span">Baaki profile details sirf Admin/HOD/Master Admin se change ki ja sakti hain.</div><button className="primary full-span"><Check size={16}/> Save Contact</button></form></section>
      <section className="panel"><PanelTitle title="Change Password"/><form className="form-grid" onSubmit={changePassword}><label>Current Password<input type="password" value={pwd.currentPassword} onChange={e=>setPwd({...pwd,currentPassword:e.target.value})} required/></label><label>New Password<input type="password" value={pwd.newPassword} onChange={e=>setPwd({...pwd,newPassword:e.target.value})} minLength={4} required/></label><label>Confirm New Password<input type="password" value={pwd.confirm} onChange={e=>setPwd({...pwd,confirm:e.target.value})} minLength={4} required/></label><button className="primary full-span"><KeyRound size={16}/> Change Password</button></form></section>
    </div></div>
}
function formatTime(v){const s=String(v||"").slice(0,5);if(!/^\d{2}:\d{2}$/.test(s))return v||"-";let [h,m]=s.split(":").map(Number);const ap=h>=12?"PM":"AM";h=h%12||12;return `${String(h).padStart(2,"0")}:${String(m).padStart(2,"0")} ${ap}`}

function Tasks({session,notify}) {
  const isAdmin=["MASTER_ADMIN","ADMIN","HOD"].includes(session.role);
  const [rows,setRows]=useState([]), [busy,setBusy]=useState(true), [search,setSearch]=useState("");
  const nowTask=new Date(); const [taskMonth,setTaskMonth]=useState(String(nowTask.getMonth()+1).padStart(2,"0")); const [taskYear,setTaskYear]=useState(String(nowTask.getFullYear()));
  const load=async()=>{setBusy(true);try{const syncKey=`otr_tasks_${session.username}_${taskYear}_${taskMonth}`;const since=localStorage.getItem(syncKey)||"";const r=await api("tasks",{session,month:taskMonth,year:taskYear,since});let merged=r.rows||[];if(r.incremental){try{const cached=JSON.parse(localStorage.getItem(`${syncKey}_rows`)||"[]");const byId=new Map(cached.map(x=>[String(x.id),x]));(r.rows||[]).forEach(x=>byId.set(String(x.id),x));merged=Array.from(byId.values()).sort((a,b)=>String(b.date).localeCompare(String(a.date)));}catch(_e){}}setRows(merged);try{localStorage.setItem(`${syncKey}_rows`,JSON.stringify(merged));}catch(_e){}if(r.syncAt)localStorage.setItem(syncKey,r.syncAt)}catch(e){notify("error",e.message)}finally{setBusy(false)}};
  useEffect(()=>{load(); const t=setInterval(()=>load(),10000); return()=>clearInterval(t)},[taskMonth,taskYear,session.username]);
  const action=async(id,type)=>{try{await api("taskAction",{session,taskId:id,type});notify("success",type==="START"?"Task started.":"Task updated.");load()}catch(e){notify("error",e.message)}};
  const saveMinutes=async(r)=>{const value=window.prompt(`Task "${r.name}" ke liye kitne minutes lage?`,"");if(value===null)return;const n=Number(value);if(!Number.isFinite(n)||n<=0){notify("error","Valid minutes add karein.");return}try{await api("taskAction",{session,taskId:r.id,type:"SAVE_TIME",minutes:n});notify("success",`${n} minutes save ho gaye.`);load()}catch(e){notify("error",e.message)}};
  const progress=async(r)=>{const value=window.prompt(`Progress % for ${r.name}`,String(r.progress||0));if(value===null)return;const note=window.prompt("Progress note (optional)",String(r.progressNote||""));try{await api("taskProgress",{session,taskId:r.id,progress:Number(value),note:note||""});notify("success","Task progress updated.");load()}catch(e){notify("error",e.message)}};
  const filtered=rows.filter(r=>Object.values(r).join(" ").toLowerCase().includes(search.toLowerCase()));
  return <div>
    <PageHead title={isAdmin?"Task Management":"Check Task"} subtitle={isAdmin?"Check assigned tasks, status, timing and progress.":"Check your assigned tasks, progress, timing and completion."} action={<div className="task-filter-bar"><select value={taskMonth} onChange={e=>setTaskMonth(e.target.value)}>{Array.from({length:12},(_,i)=>{const v=String(i+1).padStart(2,"0");return <option key={v} value={v}>{new Date(2000,i,1).toLocaleString("en-IN",{month:"long"})}</option>})}</select><select value={taskYear} onChange={e=>setTaskYear(e.target.value)}>{Array.from({length:13},(_,i)=>{const y=String(new Date().getFullYear()-5+i);return <option key={y} value={y}>{y}</option>})}</select><div className="search"><Search size={16}/><input placeholder="Search tasks…" value={search} onChange={e=>setSearch(e.target.value)}/></div></div>}/>
    <section className="panel">{busy?<Loader/>:<div className="table-wrap"><table><thead><tr><th>Task</th><th>Priority</th><th>Date To Date</th><th>Status</th><th>Progress</th><th>Actual</th><th>Action</th></tr></thead><tbody>{filtered.length?filtered.map(r=><tr key={r.id}><td><b>{r.name}</b><small>{r.category}</small></td><td><span className={`priority ${String(r.priority).toLowerCase()}`}>{r.priority}</span></td><td>{formatTaskPeriod(r)}</td><td><span className="status">{r.status}</span></td><td><div className="progress-cell"><b>{Number(r.progress||0)}%</b><div className="progress-track"><span style={{width:`${Math.min(100,Math.max(0,Number(r.progress||0)))}%`}}/></div></div></td><td>{r.actualMinutes||0} min</td><td><div className="table-actions">{r.status==="ASSIGNED"&&<button className="mini primary" onClick={()=>action(r.id,"START")}>Start</button>}{r.status==="IN_PROGRESS"&&<button className="mini primary" onClick={()=>saveMinutes(r)}>Save Minutes</button>}{!['COMPLETED','CANCELLED'].includes(r.status)&&<button className="mini secondary" onClick={()=>progress(r)}>Progress</button>}</div></td></tr>):<tr><td colSpan="7" className="empty">No tasks found.</td></tr>}</tbody></table></div>}</section>
  </div>
}

function Templates({session,notify}) {
  const empty={name:"",details:"",category:"",priority:"Normal",fromDay:"1",toDay:"1"};
  const [rows,setRows]=useState([]);
  const [form,setForm]=useState(empty);
  const [editing,setEditing]=useState(false);
  const [users,setUsers]=useState([]);
  const [assign,setAssign]=useState({templateId:"",username:"",month:String(new Date().getMonth()+1).padStart(2,"0"),year:String(new Date().getFullYear()),fromDay:"1",toDay:"1"});
  const [busy,setBusy]=useState(false);

  const taskCategories = {
    "Followup":["Monthly Report","Hind Mushawarat Task","HOD-Department Points Task","Data Required","Other"],
    "File Work":["Analise","Errors Cheking","Application Required Data","Other"],
    "Outdoor":["Office Related","Journey","Tarbiyati Ijtima","3 Din Qafila","Other"],
    "Meeting":["Online Meeting","Physicall Meeting","Other"],
    "Other":["Other"]
  };
  const taskNames=["Followup","File Work","Outdoor","Meeting","Other"];
  const categories=taskCategories[form.name] || [];

  const load=async()=>{try{
    const r=await api("templates",{session});setRows(r.rows||[]);
    const u=await api("employees",{session});setUsers((u.rows||[]).filter(x=>x.status==="ACTIVE" && ["EMPLOYEE","HOD"].includes(x.role)));
  }catch(e){notify("error",e.message)}};
  useEffect(()=>{load()},[]);

  const changeTask=(name)=>{
    const list=taskCategories[name]||[];
    setForm({...form,name,category:list[0]||""});
  };

  const reset=()=>{setForm(empty);setEditing(false)};
  const edit=(r)=>{
    setForm({name:r.name||"",details:r.details||"",category:r.category||"",priority:r.priority||"Normal",fromDay:String(r.fromDay||1),toDay:String(r.toDay||r.fromDay||1),id:r.id});
    setEditing(true);
    window.scrollTo({top:0,behavior:"smooth"});
  };

  const save=async(e)=>{
    e.preventDefault(); setBusy(true);
    const from=Number(form.fromDay), to=Number(form.toDay);
    if(from<1 || to>31 || from>to){notify("error","Task period 1 se 31 ke beech aur From Day, To Day se chhota/equal hona chahiye.");setBusy(false);return;}
    try{
      if(editing){
        await api("updateTemplate",{session,template:{...form,fromDay:from,toDay:to}});
        notify("success","Task template updated.");
      }else{
        await api("createTemplate",{session,template:{...form,fromDay:from,toDay:to}});
        notify("success","Task template created.");
      }
      reset(); load();
    }catch(e){notify("error",e.message)}finally{setBusy(false)}
  };

  const remove=async(r)=>{
    if(!window.confirm(`Delete template "${r.name} - ${r.category}"?`)) return;
    try{await api("deleteTemplate",{session,templateId:r.id});notify("success","Template deleted.");load();}
    catch(e){notify("error",e.message)}
  };

  const assignTemplate=async(e)=>{
    e.preventDefault();
    if(!assign.templateId || !assign.username){notify("error","Template aur Employee/HOD select karein.");return;}
    try{
      const r=await api("assignTemplate",{session,templateId:assign.templateId,username:assign.username,month:Number(assign.month),year:Number(assign.year),fromDay:Number(assign.fromDay),toDay:Number(assign.toDay)});
      notify("success",`Template ${r.name} ko ${r.fromDay}-${r.toDay}/${r.month}/${r.year} ke liye ${r.count} daily task(s) assign ho gaye.`);
      setAssign({...assign,templateId:"",username:""});
    }catch(e){notify("error",e.message)}
  };

  return <div>
    <PageHead title="Task Templates" subtitle="Reusable work templates for Admin/HOD." action={<button className="secondary" onClick={load}><RefreshCw size={16}/> Sync</button>}/>
    <div className="two-col">
      <form className="panel form-grid" onSubmit={save}>
        <PanelTitle title={editing ? "Edit Template" : "Create Template"} action={editing && <button type="button" className="secondary" onClick={reset}>Cancel</button>}/>
        <label>Task Name
          <select value={form.name} onChange={e=>changeTask(e.target.value)} required>
            <option value="">Select Task Name</option>
            {taskNames.map(x=><option key={x} value={x}>{x}</option>)}
          </select>
        </label>
        <label>Category
          <select value={form.category} onChange={e=>setForm({...form,category:e.target.value})} disabled={!form.name} required>
            <option value="">{form.name ? "Select Category" : "Select Task Name First"}</option>
            {categories.map(x=><option key={x} value={x}>{x}</option>)}
          </select>
        </label>
        <label>Details
          <textarea value={form.details} onChange={e=>setForm({...form,details:e.target.value})} placeholder="Task details / instructions"/>
        </label>
        <label>Priority
          <select value={form.priority} onChange={e=>setForm({...form,priority:e.target.value})}>
            <option>Low</option><option>Normal</option><option>High</option><option>Urgent</option>
          </select>
        </label>
        <div className="template-period full-span">
          <label>From Day<select value={form.fromDay} onChange={e=>setForm({...form,fromDay:e.target.value})}>{Array.from({length:31},(_,i)=><option key={i+1} value={i+1}>{i+1}</option>)}</select></label>
          <div className="period-arrow">TO</div>
          <label>To Day<select value={form.toDay} onChange={e=>setForm({...form,toDay:e.target.value})}>{Array.from({length:31},(_,i)=><option key={i+1} value={i+1}>{i+1}</option>)}</select></label>
        </div>
        <div className="period-note full-span"><CalendarDays size={15}/> Monthly task period: <b>{form.fromDay} to {form.toDay} date</b>. 30-day months will naturally have no 31st date.</div>
        <button className="primary full-span" disabled={busy}>{editing ? <><CheckCircle2 size={16}/> Update Template</> : <><Plus size={16}/> Create Template</>}</button>
      </form>
      <section className="panel">
        <PanelTitle title="Assign Task for Month / Date Range"/>
        <form className="form-grid compact-form" onSubmit={assignTemplate}>
          <label>Template
            <select value={assign.templateId} onChange={e=>{const id=e.target.value;const t=rows.find(x=>String(x.id)===String(id));setAssign({...assign,templateId:id,fromDay:String(t?.fromDay||1),toDay:String(t?.toDay||t?.fromDay||1)})}} required>
              <option value="">Select Template</option>
              {rows.map(r=><option key={r.id} value={r.id}>{r.name} · {r.category}</option>)}
            </select>
          </label>
          <label>Month<select value={assign.month} onChange={e=>setAssign({...assign,month:e.target.value})}>{Array.from({length:12},(_,i)=>{const v=String(i+1).padStart(2,"0");return <option key={v} value={v}>{new Date(2000,i,1).toLocaleString("en-IN",{month:"long"})}</option>})}</select></label>
          <label>Year<select value={assign.year} onChange={e=>setAssign({...assign,year:e.target.value})}>{Array.from({length:13},(_,i)=>{const y=String(new Date().getFullYear()-2+i);return <option key={y} value={y}>{y}</option>})}</select></label>
          <label>From Day<select value={assign.fromDay} onChange={e=>setAssign({...assign,fromDay:e.target.value})}>{Array.from({length:31},(_,i)=><option key={i+1} value={i+1}>{i+1}</option>)}</select></label>
          <label>To Day<select value={assign.toDay} onChange={e=>setAssign({...assign,toDay:e.target.value})}>{Array.from({length:31},(_,i)=><option key={i+1} value={i+1}>{i+1}</option>)}</select></label>
          <label>Employee / HOD
            <select value={assign.username} onChange={e=>setAssign({...assign,username:e.target.value})} required>
              <option value="">Select Employee / HOD</option>
              {users.map(u=><option key={u.username} value={u.username}>{u.name} · {u.role} · {u.employeeCode}</option>)}
            </select>
          </label>
          <div className="assign-note full-span">Month/Year aur Date To Date select karke ek hi task period assign karein. Task poore selected period ke liye rahega; poore saal automatically repeat nahi hoga.</div>
          <button className="primary full-span"><CheckCircle2 size={16}/> Assign for Selected Period</button>
        </form>
        <div className="section-divider"/>
        <PanelTitle title="Template Library"/>
        <div className="table-wrap"><table><thead><tr><th>Task Name</th><th>Category</th><th>Period</th><th>Priority</th><th>Action</th></tr></thead><tbody>
          {rows.length ? rows.map(r=><tr key={r.id}><td>{r.name}</td><td>{r.category}</td><td>{`${r.fromDay || 1}–${r.toDay || r.fromDay || 1}`}</td><td>{r.priority}</td><td><div className="table-actions"><button className="approve" title="Edit" onClick={()=>edit(r)}><Pencil size={14}/></button><button className="reject" title="Delete" onClick={()=>remove(r)}><X size={14}/></button></div></td></tr>) : <tr><td colSpan="5" className="empty">No records found.</td></tr>}
        </tbody></table></div>
      </section>
    </div>
  </div>
}

async function compressProfilePhoto(file){
  if(!file) return null;
  return new Promise((resolve,reject)=>{
    const reader=new FileReader();
    reader.onerror=()=>reject(new Error("Photo read nahi ho saki."));
    reader.onload=()=>{
      const img=new Image();
      img.onerror=()=>reject(new Error("Photo load nahi ho saki."));
      img.onload=()=>{
        const max=900, scale=Math.min(1,max/Math.max(img.width,img.height));
        const canvas=document.createElement("canvas"); canvas.width=Math.max(1,Math.round(img.width*scale)); canvas.height=Math.max(1,Math.round(img.height*scale));
        const ctx=canvas.getContext("2d"); ctx.drawImage(img,0,0,canvas.width,canvas.height);
        canvas.toBlob(blob=>{if(!blob){reject(new Error("Photo compress nahi ho saki."));return;} const name=(file.name||"profile").replace(/\.[^.]+$/i,"")+".jpg"; const reader2=new FileReader(); reader2.onload=()=>resolve({photoDataUrl:reader2.result,photoFileName:name,photoMimeType:"image/jpeg"}); reader2.onerror=()=>reject(new Error("Photo prepare nahi ho saki.")); reader2.readAsDataURL(blob);},"image/jpeg",0.82);
      };
      img.src=String(reader.result||"");
    };
    reader.readAsDataURL(file);
  });
}

function Employees({session,notify}) {
  const empty={name:"",code:"",username:"",password:"",role:"EMPLOYEE",department:"",designation:"",phone:"",whatsapp:"",office:"",address:"",officeInTime:"",officeOutTime:"",weekoffDay:"",country:"",region:"",state:"",division:"",district:""};
  const [rows,setRows]=useState([]); const [form,setForm]=useState(empty); const [editing,setEditing]=useState(false); const [busy,setBusy]=useState(false); const [profile,setProfile]=useState(null); const [photoFile,setPhotoFile]=useState(null); const [photoPreview,setPhotoPreview]=useState("");
  const weekDays=["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"];
  const load=async()=>{try{const r=await api("employees",{session});setRows(r.rows||[])}catch(e){notify("error",e.message)}};
  useEffect(()=>{load()},[]);
  const edit=(r)=>{setForm({...empty,name:r.name||"",code:r.code||"",username:r.username||"",password:"",role:r.role||"EMPLOYEE",department:r.department||"",designation:r.designation||"",phone:r.phone||"",whatsapp:r.whatsapp||"",office:r.office||"",address:r.address||"",officeInTime:r.officeInTime||"",officeOutTime:r.officeOutTime||"",weekoffDay:String(r.weekoffDay??""),country:r.country||"",region:r.region||"",state:r.state||"",division:r.division||"",district:r.district||""});setPhotoFile(null);setPhotoPreview("");setEditing(true);window.scrollTo({top:0,behavior:"smooth"})};
  const reset=()=>{setForm(empty);setPhotoFile(null);setPhotoPreview("");setEditing(false)};
  const selectEmployeePhoto=(e)=>{const f=e.target.files?.[0];if(!f)return;if(!f.type.startsWith("image/")){notify("error","Sirf image file upload karein.");e.target.value="";return}if(f.size>2*1024*1024){notify("error","Photo 2 MB se chhoti honi chahiye.");e.target.value="";return}setPhotoFile(f);const reader=new FileReader();reader.onload=()=>setPhotoPreview(String(reader.result||""));reader.readAsDataURL(f)};
  const readPhotoForSave=async(file)=>compressProfilePhoto(file);
  const save=async(e)=>{e.preventDefault();setBusy(true);try{const photo=await readPhotoForSave(photoFile);const employeePayload={...form,...(photo||{})};if(editing){await api("updateEmployee",{session,employee:employeePayload});notify("success","User updated and synced.")}else{const r=await api("createEmployee",{session,employee:employeePayload});notify("success",`User created successfully. Password: ${r.temporaryPassword||"set"}`)}reset();load()}catch(e){notify("error",e?.message||"User save nahi ho saka.")}finally{setBusy(false)}};
  const viewProfile=async(r)=>{try{const x=await api("employeeProfile",{session,username:r.username});setProfile(x.profile)}catch(e){notify("error",e.message)}};
  const uploadEmployeePhoto=async(e)=>{const f=e.target.files?.[0]; if(!f||!profile)return; if(!f.type.startsWith("image/")){notify("error","Sirf image file upload karein.");e.target.value="";return} if(f.size>5*1024*1024){notify("error","Photo 5 MB se chhoti honi chahiye.");e.target.value="";return} try{const photo=await compressProfilePhoto(f); const employee={name:profile.name||"",employeeId:profile.employeeId||"",code:profile.employeeId||"",username:profile.username||"",role:profile.role||"EMPLOYEE",department:profile.department||"",designation:profile.designation||"",phone:profile.phone||"",whatsapp:profile.whatsapp||"",office:profile.officeCity||"",address:profile.officeAddress||"",officeInTime:profile.officeInTime||"",officeOutTime:profile.officeOutTime||"",weekoffDay:profile.weekoffDay||"",country:profile.country||"",region:profile.region||"",state:profile.state||"",division:profile.division||"",district:profile.district||"",...photo}; await api("updateEmployee",{session,employee}); const refreshed=await api("employeeProfile",{session,username:profile.username}); setProfile(refreshed.profile); notify("success","Employee profile photo update ho gayi.")}catch(err){notify("error",err.message)}finally{e.target.value=""}};
  const remove=async(r)=>{if(r.username==="admin"){notify("error","Master Admin cannot be deleted.");return}if(!window.confirm(`Delete ${r.name} (${r.code})?`))return;try{await api("deleteEmployee",{session,username:r.username});notify("success","User deleted/deactivated and Google Sheet synced.");load()}catch(e){notify("error",e.message)}};
  return <div>
    <PageHead title="User Create" subtitle="Employee Id ke saath office timing aur weekoff bhi Admin set/edit kar sakta hai." action={<button className="secondary" onClick={load}><RefreshCw size={16}/> Sync</button>}/>
    <div className="two-col user-create-layout">
      <form className="panel form-grid" onSubmit={save}>
        <PanelTitle title={editing ? "Edit User" : "User Create"} action={editing && <button type="button" className="secondary" onClick={reset}>Cancel</button>}/>
        <label>Employee Name<input value={form.name} onChange={e=>setForm({...form,name:e.target.value})} required/></label>
        <label>Employee Id<input value={form.code} onChange={e=>setForm({...form,code:e.target.value})} required disabled={editing}/></label>
        <label>Username<input value={form.username} onChange={e=>setForm({...form,username:e.target.value})} required disabled={editing} autoComplete="username"/></label>
        <label>{editing ? "New Password (Optional)" : "Password"}<input value={form.password} onChange={e=>setForm({...form,password:e.target.value})} type="password" autoComplete="new-password" required={!editing} minLength={4} placeholder={editing ? "Leave blank to keep current password" : "Enter login password"}/></label>
        <label>Profile Photo<div className="admin-photo-field"><div className="admin-photo-preview">{photoPreview?<img src={photoPreview} alt="Selected profile"/>:<UserCircle2 size={34}/>}</div><input type="file" accept="image/*" onChange={selectEmployeePhoto}/><small>Admin/Master Admin se upload hoga · Max 2 MB</small></div></label>
        <label>Department<input value={form.department} onChange={e=>setForm({...form,department:e.target.value})}/></label>
        <label>Designation<input value={form.designation} onChange={e=>setForm({...form,designation:e.target.value})}/></label>
        <label>Role<select value={form.role} onChange={e=>setForm({...form,role:e.target.value})}><option>EMPLOYEE</option><option>HOD</option><option>ADMIN</option></select></label>
        <label>Contact<input value={form.phone} onChange={e=>setForm({...form,phone:e.target.value})}/></label>
        <label>WhatsApp<input value={form.whatsapp} onChange={e=>setForm({...form,whatsapp:e.target.value})}/></label>
        <div className="location-heading full-span">Office Settings</div>
        <label>Office City<input value={form.office} onChange={e=>setForm({...form,office:e.target.value})} placeholder="Office City"/></label>
        <label>Office Address<input value={form.address} onChange={e=>setForm({...form,address:e.target.value})} placeholder="Office Address"/></label>
        <label>Office In Time<input type="time" value={form.officeInTime} onChange={e=>setForm({...form,officeInTime:e.target.value})}/></label>
        <label>Office Out Time<input type="time" value={form.officeOutTime} onChange={e=>setForm({...form,officeOutTime:e.target.value})}/></label>
        <label>Weekoff
          <select value={form.weekoffDay} onChange={e=>setForm({...form,weekoffDay:e.target.value})}>
            <option value="">Default System Weekoff</option>
            {weekDays.map((x,i)=><option key={i} value={i}>{x}</option>)}
          </select>
        </label>
        <div className="assign-note full-span">Leave aur approved Weekoff Adjustment attendance mein automatically non-working maana jayega. Employee ko manually select nahi karna hoga.</div>
        <div className="location-heading full-span">Employee Location / Reporting Area</div>
        <label>Country<input value={form.country} onChange={e=>setForm({...form,country:e.target.value})} placeholder="Country"/></label>
        <label>Region<input value={form.region} onChange={e=>setForm({...form,region:e.target.value})} placeholder="Region"/></label>
        <label>State<input value={form.state} onChange={e=>setForm({...form,state:e.target.value})} placeholder="State"/></label>
        <label>Division<input value={form.division} onChange={e=>setForm({...form,division:e.target.value})} placeholder="Division"/></label>
        <label>District<input value={form.district} onChange={e=>setForm({...form,district:e.target.value})} placeholder="District"/></label>
        <button className="primary full-span" disabled={busy}>{editing ? <><CheckCircle2 size={16}/> Update User</> : <><Plus size={16}/> User Create</>}</button>
      </form>
      <section className="panel user-list-panel"><PanelTitle title="All Users" action={<button className="secondary" onClick={load}><RefreshCw size={15}/> Refresh</button>}/><div className="table-wrap"><table><thead><tr><th>Name</th><th>Employee Id</th><th>Username</th><th>Office City</th><th>Office In</th><th>Office Out</th><th>Weekoff</th><th>Role</th><th>Status</th><th>Action</th></tr></thead><tbody>{rows.map((r,i)=><tr key={i}><td><b>{r.name}</b></td><td><button className="link-button" onClick={()=>viewProfile(r)} title="View Employee Profile">{r.code}</button></td><td>{r.username}</td><td>{r.office}</td><td>{formatTime(r.officeInTime)}</td><td>{formatTime(r.officeOutTime)}</td><td>{r.weekoffLabel||"Default"}</td><td>{r.role}</td><td><span className="status">{r.status}</span></td><td><div className="table-actions"><button className="approve" title="View Profile" onClick={()=>viewProfile(r)}><UserCircle2 size={14}/></button><button className="approve" title="Edit User" onClick={()=>edit(r)}><Pencil size={14}/></button><button className="reject" title="Delete" onClick={()=>remove(r)} disabled={r.username==="admin"}><X size={14}/></button></div></td></tr>)}</tbody></table></div></section>
    </div>
    {profile && <div className="modal-backdrop" onClick={()=>setProfile(null)}><div className="profile-modal" onClick={e=>e.stopPropagation()}><div className="modal-head"><div><h3>Employee Profile</h3><p>{profile.name}</p></div><button className="icon-btn" onClick={()=>setProfile(null)}><X size={18}/></button></div><div className="profile-head"><div className="profile-photo">{(profile.photoDataUrl||profile.photoUrl)?<img src={profile.photoDataUrl||profile.photoUrl} alt="Profile"/>:<UserCircle2 size={72}/>}<label className="photo-upload admin-photo-upload" title="Upload Employee Photo"><Camera size={14}/><input type="file" accept="image/*" onChange={uploadEmployeePhoto}/></label></div><div><h2>{profile.name}</h2><p>{profile.designation||profile.role}</p><span className="tag">Employee ID: {profile.employeeId||"-"}</span><small className="profile-note">Photo upload: Admin only</small></div></div><div className="profile-grid"><div><span>Office City</span><b>{profile.officeCity||"-"}</b></div><div><span>Office Address</span><b>{profile.officeAddress||"-"}</b></div><div><span>Office Time</span><b>{profile.officeInTime&&profile.officeOutTime?`${formatTime(profile.officeInTime)} To ${formatTime(profile.officeOutTime)}`:"-"}</b></div><div><span>Weekoff</span><b>{profile.weekoffLabel||"-"}</b></div><div><span>Department</span><b>{profile.department||"-"}</b></div><div><span>Contact</span><b>{profile.phone||"-"}</b></div></div></div></div>}
  </div>
}

function Approvals({session,notify}) {
  const [data,setData]=useState({tasks:[],attendance:[],taskPendingCount:0,attendancePendingCount:0});
  const [selectedTasks,setSelectedTasks]=useState([]), [selectedAttendance,setSelectedAttendance]=useState([]), [busy,setBusy]=useState(false);
  const [requestRows,setRequestRows]=useState([]);
  const load=async()=>{try{const [r,p]=await Promise.all([api("approvalCenter",{session}),api("approvals",{session})]);setData(r||{});setRequestRows(p.rows||[]);setSelectedTasks([]);setSelectedAttendance([])}catch(e){notify("error",e.message)}};
  useEffect(()=>{load()},[]);
  const toggle=(set,list,key)=>set(list.includes(key)?list.filter(x=>x!==key):[...list,key]);
  const approveSelected=async(kind)=>{const list=kind==='TASK'?selectedTasks:selectedAttendance;if(!list.length){notify("error","Pehle record select karein.");return;}setBusy(true);try{const source=kind==='TASK'?data.tasks:data.attendance;const items=source.filter(x=>list.includes(x.key)).map(x=>({username:x.username,id:kind==='TASK'?x.id:x.date}));const r=await api("bulkApproval",{session,kind,items,status:"APPROVED",all:false});notify("success",`${r.count||0} ${kind==='TASK'?'task':'attendance'} approve ho gayi.`);await load()}catch(e){notify("error",e.message)}finally{setBusy(false)}};
  const approveAll=async(kind)=>{if(!window.confirm(`Saare pending ${kind==='TASK'?'Tasks':'Attendance'} approve karne hain?`))return;setBusy(true);try{const r=await api("bulkApproval",{session,kind,status:"APPROVED",all:true});notify("success",`${r.count||0} ${kind==='TASK'?'task':'attendance'} approve ho gayi.`);await load()}catch(e){notify("error",e.message)}finally{setBusy(false)}};
  const approveOne=async(kind,row)=>{setBusy(true);try{await api("approvalItem",{session,kind,username:row.username,id:kind==='TASK'?row.id:row.date,status:"APPROVED"});notify("success",`${kind==='TASK'?'Task':'Attendance'} approve ho gayi.`);await load()}catch(e){notify("error",e.message)}finally{setBusy(false)}};
  return <div>
    <PageHead title="Approvals" subtitle="Task aur Attendance ko one-by-one ya ek saath approve karein." action={<button className="secondary" onClick={load} disabled={busy}><RefreshCw size={15}/> Refresh</button>}/>
    <section className="panel"><PanelTitle title={`Task Approvals (${data.taskPendingCount||0})`} action={<div className="table-actions"><button className="secondary" onClick={()=>setSelectedTasks((data.tasks||[]).map(x=>x.key))}>Select All</button><button className="primary" onClick={()=>approveSelected('TASK')} disabled={busy||!selectedTasks.length}><Check size={14}/> Approve Selected</button><button className="approve" onClick={()=>approveAll('TASK')} disabled={busy||!(data.taskPendingCount>0)}><CheckCircle2 size={14}/> Approve All</button></div>}/><div className="table-wrap"><table><thead><tr><th>Select</th><th>Employee</th><th>Task</th><th>Date To Date</th><th>Status</th><th>Minutes</th><th>Approval</th><th>Action</th></tr></thead><tbody>{(data.tasks||[]).length?(data.tasks||[]).map(r=><tr key={r.key}><td><input type="checkbox" checked={selectedTasks.includes(r.key)} onChange={()=>toggle(setSelectedTasks,selectedTasks,r.key)}/></td><td>{r.employee}<small>{r.employeeId}</small></td><td><b>{r.name}</b><small>{r.category}</small></td><td>{formatTaskPeriod({assignmentYear:String(r.period||'').slice(0,4),assignmentMonth:String(r.period||'').slice(5,7),assignmentFrom:String(r.period||'').slice(8,10),assignmentTo:String(r.period||'').slice(-2)})}</td><td>{r.status}</td><td>{r.actualMinutes||0} min</td><td>Not Approve</td><td><button className="approve" onClick={()=>approveOne('TASK',r)} disabled={busy}><Check size={14}/> Approve</button></td></tr>):<tr><td colSpan="8" className="empty">No pending task approvals.</td></tr>}</tbody></table></div></section>
    <section className="panel"><PanelTitle title={`Attendance Approvals (${data.attendancePendingCount||0})`} action={<div className="table-actions"><button className="secondary" onClick={()=>setSelectedAttendance((data.attendance||[]).map(x=>x.key))}>Select All</button><button className="primary" onClick={()=>approveSelected('ATTENDANCE')} disabled={busy||!selectedAttendance.length}><Check size={14}/> Approve Selected</button><button className="approve" onClick={()=>approveAll('ATTENDANCE')} disabled={busy||!(data.attendancePendingCount>0)}><CheckCircle2 size={14}/> Approve All</button></div>}/><div className="table-wrap"><table><thead><tr><th>Select</th><th>Employee</th><th>Date</th><th>In Time</th><th>Out Time</th><th>Office Minutes</th><th>Status</th><th>Approval</th><th>Action</th></tr></thead><tbody>{(data.attendance||[]).length?(data.attendance||[]).map(r=><tr key={r.key}><td><input type="checkbox" checked={selectedAttendance.includes(r.key)} onChange={()=>toggle(setSelectedAttendance,selectedAttendance,r.key)}/></td><td>{r.employee}<small>{r.employeeId}</small></td><td>{formatAttendanceDate(r.date)}</td><td>{formatAttendanceTime(r.in)}</td><td>{formatAttendanceTime(r.out)}</td><td>{r.officeMinutes||0}</td><td>{r.status}</td><td>Not Approve</td><td><button className="approve" onClick={()=>approveOne('ATTENDANCE',r)} disabled={busy}><Check size={14}/> Approve</button></td></tr>):<tr><td colSpan="9" className="empty">No pending attendance approvals.</td></tr>}</tbody></table></div></section>
    <section className="panel"><PanelTitle title="Request Approvals"/><SimpleTable columns={["Type","Employee","Date","Details","Status","Action"]} rows={(requestRows||[]).map(r=>[r.type,r.employeeName||r.employee,r.date,r.details,r.status,<div className="table-actions"><button className="approve" onClick={async()=>{try{await api("approvalAction",{session,id:r.id,status:"APPROVED"});notify("success","Request approve ho gayi.");load()}catch(e){notify("error",e.message)}}}><Check size={14}/></button><button className="reject" onClick={async()=>{try{await api("approvalAction",{session,id:r.id,status:"REJECTED"});notify("success","Request reject ho gayi.");load()}catch(e){notify("error",e.message)}}}><X size={14}/></button></div>])}/></section>
  </div>
}

function Reports({session,notify}) {
  const [data,setData]=useState(null), [advanced,setAdvanced]=useState(null), [range,setRange]=useState("month"), [busy,setBusy]=useState(false);
  const load=async()=>{setBusy(true);try{const [r,a]=await Promise.all([api("reports",{session,range}),api("advancedReports",{session,range})]);setData(r);setAdvanced(a)}catch(e){notify("error",e.message)}finally{setBusy(false)}};
  useEffect(()=>{load()},[range]);
  const downloadCSV=async()=>{try{const r=await api("exportReport",{session,format:"CSV",range});const a=document.createElement("a");a.href=URL.createObjectURL(new Blob([r.content],{type:"text/csv;charset=utf-8"}));a.download=r.fileName;a.click();}catch(e){notify("error",e.message)}};
  const downloadExcel=async()=>{try{const XLSX=await getXLSX();const r=await api("exportReport",{session,format:"XLSX",range});const ws=XLSX.utils.aoa_to_sheet(r.rows);const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"Report");XLSX.writeFile(wb,r.fileName)}catch(e){notify("error",e.message)}};
  const printPDF=()=>window.print();
  return <div><PageHead title="Reports & Analytics" subtitle="Role-based employee progress, attendance and task analytics." action={<div className="table-actions"><select value={range} onChange={e=>setRange(e.target.value)}><option value="month">This Month</option><option value="year">This Year</option><option value="all">All Data</option></select><button className="secondary" onClick={downloadCSV}><Download size={15}/> CSV</button><button className="secondary" onClick={downloadExcel}><FileSpreadsheet size={15}/> Excel</button><button className="secondary" onClick={printPDF}><FileText size={15}/> PDF</button></div>}/>
    {busy&&!data?<Loader/>:<><div className="kpi-grid"><Kpi label="Employees" value={data?.employees||0}/><Kpi label="Tasks" value={data?.tasks||0}/><Kpi label="Completed" value={data?.completed||0}/><Kpi label="Pending" value={data?.pending||0}/><Kpi label="Present Today" value={data?.present||0}/></div>
    <div className="two-col"><section className="panel"><PanelTitle title="Task Trend"/><div className="chart"><ResponsiveContainer width="100%" height={280}><BarChart data={data?.trend||[]}><CartesianGrid strokeDasharray="3 3" vertical={false}/><XAxis dataKey="label"/><YAxis/><Tooltip/><Bar dataKey="completed" fill="#0f766e" radius={[5,5,0,0]}/><Bar dataKey="pending" fill="#f59e0b" radius={[5,5,0,0]}/></BarChart></ResponsiveContainer></div></section><section className="panel"><PanelTitle title="Employee Progress"/><SimpleTable columns={["Employee","Department","Tasks","Completed","Pending","Progress"]} rows={(data?.employeesRows||[]).map(r=>[r.name,r.department||"-",r.tasks,r.completed,r.pending,`${r.progress}%`])}/></section></div>
    <div className="two-col"><section className="panel"><PanelTitle title="Attendance Summary"/><SimpleTable columns={["Status","Total"]} rows={Object.entries(advanced?.attendanceSummary||{}).map(([k,v])=>[k,v])}/></section><section className="panel"><PanelTitle title="Break Analysis"/><SimpleTable columns={["Break","Total Minutes"]} rows={Object.entries(advanced?.breakTotals||{}).map(([k,v])=>[k,v])}/></section></div>
    <section className="panel"><PanelTitle title="Employee Attendance & Break Summary"/><SimpleTable columns={["Employee","Department","Present","Leave","Holiday","Weekoff","Absent","Break Minutes","Extra Break Minutes","Late","Early","Overtime"]} rows={(advanced?.employeeBreakRows||[]).map(r=>[r.employee,r.department||"-",r.present,r.leave,r.holiday,r.weekoff,r.absent,r.totalBreakMinutes,r.extraBreakMinutes||0,r.late,r.early,r.overtime])}/></section>
    <section className="panel"><PanelTitle title="Overdue Tasks" action={<span className="muted">Open / delayed tasks only</span>}/>{(advanced?.overdueTasks||[]).length?<SimpleTable columns={["Employee","Task","Due","Status"]} rows={advanced.overdueTasks.map(r=>[r.employee,r.task,formatAttendanceDate(r.due),r.status])}/>:<div className="empty">No overdue tasks for selected range.</div>}</section>
    </>}
  </div>
}

function Notifications({session,notify}) {
  const [rows,setRows]=useState([]);
  useEffect(()=>{const load=()=>api("notifications",{session}).then(r=>setRows(r.rows||[])).catch(e=>notify("error",e.message)); load(); const t=setInterval(load,5000); const open=()=>load(); window.addEventListener("open-notifications",open); return()=>{clearInterval(t);window.removeEventListener("open-notifications",open)}},[session.username]);
  const markRead=async()=>{try{await api("markNotificationsRead",{session});setRows(rows.map(x=>({...x,read:"Y"})))}catch(e){notify("error",e.message)}};
  return <div><PageHead title="Notifications" subtitle="Reminders, approvals and system alerts." action={<button className="secondary" onClick={markRead}><Check size={15}/> Mark all read</button>}/><section className="panel">{rows.length?rows.map((n,i)=><div className="notice" key={i}><div className="notice-icon"><Bell size={17}/></div><div><b>{n.title}</b><p>{n.message}</p><small>{n.time}</small></div></div>):<Empty text="No new notifications."/>}</section></div>
}

function RequestsCenter({session,notify}) {
  const isAdmin=["MASTER_ADMIN","ADMIN","HOD"].includes(session.role);
  const canRequest=Boolean(session.code || session.employeeId);
  const [data,setData]=useState({requests:[],holidays:[],approvals:[],weekoff:"0"});
  const [form,setForm]=useState({type:"Leave",date:"",details:"",adjustmentField:"IN",adjustmentTime:"",meetingMode:"Online",startTime:"",endTime:""});
  const [holiday,setHoliday]=useState({date:"",name:"",type:"PUBLIC"});
  const [weekoff,setWeekoff]=useState("0");
  const [busy,setBusy]=useState(false);
  const requestTypes=["Leave","Attendance Correction","Weekoff Adjustment","Time Adjustment","Meeting","Meeting Journey","3 Days Qafila","Tarbiyati Ijtima","Others"];
  const weekDays=["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"];
  const load=async()=>{try{const r=await api("advanced",{session});setData(r);setWeekoff(String(r.weekoff||"0"))}catch(e){notify("error",e.message)}};
  useEffect(()=>{load()},[]);
  const submitRequest=async(e)=>{e.preventDefault();setBusy(true);try{await api("createRequest",{session,request:form});notify("success","Request submit ho gayi. Approval ke liye bhej di gayi.");setForm({type:"Leave",date:"",details:"",adjustmentField:"IN",adjustmentTime:"",meetingMode:"Online",startTime:"",endTime:""});load()}catch(e){notify("error",e.message)}finally{setBusy(false)}};
  const addHoliday=async(e)=>{e.preventDefault();try{await api("createHoliday",{session,holiday});notify("success","Holiday add ho gayi.");setHoliday({date:"",name:"",type:"PUBLIC"});load()}catch(e){notify("error",e.message)}};
  const delHoliday=async(id)=>{if(!window.confirm("Is holiday ko delete karna hai?"))return;try{await api("deleteHoliday",{session,id});notify("success","Holiday delete ho gayi.");load()}catch(e){notify("error",e.message)}};
  const saveWeekoff=async()=>{try{await api("saveWeekoff",{session,day:weekoff});notify("success","Weekoff setting save ho gayi.")}catch(e){notify("error",e.message)}};
  const approve=async(id,status)=>{try{await api("approvalAction",{session,id,status});notify("success",`Request ${status.toLowerCase()} ho gayi.`);load()}catch(e){notify("error",e.message)}};
  const exportRequests=()=>{const rows=data.requests||[];const csv=["Type,Date,Details,Status,Created At",...rows.map(r=>[r.type,r.date,r.details,r.status,r.createdAt].map(v=>`"${String(v??"").replaceAll('"','""')}"`).join(","))].join("\n");const a=document.createElement("a");a.href=URL.createObjectURL(new Blob([csv],{type:"text/csv;charset=utf-8"}));a.download="My_Requests.csv";a.click();};
  return <div>
    <PageHead title="Requests & Advanced" subtitle="Leave, attendance correction, schedule requests, holidays and advanced controls." action={<button className="secondary" onClick={load}><RefreshCw size={16}/> Sync</button>}/>
    <div className="two-col">
      <section className="panel"><PanelTitle title="Create Request"/>
        {!canRequest && <div className="assign-note">Master Admin account ke liye personal request ki zarurat nahi hai. Employee/HOD accounts yahan request submit kar sakte hain.</div>}
        <form className="form-grid" onSubmit={submitRequest}>
          <label>Request Type<select value={form.type} onChange={e=>setForm({...form,type:e.target.value,adjustmentField:"IN",adjustmentTime:"",meetingMode:"Online",startTime:"",endTime:""})}>{requestTypes.map(x=><option key={x}>{x}</option>)}</select></label>
          <label>Attendance Date<input type="date" value={form.date} onChange={e=>setForm({...form,date:e.target.value})} required/></label>
          {form.type==="Time Adjustment" && <><label>Adjustment<select value={form.adjustmentField} onChange={e=>setForm({...form,adjustmentField:e.target.value})}><option value="IN">In Time</option><option value="OUT">Out Time</option></select></label><label>{form.adjustmentField} Time<input type="time" value={form.adjustmentTime} onChange={e=>setForm({...form,adjustmentTime:e.target.value})} required/></label></>}
          {form.type==="Meeting" && <label>Meeting Type<select value={form.meetingMode} onChange={e=>setForm({...form,meetingMode:e.target.value})}><option>Online</option><option>Physically</option></select></label>}
          {form.type==="Meeting Journey" && <><label>From Time<input type="time" value={form.startTime} onChange={e=>setForm({...form,startTime:e.target.value})} required/></label><label>To Time<input type="time" value={form.endTime} onChange={e=>setForm({...form,endTime:e.target.value})} required/></label></>}
          <label className="full-span">Details<textarea value={form.details} onChange={e=>setForm({...form,details:e.target.value})} placeholder="Request details / reason" required/></label>
          <button className="primary full-span" disabled={busy || !canRequest}><Plus size={16}/> Submit Request</button>
        </form>
      </section>
      <section className="panel"><PanelTitle title="My Requests"/><SimpleTable columns={["Type","Date","Details","Status"]} rows={(data.requests||[]).map(r=>[r.type,r.date,r.details,r.status])}/></section>
    </div>
    {isAdmin && <>
      <div className="two-col">
        <section className="panel"><PanelTitle title="Holiday Management"/>
          <form className="form-grid compact-form" onSubmit={addHoliday}><label>Date<input type="date" value={holiday.date} onChange={e=>setHoliday({...holiday,date:e.target.value})} required/></label><label>Holiday Name<input value={holiday.name} onChange={e=>setHoliday({...holiday,name:e.target.value})} required/></label><label>Type<select value={holiday.type} onChange={e=>setHoliday({...holiday,type:e.target.value})}><option>PUBLIC</option><option>OPTIONAL</option><option>OFFICE</option></select></label><button className="primary"><Plus size={15}/> Add Holiday</button></form>
          <div className="table-wrap"><table><thead><tr><th>Date</th><th>Name</th><th>Type</th><th></th></tr></thead><tbody>{(data.holidays||[]).map(h=><tr key={h.id}><td>{h.date}</td><td>{h.name}</td><td>{h.type}</td><td><button className="reject" onClick={()=>delHoliday(h.id)}><X size={14}/></button></td></tr>)}</tbody></table></div>
        </section>
        <section className="panel"><PanelTitle title="Weekoff Management"/><div className="form-grid compact-form"><label>Weekly Off Day<select value={weekoff} onChange={e=>setWeekoff(e.target.value)}>{weekDays.map((d,i)=><option key={i} value={i}>{d}</option>)}</select></label><button className="primary" onClick={saveWeekoff}><Check size={15}/> Save Weekoff</button></div><div className="assign-note">Weekoff aur holiday rules ko attendance reports mein future validation ke liye use kiya ja sakta hai.</div></section>
      </div>
      <section className="panel"><PanelTitle title="Pending Approvals"/><SimpleTable columns={["Employee","Employee Id","Type","Date","Details","Action"]} rows={(data.approvals||[]).map(r=>[r.employeeName,r.employeeId,r.type,r.date,r.details,<div className="table-actions"><button className="approve" onClick={()=>approve(r.id,"APPROVED")}><Check size={14}/></button><button className="reject" onClick={()=>approve(r.id,"REJECTED")}><X size={14}/></button></div>])}/></section>
      <section className="panel"><PanelTitle title="Audit Log"/><SimpleTable columns={["Time","User","Action","Module","Details"]} rows={(data.audit||[]).map(r=>[r.time,r.username,r.action,r.module,r.details])}/></section>
    </>}
  </div>
}

function SettingsPage({session,notify}){
  const isMaster=session.role==='MASTER_ADMIN'; const [auto,setAuto]=useState(null); const [busy,setBusy]=useState(false);
  const load=async()=>{if(!['MASTER_ADMIN','ADMIN','HOD'].includes(session.role))return;try{setAuto(await api("automation",{session}))}catch(e){notify("error",e.message)}};
  useEffect(()=>{load()},[]);
  const install=async()=>{setBusy(true);try{await api("installAutomation",{session});notify("success","Hourly automation enabled.");load()}catch(e){notify("error",e.message)}finally{setBusy(false)}};
  const run=async()=>{setBusy(true);try{const r=await api("runAutomation",{session});notify("success",`Automation run: ${r.reminders||0} reminders, ${r.approvals||0} approvals.`);load()}catch(e){notify("error",e.message)}finally{setBusy(false)}};
  return <div><PageHead title="Settings & Automation" subtitle="System configuration, performance and scheduled automation."/><section className="panel"><div className="setting-row"><div><b>System Architecture</b><p>Cloudflare Worker / same-origin API + Google Apps Script + Google Drive yearly employee files.</p></div><span className="tag">V.54</span></div><div className="setting-row"><div><b>Employee File Rule</b><p>One Google Sheet per employee per year: Name_EmployeeId_Year. Next year is created automatically when accessed.</p></div><span className="tag">Jan–Dec</span></div><div className="setting-row"><div><b>Automation</b><p>Hourly pending-task, approval reminders and yearly rollover checks.</p><small>Status: {auto?.enabled?"Enabled":"Not Enabled"}{auto?.lastRun?` · Last run ${auto.lastRun}`:""}</small></div><div className="table-actions"><button className="secondary" onClick={run} disabled={busy||!auto}>Run Now</button>{isMaster&&<button className="primary" onClick={install} disabled={busy}>{auto?.enabled?"Reinstall Hourly Trigger":"Enable Hourly Automation"}</button>}</div></div></section></div>
}

function PageHead({title,subtitle,action}){return <div className="page-head"><div><h1>{title}</h1><p>{subtitle}</p></div>{action}</div>}
function PanelTitle({title,action}){return <div className="panel-title"><h3>{title}</h3>{action}</div>}
function Kpi({icon,label,value}){return <div className="kpi"><div className="kpi-icon">{icon}</div><div><span>{label}</span><strong>{value}</strong></div></div>}
function ActivityList({items=[]}){return <div className="activity-list">{items.length?items.map((x,i)=><div className="activity" key={i}><span className="activity-dot"/><div><b>{x.title}</b><small>{x.time}</small></div></div>):<Empty text="No activity yet."/>}</div>}
function ReminderList({items=[]}){return <div className="reminders">{items.length?items.map((x,i)=><div className="reminder" key={i}><CalendarDays size={17}/><div><b>{x.title}</b><span>{x.when}</span></div></div>):<Empty text="No upcoming reminders."/>}</div>}
function TaskTable({rows,onAction}){return <div className="table-wrap"><table><thead><tr><th>Task</th><th>Priority</th><th>Date To Date</th><th>Status</th><th>Actual</th><th></th></tr></thead><tbody>{rows.map((r,i)=><tr key={i}><td><b>{r.name}</b><small>{r.category}</small></td><td><span className={`priority ${String(r.priority).toLowerCase()}`}>{r.priority}</span></td><td>{formatTaskPeriod(r)}</td><td><span className="status">{r.status}</span></td><td>{r.actualMinutes||0} min</td><td>{r.status==="ASSIGNED"?<button className="mini primary" onClick={()=>onAction(r.id,"START")}>Start</button>:r.status==="IN_PROGRESS"?<button className="mini primary" onClick={()=>onAction(r.id,"COMPLETE")}>Complete</button>:null}</td></tr>)}</tbody></table></div>}
function SimpleTable({columns,rows}){if(!rows.length)return <Empty text="No records found."/>;return <div className="table-wrap"><table><thead><tr>{columns.map(c=><th key={c}>{c}</th>)}</tr></thead><tbody>{rows.map((r,i)=><tr key={i}>{r.map((c,j)=><td key={j}>{c}</td>)}</tr>)}</tbody></table></div>}
function Empty({text}){return <div className="empty">{text}</div>}
function Loader({text="Loading…"}){return <div className="loader"><RefreshCw size={17} className="spin"/> {text}</div>}

createRoot(document.getElementById("root")).render(<App />);
