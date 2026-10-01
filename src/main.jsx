import React, { useEffect, useMemo, useState } from "react";
import { createRoot } from "react-dom/client";
import {
  Bell, CalendarDays, CheckCircle2, Clock3, FileText, LayoutDashboard,
  LogOut, Menu, Settings, ShieldCheck, Users, ClipboardList, Search,
  UserCircle2, AlertCircle, Check, X, Plus, RefreshCw
} from "lucide-react";
import { AreaChart, Area, CartesianGrid, XAxis, YAxis, Tooltip, ResponsiveContainer, BarChart, Bar } from "recharts";
import "./styles.css";

const API_URL = import.meta.env.VITE_APPS_SCRIPT_URL || "";

async function api(action, payload = {}) {
  if (!API_URL) throw new Error("VITE_APPS_SCRIPT_URL is not configured.");
  const res = await fetch(API_URL, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({ action, ...payload })
  });
  const data = await res.json();
  if (!data.ok) throw new Error(data.message || "Request failed");
  return data;
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
          {view === "settings" && <SettingsPage session={session} notify={notify} />}
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
    ["tasks", "My Tasks", ClipboardList],
    ["templates", "Task Templates", FileText],
    ["employees", "Employees", Users],
    ["approvals", "Approvals", CheckCircle2],
    ["reports", "Progress Reports", CalendarDays],
    ["notifications", "Notifications", Bell],
    ["settings", "Settings", Settings],
  ];
  const canAdmin = ["MASTER_ADMIN","ADMIN","HOD"].includes(session.role);
  return <aside className={`sidebar ${open ? "open" : ""}`}>
    <div className="side-brand"><div className="brand-mark small"><ClipboardList size={21}/></div><div><b>Office Task</b><span>Report V.1</span></div></div>
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
  const [now,setNow]=useState(new Date());
  useEffect(()=>{const t=setInterval(()=>setNow(new Date()),1000);return()=>clearInterval(t)},[]);
  return <header className="topbar">
    <button className="icon-btn mobile-menu" onClick={onMenu}><Menu/></button>
    <div><div className="top-title">Good day, {session.name.split(" ")[0]}</div><div className="top-date">{now.toLocaleDateString("en-IN",{weekday:"long",day:"2-digit",month:"short",year:"numeric"})} · {now.toLocaleTimeString("en-IN")}</div></div>
    <div className="top-actions"><div className="live-pill"><span/> Live</div><button className="icon-btn"><Bell size={19}/></button><button className="avatar sm" onClick={onLogout}>{session.name.slice(0,1).toUpperCase()}</button></div>
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

function Attendance({session,notify}) {
  const [data,setData]=useState(null); const [busy,setBusy]=useState(false);
  const load=async()=>{setBusy(true);try{setData(await api("attendance",{session}))}catch(e){notify("error",e.message)}finally{setBusy(false)}};
  useEffect(()=>{load()},[]);
  const punch=async(type)=>{try{await api("punch",{session,type});notify("success",`${type} time recorded.`);load()}catch(e){notify("error",e.message)}};
  return <div><PageHead title="Attendance" subtitle="Track IN, OUT, breaks and office time." action={<button className="secondary" onClick={load}><RefreshCw size={16}/> Refresh</button>}/>
    <div className="attendance-hero panel">
      <div><span className="status-dot"/> Today</div>
      <div className="big-time">{data?.today?.in || "--:--"} <span>→</span> {data?.today?.out || "--:--"}</div>
      <div className="muted">Office Time: <b>{data?.today?.officeMinutes || 0} min</b> · Break: <b>{data?.today?.breakMinutes || 0} min</b></div>
      <div className="action-row"><button className="primary" onClick={()=>punch("IN")} disabled={!!data?.today?.in}>IN Time</button><button className="secondary" onClick={()=>punch("OUT")} disabled={!data?.today?.in || !!data?.today?.out}>OUT Time</button></div>
    </div>
    <section className="panel"><PanelTitle title="Recent Attendance"/>{busy&&!data?<Loader/>:<SimpleTable columns={["Date","IN","OUT","Office Minutes","Status"]} rows={(data?.rows||[]).map(r=>[r.date,r.in,r.out,r.officeMinutes,r.status])}/>}</section>
  </div>
}

function Tasks({session,notify}) {
  const [rows,setRows]=useState([]); const [busy,setBusy]=useState(true); const [search,setSearch]=useState("");
  const load=async()=>{setBusy(true);try{const r=await api("tasks",{session});setRows(r.rows||[])}catch(e){notify("error",e.message)}finally{setBusy(false)}};
  useEffect(()=>{load()},[]);
  const action=async(id,type)=>{try{await api("taskAction",{session,taskId:id,type});notify("success",type==="START"?"Task started.":"Task marked complete.");load()}catch(e){notify("error",e.message)}};
  const filtered=rows.filter(r=>Object.values(r).join(" ").toLowerCase().includes(search.toLowerCase()));
  return <div><PageHead title="My Tasks" subtitle="Assigned work, progress and actual time." action={<div className="search"><Search size={16}/><input placeholder="Search tasks…" value={search} onChange={e=>setSearch(e.target.value)}/></div>}/>
    <section className="panel">{busy?<Loader/>:<TaskTable rows={filtered} onAction={action}/>}</section>
  </div>
}

function Templates({session,notify}) {
  const [rows,setRows]=useState([]); const [form,setForm]=useState({name:"",details:"",category:"General",priority:"Normal",minutes:30});
  const load=async()=>{try{const r=await api("templates",{session});setRows(r.rows||[])}catch(e){notify("error",e.message)}};
  useEffect(()=>{load()},[]);
  const save=async(e)=>{e.preventDefault();try{await api("createTemplate",{session,template:form});setForm({name:"",details:"",category:"General",priority:"Normal",minutes:30});notify("success","Template created.");load()}catch(e){notify("error",e.message)}};
  return <div><PageHead title="Task Templates" subtitle="Reusable work templates for Admin/HOD."/>
    <div className="two-col"><form className="panel form-grid" onSubmit={save}><PanelTitle title="Create Template"/><label>Task Name<input value={form.name} onChange={e=>setForm({...form,name:e.target.value})} required/></label><label>Details<textarea value={form.details} onChange={e=>setForm({...form,details:e.target.value})}/></label><label>Category<input value={form.category} onChange={e=>setForm({...form,category:e.target.value})}/></label><label>Priority<select value={form.priority} onChange={e=>setForm({...form,priority:e.target.value})}><option>Low</option><option>Normal</option><option>High</option><option>Urgent</option></select></label><label>Expected Minutes<input type="number" min="1" value={form.minutes} onChange={e=>setForm({...form,minutes:e.target.value})}/></label><button className="primary"><Plus size={16}/> Create Template</button></form>
    <section className="panel"><PanelTitle title="Template Library"/><SimpleTable columns={["Name","Category","Priority","Minutes"]} rows={rows.map(r=>[r.name,r.category,r.priority,r.minutes])}/></section></div>
  </div>
}

function Employees({session,notify}) {
  const [rows,setRows]=useState([]); const [form,setForm]=useState({name:"",code:"",username:"",role:"EMPLOYEE",department:"",designation:"",phone:"",whatsapp:"",office:"",address:""});
  const load=async()=>{try{const r=await api("employees",{session});setRows(r.rows||[])}catch(e){notify("error",e.message)}};
  useEffect(()=>{load()},[]);
  const save=async(e)=>{e.preventDefault();try{await api("createEmployee",{session,employee:form});notify("success","Employee created and yearly file initialized.");setForm({name:"",code:"",username:"",role:"EMPLOYEE",department:"",designation:"",phone:"",whatsapp:"",office:"",address:""});load()}catch(e){notify("error",e.message)}};
  return <div><PageHead title="Employees" subtitle="Manage employee profiles, roles and access."/>
    <div className="two-col"><form className="panel form-grid" onSubmit={save}><PanelTitle title="Add Employee"/>{Object.entries({name:"Employee Name",code:"Employee Code",username:"Username",department:"Department",designation:"Designation",phone:"Contact",whatsapp:"WhatsApp",office:"Office Location",address:"Office Address"}).map(([k,l])=><label key={k}>{l}<input value={form[k]} onChange={e=>setForm({...form,[k]:e.target.value})} required={["name","code","username"].includes(k)}/></label>)}<label>Role<select value={form.role} onChange={e=>setForm({...form,role:e.target.value})}><option>EMPLOYEE</option><option>HOD</option><option>ADMIN</option></select></label><button className="primary"><Plus size={16}/> Create Employee</button></form>
    <section className="panel"><PanelTitle title="Employee Directory"/><SimpleTable columns={["Name","Code","Department","Role","Office","Status"]} rows={rows.map(r=>[r.name,r.code,r.department,r.role,r.office,r.status])}/></section></div>
  </div>
}

function Approvals({session,notify}) {
  const [rows,setRows]=useState([]);
  const load=async()=>{try{const r=await api("approvals",{session});setRows(r.rows||[])}catch(e){notify("error",e.message)}};
  useEffect(()=>{load()},[]);
  const act=async(id,status)=>{try{await api("approvalAction",{session,id,status});notify("success",`Request ${status.toLowerCase()}.`);load()}catch(e){notify("error",e.message)}};
  return <div><PageHead title="Approvals" subtitle="Leave, attendance and task approvals."/><section className="panel"><SimpleTable columns={["Type","Employee","Date","Details","Status","Action"]} rows={rows.map(r=>[r.type,r.employee,r.date,r.details,r.status,<div className="table-actions"><button className="approve" onClick={()=>act(r.id,"APPROVED")}><Check size={14}/></button><button className="reject" onClick={()=>act(r.id,"REJECTED")}><X size={14}/></button></div>])}/></section></div>
}

function Reports({session,notify}) {
  const [data,setData]=useState(null);
  useEffect(()=>{api("reports",{session}).then(setData).catch(e=>notify("error",e.message))},[]);
  return <div><PageHead title="Progress Reports" subtitle="Attendance, tasks and productivity insights."/>
    <div className="kpi-grid"><Kpi label="Employees" value={data?.employees||0}/><Kpi label="Tasks" value={data?.tasks||0}/><Kpi label="Completed" value={data?.completed||0}/><Kpi label="Pending" value={data?.pending||0}/></div>
    <div className="two-col"><section className="panel"><PanelTitle title="Task Trend"/><div className="chart"><ResponsiveContainer width="100%" height={280}><BarChart data={data?.trend||[]}><CartesianGrid strokeDasharray="3 3" vertical={false}/><XAxis dataKey="label"/><YAxis/><Tooltip/><Bar dataKey="completed" fill="#0f766e" radius={[5,5,0,0]}/><Bar dataKey="pending" fill="#f59e0b" radius={[5,5,0,0]}/></BarChart></ResponsiveContainer></div></section><section className="panel"><PanelTitle title="Employee Progress"/><SimpleTable columns={["Employee","Tasks","Completed","Pending","Progress"]} rows={(data?.employeesRows||[]).map(r=>[r.name,r.tasks,r.completed,r.pending,`${r.progress}%`])}/></section></div>
  </div>
}

function Notifications({session,notify}) {
  const [rows,setRows]=useState([]);
  useEffect(()=>{api("notifications",{session}).then(r=>setRows(r.rows||[])).catch(e=>notify("error",e.message))},[]);
  return <div><PageHead title="Notifications" subtitle="Reminders, approvals and system alerts."/><section className="panel">{rows.length?rows.map((n,i)=><div className="notice" key={i}><div className="notice-icon"><Bell size={17}/></div><div><b>{n.title}</b><p>{n.message}</p><small>{n.time}</small></div></div>):<Empty text="No new notifications."/>}</section></div>
}

function SettingsPage(){return <div><PageHead title="Settings" subtitle="System preferences and configuration."/><section className="panel"><div className="setting-row"><div><b>System Architecture</b><p>Vercel frontend + Google Apps Script + Google Drive yearly employee files.</p></div><span className="tag">V.1</span></div><div className="setting-row"><div><b>Employee File Rule</b><p>One Google Sheet per employee per year: Name_EmployeeCode_Year</p></div><span className="tag">Jan–Dec</span></div></section></div>}

function PageHead({title,subtitle,action}){return <div className="page-head"><div><h1>{title}</h1><p>{subtitle}</p></div>{action}</div>}
function PanelTitle({title,action}){return <div className="panel-title"><h3>{title}</h3>{action}</div>}
function Kpi({icon,label,value}){return <div className="kpi"><div className="kpi-icon">{icon}</div><div><span>{label}</span><strong>{value}</strong></div></div>}
function ActivityList({items=[]}){return <div className="activity-list">{items.length?items.map((x,i)=><div className="activity" key={i}><span className="activity-dot"/><div><b>{x.title}</b><small>{x.time}</small></div></div>):<Empty text="No activity yet."/>}</div>}
function ReminderList({items=[]}){return <div className="reminders">{items.length?items.map((x,i)=><div className="reminder" key={i}><CalendarDays size={17}/><div><b>{x.title}</b><span>{x.when}</span></div></div>):<Empty text="No upcoming reminders."/>}</div>}
function TaskTable({rows,onAction}){return <div className="table-wrap"><table><thead><tr><th>Task</th><th>Priority</th><th>Due</th><th>Status</th><th>Actual</th><th></th></tr></thead><tbody>{rows.map((r,i)=><tr key={i}><td><b>{r.name}</b><small>{r.category}</small></td><td><span className={`priority ${String(r.priority).toLowerCase()}`}>{r.priority}</span></td><td>{r.due}</td><td><span className="status">{r.status}</span></td><td>{r.actualMinutes||0} min</td><td>{r.status==="ASSIGNED"?<button className="mini primary" onClick={()=>onAction(r.id,"START")}>Start</button>:r.status==="IN_PROGRESS"?<button className="mini primary" onClick={()=>onAction(r.id,"COMPLETE")}>Complete</button>:null}</td></tr>)}</tbody></table></div>}
function SimpleTable({columns,rows}){if(!rows.length)return <Empty text="No records found."/>;return <div className="table-wrap"><table><thead><tr>{columns.map(c=><th key={c}>{c}</th>)}</tr></thead><tbody>{rows.map((r,i)=><tr key={i}>{r.map((c,j)=><td key={j}>{c}</td>)}</tr>)}</tbody></table></div>}
function Empty({text}){return <div className="empty">{text}</div>}
function Loader({text="Loading…"}){return <div className="loader"><RefreshCw size={17} className="spin"/> {text}</div>}

createRoot(document.getElementById("root")).render(<App />);
