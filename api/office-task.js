const APPS_SCRIPT_URL =
  process.env.OFFICE_TASK_APPS_SCRIPT_URL ||
  "https://script.google.com/macros/s/AKfycbyfL_lBPCkUlHzvv0iWvbUyNmhjSE7dVh6yqHp0L4E9JAs8S6e9jDx3v2dKku3ClK6M/exec";

export default async function handler(req, res) {
  if (req.method === "OPTIONS") {
    res.setHeader("Access-Control-Allow-Origin", "*");
    res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
    res.setHeader("Access-Control-Allow-Headers", "Content-Type");
    return res.status(204).end();
  }
  if (req.method !== "POST") return res.status(405).json({ ok:false, message:"Method not allowed. Use POST." });

  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 30000);
  try {
    const payload = typeof req.body === "string" ? JSON.parse(req.body || "{}") : (req.body || {});
    const upstream = await fetch(APPS_SCRIPT_URL, {
      method:"POST",
      headers:{"Content-Type":"text/plain;charset=utf-8","Accept":"application/json,text/plain,*/*"},
      body:JSON.stringify(payload),
      redirect:"follow",
      signal:controller.signal,
      cache:"no-store"
    });
    const text = await upstream.text();
    let data;
    try { data=JSON.parse(text); }
    catch (_) { return res.status(502).json({ok:false,message:"Apps Script returned a non-JSON response.",status:upstream.status,detail:String(text||"").slice(0,1000)}); }
    res.setHeader("Cache-Control","no-store, max-age=0");
    res.setHeader("Content-Type","application/json; charset=utf-8");
    return res.status(upstream.status).send(JSON.stringify(data));
  } catch (error) {
    const message=error && error.name === "AbortError" ? "Apps Script backend timed out after 30 seconds." : "Proxy Apps Script se connect nahi kar paaya.";
    return res.status(502).json({ok:false,message,detail:String(error && error.message ? error.message : error)});
  } finally { clearTimeout(timer); }
};
