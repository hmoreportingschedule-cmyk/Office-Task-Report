const APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbyfL_lBPCkUlHzvv0iWvbUyNmhjSE7dVh6yqHp0L4E9JAs8S6e9jDx3v2dKku3ClK6M/exec";
const UPSTREAM_TIMEOUT_MS = 30000;

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    if (url.pathname === "/api/office-task") {
      if (request.method === "OPTIONS") return new Response(null,{status:204,headers:corsHeaders()});
      if (request.method === "GET" && url.searchParams.get("health") === "1") return json({ok:true,app:"Office Task Report",worker:"V.83",proxy:true,appsScript:APPS_SCRIPT_URL});
      if (request.method !== "POST") return json({ok:false,message:"Method not allowed. Use POST."},405);
      try {
        const body=await request.text();
        const upstream=await fetchAppsScript_(body);
        const text=await upstream.text();
        let data;
        try { data=JSON.parse(text); }
        catch (_) { return json({ok:false,message:"Apps Script returned a non-JSON response.",status:upstream.status,contentType:upstream.headers.get("content-type")||"",finalUrl:upstream.url||APPS_SCRIPT_URL,detail:String(text||"").replace(/\s+/g," ").slice(0,1000)},502); }
        return json(data,upstream.status);
      } catch(e) {
        const msg=String(e?.message||e);
        return json({ok:false,message:msg.includes("timed out")?"Apps Script backend timed out after 30 seconds.":"Cloudflare Worker could not connect to Apps Script.",detail:msg},502);
      }
    }
    return env.ASSETS.fetch(request);
  }
};

async function fetchAppsScript_(body) {
  const controller=new AbortController();
  const timer=setTimeout(()=>controller.abort(),UPSTREAM_TIMEOUT_MS);
  try {
    return await fetch(APPS_SCRIPT_URL,{method:"POST",headers:{"Content-Type":"text/plain;charset=utf-8","Accept":"application/json,text/plain,*/*"},body,redirect:"follow",signal:controller.signal,cache:"no-store"});
  } catch(e) {
    if(controller.signal.aborted) throw new Error("Apps Script request timed out after 30 seconds.");
    throw e;
  } finally { clearTimeout(timer); }
}
function corsHeaders(){return {"Access-Control-Allow-Origin":"*","Access-Control-Allow-Methods":"GET,POST,OPTIONS","Access-Control-Allow-Headers":"Content-Type","Access-Control-Max-Age":"86400"};}
function json(data,status=200){return new Response(JSON.stringify(data),{status,headers:{"Content-Type":"application/json; charset=utf-8","Cache-Control":"no-store",...corsHeaders()}});}
