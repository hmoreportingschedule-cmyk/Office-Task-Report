const APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbyfL_lBPCkUlHzvv0iWvbUyNmhjSE7dVh6yqHp0L4E9JAs8S6e9jDx3v2dKku3ClK6M/exec";

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    if (url.pathname === "/api/office-task") {
      if (request.method === "OPTIONS") {
        return new Response(null, { status: 204, headers: corsHeaders() });
      }
      if (request.method !== "POST") {
        return json({ok:false,message:"Method not allowed. Use POST."},405);
      }
      try {
        const body = await request.text();
        let target = APPS_SCRIPT_URL;
        let upstream;
        for (let i=0;i<5;i++) {
          upstream = await fetch(target,{method:"POST",headers:{"Content-Type":"text/plain;charset=utf-8"},body,redirect:"manual"});
          const location = upstream.headers.get("location");
          if (![301,302,303,307,308].includes(upstream.status) || !location) break;
          target = new URL(location,target).toString();
        }
        const text = await upstream.text();
        if ([301,302,303,307,308].includes(upstream.status)) return json({ok:false,message:"Apps Script redirect chain could not be completed.",status:upstream.status},502);
        let data;
        try { data=JSON.parse(text); } catch { return json({ok:false,message:"Apps Script returned a non-JSON response.",detail:text.slice(0,500)},502); }
        return json(data, upstream.status);
      } catch (e) {
        return json({ok:false,message:"Cloudflare Worker could not connect to Apps Script.",detail:String(e?.message||e)},502);
      }
    }
    return env.ASSETS.fetch(request);
  }
};

function corsHeaders(){return {"Access-Control-Allow-Origin":"*","Access-Control-Allow-Methods":"POST,OPTIONS","Access-Control-Allow-Headers":"Content-Type"};}
function json(data,status=200){return new Response(JSON.stringify(data),{status,headers:{"Content-Type":"application/json; charset=utf-8","Cache-Control":"no-store","Access-Control-Allow-Origin":"*"}});}
