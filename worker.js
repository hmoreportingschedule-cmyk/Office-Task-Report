export default {
  async fetch(request, env) {
    const APPS_SCRIPT_URL = env.APPS_SCRIPT_URL;
    if (!APPS_SCRIPT_URL && new URL(request.url).pathname.startsWith("/api/office-task")) {
      return json({ok:false,message:"APPS_SCRIPT_URL secret is not configured."}, 500, request);
    }
    const url = new URL(request.url);

    if (url.pathname === "/api/office-task") {
      if (request.method === "OPTIONS") {
        return new Response(null, { status: 204, headers: corsHeaders(request) });
      }

      if (request.method !== "POST") {
        return json({ ok:false, message:"Method not allowed. Use POST." }, 405, request);
      }

      try {
        const body = await request.text();

        if (!body) {
          return json({ ok:false, message:"Request body is empty." }, 400, request);
        }

        // Let the Fetch runtime follow the Apps Script redirect.
        // The request body is supplied as a string so it can be replayed by the runtime.
        const upstream = await fetch(APPS_SCRIPT_URL, {
          method: "POST",
          headers: {
            "Content-Type": "text/plain;charset=utf-8",
            "Accept": "application/json,text/plain,*/*",
            "Cache-Control": "no-cache"
          },
          body,
          redirect: "follow"
        });

        const text = await upstream.text();

        let data;
        try {
          data = JSON.parse(text);
        } catch (_) {
          return json({
            ok:false,
            message:"Apps Script returned a non-JSON response.",
            status: upstream.status,
            detail:text.slice(0,1000)
          }, 502, request);
        }

        return json(data, upstream.status, request);
      } catch (e) {
        return json({
          ok:false,
          message:"Cloudflare Worker could not complete the Apps Script login request.",
          detail:String(e?.message || e)
        }, 502, request);
      }
    }

    if (url.pathname === "/api/office-task/health") {
      return json({
        ok:true,
        app:"Office Task Report",
        worker:"V.39",
        proxy:true,
        appsScriptConfigured:Boolean(APPS_SCRIPT_URL)
      }, 200, request);
    }

    return env.ASSETS.fetch(request);
  }
};

function corsHeaders(request) {
  const origin = request.headers.get("Origin") || "*";
  return {
    "Access-Control-Allow-Origin": origin,
    "Access-Control-Allow-Methods": "POST,OPTIONS,GET",
    "Access-Control-Allow-Headers": "Content-Type,Accept",
    "Access-Control-Max-Age": "86400",
    "Vary":"Origin"
  };
}

function json(data,status=200,request) {
  return new Response(JSON.stringify(data), {
    status,
    headers:{
      "Content-Type":"application/json; charset=utf-8",
      "Cache-Control":"no-store",
      ...corsHeaders(request)
    }
  });
}
