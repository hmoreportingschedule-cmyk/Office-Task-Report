const APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbyfL_lBPCkUlHzvv0iWvbUyNmhjSE7dVh6yqHp0L4E9JAs8S6e9jDx3v2dKku3ClK6M/exec";

const REDIRECT_STATUSES = new Set([301, 302, 303, 307, 308]);

export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    if (url.pathname === "/api/office-task") {
      if (request.method === "OPTIONS") {
        return new Response(null, { status: 204, headers: corsHeaders(request) });
      }

      if (request.method !== "POST") {
        return json({
          ok: false,
          message: "Method not allowed. Use POST."
        }, 405, request);
      }

      try {
        const rawBody = await request.text();

        if (!rawBody) {
          return json({
            ok: false,
            message: "Request body is empty."
          }, 400, request);
        }

        // Keep the exact POST payload while following Apps Script redirects.
        // Apps Script can redirect from script.google.com to googleusercontent.com.
        const result = await postToAppsScript_(rawBody);

        return json(result.data, result.status, request);
      } catch (error) {
        return json({
          ok: false,
          message: "Cloudflare Worker could not complete the Apps Script request.",
          detail: String(error?.message || error)
        }, 502, request);
      }
    }

    // Simple browser health endpoint for diagnosing the Worker without touching login.
    if (url.pathname === "/api/office-task/health") {
      return json({
        ok: true,
        app: "Office Task Report",
        worker: "V.36",
        proxy: true,
        appsScript: APPS_SCRIPT_URL
      }, 200, request);
    }

    return env.ASSETS.fetch(request);
  }
};

async function postToAppsScript_(body) {
  let target = APPS_SCRIPT_URL;
  let lastStatus = 0;

  for (let attempt = 0; attempt < 8; attempt++) {
    const upstream = await fetch(target, {
      method: "POST",
      headers: {
        "Content-Type": "text/plain;charset=utf-8",
        "Accept": "application/json,text/plain,*/*",
        "Cache-Control": "no-cache"
      },
      body,
      redirect: "manual"
    });

    lastStatus = upstream.status;

    if (!REDIRECT_STATUSES.has(upstream.status)) {
      const text = await upstream.text();

      let data;
      try {
        data = JSON.parse(text);
      } catch (_) {
        throw new Error(
          `Apps Script returned non-JSON response (HTTP ${upstream.status}). ` +
          text.slice(0, 300)
        );
      }

      return { status: upstream.status, data };
    }

    const location = upstream.headers.get("Location") || upstream.headers.get("location");
    if (!location) {
      throw new Error(`Apps Script returned HTTP ${upstream.status} without a redirect location.`);
    }

    target = new URL(location, target).toString();
  }

  throw new Error(`Apps Script redirect chain exceeded the maximum number of redirects (last HTTP ${lastStatus}).`);
}

function corsHeaders(request) {
  const origin = request.headers.get("Origin") || "*";
  return {
    "Access-Control-Allow-Origin": origin,
    "Access-Control-Allow-Methods": "POST,OPTIONS,GET",
    "Access-Control-Allow-Headers": "Content-Type,Accept",
    "Access-Control-Max-Age": "86400",
    "Vary": "Origin"
  };
}

function json(data, status = 200, request) {
  return new Response(JSON.stringify(data), {
    status,
    headers: {
      "Content-Type": "application/json; charset=utf-8",
      "Cache-Control": "no-store",
      ...corsHeaders(request)
    }
  });
}
