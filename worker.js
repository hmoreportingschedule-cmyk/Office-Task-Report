const APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbyfL_lBPCkUlHzvv0iWvbUyNmhjSE7dVh6yqHp0L4E9JAs8S6e9jDx3v2dKku3ClK6M/exec";

export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    if (url.pathname === "/api/office-task") {
      if (request.method === "OPTIONS") {
        return new Response(null, { status: 204, headers: corsHeaders() });
      }

      if (request.method !== "POST") {
        return json({ ok: false, message: "Method not allowed. Use POST." }, 405);
      }

      try {
        const body = await request.text();
        const upstream = await postToAppsScript_(body);
        const text = await upstream.text();

        let data;
        try {
          data = JSON.parse(text);
        } catch (_) {
          return json({
            ok: false,
            message: "Apps Script returned a non-JSON response.",
            status: upstream.status,
            contentType: upstream.headers.get("content-type") || "",
            detail: text.slice(0, 800)
          }, 502);
        }

        return json(data, upstream.status);
      } catch (e) {
        return json({
          ok: false,
          message: "Cloudflare Worker could not connect to Apps Script.",
          detail: String(e?.message || e)
        }, 502);
      }
    }

    return env.ASSETS.fetch(request);
  }
};

async function postToAppsScript_(body) {
  let target = APPS_SCRIPT_URL;

  for (let hop = 0; hop < 6; hop++) {
    const response = await fetch(target, {
      method: "POST",
      headers: {
        "Content-Type": "text/plain;charset=utf-8"
      },
      body,
      redirect: "manual"
    });

    const location = response.headers.get("location");
    if (![301, 302, 303, 307, 308].includes(response.status) || !location) {
      return response;
    }

    const next = new URL(location, target).toString();

    // Google Apps Script ContentService normally answers the published
    // /exec URL with a 302/303 to a one-time googleusercontent.com URL.
    // That redirected response is retrieved with GET. Resending POST to the
    // one-time URL can produce HTML/405 instead of the JSON returned by doPost.
    if ([301, 302, 303].includes(response.status)) {
      const redirected = await fetch(next, {
        method: "GET",
        redirect: "manual"
      });

      // Normally the one-time URL is the final response. If Google adds
      // another redirect, continue following it as GET without replaying the
      // original login payload.
      if (![301, 302, 303, 307, 308].includes(redirected.status)) {
        return redirected;
      }

      target = next;
      let current = redirected;
      for (let inner = 0; inner < 5; inner++) {
        const innerLocation = current.headers.get("location");
        if (!innerLocation) return current;
        const innerTarget = new URL(innerLocation, target).toString();
        current = await fetch(innerTarget, { method: "GET", redirect: "manual" });
        if (![301, 302, 303, 307, 308].includes(current.status)) return current;
        target = innerTarget;
      }
      return current;
    }

    // 307/308 preserve method and body by definition, so replay the same POST.
    target = next;
  }

  return new Response(JSON.stringify({
    ok: false,
    message: "Apps Script redirect chain could not be completed."
  }), {
    status: 502,
    headers: { "Content-Type": "application/json; charset=utf-8" }
  });
}

function corsHeaders() {
  return {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "POST,OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type"
  };
}

function json(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: {
      "Content-Type": "application/json; charset=utf-8",
      "Cache-Control": "no-store",
      "Access-Control-Allow-Origin": "*"
    }
  });
}
