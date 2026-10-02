// Office Task Report V.30 - Vercel serverless proxy
// This keeps the Google Apps Script URL on the server side and avoids the
// browser CORS/redirect problem seen when calling Apps Script directly.

const APPS_SCRIPT_URL =
  process.env.OFFICE_TASK_APPS_SCRIPT_URL ||
  "https://script.google.com/macros/s/AKfycbyfL_lBPCkUlHzvv0iWvbUyNmhjSE7dVh6yqHp0L4E9JAs8S6e9jDx3v2dKku3ClK6M/exec";

module.exports = async function handler(req, res) {
  if (req.method === "OPTIONS") {
    res.setHeader("Access-Control-Allow-Origin", "*");
    res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
    res.setHeader("Access-Control-Allow-Headers", "Content-Type");
    return res.status(204).end();
  }

  if (req.method !== "POST") {
    return res.status(405).json({ ok: false, message: "Method not allowed. Use POST." });
  }

  try {
    const payload = typeof req.body === "string"
      ? JSON.parse(req.body || "{}")
      : (req.body || {});
    const body = JSON.stringify(payload);

    // Apps Script Web Apps commonly return a 302 redirect from the
    // script.google.com URL to a script.googleusercontent.com URL.
    // IMPORTANT: using fetch(..., redirect:"follow") can turn the redirected
    // POST into a GET, which means the Apps Script doPost() never receives the
    // login request. We therefore follow redirects manually and resend the
    // SAME POST body at every redirect target.
    let target = APPS_SCRIPT_URL;
    let upstream;

    for (let i = 0; i < 5; i++) {
      upstream = await fetch(target, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body,
        redirect: "manual",
      });

      const location = upstream.headers.get("location");
      if (![301, 302, 303, 307, 308].includes(upstream.status) || !location) {
        break;
      }

      target = new URL(location, target).toString();
    }

    const text = await upstream.text();
    res.setHeader("Cache-Control", "no-store, max-age=0");
    res.setHeader("Content-Type", "application/json; charset=utf-8");

    // If Apps Script still returned a redirect after the safety limit, return
    // a useful diagnostic instead of leaving the browser waiting.
    if ([301, 302, 303, 307, 308].includes(upstream.status)) {
      return res.status(502).json({
        ok: false,
        message: "Apps Script redirect chain could not be completed.",
        status: upstream.status,
        location: upstream.headers.get("location") || ""
      });
    }

    return res.status(upstream.status).send(text);
  } catch (error) {
    return res.status(502).json({
      ok: false,
      message: "Vercel proxy Apps Script se connect nahi kar paaya.",
      detail: String(error && error.message ? error.message : error),
    });
  }
};
