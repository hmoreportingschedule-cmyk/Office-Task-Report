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
    const payload = typeof req.body === "string" ? JSON.parse(req.body || "{}") : (req.body || {});

    const upstream = await fetch(APPS_SCRIPT_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(payload),
      redirect: "follow",
    });

    const text = await upstream.text();
    res.setHeader("Cache-Control", "no-store, max-age=0");
    res.setHeader("Content-Type", "application/json; charset=utf-8");

    // Apps Script should return JSON. Preserve its exact response so the
    // frontend receives the real backend error instead of a generic proxy error.
    return res.status(upstream.status).send(text);
  } catch (error) {
    return res.status(502).json({
      ok: false,
      message: "Vercel proxy Apps Script se connect nahi kar paaya.",
      detail: String(error && error.message ? error.message : error),
    });
  }
};
