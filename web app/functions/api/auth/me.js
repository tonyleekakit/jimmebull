const SERVER_VERSION = "auth-2025-12-30-1";

function jsonResponse(data, init) {
  const headers = new Headers(init?.headers || {});
  if (!headers.has("Content-Type")) headers.set("Content-Type", "application/json; charset=utf-8");
  if (!headers.has("X-Server-Version")) headers.set("X-Server-Version", SERVER_VERSION);
  return new Response(JSON.stringify(data), { ...init, headers });
}

function getCookie(request, name) {
  const cookie = request.headers.get("Cookie") || "";
  const parts = cookie.split(";").map((p) => p.trim());
  const prefix = `${name}=`;
  for (const p of parts) {
    if (p.startsWith(prefix)) return p.slice(prefix.length);
  }
  return "";
}

export async function onRequestGet(ctx) {
  const { request, env } = ctx;
  const db = env?.DB;
  if (!db) return jsonResponse({ error: "DB not configured" }, { status: 500 });

  const sid = getCookie(request, "tm_session");
  if (!sid) return jsonResponse({ ok: true, user: null });

  const now = Date.now();
  const row = await db
    .prepare(
      "SELECT u.id AS user_id, u.email AS email FROM sessions s JOIN users u ON u.id = s.user_id WHERE s.id = ?1 AND s.expires_at > ?2 LIMIT 1",
    )
    .bind(sid, now)
    .first();

  if (!row?.user_id) return jsonResponse({ ok: true, user: null });
  return jsonResponse({ ok: true, user: { email: row.email } });
}

