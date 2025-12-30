const SERVER_VERSION = "state-2025-12-30-1";

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

async function requireUserId(ctx) {
  const { request, env } = ctx;
  const db = env?.DB;
  if (!db) return { ok: false, status: 500, error: "DB not configured" };

  const sid = getCookie(request, "tm_session");
  if (!sid) return { ok: false, status: 401, error: "Not logged in" };

  const now = Date.now();
  const row = await db
    .prepare("SELECT user_id FROM sessions WHERE id = ?1 AND expires_at > ?2 LIMIT 1")
    .bind(sid, now)
    .first();

  const userId = String(row?.user_id || "");
  if (!userId) return { ok: false, status: 401, error: "Not logged in" };
  return { ok: true, userId, db };
}

function isValidStatePayload(state) {
  if (!state || typeof state !== "object") return false;
  if (!Array.isArray(state.weeks) || state.weeks.length !== 52) return false;
  return true;
}

export async function onRequestGet(ctx) {
  const auth = await requireUserId(ctx);
  if (!auth.ok) return jsonResponse({ error: auth.error }, { status: auth.status });

  const row = await auth.db
    .prepare("SELECT state_json, updated_at FROM user_state WHERE user_id = ?1 LIMIT 1")
    .bind(auth.userId)
    .first();

  if (!row?.state_json) return jsonResponse({ ok: true, state: null, updatedAt: null });

  let parsed = null;
  try {
    parsed = JSON.parse(String(row.state_json));
  } catch {
    parsed = null;
  }

  if (!isValidStatePayload(parsed)) return jsonResponse({ ok: true, state: null, updatedAt: null });
  return jsonResponse({ ok: true, state: parsed, updatedAt: Number(row.updated_at) || null });
}

export async function onRequestPut(ctx) {
  const auth = await requireUserId(ctx);
  if (!auth.ok) return jsonResponse({ error: auth.error }, { status: auth.status });

  let body = null;
  try {
    body = await ctx.request.json();
  } catch {
    body = null;
  }
  const nextState = body?.state;
  if (!isValidStatePayload(nextState)) return jsonResponse({ error: "Invalid state payload" }, { status: 400 });

  const updatedAt = Date.now();
  const json = JSON.stringify(nextState);
  await auth.db
    .prepare(
      "INSERT INTO user_state (user_id, state_json, updated_at) VALUES (?1, ?2, ?3) ON CONFLICT(user_id) DO UPDATE SET state_json = excluded.state_json, updated_at = excluded.updated_at",
    )
    .bind(auth.userId, json, updatedAt)
    .run();

  return jsonResponse({ ok: true, updatedAt });
}

