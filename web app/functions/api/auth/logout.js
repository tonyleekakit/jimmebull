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

function setCookie(headers, name, value, opts) {
  const parts = [`${name}=${value}`, "Path=/", "HttpOnly", "SameSite=Lax"];
  if (opts?.maxAgeSec !== undefined) parts.push(`Max-Age=${Math.max(0, Math.floor(Number(opts.maxAgeSec) || 0))}`);
  if (opts?.secure) parts.push("Secure");
  headers.append("Set-Cookie", parts.join("; "));
}

export async function onRequestPost(ctx) {
  const { request, env } = ctx;
  const db = env?.DB;
  if (!db) return jsonResponse({ error: "DB not configured" }, { status: 500 });

  const sid = getCookie(request, "tm_session");
  if (sid) {
    await db.prepare("DELETE FROM sessions WHERE id = ?1").bind(sid).run();
  }

  const headers = new Headers();
  const secure = new URL(request.url).protocol === "https:";
  setCookie(headers, "tm_session", "", { maxAgeSec: 0, secure });

  return jsonResponse({ ok: true }, { headers });
}

