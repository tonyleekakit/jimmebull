const SERVER_VERSION = "auth-2025-12-30-1";

function jsonResponse(data, init) {
  const headers = new Headers(init?.headers || {});
  if (!headers.has("Content-Type")) headers.set("Content-Type", "application/json; charset=utf-8");
  if (!headers.has("X-Server-Version")) headers.set("X-Server-Version", SERVER_VERSION);
  return new Response(JSON.stringify(data), { ...init, headers });
}

function normalizeEmail(email) {
  return String(email || "")
    .trim()
    .toLowerCase()
    .slice(0, 254);
}

function fromBase64(s) {
  const raw = atob(String(s || ""));
  const out = new Uint8Array(raw.length);
  for (let i = 0; i < raw.length; i++) out[i] = raw.charCodeAt(i);
  return out;
}

function toBase64(bytes) {
  let s = "";
  const arr = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  for (let i = 0; i < arr.length; i++) s += String.fromCharCode(arr[i]);
  return btoa(s);
}

async function hashPassword(password, saltBytes, iterations) {
  const enc = new TextEncoder();
  const key = await crypto.subtle.importKey("raw", enc.encode(String(password || "")), "PBKDF2", false, ["deriveBits"]);
  const bits = await crypto.subtle.deriveBits(
    { name: "PBKDF2", hash: "SHA-256", salt: saltBytes, iterations },
    key,
    256,
  );
  return toBase64(new Uint8Array(bits));
}

function timingSafeEqual(a, b) {
  const sa = String(a || "");
  const sb = String(b || "");
  if (sa.length !== sb.length) return false;
  let out = 0;
  for (let i = 0; i < sa.length; i++) out |= sa.charCodeAt(i) ^ sb.charCodeAt(i);
  return out === 0;
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

  let body = null;
  try {
    body = await request.json();
  } catch {
    body = null;
  }

  const email = normalizeEmail(body?.email);
  const password = String(body?.password || "");
  if (!email || !password) return jsonResponse({ error: "Invalid credentials" }, { status: 400 });

  const user = await db
    .prepare("SELECT id, email, pass_hash, pass_salt, pass_iters FROM users WHERE email = ?1 LIMIT 1")
    .bind(email)
    .first();
  if (!user?.id) return jsonResponse({ error: "Invalid credentials" }, { status: 401 });

  const saltBytes = fromBase64(user.pass_salt);
  const iterations = Number(user.pass_iters) || 310000;
  const expectedHash = String(user.pass_hash || "");
  const actualHash = await hashPassword(password, saltBytes, iterations);
  if (!timingSafeEqual(actualHash, expectedHash)) return jsonResponse({ error: "Invalid credentials" }, { status: 401 });

  const sessionId = crypto.randomUUID();
  const now = Date.now();
  const maxAgeSec = 60 * 60 * 24 * 30;
  const expiresAt = now + maxAgeSec * 1000;

  await db.prepare("INSERT INTO sessions (id, user_id, expires_at) VALUES (?1, ?2, ?3)").bind(sessionId, user.id, expiresAt).run();

  const headers = new Headers();
  const secure = new URL(request.url).protocol === "https:";
  setCookie(headers, "tm_session", sessionId, { maxAgeSec, secure });

  return jsonResponse({ ok: true, email: user.email }, { headers });
}

