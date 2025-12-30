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

function isValidEmail(email) {
  const s = normalizeEmail(email);
  if (!s) return false;
  if (s.length < 3 || s.length > 254) return false;
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s);
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
  if (!isValidEmail(email)) return jsonResponse({ error: "Invalid email" }, { status: 400 });
  if (password.length < 8 || password.length > 200) return jsonResponse({ error: "Invalid password" }, { status: 400 });

  const existing = await db.prepare("SELECT id FROM users WHERE email = ?1 LIMIT 1").bind(email).first();
  if (existing?.id) return jsonResponse({ error: "Email already registered" }, { status: 409 });

  const saltBytes = crypto.getRandomValues(new Uint8Array(16));
  const iterations = 310000;
  const passHash = await hashPassword(password, saltBytes, iterations);
  const passSalt = toBase64(saltBytes);
  const id = crypto.randomUUID();
  const createdAt = Date.now();

  await db
    .prepare("INSERT INTO users (id, email, pass_hash, pass_salt, pass_iters, created_at) VALUES (?1, ?2, ?3, ?4, ?5, ?6)")
    .bind(id, email, passHash, passSalt, iterations, createdAt)
    .run();

  return jsonResponse({ ok: true });
}

