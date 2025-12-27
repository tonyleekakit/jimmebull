function jsonResponse(data, init) {
  const headers = new Headers(init?.headers || {});
  if (!headers.has("Content-Type")) headers.set("Content-Type", "application/json; charset=utf-8");
  return new Response(JSON.stringify(data), { ...init, headers });
}

function clamp(n, min, max) {
  return Math.max(min, Math.min(max, n));
}

function extractFirstJsonBlock(text) {
  const s = String(text || "");
  const m = s.match(/```(?:json)?\s*([\s\S]*?)```/i);
  if (!m) return s;
  return String(m[1] || "");
}

function normalizeJsonLike(text) {
  let s = String(text || "");
  s = s.replace(/^\uFEFF/, "");
  s = s.replace(/[\u200B-\u200D\uFEFF]/g, "");
  s = s.replace(/\u2028|\u2029/g, "\n");
  s = s.replace(/[“”]/g, "\"").replace(/[‘’]/g, "'");
  s = s.replace(/,\s*([}\]])/g, "$1");
  return s.trim();
}

function parseFirstValidUpdates(text) {
  const block = normalizeJsonLike(extractFirstJsonBlock(text));
  try {
    const direct = JSON.parse(block);
    if (direct && typeof direct === "object" && Array.isArray(direct.weeks)) return direct;
  } catch {}

  const s = block;
  const n = s.length;
  for (let i = 0; i < n; i++) {
    if (s[i] !== "{") continue;
    let depth = 0;
    let inStr = false;
    let esc = false;
    for (let j = i; j < n; j++) {
      const ch = s[j];
      if (inStr) {
        if (esc) {
          esc = false;
        } else if (ch === "\\") {
          esc = true;
        } else if (ch === "\"") {
          inStr = false;
        }
        continue;
      }
      if (ch === "\"") {
        inStr = true;
        continue;
      }
      if (ch === "{") depth++;
      if (ch === "}") depth--;
      if (depth === 0) {
        const candidate = normalizeJsonLike(s.slice(i, j + 1));
        try {
          const obj = JSON.parse(candidate);
          if (obj && typeof obj === "object" && Array.isArray(obj.weeks)) return obj;
        } catch {}
        break;
      }
    }
  }
  return null;
}

async function checkRateLimit(request, env) {
  const kv = env?.KV_RATE_LIMIT;
  if (!kv || typeof kv.get !== "function" || typeof kv.put !== "function") return null;

  const ip =
    request.headers.get("CF-Connecting-IP") ||
    String(request.headers.get("X-Forwarded-For") || "")
      .split(",")[0]
      .trim() ||
    "unknown";

  const perMinute = clamp(Number(env?.AI_RATE_LIMIT_PER_MINUTE) || 10, 1, 120);
  const windowKey = Math.floor(Date.now() / 60000);
  const key = `ai_plan:${ip}:${windowKey}`;

  const current = Number(await kv.get(key)) || 0;
  if (current >= perMinute) return { ok: false, retryAfterSec: 60 };
  await kv.put(key, String(current + 1), { expirationTtl: 70 });
  return { ok: true };
}

async function runGemini(prompt, env) {
  if (!env || !env.GEMINI_API_KEY) {
    return { ok: false, status: 500, error: "Missing GEMINI_API_KEY" };
  }

  const model = typeof env.GEMINI_MODEL === "string" && env.GEMINI_MODEL.trim() ? env.GEMINI_MODEL.trim() : "gemini-1.5-flash";
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent?key=${encodeURIComponent(env.GEMINI_API_KEY)}`;

  const resp = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      contents: [{ role: "user", parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0.4 },
    }),
  });

  let data = null;
  try {
    data = await resp.json();
  } catch {
    data = null;
  }

  if (!resp.ok) {
    const message =
      (typeof data?.error?.message === "string" && data.error.message) ||
      (typeof data?.message === "string" && data.message) ||
      "Gemini API error";
    return { ok: false, status: 502, error: message, upstreamStatus: resp.status };
  }

  const text =
    data?.candidates?.[0]?.content?.parts?.map((p) => (typeof p?.text === "string" ? p.text : "")).join("") ||
    data?.candidates?.[0]?.content?.parts?.[0]?.text ||
    "";

  return { ok: true, text };
}

async function runWorkersAi(prompt, env) {
  const ai = env?.AI;
  if (!ai || typeof ai.run !== "function") {
    return { ok: false, status: 500, error: "Missing Workers AI binding (AI)" };
  }

  const model =
    typeof env?.CF_AI_MODEL === "string" && env.CF_AI_MODEL.trim() ? env.CF_AI_MODEL.trim() : "@cf/meta/llama-3.1-8b-instruct";

  let out = null;
  try {
    out = await ai.run(model, {
      messages: [
        { role: "system", content: "Return only valid JSON. Do not use markdown or code fences. Do not add extra text." },
        { role: "user", content: prompt },
      ],
    });
  } catch (e) {
    const msg = typeof e?.message === "string" && e.message.trim() ? e.message.trim() : "Workers AI error";
    return { ok: false, status: 502, error: msg };
  }

  const text =
    typeof out === "string"
      ? out
      : typeof out?.response === "string"
        ? out.response
        : typeof out?.result === "string"
          ? out.result
          : JSON.stringify(out);

  return { ok: true, text };
}

async function generateText(prompt, env) {
  const provider = typeof env?.AI_PROVIDER === "string" ? env.AI_PROVIDER.trim().toLowerCase() : "gemini";
  if (provider === "cf" || provider === "workers_ai" || provider === "workers-ai") return runWorkersAi(prompt, env);
  return runGemini(prompt, env);
}

export async function onRequestPost(ctx) {
  const { request, env } = ctx;
  const rl = await checkRateLimit(request, env);
  if (rl && !rl.ok) {
    return jsonResponse(
      { error: "Too many requests" },
      { status: 429, headers: { "Retry-After": String(rl.retryAfterSec || 60) } },
    );
  }

  let body = null;
  try {
    body = await request.json();
  } catch {
    body = null;
  }

  const inputState = body?.state;
  const notes = typeof body?.notes === "string" ? body.notes.trim().slice(0, 2000) : "";
  if (!inputState || typeof inputState !== "object" || !Array.isArray(inputState.weeks) || inputState.weeks.length !== 52) {
    return jsonResponse({ error: "Invalid state payload" }, { status: 400 });
  }

  const phaseOptions = ["Aerobic Endurance", "Tempo", "Threshold", "VO2Max", "Anaerobic", "Peaking", "Deload"];
  const blockOptions = ["Base", "Build", "Peak", "Deload", "Transition"];

  const prompt = [
    "你係一個跑步訓練教練助手。你會收到用家 52 週訓練表現況（JSON），包括每週 Monday、比賽(名稱/日期)、優先級(A/B/C)同現有欄位。",
    "請你按比賽重要程度以及日期，幫用家：",
    "1) 於每週設定 block（Base/Build/Peak/Deload/Transition）",
    `2) 於每週設定 phases（只可用以下英文 key，最多可多選）：${phaseOptions.join(", ")}`,
    "3) 為每週編排訓練量 volumeHrs（字串，1 位小數，例如 \"6.5\"）",
    "4) 為每週 7 日訓練，預填每個 Day 的 durationMinutes（整數，分鐘）及 rpe（1-10），並設定 zone（1-6）",
    "",
    "限制：",
    `- block 只可為：${blockOptions.join(", ")}`,
    `- phases 只可為：${phaseOptions.join(", ")}`,
    "- zones 只可為 1-6",
    "- 盡量令每週總 minutes 約等於 volumeHrs * 60",
    "- 比賽當週要反映優先級：A 賽前要有減量（Deload/Peaking），比賽日可安排高強度但總量要合理",
    "- 不要更改 races / priority / monday 等資料",
    "",
    "回覆格式：只輸出 JSON（不要 markdown、不要 code block、不要 ```）。",
    "JSON 結構：{ weeks: [ { index:number(0-51), block:string, phases:string[], volumeHrs:string, sessions:[{ dayIndex:number(0-6), durationMinutes:number, rpe:number(1-10), zone:number(1-6) }] } ] }",
    "",
    notes ? `用家補充要求：${notes}` : "",
    "現況 JSON：",
    JSON.stringify(inputState),
  ]
    .filter(Boolean)
    .join("\n");

  const gen = await generateText(prompt, env);
  if (!gen || !gen.ok) {
    const status = Number(gen?.status) || 502;
    const upstreamStatus = Number.isFinite(Number(gen?.upstreamStatus)) ? Number(gen.upstreamStatus) : undefined;
    return jsonResponse(
      { error: typeof gen?.error === "string" ? gen.error : "AI service error", upstreamStatus },
      { status },
    );
  }

  const text = typeof gen?.text === "string" ? gen.text : "";

  const updates = parseFirstValidUpdates(text);
  if (!updates) return jsonResponse({ error: "Invalid updates JSON", rawText: String(text).slice(0, 3000) }, { status: 502 });

  const sanitized = {
    weeks: updates.weeks
      .map((w) => {
        const index = clamp(Number(w?.index), 0, 51);
        const block = blockOptions.includes(String(w?.block || "")) ? String(w.block) : null;
        const phases = Array.isArray(w?.phases) ? w.phases.filter((p) => phaseOptions.includes(String(p))) : [];
        const volumeHrs = typeof w?.volumeHrs === "string" ? w.volumeHrs.trim() : "";
        const sessions = Array.isArray(w?.sessions)
          ? w.sessions
              .map((s) => {
                const dayIndex = clamp(Number(s?.dayIndex), 0, 6);
                const durationMinutes = Math.max(0, Math.round(Number(s?.durationMinutes) || 0));
                const rpe = clamp(Number(s?.rpe) || 1, 1, 10);
                const zone = clamp(Number(s?.zone) || 1, 1, 6);
                return { dayIndex, durationMinutes, rpe, zone };
              })
              .slice(0, 7)
          : [];
        return { index, block, phases, volumeHrs, sessions };
      })
      .filter((w) => Number.isFinite(w.index)),
  };

  return jsonResponse({ updates: sanitized }, { status: 200 });
}
