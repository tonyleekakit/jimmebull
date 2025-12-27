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

function normalizeUpdatesShape(obj) {
  if (!obj) return null;
  if (Array.isArray(obj)) return { weeks: obj };
  if (typeof obj !== "object") return null;

  if (Array.isArray(obj.weeks)) return obj;
  if (obj.updates && typeof obj.updates === "object" && Array.isArray(obj.updates.weeks)) return obj.updates;
  if (obj.data && typeof obj.data === "object" && Array.isArray(obj.data.weeks)) return obj.data;
  if (obj.result && typeof obj.result === "object" && Array.isArray(obj.result.weeks)) return obj.result;

  return null;
}

function stripToLikelyJson(text) {
  const s = String(text || "");
  const iObj = s.indexOf("{");
  const iArr = s.indexOf("[");
  if (iObj === -1 && iArr === -1) return s;
  if (iObj === -1) return s.slice(iArr);
  if (iArr === -1) return s.slice(iObj);
  return s.slice(Math.min(iObj, iArr));
}

function repairJsonLike(text) {
  let s = normalizeJsonLike(stripToLikelyJson(text));
  if (!s) return s;

  s = s.replace(/\bNone\b/g, "null").replace(/\bTrue\b/g, "true").replace(/\bFalse\b/g, "false");

  s = s.replace(/(^|[{,]\s*)([A-Za-z_][A-Za-z0-9_]*)(\s*:)/g, '$1"$2"$3');

  s = s.replace(/'([^'\\]*(?:\\.[^'\\]*)*)'/g, (_, inner) => `"${String(inner).replace(/"/g, '\\"')}"`);

  s = s.replace(/([\[,]\s*)([A-Za-z][A-Za-z0-9 _-]*)(\s*(?=,|\]))/g, (m, p1, word, p3) => {
    const v = String(word).trim();
    if (!v) return m;
    if (v === "true" || v === "false" || v === "null") return `${p1}${v}${p3}`;
    if (/^-?\d+(\.\d+)?$/.test(v)) return `${p1}${v}${p3}`;
    return `${p1}"${v.replace(/"/g, '\\"')}"${p3}`;
  });

  const stack = [];
  let inStr = false;
  let quote = "";
  let esc = false;

  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (inStr) {
      if (esc) {
        esc = false;
        continue;
      }
      if (ch === "\\") {
        esc = true;
        continue;
      }
      if (ch === quote) {
        inStr = false;
        quote = "";
      }
      continue;
    }

    if (ch === "\"" || ch === "'") {
      inStr = true;
      quote = ch;
      continue;
    }
    if (ch === "{" || ch === "[") stack.push(ch);
    if (ch === "}" || ch === "]") {
      const top = stack[stack.length - 1];
      if ((ch === "}" && top === "{") || (ch === "]" && top === "[")) stack.pop();
    }
  }

  while (stack.length) {
    const top = stack.pop();
    s += top === "{" ? "}" : "]";
  }

  return normalizeJsonLike(s);
}

function parseFirstValidUpdates(text) {
  const rawBlock = extractFirstJsonBlock(text);
  const block = normalizeJsonLike(rawBlock);
  const repaired = repairJsonLike(rawBlock);

  const tryParseOne = (candidate) => {
    if (!candidate) return null;
    try {
      const obj = JSON.parse(candidate);
      return normalizeUpdatesShape(obj);
    } catch {
      return null;
    }
  };

  const direct = tryParseOne(block) || tryParseOne(repaired);
  if (direct) return direct;

  const s = repaired || block;
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
        const candidate = repairJsonLike(s.slice(i, j + 1));
        const parsed = tryParseOne(candidate);
        if (parsed) return parsed;
        break;
      }
    }
  }
  return null;
}

function buildAiStateSlice(inputState, startIndex, endIndex) {
  const weeks = Array.isArray(inputState?.weeks) ? inputState.weeks : [];
  const races = [];
  weeks.forEach((w) => {
    const idx = Number(w?.index);
    if (!Number.isFinite(idx)) return;
    const pr = typeof w?.priority === "string" ? w.priority : "";
    const arr = Array.isArray(w?.races) ? w.races : [];
    arr.forEach((r) => {
      const name = typeof r?.name === "string" ? r.name.trim() : "";
      const date = typeof r?.date === "string" ? r.date.trim() : "";
      if (!name || !date) return;
      races.push({ index: idx, weekNo: w?.weekNo, monday: w?.monday, date, name, priority: pr });
    });
  });
  races.sort((a, b) => String(a.date).localeCompare(String(b.date)) || Number(a.index) - Number(b.index));

  const slice = weeks
    .slice(startIndex, endIndex + 1)
    .map((w) => ({
      index: w?.index,
      weekNo: w?.weekNo,
      monday: w?.monday,
      priority: typeof w?.priority === "string" ? w.priority : "",
      races: Array.isArray(w?.races)
        ? w.races
            .map((r) => ({
              name: typeof r?.name === "string" ? r.name.trim() : "",
              date: typeof r?.date === "string" ? r.date.trim() : "",
            }))
            .filter((r) => r.name && r.date)
        : [],
      block: typeof w?.block === "string" ? w.block : "",
      phases: Array.isArray(w?.phases) ? w.phases : [],
      volumeHrs: typeof w?.volumeHrs === "string" ? w.volumeHrs : "",
    }))
    .filter((w) => Number.isFinite(Number(w.index)));

  return { startDate: inputState?.startDate, races, weeks: slice, range: { startIndex, endIndex } };
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

  const baseInstructions = [
    "你係一個跑步訓練教練助手。",
    "你會收到用家訓練表狀態（JSON），包括 52 週資料同全季賽事清單。",
    "你要按比賽重要程度(A/B/C)以及日期，幫用家為每週設定 block/phases/volumeHrs，同埋為每週 7 日預填 durationMinutes/rpe/zone。",
    "",
    "限制：",
    `- block 只可為：${blockOptions.join(", ")}`,
    `- phases 只可為：${phaseOptions.join(", ")}`,
    "- zones 只可為 1-6；rpe 只可為 1-10；durationMinutes 係整數分鐘",
    "- 盡量令每週總 minutes 約等於 volumeHrs * 60",
    "- 比賽當週要反映優先級：A 賽前要有減量（Deload/Peaking），比賽日可安排高強度但總量要合理",
    "- 不要更改 races / priority / monday 等資料",
    "",
    "回覆格式：只輸出 JSON（不要 markdown、不要 code block、不要 ```）。",
    "JSON 結構：{ weeks: [ { index:number(0-51), block:string, phases:string[], volumeHrs:string, sessions:[{ dayIndex:number(0-6), durationMinutes:number, rpe:number(1-10), zone:number(1-6) }] } ] }",
    "請輸出最短 JSON（可以唔換行）。",
  ];

  const chunkSize = clamp(Number(env?.AI_CHUNK_WEEKS) || 13, 4, 26);
  const allWeeks = [];
  for (let startIndex = 0; startIndex < 52; startIndex += chunkSize) {
    const endIndex = Math.min(51, startIndex + chunkSize - 1);
    const sliceState = buildAiStateSlice(inputState, startIndex, endIndex);
    const prompt = [
      ...baseInstructions,
      "",
      `你只需要輸出 index ${startIndex} 至 ${endIndex} 呢段 weeks 的 updates（weeks array 只包含呢段）。`,
      notes ? `用家補充要求：${notes}` : "",
      "現況 JSON（包含全季 races + 目標範圍 weeks）：",
      JSON.stringify(sliceState),
    ]
      .filter(Boolean)
      .join("\n");

    const gen = await generateText(prompt, env);
    if (!gen || !gen.ok) {
      const status = Number(gen?.status) || 502;
      const upstreamStatus = Number.isFinite(Number(gen?.upstreamStatus)) ? Number(gen.upstreamStatus) : undefined;
      return jsonResponse(
        { error: typeof gen?.error === "string" ? gen.error : "AI service error", upstreamStatus, range: { startIndex, endIndex } },
        { status },
      );
    }

    const text = typeof gen?.text === "string" ? gen.text : "";
    const updates = parseFirstValidUpdates(text);
    if (!updates) {
      return jsonResponse(
        { error: "Invalid updates JSON", rawText: String(text).slice(0, 3000), range: { startIndex, endIndex } },
        { status: 502 },
      );
    }

    const inRangeWeeks = Array.isArray(updates.weeks)
      ? updates.weeks.filter((w) => {
          const idx = Number(w?.index);
          return Number.isFinite(idx) && idx >= startIndex && idx <= endIndex;
        })
      : [];
    allWeeks.push(...inRangeWeeks);
  }

  const sanitized = {
    weeks: allWeeks
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
