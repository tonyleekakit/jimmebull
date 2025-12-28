const MS_PER_DAY = 24 * 60 * 60 * 1000;
     
function pad2(n) {
  return String(n).padStart(2, "0");
}

function formatMD(d) {
  const m = d.getMonth() + 1;
  const day = d.getDate();
  return `${day}/${m}`;
}

function formatYMD(d) {
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
}

function formatWeekdayEnShort(d) {
  const days = ["週日", "週一", "週二", "週三", "週四", "週五", "週六"];
  return days[d.getDay()] || "";
}

function parseYMD(value) {
  if (typeof value !== "string") return null;
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value.trim());
  if (!m) return null;
  const y = Number(m[1]);
  const mo = Number(m[2]);
  const d = Number(m[3]);
  if (!Number.isFinite(y) || !Number.isFinite(mo) || !Number.isFinite(d)) return null;
  if (mo < 1 || mo > 12) return null;
  if (d < 1 || d > 31) return null;
  const dt = new Date(`${m[1]}-${m[2]}-${m[3]}T00:00:00`);
  if (Number.isNaN(dt.getTime())) return null;
  return dt;
}

function clamp(n, min, max) {
  return Math.max(min, Math.min(max, n));
}

function lockScrollPosition(updateFn) {
  const x = window.scrollX;
  const y = window.scrollY;
  updateFn();
  requestAnimationFrame(() => {
    requestAnimationFrame(() => {
      window.scrollTo(x, y);
    });
  });
}

function computePastMean(values, endIndex, lookback) {
  const end = clamp(Number(endIndex) || 0, 0, values.length);
  const back = Math.max(0, Math.round(Number(lookback) || 0));
  const start = Math.max(0, end - back);
  let sum = 0;
  let count = 0;
  for (let i = start; i < end; i++) {
    sum += Number(values[i]) || 0;
    count++;
  }
  return count ? sum / count : 0;
}

function formatVolumeHrs(n) {
  if (!Number.isFinite(n) || n <= 0) return "";
  return (Math.round(n * 10) / 10).toFixed(1);
}

function recomputeFormulaVolumes() {
  if (!state || !Array.isArray(state.weeks)) return;
  const effective = [];
  for (let i = 0; i < state.weeks.length; i++) {
    const w = state.weeks[i];
    if (!w || typeof w !== "object") {
      effective[i] = 0;
      continue;
    }
    const mode = w.volumeMode === "formula" ? "formula" : "direct";
    w.volumeMode = mode;
    if (mode === "formula") {
      const f = Number(w.volumeFactor);
      const factor = Number.isFinite(f) ? f : 1;
      w.volumeFactor = factor;
      const base = computePastMean(effective, i, 4);
      const out = formatVolumeHrs(base * factor);
      w.volumeHrs = out;
      effective[i] = Number(out) || 0;
    } else {
      if (!Number.isFinite(Number(w.volumeFactor))) w.volumeFactor = 1;
      effective[i] = Number(w.volumeHrs) || 0;
      if (typeof w.volumeHrs !== "string") w.volumeHrs = w.volumeHrs ? String(w.volumeHrs) : "";
    }
  }
}

function seededRandom(seed) {
  let x = Math.sin(seed) * 10000;
  return () => {
    x = Math.sin(x) * 10000;
    return x - Math.floor(x);
  };
}

const PACE_TEST_OPTIONS = [
  { id: "1mi", label: "1 英里", meters: 1609.344 },
  { id: "3k", label: "3 公里", meters: 3000 },
  { id: "5k", label: "5 公里", meters: 5000 },
  { id: "10k", label: "10 公里", meters: 10000 },
  { id: "hm", label: "半馬（21.0975 公里）", meters: 21097.5 },
  { id: "m", label: "全馬（42.195 公里）", meters: 42195 },
];

const ACTIVE_TAB_KEY = "activeTab";

function parseHmsToSeconds(raw) {
  if (typeof raw !== "string") return null;
  const s = raw.trim();
  if (!s) return null;
  const parts = s.split(":").map((x) => x.trim());
  if (parts.some((x) => !/^\d+$/.test(x))) return null;
  const nums = parts.map((x) => Number(x));
  if (nums.some((x) => !Number.isFinite(x))) return null;
  if (nums.length === 3) {
    const [h, m, sec] = nums;
    return h * 3600 + m * 60 + sec;
  }
  if (nums.length === 2) {
    const [m, sec] = nums;
    return m * 60 + sec;
  }
  if (nums.length === 1) return nums[0];
  return null;
}

function getTabKeyFromHash() {
  const h = String(window.location.hash || "").trim();
  if (!h) return "";
  return h.replace(/^#/, "").trim();
}

function activateTab(nextKey, options) {
  const tabs = Array.from(document.querySelectorAll(".tab"));
  const keys = tabs
    .map((t) => String(t.getAttribute("data-tab") || "").trim())
    .filter(Boolean);
  const key = keys.includes(String(nextKey || "")) ? String(nextKey || "") : keys[0] || "plan";

  tabs.forEach((t) => {
    const k = String(t.getAttribute("data-tab") || "").trim();
    t.classList.toggle("is-active", k === key);
  });

  document.querySelectorAll(".tabPanel").forEach((p) => p.classList.remove("is-active"));
  const panel = document.getElementById(`tab-${key}`);
  if (panel) panel.classList.add("is-active");

  if (options?.persist) {
    try {
      localStorage.setItem(ACTIVE_TAB_KEY, key);
    } catch {}
  }

  if (options?.updateHash) {
    const nextHash = `#${key}`;
    if (window.location.hash !== nextHash) window.location.hash = nextHash;
  }

  if (key === "charts") renderCharts();
}

function readPaceTestTimeSeconds() {
  const hEl = document.getElementById("paceTestHours");
  const mEl = document.getElementById("paceTestMinutes");
  const sEl = document.getElementById("paceTestSeconds");
  if (hEl || mEl || sEl) {
    const hs = String(hEl?.value ?? "").trim();
    const ms = String(mEl?.value ?? "").trim();
    const ss = String(sEl?.value ?? "").trim();
    if (!hs && !ms && !ss) return null;

    const h = hs ? Number(hs) : 0;
    const m = ms ? Number(ms) : 0;
    const sec = ss ? Number(ss) : 0;
    if (![h, m, sec].every((x) => Number.isFinite(x))) return null;
    if (h < 0) return null;
    if (m < 0 || m > 59) return null;
    if (sec < 0 || sec > 59) return null;
    return Math.floor(h) * 3600 + Math.floor(m) * 60 + Math.floor(sec);
  }

  const timeInput = document.getElementById("paceTestTime");
  if (!timeInput) return null;
  return parseHmsToSeconds(String(timeInput.value || ""));
}

function formatHms(seconds) {
  const total = Math.max(0, Math.round(Number(seconds) || 0));
  const h = Math.floor(total / 3600);
  const m = Math.floor((total % 3600) / 60);
  const s = total % 60;
  return `${pad2(h)}:${pad2(m)}:${pad2(s)}`;
}

function formatPaceFromSecondsPerKm(secondsPerKm) {
  const s = Math.max(0, Math.round(Number(secondsPerKm) || 0));
  const m = Math.floor(s / 60);
  const r = s % 60;
  return `${m}:${pad2(r)}`;
}

function vdotFromRace(distanceMeters, timeSeconds) {
  const dist = Number(distanceMeters);
  const timeSec = Number(timeSeconds);
  if (!Number.isFinite(dist) || dist <= 0) return null;
  if (!Number.isFinite(timeSec) || timeSec <= 0) return null;

  const t = timeSec / 60;
  const v = dist / t;
  const vo2 = -4.6 + 0.182258 * v + 0.000104 * v * v;
  const pct = 0.8 + 0.1894393 * Math.exp(-0.012778 * t) + 0.2989558 * Math.exp(-0.1932605 * t);
  if (!Number.isFinite(vo2) || !Number.isFinite(pct) || pct <= 0) return null;
  return vo2 / pct;
}

function solveRaceTimeSeconds(distanceMeters, vdot) {
  const dist = Number(distanceMeters);
  const target = Number(vdot);
  if (!Number.isFinite(dist) || dist <= 0) return null;
  if (!Number.isFinite(target) || target <= 0) return null;

  const calc = (tSec) => vdotFromRace(dist, tSec) || 0;
  let lo = 30;
  let hi = 600;
  while (hi < 60 * 60 * 12 && calc(hi) > target) hi = Math.round(hi * 1.35);
  if (calc(hi) > target) return null;

  for (let i = 0; i < 64; i++) {
    const mid = (lo + hi) / 2;
    const v = calc(mid);
    if (v > target) lo = mid;
    else hi = mid;
  }
  return hi;
}

function speedMetersPerMinuteFromVo2(vo2) {
  const y = Number(vo2);
  if (!Number.isFinite(y)) return null;
  const a = 0.000104;
  const b = 0.182258;
  const c = -4.6 - y;
  const disc = b * b - 4 * a * c;
  if (!Number.isFinite(disc) || disc < 0) return null;
  const v = (-b + Math.sqrt(disc)) / (2 * a);
  return Number.isFinite(v) && v > 0 ? v : null;
}

function paceSecondsPerKmFromVdotFraction(vdot, fraction) {
  const f = Number(fraction);
  const base = Number(vdot);
  if (!Number.isFinite(f) || f <= 0) return null;
  if (!Number.isFinite(base) || base <= 0) return null;
  const v = speedMetersPerMinuteFromVo2(base * f);
  if (!Number.isFinite(v) || v <= 0) return null;
  return (60 * 1000) / v;
}

function renderPaceTable(root, rows) {
  if (!root) return;
  root.replaceChildren();
  const header = el("div", "paceTableRow paceTableRow--header");
  header.appendChild(el("div", "paceTableCell", "項目"));
  header.appendChild(el("div", "paceTableCell paceTableCell--right", "結果"));
  root.appendChild(header);
  rows.forEach((r) => {
    const row = el("div", "paceTableRow");
    row.appendChild(el("div", "paceTableCell", r.label));
    row.appendChild(el("div", "paceTableCell paceTableCell--right", r.value));
    root.appendChild(row);
  });
}

function computeAndRenderPaceCalculator() {
  const distSelect = document.getElementById("paceTestDistance");
  const meta = document.getElementById("paceMeta");
  const zonesRoot = document.getElementById("paceZones");
  const predsRoot = document.getElementById("pacePredictions");
  if (!distSelect || !meta || !zonesRoot || !predsRoot) return;

  const distMeters = Number(distSelect.value);
  const tSec = readPaceTestTimeSeconds();
  if (!Number.isFinite(distMeters) || distMeters <= 0 || !Number.isFinite(tSec) || tSec <= 0) {
    showToast("請輸入有效的測試距離及時間", { variant: "warn", durationMs: 1800 });
    return;
  }

  const vdot = vdotFromRace(distMeters, tSec);
  if (!Number.isFinite(vdot) || vdot <= 0) {
    showToast("無法計算（請檢查時間格式）", { variant: "warn", durationMs: 1800 });
    return;
  }

  const testPaceSecPerKm = tSec / (distMeters / 1000);
  meta.textContent = `測試配速：${formatPaceFromSecondsPerKm(testPaceSecPerKm)} / 公里 · VDOT：${vdot.toFixed(1)}`;

  const zoneDefs = [
    { label: "Z1 恢復", lo: 0.59, hi: 0.69 },
    { label: "Z2 有氧", lo: 0.69, hi: 0.78 },
    { label: "Z3 節奏", lo: 0.78, hi: 0.86 },
    { label: "Z4 乳酸", lo: 0.86, hi: 0.92 },
    { label: "Z5 最大攝氧量", lo: 0.92, hi: 0.99 },
    { label: "Z6 無氧", lo: 0.99, hi: 1.06 },
  ];

  const zoneRows = zoneDefs.map((z) => {
    const slow = paceSecondsPerKmFromVdotFraction(vdot, z.lo);
    const fast = paceSecondsPerKmFromVdotFraction(vdot, z.hi);
    if (!Number.isFinite(slow) || !Number.isFinite(fast)) return { label: z.label, value: "—" };
    const a = formatPaceFromSecondsPerKm(Math.max(slow, fast));
    const b = formatPaceFromSecondsPerKm(Math.min(slow, fast));
    return { label: z.label, value: `${a}–${b} / 公里` };
  });
  renderPaceTable(zonesRoot, zoneRows);

  const predDists = [
    { label: "5 公里", meters: 5000 },
    { label: "10 公里", meters: 10000 },
    { label: "半馬", meters: 21097.5 },
    { label: "全馬", meters: 42195 },
  ];

  const predRows = predDists.map((d) => {
    const sec = solveRaceTimeSeconds(d.meters, vdot);
    if (!Number.isFinite(sec) || sec <= 0) return { label: d.label, value: "—" };
    const pace = sec / (d.meters / 1000);
    return { label: d.label, value: `${formatHms(sec)}（${formatPaceFromSecondsPerKm(pace)} / 公里）` };
  });
  renderPaceTable(predsRoot, predRows);
}

function resetPaceCalculator() {
  const distSelect = document.getElementById("paceTestDistance");
  const hEl = document.getElementById("paceTestHours");
  const mEl = document.getElementById("paceTestMinutes");
  const sEl = document.getElementById("paceTestSeconds");
  const timeInput = document.getElementById("paceTestTime");
  const meta = document.getElementById("paceMeta");
  const zonesRoot = document.getElementById("paceZones");
  const predsRoot = document.getElementById("pacePredictions");
  if (distSelect) distSelect.selectedIndex = 2;
  if (hEl) hEl.value = "";
  if (mEl) mEl.value = "";
  if (sEl) sEl.value = "";
  if (timeInput) timeInput.value = "";
  if (meta) meta.textContent = "";
  if (zonesRoot) zonesRoot.replaceChildren();
  if (predsRoot) predsRoot.replaceChildren();
}

function startOfMonday(date) {
  const d = new Date(date);
  const day = d.getDay();
  const diff = (day + 6) % 7;
  d.setHours(0, 0, 0, 0);
  d.setTime(d.getTime() - diff * MS_PER_DAY);
  return d;
}

function addDays(date, days) {
  const d = new Date(date);
  d.setTime(d.getTime() + days * MS_PER_DAY);
  return d;
}

function normalizeBlockValue(value) {
  const v = (value || "").trim();
  if (v === "Taper") return "Deload";
  if (v === "Off") return "Transition";
  return v;
}

const PHASE_OPTIONS = ["Aerobic Endurance", "Tempo", "Threshold", "VO2Max", "Anaerobic", "Peaking", "Deload"];

const BLOCK_LABELS_ZH = {
  "": "—",
  Base: "基礎",
  Build: "建立",
  Peak: "巔峰",
  Deload: "減量",
  Transition: "過渡",
};

function computeCoachBlockByRules() {
  const out = new Array(52).fill("");
  const rank = { "": 0, Base: 0, Deload: 0, Build: 1, Transition: 2, Peak: 3 };

  const setBlock = (idx, block) => {
    if (!Number.isFinite(idx) || idx < 0 || idx > 51) return;
    const next = normalizeBlockValue(block || "");
    if (!Object.prototype.hasOwnProperty.call(BLOCK_LABELS_ZH, next)) return;
    const cur = out[idx] || "";
    if ((rank[next] || 0) >= (rank[cur] || 0)) out[idx] = next;
  };

  state.weeks.forEach((w) => {
    if (!w) return;
    const p = String(w.priority || "").trim().toUpperCase();
    if (p !== "A" && p !== "B") return;
    if (!Array.isArray(w.races) || w.races.length === 0) return;
    const idx = clamp(Number(w.index), 0, 51);

    if (p === "A") {
      setBlock(idx, "Peak");
      setBlock(idx - 1, "Peak");
      setBlock(idx + 1, "Transition");
      for (let d = 2; d <= 4; d++) setBlock(idx - d, "Build");
      return;
    }

    setBlock(idx, "Peak");
    for (let d = 1; d <= 3; d++) setBlock(idx - d, "Build");
  });

  let baseCount = 0;
  for (let i = 0; i < 52; i++) {
    if (out[i]) {
      baseCount = 0;
      continue;
    }
    if (baseCount < 3) {
      out[i] = "Base";
      baseCount++;
    } else {
      out[i] = "Deload";
      baseCount = 0;
    }
  }

  return out;
}

function applyCoachBlockRules() {
  if (!state || !Array.isArray(state.weeks) || state.weeks.length !== 52) return false;
  const blocks = computeCoachBlockByRules();
  let changed = false;
  for (let i = 0; i < 52; i++) {
    const w = state.weeks[i];
    if (!w) continue;
    const next = blocks[i] || "Base";
    if (normalizeBlockValue(w.block || "") !== next) {
      w.block = next;
      changed = true;
    }
  }
  return changed;
}

function phasesForRaceDistance(distanceKm, kind) {
  const d = Number(distanceKm);
  const k = String(kind || "").trim();
  if (!Number.isFinite(d) || d <= 0) return ["Tempo", "Threshold"];
  if (k === "trail") {
    if (d <= 12) return ["Threshold", "VO2Max"];
    if (d <= 25) return ["Aerobic Endurance", "Threshold"];
    if (d <= 45) return ["Aerobic Endurance", "Tempo"];
    return ["Aerobic Endurance", "Tempo"];
  }
  if (d <= 5) return ["VO2Max", "Anaerobic"];
  if (d <= 12) return ["Threshold", "VO2Max"];
  if (d <= 25) return ["Tempo", "Threshold"];
  return ["Aerobic Endurance", "Tempo"];
}

function intensityRelevanceForRace(distanceKm, kind) {
  const d = Number(distanceKm);
  const k = String(kind || "").trim();
  if (!Number.isFinite(d) || d <= 0) return ["Tempo", "Threshold", "VO2Max", "Anaerobic"];
  if (k === "trail") {
    if (d <= 12) return ["Threshold", "VO2Max", "Tempo", "Anaerobic"];
    if (d <= 25) return ["Threshold", "Tempo", "VO2Max", "Anaerobic"];
    return ["Tempo", "Threshold", "VO2Max", "Anaerobic"];
  }
  if (d <= 5) return ["VO2Max", "Anaerobic", "Threshold", "Tempo"];
  if (d <= 12) return ["Threshold", "VO2Max", "Tempo", "Anaerobic"];
  if (d <= 25) return ["Tempo", "Threshold", "VO2Max", "Anaerobic"];
  return ["Tempo", "Threshold", "VO2Max", "Anaerobic"];
}

function computeCoachPhasesByRules() {
  const out = new Array(52).fill(null).map(() => []);

  const nextTargetRaceByWeek = new Array(52).fill(null);
  const nextTargetRaceWeekIndex = new Array(52).fill(null);
  for (let i = 0; i < 52; i++) {
    for (let j = i; j < 52; j++) {
      const w = state.weeks[j];
      if (!w) continue;
      const pr = String(w.priority || "").trim().toUpperCase();
      if (pr !== "A" && pr !== "B") continue;
      const races = Array.isArray(w.races) ? w.races : [];
      const candidates = races
        .map((r) => {
          const date = String(r?.date || "").trim();
          const dist = Number(r?.distanceKm);
          const distanceKm = Number.isFinite(dist) && dist > 0 ? dist : null;
          const kind = typeof r?.kind === "string" ? r.kind : "";
          return { date, distanceKm, kind };
        })
        .filter((r) => r.date && r.distanceKm);
      if (!candidates.length) continue;
      candidates.sort((a, b) => a.date.localeCompare(b.date));
      nextTargetRaceByWeek[i] = candidates[0];
      nextTargetRaceWeekIndex[i] = j;
      break;
    }
  }

  for (let i = 0; i < 52; i++) {
    const w = state.weeks[i];
    const block = normalizeBlockValue(w?.block || "");
    if (block === "Base") {
      const race = nextTargetRaceByWeek[i];
      const raceWeekIdx = nextTargetRaceWeekIndex[i];
      const buildPhases = normalizePhases(phasesForRaceDistance(race?.distanceKm, race?.kind));
      const relevance = intensityRelevanceForRace(race?.distanceKm, race?.kind);
      const candidates = relevance.filter((p) => !buildPhases.includes(p));
      const weeksToRace = Number.isFinite(raceWeekIdx) ? Math.max(0, raceWeekIdx - i) : null;
      const pickIdx = Number.isFinite(weeksToRace) ? clamp(Math.floor(weeksToRace / 4), 0, candidates.length - 1) : 0;
      const extra = candidates[pickIdx] || candidates[0] || "";
      out[i] = extra ? ["Aerobic Endurance", extra] : ["Aerobic Endurance"];
      continue;
    }
    if (block === "Deload") {
      out[i] = ["Deload"];
      continue;
    }
    if (block === "Peak") {
      out[i] = ["Peaking"];
      continue;
    }
    if (block === "Build") {
      const race = nextTargetRaceByWeek[i];
      out[i] = phasesForRaceDistance(race?.distanceKm, race?.kind);
      continue;
    }
    out[i] = [];
  }

  return out.map((p) => normalizePhases(p));
}

function applyCoachPhaseRules() {
  if (!state || !Array.isArray(state.weeks) || state.weeks.length !== 52) return false;
  const phases = computeCoachPhasesByRules();
  let changed = false;
  for (let i = 0; i < 52; i++) {
    const w = state.weeks[i];
    if (!w) continue;
    const cur = normalizePhases(w.phases);
    const next = normalizePhases(phases[i]);
    if (cur.length !== next.length || cur.some((x, idx) => x !== next[idx])) {
      w.phases = next;
      changed = true;
    }
  }
  return changed;
}

const AUTO_NOTE_START = "\u001F";
const AUTO_NOTE_END = "\u001E";
const LEGACY_AUTO_NOTE_START = "【自動課表】";
const LEGACY_AUTO_NOTE_END = "【/自動課表】";

function splitAutoNote(raw) {
  const note = typeof raw === "string" ? raw : "";
  const start = note.indexOf(AUTO_NOTE_START);
  const end = note.indexOf(AUTO_NOTE_END);
  if (start >= 0 && end >= 0 && end > start) {
    const prefix = note.slice(0, start);
    const auto = note.slice(start + AUTO_NOTE_START.length, end);
    const suffix = note.slice(end + AUTO_NOTE_END.length);
    return { has: true, legacy: false, prefix, auto, suffix };
  }

  const legacyStart = note.indexOf(LEGACY_AUTO_NOTE_START);
  const legacyEnd = note.indexOf(LEGACY_AUTO_NOTE_END);
  if (legacyStart < 0 || legacyEnd < 0 || legacyEnd < legacyStart) return { has: false, legacy: false, prefix: note, auto: "", suffix: "" };
  const prefix = note.slice(0, legacyStart);
  const auto = note.slice(legacyStart + LEGACY_AUTO_NOTE_START.length, legacyEnd);
  const suffix = note.slice(legacyEnd + LEGACY_AUTO_NOTE_END.length);
  return { has: true, legacy: true, prefix, auto, suffix };
}

function mergeAutoNote(raw, autoBody) {
  const parts = splitAutoNote(raw);
  const body = typeof autoBody === "string" ? autoBody.trim() : "";
  const wrapped = body ? `${AUTO_NOTE_START}${body}${AUTO_NOTE_END}` : "";
  if (!parts.has) {
    const base = typeof raw === "string" ? raw.trimEnd() : "";
    if (!wrapped) return base.trim();
    return base ? `${base}\n\n${wrapped}` : wrapped;
  }
  const prefix = parts.prefix || "";
  const suffix = parts.suffix || "";
  if (!wrapped) {
    const out = `${prefix.trimEnd()}${suffix.trimStart() ? `\n\n${suffix.trimStart()}` : ""}`;
    return out.trim();
  }
  const out = `${prefix.trimEnd()}${prefix.trimEnd() ? "\n\n" : ""}${wrapped}${suffix.trimStart() ? `\n\n${suffix.trimStart()}` : ""}`;
  return out.trimEnd();
}

function isAutoOnlyNote(raw) {
  const p = splitAutoNote(raw);
  if (!p.has) return false;
  return `${p.prefix || ""}${p.suffix || ""}`.trim().length === 0;
}

function plannedMinutesForWeek(week) {
  const v = Number(week?.volumeHrs);
  if (!Number.isFinite(v) || v <= 0) return 0;
  return Math.max(0, Math.round(v * 60));
}

function chronicAvgDailyLoadForWeekIndex(weekIndex) {
  const idx = clamp(Number(weekIndex) || 0, 0, 51);
  if (idx <= 0) return null;
  let total = 0;
  let days = 0;
  for (let k = 1; k <= 4; k++) {
    const w = state.weeks[idx - k];
    if (!w) break;
    const sessions = getWeekSessions(w);
    sessions.forEach((s) => {
      const t = getSessionTotals(s);
      total += t.load;
      days += 1;
    });
  }
  if (days <= 0) return null;
  const out = total / days;
  return Number.isFinite(out) && out > 0 ? out : null;
}

function monotonyFromDailyLoads(dailyLoads) {
  const loads = Array.isArray(dailyLoads) ? dailyLoads.map((x) => (Number.isFinite(Number(x)) ? Number(x) : 0)) : [];
  const n = 7;
  while (loads.length < n) loads.push(0);
  if (loads.length > n) loads.length = n;
  const total = loads.reduce((a, b) => a + b, 0);
  const mean = total / n;
  const variance = loads.reduce((sum, x) => sum + (x - mean) ** 2, 0) / n;
  const stdDev = Math.sqrt(variance);
  if (stdDev <= 0) return null;
  return mean / stdDev;
}

function loadMetricsFromPlans(plans, chronicAvgDaily) {
  const dailyLoads = new Array(7).fill(0);
  for (let i = 0; i < 7; i++) {
    const p = plans?.[i];
    const minutes = Math.max(0, Math.round(Number(p?.minutes) || 0));
    const rpe = clamp(Number(p?.rpeOverride) || 1, 1, 10);
    dailyLoads[i] = minutes > 0 ? minutes * rpe : 0;
  }
  const meanLoad = dailyLoads.reduce((a, b) => a + b, 0) / 7;
  const monotony = monotonyFromDailyLoads(dailyLoads);
  const chronic = Number(chronicAvgDaily);
  const acwr = Number.isFinite(chronic) && chronic > 0 ? meanLoad / chronic : null;
  return { dailyLoads, meanLoad, monotony, acwr };
}

function rpeBoundsForPlan(plan, block) {
  const t = String(plan?.type || "").trim();
  const b = String(block || "").trim();
  if (t === "Rest") return { min: 1, max: 1 };
  if (t === "Race") return { min: 9, max: 10 };
  if (t === "Easy" || t === "Long") {
    if (b === "Deload" || b === "Transition") return { min: 2, max: 4 };
    return { min: 3, max: 4 };
  }
  const phase = String(plan?.phase || "").trim();
  if (phase === "Tempo") return { min: 5, max: 6 };
  if (phase === "Threshold") return { min: 7, max: 8 };
  if (phase === "VO2Max") return { min: 8, max: 9 };
  if (phase === "Anaerobic") return { min: 9, max: 10 };
  return { min: 5, max: 8 };
}

function defaultRpeForPlan(plan, block) {
  const t = String(plan?.type || "").trim();
  if (t === "Rest") return 1;
  if (t === "Race") return 9;
  const b = String(block || "").trim();
  if (t === "Easy" || t === "Long") return b === "Deload" || b === "Transition" ? 2 : 3;
  const phase = String(plan?.phase || "").trim();
  if (phase === "Tempo") return 5;
  if (phase === "Threshold") return 7;
  if (phase === "VO2Max") return 8;
  if (phase === "Anaerobic") return 9;
  return 6;
}

function minuteCapsForPlan(plan, options) {
  const t = String(plan?.type || "").trim();
  const minLong = Number.isFinite(options?.minLong) ? options.minLong : 30;
  const maxLong = Number.isFinite(options?.maxLong) ? options.maxLong : 180;
  const minEasy = Number.isFinite(options?.minEasy) ? options.minEasy : 0;
  const maxEasy = Number.isFinite(options?.maxEasy) ? options.maxEasy : 120;
  const minQuality = Number.isFinite(options?.minQuality) ? options.minQuality : 30;
  const maxQuality = Number.isFinite(options?.maxQuality) ? options.maxQuality : 110;
  if (t === "Long") return { min: minLong, max: maxLong };
  if (t === "Easy") return { min: minEasy, max: maxEasy };
  if (t === "Quality") return { min: minQuality, max: maxQuality };
  if (t === "Race") return { min: 30, max: 240 };
  return { min: 0, max: 0 };
}

function applyLoadConstraintsToPlans(weekIndex, plans, ctx) {
  const idx = clamp(Number(weekIndex) || 0, 0, 51);
  const block = String(ctx?.block || "").trim();
  const isRaceWeek = Boolean(ctx?.isRaceWeek);
  const raceDay = Number.isFinite(Number(ctx?.raceDay)) ? clamp(Number(ctx.raceDay), 0, 6) : null;
  const options = ctx?.minuteOptions || {};

  const base = (Array.isArray(plans) ? plans : []).slice(0, 7).map((p) => ({ ...p, rpeOverride: defaultRpeForPlan(p, block) }));
  while (base.length < 7) base.push({ type: "Rest", minutes: 0, phase: "", race: null, rpeOverride: 1 });

  const chronic = chronicAvgDailyLoadForWeekIndex(idx);
  const chronicFallback = (() => {
    if (idx <= 0) return null;
    const m = computeWeekMetrics(idx - 1);
    return Number.isFinite(m?.meanLoad) && m.meanLoad > 0 ? m.meanLoad : null;
  })();
  const chronicDaily = chronic || chronicFallback || null;

  const limitMonotony = isRaceWeek ? 1 : 2;

  const clampMinutesByCaps = () => {
    for (let i = 0; i < 7; i++) {
      const p = base[i];
      const cap = minuteCapsForPlan(p, options);
      p.minutes = clamp(Math.round(Number(p.minutes) || 0), cap.min, cap.max);
      const b = rpeBoundsForPlan(p, block);
      p.rpeOverride = clamp(Number(p.rpeOverride) || b.min, b.min, b.max);
    }
  };

  const compute = () => loadMetricsFromPlans(base, chronicDaily);

  const pickRestCandidate = () => {
    let best = -1;
    let bestScore = Infinity;
    for (let i = 0; i < 7; i++) {
      const p = base[i];
      if (i === raceDay) continue;
      if (String(p.type || "") !== "Easy") continue;
      const minutes = Math.max(0, Math.round(Number(p.minutes) || 0));
      if (!minutes) continue;
      const rpe = clamp(Number(p.rpeOverride) || 1, 1, 10);
      const score = minutes * rpe;
      if (score < bestScore) {
        bestScore = score;
        best = i;
      }
    }
    return best;
  };

  const redistributeMinutes = (fromIndex, minutes) => {
    let remaining = Math.max(0, Math.round(Number(minutes) || 0));
    if (!remaining) return;
    const candidates = [];
    for (let i = 0; i < 7; i++) {
      if (i === fromIndex) continue;
      const p = base[i];
      if (String(p.type || "") === "Rest") continue;
      candidates.push(i);
    }
    candidates.sort((a, b) => {
      const pa = base[a];
      const pb = base[b];
      const ta = String(pa.type || "");
      const tb = String(pb.type || "");
      const wA = ta === "Long" ? 0 : ta === "Quality" ? 1 : ta === "Easy" ? 2 : 3;
      const wB = tb === "Long" ? 0 : tb === "Quality" ? 1 : tb === "Easy" ? 2 : 3;
      return wA - wB;
    });
    for (const i of candidates) {
      if (remaining <= 0) break;
      const p = base[i];
      const cap = minuteCapsForPlan(p, options);
      const cur = Math.max(0, Math.round(Number(p.minutes) || 0));
      const room = Math.max(0, cap.max - cur);
      const take = Math.min(room, remaining);
      if (take > 0) {
        p.minutes = cur + take;
        remaining -= take;
      }
    }
  };

  const forceLowRpeWherePossible = () => {
    for (let i = 0; i < 7; i++) {
      const p = base[i];
      const b = rpeBoundsForPlan(p, block);
      p.rpeOverride = clamp(Number(p.rpeOverride) || b.min, b.min, b.max);
      p.rpeOverride = b.min;
    }
  };

  const scaleDownMinutes = (factor) => {
    const f = clamp(Number(factor) || 1, 0.1, 1);
    for (let i = 0; i < 7; i++) {
      const p = base[i];
      if (String(p.type || "") === "Rest") continue;
      if (isRaceWeek && i === raceDay) continue;
      const cur = Math.max(0, Math.round(Number(p.minutes) || 0));
      p.minutes = Math.max(0, Math.round(cur * f));
    }
    clampMinutesByCaps();
  };

  const updateVolumeOverride = () => {
    const totalMinutes = base.reduce((sum, p) => sum + Math.max(0, Math.round(Number(p.minutes) || 0)), 0);
    const hrs = totalMinutes > 0 ? formatVolumeHrs(totalMinutes / 60) : "";
    return hrs;
  };

  if (isRaceWeek && Number.isFinite(raceDay)) {
    const easyCapOptions = { ...options, maxEasy: 90, minEasy: 0 };

    for (let i = 0; i < 7; i++) {
      if (i === raceDay) continue;
      base[i] = { type: "Rest", minutes: 0, phase: "", race: null, rpeOverride: 1 };
    }
    base[raceDay].type = "Race";
    base[raceDay].rpeOverride = 9;
    clampMinutesByCaps();

    const matchRaceWeekTotalLoad = (maxIter) => {
      if (!chronicDaily) return;
      const desiredTotalLoad = chronicDaily * 7;
      for (let guard = 0; guard < maxIter; guard++) {
        const m = compute();
        const totalLoad = m.dailyLoads.reduce((a, b) => a + b, 0);
        if (Math.abs(totalLoad - desiredTotalLoad) < Math.max(6, desiredTotalLoad * 0.01)) break;

        if (totalLoad > desiredTotalLoad) {
          const p = base[raceDay];
          const cap = minuteCapsForPlan(p, options);
          const rpe = clamp(Number(p.rpeOverride) || 9, 9, 10);
          const targetMinutes = Math.floor(desiredTotalLoad / rpe);
          p.minutes = clamp(targetMinutes, cap.min, Math.max(cap.min, Math.round(Number(p.minutes) || 0)));
          clampMinutesByCaps();
          continue;
        }

        const deficit = Math.max(0, desiredTotalLoad - totalLoad);
        const easyDays = [2, 4, 1, 3, 0, 6, 5].filter((d) => d !== raceDay);
        let remaining = deficit;
        for (const d of easyDays) {
          if (remaining <= 0) break;
          const p = base[d];
          p.type = "Easy";
          p.phase = "";
          p.race = null;
          const b = rpeBoundsForPlan(p, block === "Deload" ? "Deload" : "Transition");
          p.rpeOverride = b.min;
          const cap = minuteCapsForPlan(p, easyCapOptions);
          const want = Math.min(cap.max, Math.max(0, Math.round(remaining / p.rpeOverride)));
          p.minutes = clamp((Number(p.minutes) || 0) + want, cap.min, cap.max);
          remaining = Math.max(0, desiredTotalLoad - compute().dailyLoads.reduce((a, b2) => a + b2, 0));
        }
        clampMinutesByCaps();
      }
    };

    matchRaceWeekTotalLoad(12);

    for (let iter = 0; iter < 6; iter++) {
      const m = compute();
      if (m.monotony !== null && m.monotony < limitMonotony) break;
      const extra = pickRestCandidate();
      if (extra < 0) break;

      const moved = Math.max(0, Math.round(Number(base[extra].minutes) || 0));
      base[extra] = { type: "Rest", minutes: 0, phase: "", race: null, rpeOverride: 1 };

      let receiver = -1;
      let best = -1;
      for (let i = 0; i < 7; i++) {
        if (i === extra || i === raceDay) continue;
        if (String(base[i]?.type || "") !== "Easy") continue;
        const mins = Math.max(0, Math.round(Number(base[i]?.minutes) || 0));
        if (mins > best) {
          best = mins;
          receiver = i;
        }
      }

      if (receiver >= 0 && moved > 0) {
        const p = base[receiver];
        const cap = minuteCapsForPlan(p, easyCapOptions);
        const cur = Math.max(0, Math.round(Number(p.minutes) || 0));
        const room = Math.max(0, cap.max - cur);
        const take = Math.min(room, moved);
        if (take > 0) p.minutes = cur + take;
        const remaining = moved - take;
        if (remaining > 0) redistributeMinutes(extra, remaining);
      } else {
        redistributeMinutes(extra, moved);
      }

      clampMinutesByCaps();
      matchRaceWeekTotalLoad(8);
    }

    return { plans: base, volumeHrsOverride: updateVolumeOverride() };
  }

  if (block === "Deload") {
    for (let i = 0; i < 7; i++) {
      const p = base[i];
      if (String(p.type || "") === "Quality" || String(p.type || "") === "Long") {
        base[i] = { type: "Rest", minutes: 0, phase: "", race: null, rpeOverride: 1 };
      }
    }
    for (let i = 0; i < 7; i++) {
      const p = base[i];
      if (String(p.type || "") === "Easy") p.rpeOverride = 2;
    }
    clampMinutesByCaps();
  }

  let changed = false;
  for (let iter = 0; iter < 10; iter++) {
    const m = compute();
    const needMonotonyFix = m.monotony === null || m.monotony >= limitMonotony;
    if (!needMonotonyFix) break;

    const restIdx = pickRestCandidate();
    if (restIdx < 0) break;
    const moved = Math.max(0, Math.round(Number(base[restIdx].minutes) || 0));
    base[restIdx] = { type: "Rest", minutes: 0, phase: "", race: null, rpeOverride: 1 };
    redistributeMinutes(restIdx, moved);
    clampMinutesByCaps();
    changed = true;
  }

  if (chronicDaily) {
    for (let iter = 0; iter < 8; iter++) {
      const m = compute();
      if (m.acwr === null) break;
      if (block === "Deload") {
        if (m.acwr < 0.8) break;
      } else {
        if (m.acwr < 1.5) break;
      }

      forceLowRpeWherePossible();
      clampMinutesByCaps();
      const m2 = compute();
      if (m2.acwr === null) break;
      const cap = block === "Deload" ? 0.78 : 1.45;
      if (m2.acwr < (block === "Deload" ? 0.8 : 1.5)) break;
      const totalLoad = m2.dailyLoads.reduce((a, b) => a + b, 0);
      const desiredTotalLoad = chronicDaily * 7 * cap;
      const factor = totalLoad > 0 ? desiredTotalLoad / totalLoad : 1;
      if (factor >= 1) break;
      scaleDownMinutes(factor);
      changed = true;
    }
  }

  const final = compute();
  const needVolumeOverride =
    (block === "Deload" && chronicDaily && final.acwr !== null && final.acwr >= 0.8) ||
    (chronicDaily && final.acwr !== null && final.acwr >= 1.5);
  const volumeHrsOverride = changed || needVolumeOverride ? updateVolumeOverride() : null;
  return { plans: base, volumeHrsOverride };
}

function raceEntriesForWeek(week) {
  if (!week) return [];
  const monday = week.monday instanceof Date ? week.monday : null;
  if (!monday) return [];
  const races = Array.isArray(week.races) ? week.races : [];
  const out = [];
  for (const r of races) {
    const name = String(r?.name || "").trim();
    const date = String(r?.date || "").trim();
    if (!name || !date) continue;
    const d = parseYMD(date);
    if (!d) continue;
    const dayIndex = Math.round((d.getTime() - monday.getTime()) / MS_PER_DAY);
    if (!Number.isFinite(dayIndex) || dayIndex < 0 || dayIndex > 6) continue;
    const dist = Number(r?.distanceKm);
    const distanceKm = Number.isFinite(dist) && dist > 0 ? dist : null;
    const kind = typeof r?.kind === "string" ? r.kind : "";
    out.push({ name, date, dayIndex, distanceKm, kind });
  }
  out.sort((a, b) => a.date.localeCompare(b.date) || a.name.localeCompare(b.name));
  return out;
}

function nextTargetRaceFromWeekIndex(fromWeekIndex) {
  const start = clamp(Number(fromWeekIndex) || 0, 0, 51);
  for (let j = start; j < 52; j++) {
    const w = state.weeks[j];
    if (!w) continue;
    const pr = String(w.priority || "").trim().toUpperCase();
    if (pr !== "A" && pr !== "B") continue;
    const races = Array.isArray(w.races) ? w.races : [];
    const candidates = races
      .map((r) => {
        const date = String(r?.date || "").trim();
        const dist = Number(r?.distanceKm);
        const distanceKm = Number.isFinite(dist) && dist > 0 ? dist : null;
        const kind = typeof r?.kind === "string" ? r.kind : "";
        return { date, distanceKm, kind };
      })
      .filter((r) => r.date && r.distanceKm);
    if (!candidates.length) continue;
    candidates.sort((a, b) => a.date.localeCompare(b.date));
    return { weekIndex: j, race: candidates[0] };
  }
  return null;
}

function phaseStreakIndex(weekIndex, phase) {
  const p = String(phase || "").trim();
  if (!PHASE_OPTIONS.includes(p)) return 0;
  let streak = 0;
  for (let i = clamp(weekIndex, 0, 51); i >= 0; i--) {
    const w = state.weeks[i];
    if (!w) break;
    const phases = normalizePhases(w.phases);
    if (!phases.includes(p)) break;
    streak++;
  }
  return Math.max(0, streak - 1);
}

function estimateRaceMinutes(distanceKm, kind) {
  const d = Number(distanceKm);
  const k = String(kind || "").trim();
  if (!Number.isFinite(d) || d <= 0) return 75;
  const minPerKm = k === "trail" ? 9 : 6.5;
  const raw = d * minPerKm;
  return clamp(Math.round(raw), 20, 240);
}

function buildAutoNoteBodyForPlan(plan) {
  const lines = [];
  if (plan?.title) lines.push(String(plan.title));
  if (plan?.details && Array.isArray(plan.details)) {
    plan.details.forEach((t) => {
      const s = String(t || "").trim();
      if (s) lines.push(s);
    });
  }
  if (plan?.minutes) lines.push(`主課時長：${Math.round(Number(plan.minutes) || 0)} 分鐘`);
  if (plan?.rpeText) lines.push(`RPE：${plan.rpeText}`);
  return lines.join("\n").trim();
}

function sessionPlanForDay(weekIndex, day, ctx) {
  const minutes = Math.max(0, Math.round(Number(day?.minutes) || 0));
  const type = String(day?.type || "").trim();
  const phase = String(day?.phase || "").trim();
  const race = day?.race || null;
  const block = String(ctx?.block || "").trim();
  const rpeOverride = Number.isFinite(Number(day?.rpeOverride)) ? clamp(Number(day.rpeOverride), 1, 10) : null;

  if (type === "Rest" || minutes <= 0) {
    return { zone: 1, rpe: 1, workoutMinutes: 0, noteBody: buildAutoNoteBodyForPlan({ title: "休息／伸展", minutes: 0, rpeText: "—" }) };
  }

  if (type === "Race" && race) {
    const kindText = race.kind === "trail" ? "越野跑" : race.kind === "road" ? "路跑" : "";
    const distText = race.distanceKm ? `${race.distanceKm}km` : "";
    const meta = [distText, kindText].filter(Boolean).join(" · ");
    const title = meta ? `比賽：${race.name}（${meta}）` : `比賽：${race.name}`;
    const details = ["熱身 15' + 比賽 + 放鬆 10'"];
    const workoutMinutes = minutes;
    return { zone: 6, rpe: rpeOverride ?? 10, workoutMinutes, noteBody: buildAutoNoteBodyForPlan({ title, details, minutes: workoutMinutes, rpeText: "9–10" }) };
  }

  if (type === "Long") {
    return {
      zone: 2,
      rpe: rpeOverride ?? 3,
      workoutMinutes: minutes,
      noteBody: buildAutoNoteBodyForPlan({
        title: "長課（有氧耐力）",
        details: ["保持輕鬆，避免配速拉高"],
        minutes,
        rpeText: "3–4",
      }),
    };
  }

  if (type === "Easy") {
    const easyRpeText = block === "Deload" || block === "Transition" ? "2–4" : "3–4";
    return {
      zone: 2,
      rpe: rpeOverride ?? (block === "Deload" || block === "Transition" ? 2 : 3),
      workoutMinutes: minutes,
      noteBody: buildAutoNoteBodyForPlan({ title: "有氧耐力", minutes, rpeText: easyRpeText }),
    };
  }

  if (phase === "Tempo") {
    const main = clamp(minutes, 25, 70);
    const details = [`主課：節奏跑 ${main}'（可每週 +5'，逐步延長）`, "另加：熱身／放鬆"];
    return { zone: 3, rpe: rpeOverride ?? 6, workoutMinutes: minutes, noteBody: buildAutoNoteBodyForPlan({ title: "節奏", details, minutes, rpeText: "5–6" }) };
  }
  if (phase === "Threshold") {
    const idx = phaseStreakIndex(weekIndex, phase);
    const repMin = clamp(6 + Math.floor(idx / 2) * 2, 6, 12);
    const restMin = Math.max(1, Math.round(repMin / 4));
    const cycle = repMin + restMin;
    const sets = clamp(Math.max(2, Math.round(minutes / cycle)), 3, 6);
    const used = Math.max(0, sets * repMin + Math.max(0, sets - 1) * restMin);
    const extra = clamp(minutes - used, 0, 25);
    const details = [`主課：${sets} × ${repMin}'（跑/休 4:1，休 ${restMin}'）`];
    if (extra >= 5) details.push(`加量：閾值續跑 ${Math.round(extra)}'`);
    details.push("另加：熱身／放鬆");
    const workoutMinutes = minutes;
    return {
      zone: 4,
      rpe: rpeOverride ?? 8,
      workoutMinutes,
      noteBody: buildAutoNoteBodyForPlan({ title: "乳酸閾值", details, minutes: workoutMinutes, rpeText: "7–8" }),
    };
  }
  if (phase === "VO2Max") {
    const idx = phaseStreakIndex(weekIndex, phase);
    const workTarget = clamp(6 + idx * 2, 6, 18);
    const repMin = workTarget <= 8 ? 2 : workTarget <= 12 ? 3 : 4;
    const reps = Math.max(1, Math.round(workTarget / repMin));
    const details = [`主課：${reps} × ${repMin}'（跑/休 1:1）`, `強度時間：約 ${clamp(reps * repMin, 6, 18)}'`, "另加：熱身／放鬆"];
    const workoutMinutes = minutes;
    return {
      zone: 5,
      rpe: rpeOverride ?? 9,
      workoutMinutes,
      noteBody: buildAutoNoteBodyForPlan({ title: "最大攝氧量", details, minutes: workoutMinutes, rpeText: "8–9" }),
    };
  }
  if (phase === "Anaerobic") {
    const idx = phaseStreakIndex(weekIndex, phase);
    const workTarget = clamp(5 + idx * 2, 5, 16);
    const repSec = workTarget <= 8 ? 30 : 60;
    const repMin = repSec === 30 ? 0.5 : 1;
    const reps = Math.max(1, Math.round(workTarget / repMin));
    const details = [`主課：${reps} × ${repSec}s（跑/休 1:1）`, `強度時間：約 ${clamp(Math.round(reps * repMin), 5, 16)}'`, "另加：熱身／放鬆"];
    const workoutMinutes = minutes;
    return {
      zone: 6,
      rpe: rpeOverride ?? 10,
      workoutMinutes,
      noteBody: buildAutoNoteBodyForPlan({ title: "無氧", details, minutes: workoutMinutes, rpeText: "9–10" }),
    };
  }

  return { zone: 2, rpe: rpeOverride ?? 3, workoutMinutes: minutes, noteBody: buildAutoNoteBodyForPlan({ title: "有氧耐力", minutes, rpeText: "3–4" }) };
}

function clampDayPlanMinutes(plans, targetMinutes, options) {
  const target = Math.max(0, Math.round(Number(targetMinutes) || 0));
  const maxLong = Number.isFinite(options?.maxLong) ? options.maxLong : 180;
  const minLong = Number.isFinite(options?.minLong) ? options.minLong : 30;
  const maxEasy = Number.isFinite(options?.maxEasy) ? options.maxEasy : 120;
  const minEasy = Number.isFinite(options?.minEasy) ? options.minEasy : 0;
  const maxQuality = Number.isFinite(options?.maxQuality) ? options.maxQuality : 100;
  const minQuality = Number.isFinite(options?.minQuality) ? options.minQuality : 30;

  const kinds = plans.map((p) => String(p.type || ""));
  const mins = plans.map((p) => Math.max(0, Math.round(Number(p.minutes) || 0)));
  const caps = mins.map((v, i) => {
    const t = kinds[i];
    if (t === "Long") return { min: minLong, max: maxLong };
    if (t === "Easy") return { min: minEasy, max: maxEasy };
    if (t === "Quality") return { min: minQuality, max: maxQuality };
    if (t === "Race") return { min: 30, max: 240 };
    return { min: 0, max: 0 };
  });

  for (let i = 0; i < mins.length; i++) {
    mins[i] = clamp(mins[i], caps[i].min, caps[i].max);
  }

  const sum = () => mins.reduce((a, b) => a + b, 0);
  const orderReduce = ["Easy", "Long", "Quality", "Race"];
  const orderAdd = ["Long", "Easy", "Quality", "Race"];

  let cur = sum();
  if (cur > target) {
    let diff = cur - target;
    for (const k of orderReduce) {
      if (diff <= 0) break;
      for (let i = 0; i < mins.length; i++) {
        if (diff <= 0) break;
        if (kinds[i] !== k) continue;
        const lo = caps[i].min;
        const room = Math.max(0, mins[i] - lo);
        const take = Math.min(room, diff);
        mins[i] -= take;
        diff -= take;
      }
    }
  } else if (cur < target) {
    let diff = target - cur;
    for (const k of orderAdd) {
      if (diff <= 0) break;
      for (let i = 0; i < mins.length; i++) {
        if (diff <= 0) break;
        if (kinds[i] !== k) continue;
        const hi = caps[i].max;
        const room = Math.max(0, hi - mins[i]);
        const take = Math.min(room, diff);
        mins[i] += take;
        diff -= take;
      }
    }
  }

  return plans.map((p, i) => ({ ...p, minutes: mins[i] }));
}

function blockStreakIndex(weekIndex, block) {
  const b = normalizeBlockValue(block || "");
  if (!b) return 0;
  let streak = 0;
  for (let i = clamp(Number(weekIndex) || 0, 0, 51); i >= 0; i--) {
    const w = state.weeks[i];
    if (!w) break;
    const cur = normalizeBlockValue(w.block || "") || "Base";
    if (cur !== b) break;
    streak++;
  }
  return Math.max(0, streak - 1);
}

function rebalanceDayPlanMinutes(plans, targetMinutes, options, chunkSize) {
  const target = Math.max(0, Math.round(Number(targetMinutes) || 0));
  const chunk = Math.max(1, Math.round(Number(chunkSize) || 60));

  const mins = plans.map((p) => Math.max(0, Math.round(Number(p?.minutes) || 0)));
  const caps = plans.map((p) => minuteCapsForPlan(p, options));
  for (let i = 0; i < mins.length; i++) mins[i] = clamp(mins[i], caps[i].min, caps[i].max);

  const sum = () => mins.reduce((a, b) => a + b, 0);
  const typeAt = (i) => String(plans[i]?.type || "");
  const ordered = (kinds) => mins.map((_, i) => i).filter((i) => kinds.includes(typeAt(i)));

  const easyIdx = ordered(["Easy"]);
  const longIdx = ordered(["Long"]);
  const qualityIdx = ordered(["Quality"]);
  const raceIdx = ordered(["Race"]);

  const addTo = (i, amount) => {
    if (amount <= 0) return 0;
    const hi = caps[i].max;
    const room = Math.max(0, hi - mins[i]);
    const take = Math.min(room, amount);
    mins[i] += take;
    return take;
  };

  const takeFrom = (i, amount) => {
    if (amount <= 0) return 0;
    const lo = caps[i].min;
    const room = Math.max(0, mins[i] - lo);
    const take = Math.min(room, amount);
    mins[i] -= take;
    return take;
  };

  let cur = sum();
  if (cur < target) {
    let diff = target - cur;
    let chunks = Math.floor(diff / chunk);
    let rem = diff - chunks * chunk;

    if (easyIdx.length) {
      for (let k = 0; k < chunks; k++) {
        const i = easyIdx[k % easyIdx.length];
        const took = addTo(i, chunk);
        if (took < chunk) break;
      }
      cur = sum();
      diff = target - cur;
      chunks = Math.floor(diff / chunk);
      rem = diff - chunks * chunk;
      for (let k = 0; k < chunks; k++) {
        const i = easyIdx[k % easyIdx.length];
        const took = addTo(i, chunk);
        if (took < chunk) break;
      }
      if (rem > 0) addTo(easyIdx[chunks % easyIdx.length], rem);
    }

    cur = sum();
    diff = target - cur;
    if (diff > 0) {
      for (const group of [longIdx, qualityIdx, raceIdx]) {
        for (const i of group) {
          if (diff <= 0) break;
          diff -= addTo(i, diff);
        }
      }
    }
  } else if (cur > target) {
    let diff = cur - target;
    for (const group of [easyIdx, longIdx, qualityIdx, raceIdx]) {
      for (const i of group) {
        if (diff <= 0) break;
        diff -= takeFrom(i, diff);
      }
    }
  }

  return plans.map((p, i) => ({ ...p, minutes: mins[i] }));
}

function computeWeekDayPlans(weekIndex) {
  const idx = clamp(Number(weekIndex) || 0, 0, 51);
  const w = state.weeks[idx];
  if (!w) return null;

  const targetMinutes = plannedMinutesForWeek(w);
  if (!targetMinutes) return null;

  const block = normalizeBlockValue(w.block || "") || "Base";
  const phases = normalizePhases(w.phases);
  const races = raceEntriesForWeek(w);
  const pr = String(w.priority || "").trim().toUpperCase();
  const isRaceWeek = races.length > 0 && (pr === "A" || pr === "B");
  const next = nextTargetRaceFromWeekIndex(idx);
  const raceContext = isRaceWeek ? races[0] : next?.race || null;

  const intensityOrder = intensityRelevanceForRace(raceContext?.distanceKm, raceContext?.kind);
  const intensityPhases = phases.filter((p) => p === "Tempo" || p === "Threshold" || p === "VO2Max" || p === "Anaerobic");
  intensityPhases.sort((a, b) => (intensityOrder.indexOf(a) < 0 ? 99 : intensityOrder.indexOf(a)) - (intensityOrder.indexOf(b) < 0 ? 99 : intensityOrder.indexOf(b)));

  const day = new Array(7).fill(null).map(() => ({ type: "Rest", minutes: 0, phase: "", race: null }));

  const raceDay = isRaceWeek ? clamp(Number(races[0]?.dayIndex) || 0, 0, 6) : null;
  if (Number.isFinite(raceDay)) {
    day[raceDay] = { type: "Race", minutes: estimateRaceMinutes(races[0]?.distanceKm, races[0]?.kind), phase: "", race: races[0] };
  }

  const hasLong = (block === "Base" || block === "Build" || (block === "Peak" && !isRaceWeek)) && !isRaceWeek;
  const longDay = hasLong ? 5 : null;

  const isPeak = block === "Peak";
  const isRecoveryBlock = block === "Deload" || block === "Transition";

  if (isRecoveryBlock) {
    [1, 3, 5].forEach((d) => {
      if (day[d].type === "Rest") day[d] = { type: "Easy", minutes: 0, phase: "", race: null };
    });
  } else if (isPeak) {
    const q1 = intensityPhases[0] || "VO2Max";
    const q2 = intensityPhases[1] || (q1 === "VO2Max" ? "Threshold" : "VO2Max");
    const qDays = [1, 3].filter((d) => !Number.isFinite(raceDay) || d !== raceDay);
    if (qDays[0] !== undefined) day[qDays[0]] = { type: "Quality", minutes: 0, phase: q1, race: null };
    if (!isRaceWeek && qDays[1] !== undefined) day[qDays[1]] = { type: "Quality", minutes: 0, phase: q2, race: null };
    if (Number.isFinite(longDay)) {
      day[longDay] = { type: "Long", minutes: 0, phase: "", race: null };
    }
  } else if (block === "Build") {
    const q1 = intensityPhases[0] || "Tempo";
    const q2 = intensityPhases[1] || "Threshold";
    day[1] = { type: "Quality", minutes: 0, phase: q1, race: null };
    day[3] = { type: "Quality", minutes: 0, phase: q2, race: null };
    if (Number.isFinite(longDay)) {
      day[longDay] = { type: "Long", minutes: 0, phase: "", race: null };
    }
    [0, 2, 6].forEach((d) => {
      if (day[d].type === "Rest") day[d] = { type: "Easy", minutes: 0, phase: "", race: null };
    });
    if (day[4].type === "Rest") day[4] = { type: "Rest", minutes: 0, phase: "", race: null };
  } else {
    const q = intensityPhases[0] || phases.find((p) => p !== "Aerobic Endurance" && p !== "Deload" && p !== "Peaking") || "Tempo";
    day[1] = { type: "Quality", minutes: 0, phase: q, race: null };
    if (Number.isFinite(longDay)) {
      day[longDay] = { type: "Long", minutes: 0, phase: "", race: null };
    }
    [0, 2, 3, 4, 6].forEach((d) => {
      if (day[d].type === "Rest") day[d] = { type: "Easy", minutes: 0, phase: "", race: null };
    });
    day[4] = { type: "Rest", minutes: 0, phase: "", race: null };
  }

  const plans = day.map((d) => ({ ...d }));
  const options =
    block === "Peak"
      ? { minEasy: 0, maxEasy: 0, minLong: 30, maxLong: 120, minQuality: 30, maxQuality: 110 }
      : block === "Deload" || block === "Transition"
        ? { minEasy: 0, maxEasy: 90, minLong: 0, maxLong: 0, minQuality: 0, maxQuality: 0 }
        : { minEasy: 30, maxEasy: 120, minLong: 30, maxLong: 180, minQuality: 30, maxQuality: 110 };

  const basePlans = plans.map((p) => ({ ...p, minutes: Math.max(0, Math.round(Number(p.minutes) || 0)) }));
  const baseIdx = blockStreakIndex(idx, "Base");
  const buildIdx = blockStreakIndex(idx, "Build");

  const setMinutesForType = (t, minutesForOne) => {
    const want = Math.max(0, Math.round(Number(minutesForOne) || 0));
    for (let i = 0; i < basePlans.length; i++) {
      if (String(basePlans[i]?.type || "") === t) basePlans[i].minutes = want;
    }
  };

  if (block === "Base" && !isRaceWeek) {
    const q = clamp(60 + baseIdx * 5, 45, 95);
    const l = clamp(60 + baseIdx * 10, 60, 180);
    setMinutesForType("Quality", q);
    setMinutesForType("Long", l);
  } else if (block === "Build" && !isRaceWeek) {
    const q = clamp(60 + buildIdx * 5, 50, 95);
    const l = clamp(75 + buildIdx * 10, 70, 180);
    setMinutesForType("Quality", q);
    setMinutesForType("Long", l);
  } else if (block === "Peak") {
    const q = clamp(55, 40, 85);
    const l = clamp(70, 45, 120);
    setMinutesForType("Quality", q);
    if (!isRaceWeek) setMinutesForType("Long", l);
  }

  const easyIndices = basePlans.map((p, i) => ({ i, t: String(p?.type || "") })).filter((x) => x.t === "Easy").map((x) => x.i);
  const hasEasy = easyIndices.length > 0;
  const currentFixed = () => basePlans.reduce((sum, p) => sum + Math.max(0, Math.round(Number(p?.minutes) || 0)), 0);

  let diff = targetMinutes - currentFixed();
  if (hasEasy && diff > 0) {
    const chunks = Math.floor(diff / 60);
    const rem = diff - chunks * 60;
    for (let k = 0; k < chunks; k++) {
      const i = easyIndices[k % easyIndices.length];
      basePlans[i].minutes = (Number(basePlans[i].minutes) || 0) + 60;
    }
    if (rem > 0) {
      const i = easyIndices[chunks % easyIndices.length];
      basePlans[i].minutes = (Number(basePlans[i].minutes) || 0) + rem;
    }
  } else if (diff < 0) {
    const reduceOrder = ["Easy", "Long", "Quality", "Race"];
    let need = -diff;
    for (const t of reduceOrder) {
      if (need <= 0) break;
      for (let i = 0; i < basePlans.length; i++) {
        if (need <= 0) break;
        const p = basePlans[i];
        if (String(p.type || "") !== t) continue;
        const cap = minuteCapsForPlan(p, options);
        const cur = Math.max(0, Math.round(Number(p.minutes) || 0));
        const lo = cap.min;
        const room = Math.max(0, cur - lo);
        const take = Math.min(room, need);
        basePlans[i].minutes = cur - take;
        need -= take;
      }
    }
  }

  const rebalanced = rebalanceDayPlanMinutes(basePlans, targetMinutes, options, 60);
  rebalanced.forEach((p) => {
    if (String(p?.type || "") === "Easy" && Math.max(0, Math.round(Number(p?.minutes) || 0)) <= 0) {
      p.type = "Rest";
      p.minutes = 0;
      p.phase = "";
      p.race = null;
    }
  });

  const constrained = applyLoadConstraintsToPlans(idx, rebalanced, { block, isRaceWeek, raceDay, minuteOptions: options });
  return { block, targetMinutes, plans: constrained.plans, volumeHrsOverride: constrained.volumeHrsOverride };
}

function applyCoachDayPlanRules(range) {
  if (!state || !Array.isArray(state.weeks) || state.weeks.length !== 52) return false;
  const start = range && Number.isFinite(Number(range.start)) ? clamp(Number(range.start), 0, 51) : 0;
  const end = range && Number.isFinite(Number(range.end)) ? clamp(Number(range.end), 0, 51) : 51;
  let changed = false;

  for (let i = start; i <= end; i++) {
    const week = state.weeks[i];
    if (!week) continue;
    const result = computeWeekDayPlans(i);
    if (!result) continue;

    if (typeof result.volumeHrsOverride === "string" && result.volumeHrsOverride !== week.volumeHrs) {
      week.volumeMode = "direct";
      week.volumeFactor = 1;
      week.volumeHrs = result.volumeHrsOverride;
      changed = true;
    }

    const sessions = getWeekSessions(week);
    if (!Array.isArray(week.sessions) || !week.sessions.length) week.sessions = sessions;

    for (let d = 0; d < 7; d++) {
      const s = sessions[d];
      if (!s) continue;

      const allowOverwrite = isAutoOnlyNote(s.note) || (getSessionTotals(s).minutes === 0 && String(s.note || "").trim() === "");
      const plan = result.plans[d];
      const isQuality = plan.type === "Quality";
      const dayPlan = sessionPlanForDay(i, plan, { block: result.block });
      const durationMinutes = Math.max(0, Math.round(Number(dayPlan?.workoutMinutes ?? plan.minutes) || 0));

      if (allowOverwrite) {
        s.workoutsCount = 1;
        s.workouts = [{ duration: durationMinutes, rpe: dayPlan.rpe }];
        s.zone = dayPlan.zone;
        ensureSessionWorkouts(s);
        const nextNote = mergeAutoNote(s.note, dayPlan.noteBody);
        if (nextNote !== s.note) s.note = nextNote;
        changed = true;
      } else if (splitAutoNote(s.note).has) {
        const body = dayPlan.noteBody;
        const nextNote = mergeAutoNote(s.note, body);
        if (nextNote !== s.note) {
          s.note = nextNote;
          changed = true;
        }
      } else if (!String(s.note || "").trim() && isQuality) {
        const nextNote = mergeAutoNote("", dayPlan.noteBody);
        if (nextNote !== s.note) {
          s.note = nextNote;
          changed = true;
        }
      }
    }
  }

  return changed;
}

function applyCoachAutoRules() {
  const a = applyCoachBlockRules();
  const b = applyCoachPhaseRules();
  const c = applyAnnualVolumeRules();
  const d = applyCoachDayPlanRules();
  return a || b || c || d;
}

const PHASE_LABELS_ZH = {
  "Aerobic Endurance": "有氧耐力",
  Tempo: "節奏",
  Threshold: "乳酸閾值",
  VO2Max: "最大攝氧量",
  Anaerobic: "無氧",
  Peaking: "巔峰期",
  Deload: "減量",
};

function weekLabelZh(weekNo) {
  return `第${weekNo}週`;
}

function dayLabelZh(dayIndex) {
  return `第${dayIndex + 1}天`;
}

function normalizeDayLabelZh(label, fallbackIndex) {
  const raw = typeof label === "string" ? label.trim() : "";
  const m = /^Day\s*(\d+)$/i.exec(raw);
  if (m) {
    const n = Number(m[1]);
    if (Number.isFinite(n) && n >= 1 && n <= 7) return `第${n}天`;
  }
  if (!raw && Number.isFinite(fallbackIndex)) return dayLabelZh(fallbackIndex);
  return raw;
}

function phaseLabelZh(phase) {
  return PHASE_LABELS_ZH[phase] || phase || "";
}

function normalizePhases(value) {
  if (Array.isArray(value)) {
    const out = [];
    value.forEach((v) => {
      if (typeof v !== "string") return;
      if (!PHASE_OPTIONS.includes(v)) return;
      if (out.includes(v)) return;
      out.push(v);
    });
    return out;
  }
  if (typeof value === "string" && PHASE_OPTIONS.includes(value)) return [value];
  return [];
}

function normalizeWorkoutEntry(w) {
  return {
    duration: Math.max(0, Math.round(Number(w?.duration) || 0)),
    rpe: clamp(Number(w?.rpe) || 1, 1, 10),
  };
}

function ensureSessionWorkouts(session) {
  if (!session || typeof session !== "object") return;

  const fromArray = Array.isArray(session.workouts) ? session.workouts.map(normalizeWorkoutEntry) : [];
  const legacyDuration = Math.max(0, Math.round(Number(session.duration) || 0));
  const legacyRpe = clamp(Number(session.rpe) || 1, 1, 10);
  let workouts = fromArray.length ? fromArray : [{ duration: legacyDuration, rpe: legacyRpe }];

  const inferredCount = workouts.length || 1;
  const count = clamp(Number(session.workoutsCount) || inferredCount, 1, 10);

  workouts = workouts.slice(0, count);
  while (workouts.length < count) workouts.push({ duration: 0, rpe: 1 });

  session.workoutsCount = count;
  session.workouts = workouts;
  session.duration = workouts[0]?.duration || 0;
  session.rpe = workouts[0]?.rpe || 1;
}

function getSessionTotals(session) {
  ensureSessionWorkouts(session);
  const workouts = Array.isArray(session?.workouts) ? session.workouts : [];
  let minutes = 0;
  let load = 0;
  workouts.forEach((w) => {
    const duration = Math.max(0, Math.round(Number(w?.duration) || 0));
    const rpe = clamp(Number(w?.rpe) || 1, 1, 10);
    minutes += duration;
    load += duration * rpe;
  });
  return { minutes, load };
}

function buildDefaultSessions() {
  const sessions = [];
  for (let i = 0; i < 7; i++) {
    sessions.push({
      dayLabel: dayLabelZh(i),
      workoutsCount: 1,
      workouts: [{ duration: 0, rpe: 1 }],
      duration: 0,
      zone: 1,
      rpe: 1,
      kind: "Run",
      note: "",
    });
  }
  return sessions;
}

function getWeekSessions(week) {
  if (!week || typeof week !== "object") return buildDefaultSessions();
  if (!Array.isArray(week.sessions) || !week.sessions.length) {
    week.sessions = buildDefaultSessions();
  }
  const defaults = buildDefaultSessions();
  for (let i = 0; i < 7; i++) {
    if (!week.sessions[i]) week.sessions[i] = defaults[i];
    const s = week.sessions[i];
    s.dayLabel = normalizeDayLabelZh(s.dayLabel, i);
    if (typeof s.note !== "string") s.note = "";
    ensureSessionWorkouts(s);
  }
  if (week.sessions.length > 7) week.sessions.length = 7;
  return week.sessions;
}

function computeWeekMetrics(weekIndex) {
  const idx = clamp(Number(weekIndex) || 0, 0, 51);
  const week = state.weeks[idx];
  const sessions = getWeekSessions(week);

  const totals = sessions.map((s) => getSessionTotals(s));
  const dailyLoads = totals.map((t) => t.load);

  const totalMinutes = totals.reduce((sum, t) => sum + t.minutes, 0);
  const totalLoad = dailyLoads.reduce((sum, x) => sum + x, 0);

  const meanLoad = dailyLoads.length ? totalLoad / dailyLoads.length : 0;
  const variance = dailyLoads.length ? dailyLoads.reduce((sum, x) => sum + (x - meanLoad) ** 2, 0) / dailyLoads.length : 0;
  const stdDev = Math.sqrt(variance);
  const monotony = stdDev > 0 ? meanLoad / stdDev : null;

  let acwr = null;
  if (idx >= 4) {
    let chronicTotal = 0;
    let chronicDays = 0;
    for (let k = 1; k <= 4; k++) {
      const prevWeek = state.weeks[idx - k];
      const prevSessions = getWeekSessions(prevWeek);
      const prevTotals = prevSessions.map((s) => getSessionTotals(s));
      chronicTotal += prevTotals.reduce((sum, t) => sum + t.load, 0);
      chronicDays += prevTotals.length || 7;
    }
    const chronicAvgDaily = chronicDays ? chronicTotal / chronicDays : 0;
    if (chronicAvgDaily > 0) acwr = meanLoad / chronicAvgDaily;
  }

  return { totalMinutes, totalLoad, meanLoad, monotony, acwr };
}

function svgEl(tag, attrs) {
  const node = document.createElementNS("http://www.w3.org/2000/svg", tag);
  if (attrs) {
    Object.entries(attrs).forEach(([k, v]) => {
      if (v === undefined || v === null) return;
      node.setAttribute(k, String(v));
    });
  }
  return node;
}

function buildWeekSeries() {
  const weeks = [];
  for (let i = 0; i < 52; i++) {
    const m = computeWeekMetrics(i);
    const monotony = m.monotony === null ? 0 : m.monotony;
    weeks.push({
      weekNo: i + 1,
      volumeHrs: m.totalMinutes / 60,
      load: m.totalLoad,
      monotony,
      strain: m.totalLoad * monotony,
    });
  }
  return weeks;
}

function buildDailyLoadSeries() {
  const loads = [];
  for (let w = 0; w < 52; w++) {
    const week = state.weeks[w];
    const sessions = getWeekSessions(week);
    for (let d = 0; d < 7; d++) {
      const s = sessions[d];
      loads.push(getSessionTotals(s).load);
    }
  }
  return loads;
}

function renderBarChartSvg(values, options) {
  const width = 900;
  const height = 240;
  const pad = { l: 44, r: 14, t: 16, b: 34 };
  const innerW = width - pad.l - pad.r;
  const innerH = height - pad.t - pad.b;
  const maxVal = Math.max(0, ...values.map((v) => Number(v) || 0));
  const yMax = options?.yMax && Number.isFinite(options.yMax) ? options.yMax : maxVal || 1;

  const svg = svgEl("svg", { viewBox: `0 0 ${width} ${height}`, class: "chartSvg", role: "img" });
  svg.appendChild(svgEl("rect", { x: 0, y: 0, width, height, rx: 14, fill: "var(--surface)" }));

  const axis = svgEl("g", { class: "chartAxis" });
  axis.appendChild(svgEl("line", { x1: pad.l, y1: pad.t + innerH, x2: pad.l + innerW, y2: pad.t + innerH }));
  axis.appendChild(svgEl("line", { x1: pad.l, y1: pad.t, x2: pad.l, y2: pad.t + innerH }));
  svg.appendChild(axis);

  const bars = svgEl("g", { class: "chartBars" });
  const step = innerW / values.length;
  const barW = step * 0.68;
  values.forEach((raw, i) => {
    const v = Math.max(0, Number(raw) || 0);
    const h = yMax ? (v / yMax) * innerH : 0;
    const x = pad.l + i * step + (step - barW) / 2;
    const y = pad.t + innerH - h;
    const rect = svgEl("rect", { x, y, width: barW, height: h, rx: 4, class: "chartBar" });
    const title = svgEl("title");
    title.textContent = options?.tooltip ? options.tooltip(i, v) : `${v}`;
    rect.appendChild(title);
    bars.appendChild(rect);
  });
  svg.appendChild(bars);

  const ticks = svgEl("g", { class: "chartTicks" });
  const tickEvery = options?.tickEvery || 4;
  for (let i = 0; i < values.length; i += tickEvery) {
    const x = pad.l + i * step + step / 2;
    ticks.appendChild(svgEl("line", { x1: x, y1: pad.t + innerH, x2: x, y2: pad.t + innerH + 5 }));
    const t = svgEl("text", { x, y: pad.t + innerH + 18, "text-anchor": "middle" });
    t.textContent = String(i + 1);
    ticks.appendChild(t);
  }
  svg.appendChild(ticks);

  const label = svgEl("text", { x: pad.l, y: pad.t - 4, class: "chartUnit" });
  label.textContent = options?.unit || "";
  svg.appendChild(label);

  return svg;
}

function renderLineChartSvg(series, options) {
  const width = 900;
  const height = 240;
  const pad = { l: 44, r: 14, t: 16, b: 34 };
  const innerW = width - pad.l - pad.r;
  const innerH = height - pad.t - pad.b;
  const vals = series.map((v) => (v === null || v === undefined ? null : Number(v)));
  const present = vals.filter((v) => Number.isFinite(v));
  const maxVal = present.length ? Math.max(...present) : 1;
  let yMax = options?.yMax && Number.isFinite(options.yMax) ? options.yMax : maxVal || 1;
  if (options?.minYMax && Number.isFinite(options.minYMax)) yMax = Math.max(yMax, options.minYMax);

  const svg = svgEl("svg", { viewBox: `0 0 ${width} ${height}`, class: "chartSvg", role: "img" });
  svg.appendChild(svgEl("rect", { x: 0, y: 0, width, height, rx: 14, fill: "var(--surface)" }));

  if (Array.isArray(options?.bands) && options.bands.length) {
    const bands = svgEl("g", { class: "chartBands" });
    const toY = (v) => pad.t + innerH - (yMax ? (v / yMax) * innerH : 0);
    options.bands.forEach((b) => {
      if (!b) return;
      const fromRaw = Number(b.from);
      if (!Number.isFinite(fromRaw)) return;
      const toRaw = b.to === undefined || b.to === null ? yMax : Number(b.to);
      const high = Number.isFinite(toRaw) ? Math.min(yMax, Math.max(0, toRaw)) : yMax;
      const low = Math.min(yMax, Math.max(0, fromRaw));
      if (!(high > low)) return;
      const yTop = toY(high);
      const yBottom = toY(low);
      const rect = svgEl("rect", { x: pad.l, y: yTop, width: innerW, height: yBottom - yTop, class: b.className || "" });
      if (b.fill) rect.setAttribute("fill", String(b.fill));
      if (b.opacity !== undefined && b.opacity !== null) rect.setAttribute("opacity", String(b.opacity));
      bands.appendChild(rect);
    });
    svg.appendChild(bands);
  }

  const axis = svgEl("g", { class: "chartAxis" });
  axis.appendChild(svgEl("line", { x1: pad.l, y1: pad.t + innerH, x2: pad.l + innerW, y2: pad.t + innerH }));
  axis.appendChild(svgEl("line", { x1: pad.l, y1: pad.t, x2: pad.l, y2: pad.t + innerH }));
  svg.appendChild(axis);

  const points = [];
  for (let i = 0; i < vals.length; i++) {
    const v = vals[i];
    if (!Number.isFinite(v)) continue;
    const x = pad.l + (vals.length === 1 ? innerW / 2 : (i / (vals.length - 1)) * innerW);
    const y = pad.t + innerH - (yMax ? (v / yMax) * innerH : 0);
    points.push({ x, y, i, v });
  }

  if (points.length) {
    const d = points.map((p, idx) => `${idx === 0 ? "M" : "L"}${p.x.toFixed(2)},${p.y.toFixed(2)}`).join(" ");
    const path = svgEl("path", { d, class: "chartLine" });
    svg.appendChild(path);
  }

  const dots = svgEl("g", { class: "chartDots" });
  points.forEach((p) => {
    const c = svgEl("circle", { cx: p.x, cy: p.y, r: 2.5, class: "chartDot" });
    const title = svgEl("title");
    title.textContent = options?.tooltip ? options.tooltip(p.i, p.v) : `${p.v}`;
    c.appendChild(title);
    dots.appendChild(c);
  });
  svg.appendChild(dots);

  const ticks = svgEl("g", { class: "chartTicks" });
  const tickEvery = options?.tickEvery || 4;
  for (let i = 0; i < series.length; i += tickEvery) {
    const x = pad.l + (series.length === 1 ? innerW / 2 : (i / (series.length - 1)) * innerW);
    ticks.appendChild(svgEl("line", { x1: x, y1: pad.t + innerH, x2: x, y2: pad.t + innerH + 5 }));
    const t = svgEl("text", { x, y: pad.t + innerH + 18, "text-anchor": "middle" });
    t.textContent = String(i + 1);
    ticks.appendChild(t);
  }
  svg.appendChild(ticks);

  const label = svgEl("text", { x: pad.l, y: pad.t - 4, class: "chartUnit" });
  label.textContent = options?.unit || "";
  svg.appendChild(label);

  return svg;
}

function renderDualLineChartSvg(seriesA, seriesB, options) {
  const width = 900;
  const height = 240;
  const pad = { l: 44, r: 14, t: 16, b: 34 };
  const innerW = width - pad.l - pad.r;
  const innerH = height - pad.t - pad.b;
  const valsA = seriesA.map((v) => (v === null || v === undefined ? null : Number(v)));
  const valsB = seriesB.map((v) => (v === null || v === undefined ? null : Number(v)));
  const present = [...valsA, ...valsB].filter((v) => Number.isFinite(v));
  const maxVal = present.length ? Math.max(...present) : 1;
  const yMax = options?.yMax && Number.isFinite(options.yMax) ? options.yMax : maxVal || 1;

  const svg = svgEl("svg", { viewBox: `0 0 ${width} ${height}`, class: "chartSvg", role: "img" });
  svg.appendChild(svgEl("rect", { x: 0, y: 0, width, height, rx: 14, fill: "var(--surface)" }));

  const axis = svgEl("g", { class: "chartAxis" });
  axis.appendChild(svgEl("line", { x1: pad.l, y1: pad.t + innerH, x2: pad.l + innerW, y2: pad.t + innerH }));
  axis.appendChild(svgEl("line", { x1: pad.l, y1: pad.t, x2: pad.l, y2: pad.t + innerH }));
  svg.appendChild(axis);

  const buildPath = (vals) => {
    const pts = [];
    for (let i = 0; i < vals.length; i++) {
      const v = vals[i];
      if (!Number.isFinite(v)) continue;
      const x = pad.l + (vals.length === 1 ? innerW / 2 : (i / (vals.length - 1)) * innerW);
      const y = pad.t + innerH - (yMax ? (v / yMax) * innerH : 0);
      pts.push({ x, y });
    }
    if (!pts.length) return null;
    return pts.map((p, idx) => `${idx === 0 ? "M" : "L"}${p.x.toFixed(2)},${p.y.toFixed(2)}`).join(" ");
  };

  const dA = buildPath(valsA);
  if (dA) svg.appendChild(svgEl("path", { d: dA, class: "chartLine chartLine--a" }));
  const dB = buildPath(valsB);
  if (dB) svg.appendChild(svgEl("path", { d: dB, class: "chartLine chartLine--b" }));

  const legend = svgEl("g", { class: "chartLegend" });
  const lx = pad.l + 6;
  const ly = pad.t + 8;
  legend.appendChild(svgEl("rect", { x: lx, y: ly - 9, width: 10, height: 3, class: "chartLegend__swatch chartLine--a" }));
  const ta = svgEl("text", { x: lx + 14, y: ly - 6 });
  ta.textContent = options?.labelA || "A";
  legend.appendChild(ta);
  legend.appendChild(svgEl("rect", { x: lx + 90, y: ly - 9, width: 10, height: 3, class: "chartLegend__swatch chartLine--b" }));
  const tb = svgEl("text", { x: lx + 104, y: ly - 6 });
  tb.textContent = options?.labelB || "B";
  legend.appendChild(tb);
  svg.appendChild(legend);

  const ticks = svgEl("g", { class: "chartTicks" });
  const tickEvery = options?.tickEvery || 28;
  for (let i = 0; i < seriesA.length; i += tickEvery) {
    const x = pad.l + (seriesA.length === 1 ? innerW / 2 : (i / (seriesA.length - 1)) * innerW);
    ticks.appendChild(svgEl("line", { x1: x, y1: pad.t + innerH, x2: x, y2: pad.t + innerH + 5 }));
    const t = svgEl("text", { x, y: pad.t + innerH + 18, "text-anchor": "middle" });
    t.textContent = String(Math.floor(i / 7) + 1);
    ticks.appendChild(t);
  }
  svg.appendChild(ticks);

  const label = svgEl("text", { x: pad.l, y: pad.t - 4, class: "chartUnit" });
  label.textContent = options?.unit || "";
  svg.appendChild(label);

  return svg;
}

function renderCharts() {
  const root = document.getElementById("chartsRoot");
  if (!root) return;
  root.replaceChildren();

  const plannedVolume = state.weeks.map((w) => Math.max(0, Number(w?.volumeHrs) || 0));
  const dailyLoads = buildDailyLoadSeries();
  const fitness = [];
  const fatigue = [];
  let fit = 0;
  let fat = 0;
  dailyLoads.forEach((x) => {
    fit = fit * 0.9765 + x * 0.0235;
    fat = fat * 0.8669 + x * 0.1331;
    fitness.push(fit);
    fatigue.push(fat);
  });

  const makeCard = (title, svg) => {
    const card = el("div", "chartCard");
    card.appendChild(el("div", "chartCard__title", title));
    card.appendChild(svg);
    return card;
  };

  root.appendChild(
    makeCard(
      "計劃訓練量（小時）",
      renderBarChartSvg(plannedVolume, { unit: "小時", tickEvery: 4, tooltip: (i, v) => `${weekLabelZh(i + 1)}：${v.toFixed(1)} 小時` }),
    ),
  );

  const weekSeries = buildWeekSeries();
  const strainSeries = weekSeries.map((w) => Math.max(0, Number(w?.strain) || 0));
  const strainMax = Math.max(1, ...strainSeries);
  root.appendChild(
    makeCard(
      "壓力水平",
      renderLineChartSvg(strainSeries, {
        unit: "A.U.",
        tickEvery: 4,
        yMax: strainMax,
        tooltip: (i, v) => `${weekLabelZh(i + 1)}：${Math.round(v)} A.U.`,
      }),
    ),
  );
  root.appendChild(
    makeCard(
      "體能與疲勞",
      renderDualLineChartSvg(fitness, fatigue, { unit: "A.U.", tickEvery: 28, labelA: "體能", labelB: "疲勞" }),
    ),
  );
}

function el(tag, className, text) {
  const node = document.createElement(tag);
  if (className) node.className = className;
  if (text !== undefined) node.textContent = text;
  return node;
}

function setCellText(cell, text) {
  const span = el("span", "fitText", text);
  cell.replaceChildren(span);
}

function exportCalendarPdf() {
  renderCharts();
  const calendarShell = document.getElementById("calendarShell");
  if (!calendarShell) return;

  const w = window.open("", "_blank");
  if (!w) {
    showToast("無法開啟匯出視窗（可能被瀏覽器封鎖彈出視窗）", { variant: "warn", durationMs: 2200 });
    return;
  }

  const cssHref = new URL("./styles.css", window.location.href).href;
  const doc = w.document;
  doc.open();
  doc.write(`<!doctype html>
<html lang="zh-HK">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>匯出 PDF</title>
    <link rel="stylesheet" href="${cssHref}" />
    <style>
      :root {
        --cellW: 18px;
        --labelW: 92px;
        --cellH: 18px;
        --racesH: 96px;
        --calFont: 8px;
        --calControlH: 18px;
      }
      * {
        -webkit-print-color-adjust: exact !important;
        print-color-adjust: exact !important;
        color-adjust: exact !important;
      }
      body {
        margin: 0;
        background: #ffffff;
      }
      .printRoot {
        padding: 10mm;
      }
      .calendarWrap {
        overflow: visible !important;
        border: none !important;
        box-shadow: none !important;
      }
      .calendarShell {
        gap: 0 !important;
      }
      .calendarSizer {
        display: none !important;
      }
      .calLabel {
        position: static !important;
      }
      .calendarSummary {
        margin-bottom: 8mm;
      }
      .printCharts {
        margin-top: 10mm;
        break-before: page;
      }
      .printCharts .chartsGrid {
        grid-template-columns: 1fr;
        gap: 8mm;
      }
      .printCharts .chartCard {
        box-shadow: none;
        break-inside: avoid;
      }
      .printCharts .chartCard--clickable {
        cursor: default;
      }
      @page {
        size: A4 landscape;
        margin: 8mm;
      }
      @media print {
        .printHint {
          display: none !important;
        }
      }
    </style>
  </head>
  <body>
    <div class="printRoot">
      <div class="muted printHint">在列印視窗選擇「儲存為 PDF」即可匯出</div>
      <div id="printMount"></div>
    </div>
  </body>
</html>`);
  doc.close();

  const mount = doc.getElementById("printMount");
  if (!mount) return;

  const plannedVolume52 = document.getElementById("plannedVolume52");
  if (plannedVolume52) mount.appendChild(doc.importNode(plannedVolume52, true));

  const clone = doc.importNode(calendarShell, true);
  const sizer = clone.querySelector(".calendarSizer");
  if (sizer) sizer.remove();

  const toSpan = (text) => {
    const span = doc.createElement("span");
    span.className = "fitText";
    span.textContent = text;
    return span;
  };

  clone.querySelectorAll("input").forEach((input) => {
    const v = String(input.value || "").trim();
    input.replaceWith(toSpan(v || "—"));
  });

  clone.querySelectorAll("select").forEach((select) => {
    const opt = select.selectedOptions && select.selectedOptions[0];
    const text = opt ? String(opt.textContent || "").trim() : String(select.value || "").trim();
    select.replaceWith(toSpan(text || "—"));
  });

  clone.querySelectorAll("button").forEach((btn) => {
    const text = String(btn.textContent || "").trim();
    btn.replaceWith(toSpan(text || ""));
  });

  mount.appendChild(clone);

  const chartsRoot = document.getElementById("chartsRoot");
  if (chartsRoot) {
    const chartsWrap = doc.createElement("div");
    chartsWrap.className = "printCharts";
    const title = doc.createElement("div");
    title.className = "chartCard__title";
    title.textContent = "圖表";
    chartsWrap.appendChild(title);
    chartsWrap.appendChild(doc.importNode(chartsRoot, true));
    mount.appendChild(chartsWrap);
  }

  w.focus();
  w.setTimeout(() => {
    w.print();
  }, 250);
}

function fitTextToBox(node, minFontPx) {
  const computed = window.getComputedStyle(node);
  const start = Number.parseFloat(computed.fontSize) || 12;
  let size = start;

  const maxIters = 18;
  for (let i = 0; i < maxIters; i++) {
    const overflowW = node.scrollWidth > node.clientWidth + 0.5;
    const overflowH = node.scrollHeight > node.clientHeight + 0.5;
    if (!overflowW && !overflowH) break;
    size = Math.max(minFontPx, size - 0.5);
    node.style.fontSize = `${size}px`;
    if (size <= minFontPx) break;
  }
}

let fitRafId = null;
function scheduleFitCalendarText() {
  if (fitRafId) cancelAnimationFrame(fitRafId);
  fitRafId = requestAnimationFrame(() => {
    fitRafId = null;
    const calendar = document.getElementById("calendar");
    if (!calendar) return;
    calendar.querySelectorAll(".fitText").forEach((n) => fitTextToBox(n, 9));
    calendar.querySelectorAll(".raceV").forEach((n) => fitTextToBox(n, 9));
  });
}

const STORAGE_KEY = "trainingPlanDemo_v1";
const CALENDAR_SIZE_KEY = "trainingPlanDemo_calendarSize_v1";

let calendarBaseVars = null;

function readRootPxVar(name, fallback) {
  const raw = window.getComputedStyle(document.documentElement).getPropertyValue(name).trim();
  const m = /^(-?\d+(?:\.\d+)?)px$/.exec(raw);
  if (m) return Number(m[1]);
  return fallback;
}

function ensureCalendarBaseVars() {
  if (calendarBaseVars) return calendarBaseVars;
  calendarBaseVars = {
    cellW: readRootPxVar("--cellW", 40),
    labelW: readRootPxVar("--labelW", 120),
    cellH: readRootPxVar("--cellH", 26),
    racesH: readRootPxVar("--racesH", 132),
    calFont: readRootPxVar("--calFont", 11),
    calControlH: readRootPxVar("--calControlH", 22),
  };
  return calendarBaseVars;
}

function applyCalendarSizePercent(percent) {
  const shell = document.getElementById("calendarShell");
  if (!shell) return;
  const base = ensureCalendarBaseVars();
  const p = clamp(Number(percent) || 100, 70, 140);
  const s = p / 100;

  const cellW = clamp(Math.round(base.cellW * s), 24, 72);
  const labelW = clamp(Math.round(base.labelW * s), 88, 220);
  const cellH = clamp(Math.round(base.cellH * s), 18, 44);
  const racesH = clamp(Math.round(base.racesH * s), 80, 280);
  const calFont = clamp(Math.round(base.calFont * s), 9, 16);
  const calControlH = clamp(Math.round(base.calControlH * s), 18, 44);

  const target = shell.style;
  target.setProperty("--cellW", `${cellW}px`);
  target.setProperty("--labelW", `${labelW}px`);
  target.setProperty("--cellH", `${cellH}px`);
  target.setProperty("--racesH", `${racesH}px`);
  target.setProperty("--calFont", `${calFont}px`);
  target.setProperty("--calControlH", `${calControlH}px`);

  try {
    localStorage.setItem(CALENDAR_SIZE_KEY, String(p));
  } catch {}

  scheduleFitCalendarText();
}

function wireCalendarSizer() {
  const range = document.getElementById("calSizeRange");
  const up = document.getElementById("calSizeUp");
  const down = document.getElementById("calSizeDown");
  if (!range || !up || !down) return;

  ensureCalendarBaseVars();

  let persisted = 100;
  try {
    const raw = localStorage.getItem(CALENDAR_SIZE_KEY);
    const v = Number(raw);
    if (Number.isFinite(v)) persisted = clamp(v, 70, 140);
  } catch {}

  range.value = String(persisted);
  applyCalendarSizePercent(persisted);

  range.addEventListener("input", () => {
    applyCalendarSizePercent(range.value);
  });

  const step = 5;
  up.addEventListener("click", () => {
    const next = clamp(Number(range.value) + step, 70, 140);
    range.value = String(next);
    applyCalendarSizePercent(next);
  });
  down.addEventListener("click", () => {
    const next = clamp(Number(range.value) - step, 70, 140);
    range.value = String(next);
    applyCalendarSizePercent(next);
  });
}

function loadPersistedState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== "object") return null;
    return parsed;
  } catch {
    return null;
  }
}

function persistState() {
  try {
    const payload = {
      startDate: formatYMD(state.startDate),
      ytdVolumeHrs: Number.isFinite(state.ytdVolumeHrs) ? state.ytdVolumeHrs : null,
      weeks: state.weeks.map((w) => ({
        races: Array.isArray(w.races) ? w.races : [],
        priority: w.priority || "",
        block: w.block || "",
        season: w.season || "",
        phases: normalizePhases(w.phases),
        volumeHrs: w.volumeHrs || "",
        volumeMode: w.volumeMode || "direct",
        volumeFactor: Number.isFinite(Number(w.volumeFactor)) ? Number(w.volumeFactor) : 1,
        sessions: Array.isArray(w.sessions)
          ? w.sessions.map((s) => ({
              dayLabel: s?.dayLabel || "",
              workoutsCount: Number(s?.workoutsCount) || 1,
              workouts: Array.isArray(s?.workouts) ? s.workouts.map(normalizeWorkoutEntry) : [],
              duration: Number(s?.duration) || 0,
              zone: Number(s?.zone) || 0,
              rpe: Number(s?.rpe) || 0,
              kind: s?.kind || "",
              note: typeof s?.note === "string" ? s.note : "",
            }))
          : [],
      })),
      selectedWeekIndex: state.selectedWeekIndex,
      connected: state.connected,
    };
    localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
  } catch {}
}

const state = {
  connected: false,
  startDate: startOfMonday(new Date("2025-03-03T00:00:00")),
  ytdVolumeHrs: null,
  weeks: [],
  selectedWeekIndex: 0,
};

function recomputeMondaysFromStartDate(date) {
  const monday = startOfMonday(date);
  state.startDate = monday;
  for (let i = 0; i < state.weeks.length; i++) {
    const w = state.weeks[i];
    if (!w) continue;
    w.monday = addDays(monday, i * 7);
  }
  reassignAllRacesByDate();
}

const historyState = {
  past: [],
  future: [],
  max: 60,
};

function snapshotForHistory() {
  return {
    startDate: formatYMD(state.startDate),
    ytdVolumeHrs: Number.isFinite(state.ytdVolumeHrs) ? state.ytdVolumeHrs : null,
    selectedWeekIndex: state.selectedWeekIndex,
    weeks: state.weeks.map((w) => ({
      races: Array.isArray(w.races)
        ? w.races.map((r) => ({
            name: r?.name || "",
            date: r?.date || "",
            distanceKm: Number.isFinite(Number(r?.distanceKm)) && Number(r.distanceKm) > 0 ? Number(r.distanceKm) : null,
            kind: typeof r?.kind === "string" ? r.kind : "",
          }))
        : [],
      priority: w.priority || "",
      block: w.block || "",
      season: w.season || "",
      phases: normalizePhases(w.phases),
      volumeHrs: w.volumeHrs || "",
      volumeMode: w.volumeMode || "direct",
      volumeFactor: Number.isFinite(Number(w.volumeFactor)) ? Number(w.volumeFactor) : 1,
      sessions: Array.isArray(w.sessions)
        ? w.sessions.map((s) => ({
            dayLabel: s?.dayLabel || "",
            workoutsCount: Number(s?.workoutsCount) || 1,
            workouts: Array.isArray(s?.workouts) ? s.workouts.map(normalizeWorkoutEntry) : [],
            duration: Number(s?.duration) || 0,
            zone: Number(s?.zone) || 0,
            rpe: Number(s?.rpe) || 0,
            kind: s?.kind || "",
            note: typeof s?.note === "string" ? s.note : "",
          }))
        : [],
    })),
  };
}

function applyHistorySnapshot(snapshot) {
  if (!snapshot || !Array.isArray(snapshot.weeks) || snapshot.weeks.length !== 52) return;

  const parsedStartDate = parseYMD(snapshot.startDate);
  if (parsedStartDate) {
    recomputeMondaysFromStartDate(parsedStartDate);
  }

  state.ytdVolumeHrs = Number.isFinite(snapshot.ytdVolumeHrs) ? snapshot.ytdVolumeHrs : null;
  snapshot.weeks.forEach((p, idx) => {
    const w = state.weeks[idx];
    if (!w) return;
    w.races = Array.isArray(p.races)
      ? p.races
          .map((r) => ({
            name: typeof r?.name === "string" ? r.name : "",
            date: typeof r?.date === "string" ? r.date : "",
            distanceKm: Number.isFinite(Number(r?.distanceKm)) && Number(r.distanceKm) > 0 ? Number(r.distanceKm) : null,
            kind: typeof r?.kind === "string" ? r.kind : "",
          }))
          .filter((r) => r.name.trim() && r.date)
      : [];
    w.priority = typeof p.priority === "string" ? p.priority : "";
    w.block = typeof p.block === "string" ? p.block : "";
    w.season = typeof p.season === "string" ? p.season : "";
    w.phases = normalizePhases(p?.phases ?? p?.phase);
    w.volumeHrs = typeof p.volumeHrs === "string" ? p.volumeHrs : "";
    w.volumeMode = typeof p.volumeMode === "string" ? p.volumeMode : "direct";
    w.volumeFactor = Number.isFinite(Number(p.volumeFactor)) ? Number(p.volumeFactor) : 1;
    w.sessions = Array.isArray(p.sessions)
      ? p.sessions.map((s, i) => {
          const next = {
            dayLabel: normalizeDayLabelZh(s?.dayLabel, i),
            workoutsCount: Number(s?.workoutsCount) || 0,
            workouts: Array.isArray(s?.workouts) ? s.workouts.map(normalizeWorkoutEntry) : [],
            duration: Number(s?.duration) || 0,
            zone: Number(s?.zone) || 0,
            rpe: Number(s?.rpe) || 0,
            kind: typeof s?.kind === "string" ? s.kind : "",
            note: typeof s?.note === "string" ? s.note : "",
          };
          ensureSessionWorkouts(next);
          return next;
        })
      : [];
  });

  state.selectedWeekIndex = clamp(Number(snapshot.selectedWeekIndex) || 0, 0, 51);
}

function pushHistory() {
  historyState.past.push(snapshotForHistory());
  if (historyState.past.length > historyState.max) historyState.past.shift();
  historyState.future = [];
}

function applyAfterHistoryRestore() {
  recomputeFormulaVolumes();
  persistState();
  updateHeader();
  renderCalendar();
  renderWeekPicker();
  renderWeekDetails();
}

function undoHistory() {
  if (!historyState.past.length) return;
  historyState.future.push(snapshotForHistory());
  const prev = historyState.past.pop();
  applyHistorySnapshot(prev);
  applyAfterHistoryRestore();
}

function redoHistory() {
  if (!historyState.future.length) return;
  historyState.past.push(snapshotForHistory());
  const next = historyState.future.pop();
  applyHistorySnapshot(next);
  applyAfterHistoryRestore();
}

function buildInitialWeeks() {
  const weeks = [];
  for (let i = 0; i < 52; i++) {
    const monday = addDays(state.startDate, i * 7);
    weeks.push({
      index: i,
      weekNo: i + 1,
      monday,
      races: [],
      priority: "",
      block: "Base",
      season: "",
      phases: [],
      volumeHrs: "",
      volumeMode: "direct",
      volumeFactor: 1,
      sessions: buildDefaultSessions(),
    });
  }

  state.weeks = weeks;
}

function generatePlan() {
  const rng = seededRandom(20250303);

  for (const w of state.weeks) {
    const base = 4.5 + rng() * 2.5;
    const volume = w.priority === "A" ? base * 0.75 : base;
    w.volumeHrs = `${volume.toFixed(1)}`;
    w.volumeMode = "direct";
    w.volumeFactor = 1;
  }
}

function applyAnnualVolumeToWeeks(annualVolumeHrs) {
  if (!state || !Array.isArray(state.weeks) || state.weeks.length !== 52) return false;
  const total = Number(annualVolumeHrs);
  if (!Number.isFinite(total) || total <= 0) return false;
  state.ytdVolumeHrs = total;
  const targetTotal = Math.round(total * 10) / 10;
  const before = state.weeks.map((w) => `${w?.volumeHrs ?? ""}|${w?.volumeMode ?? ""}|${w?.volumeFactor ?? ""}`).join(";");

  const weeks = state.weeks;
  const blocks = weeks.map((w) => normalizeBlockValue(w?.block || "") || "Base");
  const hasRace = weeks.map((w) => Array.isArray(w?.races) && w.races.length > 0);
  const isPreRaceWeek = hasRace.map((_, i) => i < 51 && hasRace[i + 1]);

  const weeklyAvg = targetTotal / 52;
  const volumeLevel = clamp((weeklyAvg - 3) / 10, 0, 1);
  const lerp = (a, b, t) => a + (b - a) * clamp(t, 0, 1);

  const factorSpecForWeek = (weekIndex) => {
    const idx = clamp(Number(weekIndex) || 0, 0, 51);
    const v = normalizeBlockValue(blocks[idx] || "") || "Base";

    if (v === "Peak") {
      if (hasRace[idx]) return { min: 0.6, max: 0.6, factor: 0.6 };
      if (isPreRaceWeek[idx]) return { min: 0.8, max: 0.8, factor: 0.8 };
      return { min: 0.8, max: 0.8, factor: 0.8 };
    }
    if (v === "Base") return { min: 1.1, max: 1.5, factor: lerp(1.1, 1.5, volumeLevel) };
    if (v === "Deload") return { min: 0.5, max: 0.8, factor: lerp(0.5, 0.8, volumeLevel) };
    if (v === "Build") return { min: 1.0, max: 1.1, factor: lerp(1.0, 1.1, volumeLevel) };
    if (v === "Transition") return { min: 0.8, max: 1.1, factor: lerp(0.8, 1.1, volumeLevel) };
    return { min: 1.0, max: 1.0, factor: 1.0 };
  };

  const meanPrev = (arr, endIndex, lookback) => {
    const end = clamp(Number(endIndex) || 0, 0, arr.length);
    const back = clamp(Number(lookback) || 0, 0, 10);
    const start = Math.max(0, end - back);
    let sum = 0;
    let count = 0;
    for (let i = start; i < end; i++) {
      const v = Number(arr[i]) || 0;
      if (v > 0) {
        sum += v;
        count++;
      }
    }
    return count ? sum / count : 0;
  };

  const defaultWeekVolume = (weekIndex) => {
    const idx = clamp(Number(weekIndex) || 0, 0, 51);
    const v = normalizeBlockValue(blocks[idx] || "") || "Base";
    if (v === "Deload") return Math.max(0.1, weeklyAvg * 0.7);
    if (v === "Peak") return Math.max(0.1, weeklyAvg * (hasRace[idx] ? 0.6 : 0.8));
    if (v === "Transition") return Math.max(0.1, weeklyAvg * 0.95);
    return Math.max(0.1, weeklyAvg);
  };

  const shape = new Array(52).fill(0);
  for (let i = 0; i < 4; i++) shape[i] = defaultWeekVolume(i);
  for (let i = 4; i < 52; i++) {
    const chronic4 = meanPrev(shape, i, 4) || defaultWeekVolume(i);
    const spec = factorSpecForWeek(i);
    const lo = Math.max(0.01, chronic4 * spec.min);
    const hi = Math.max(lo, chronic4 * spec.max);
    const raw = chronic4 * spec.factor;
    shape[i] = clamp(raw, lo, hi);
  }

  const sumShape = shape.reduce((a, b) => a + b, 0);
  if (!Number.isFinite(sumShape) || sumShape <= 0) return false;
  const scaled = shape.map((v) => (v * targetTotal) / sumShape);

  const targetTenths = Math.round(targetTotal * 10);
  const rawTenths = scaled.map((v) => Math.max(0, Number(v) || 0) * 10);
  const outTenths = rawTenths.map((v) => Math.floor(v + 1e-9));
  let sumTenths = outTenths.reduce((a, b) => a + b, 0);
  let diffTenths = targetTenths - sumTenths;

  const order = rawTenths
    .map((v, i) => ({ i, frac: v - outTenths[i] }))
    .sort((a, b) => b.frac - a.frac);

  if (diffTenths > 0) {
    for (let k = 0; k < order.length && diffTenths > 0; k++) {
      outTenths[order[k].i] += 1;
      diffTenths -= 1;
    }
  } else if (diffTenths < 0) {
    const rev = [...order].reverse();
    for (let k = 0; k < rev.length && diffTenths < 0; k++) {
      if (outTenths[rev[k].i] <= 0) continue;
      outTenths[rev[k].i] -= 1;
      diffTenths += 1;
    }
  }

  weeks.forEach((w, i) => {
    const v = (outTenths[i] || 0) / 10;
    w.volumeHrs = v > 0 ? v.toFixed(1) : "";
    w.volumeMode = "direct";
    w.volumeFactor = 1;
  });

  const after = state.weeks.map((w) => `${w?.volumeHrs ?? ""}|${w?.volumeMode ?? ""}|${w?.volumeFactor ?? ""}`).join(";");
  return before !== after;
}

function applyYtdVolumeToWeeks(ytdVolumeHrs) {
  return applyAnnualVolumeToWeeks(ytdVolumeHrs);
}

function applyAnnualVolumeRules() {
  if (!Number.isFinite(state.ytdVolumeHrs) || state.ytdVolumeHrs <= 0) return false;
  return applyAnnualVolumeToWeeks(state.ytdVolumeHrs);
}

function updateHeader() {
  const dateRangeEl = document.getElementById("dateRange");
  const volumeTotalEl = document.getElementById("volumeTotal");
  const plannedVolume52El = document.getElementById("plannedVolume52");

  const endDate = addDays(state.startDate, 52 * 7 - 1);
  if (dateRangeEl) {
    dateRangeEl.textContent = `${formatMD(state.startDate)} / ${state.startDate.getFullYear()} - ${formatMD(endDate)} / ${endDate.getFullYear()}`;
  }

  const totalHrs = state.weeks.reduce((sum, w) => sum + (Number(w.volumeHrs) || 0), 0);
  if (volumeTotalEl) {
    volumeTotalEl.textContent = `訓練總量：${totalHrs ? totalHrs.toFixed(1) : "—"} 小時`;
  }
  if (plannedVolume52El) {
    if (Number.isFinite(state.ytdVolumeHrs) && state.ytdVolumeHrs > 0) {
      const target = state.ytdVolumeHrs;
      const diff = Math.round((totalHrs - target) * 10) / 10;
      plannedVolume52El.textContent = `年總訓練量：${target.toFixed(1)} 小時 · 已分配：${totalHrs ? totalHrs.toFixed(1) : "—"} 小時 · 差：${diff.toFixed(1)} 小時`;
    } else {
      plannedVolume52El.textContent = `計劃訓練量（52週）：${totalHrs ? totalHrs.toFixed(1) : "—"} 小時`;
    }
  }
}

function renderCalendar() {
  const calendar = document.getElementById("calendar");
  calendar.replaceChildren();

  const seasonOptions = ["", "Base", "Build", "Peak", "Deload", "Transition"];
  const phaseColors = {
    Deload: "#3b82f6",
    "Aerobic Endurance": "#ef4444",
    Tempo: "#f59e0b",
    Threshold: "#22c55e",
    VO2Max: "#a855f7",
    Anaerobic: "#7c3aed",
    Peaking: "#ec4899",
  };
  const phaseRows = PHASE_OPTIONS.map((p) => ({
    label: phaseLabelZh(p),
    key: "phase",
    type: "phase",
    phase: p,
  }));

  const rows = [
    { label: "週次", key: "weekNo", type: "weekBtn" },
    { label: "星期一", key: "monday", type: "date" },
    { label: "週期", key: "block", type: "blockSelect" },
    { label: "比賽", key: "races", type: "races" },
    { label: "優先級", key: "priority", type: "prioritySelect" },
    ...phaseRows,
    { label: "計劃訓練量（小時）", key: "volumeHrs", type: "text" },
    { label: "單調度", key: "monotony", type: "metric" },
    { label: "ACWR", key: "acwr", type: "metric" },
  ];

  rows.forEach((row) => {
    const rowEl = el("div", "calRow");
    rowEl.dataset.rowKey = String(row.key || "");
    if (row.type === "races") rowEl.classList.add("calRow--races");
    if (row.type === "metric") rowEl.classList.add("calRow--metric");
    const label = el("div", "calCell calLabel");
    if (row.type === "phase") {
      label.classList.add("phaseLabel");
      label.style.setProperty("--phaseBg", phaseColors[row.phase] || "var(--accent)");
    }
    label.appendChild(el("span", "fitText", row.label));
    rowEl.appendChild(label);

    for (let i = 0; i < 52; i++) {
      const w = state.weeks[i];
      const cell = el("div", "calCell");

      if (row.type === "weekBtn") {
        const btn = el("button", "calWeekBtn", String(w.weekNo));
        if (i === state.selectedWeekIndex) btn.classList.add("is-selected");
        btn.addEventListener("click", () => {
          selectWeek(i);
        });
        cell.appendChild(btn);
      } else if (row.type === "date") {
        if (i === 0) {
          const wrap = el("div", "calInputWrap");
          const input = document.createElement("input");
          input.className = "calInput";
          input.type = "date";
          input.value = formatYMD(w.monday);
          input.addEventListener("click", (e) => e.stopPropagation());
          input.addEventListener("change", () => {
            const nextRaw = input.value;
            const nextDate = parseYMD(nextRaw);
            if (!nextDate) {
              input.value = formatYMD(state.startDate);
              return;
            }

            const nextMonday = startOfMonday(nextDate);
            if (nextMonday.getTime() === state.startDate.getTime()) return;

            pushHistory();
            recomputeMondaysFromStartDate(nextMonday);
            persistState();
            updateHeader();
            renderCalendar();
            renderWeekDetails();
          });
          wrap.appendChild(input);
          cell.appendChild(wrap);
        } else {
          setCellText(cell, formatMD(w.monday));
        }
      } else if (row.type === "blockSelect") {
        const wrap = el("div", "calSelectWrap");
        wrap.classList.add("calSelectWrap--wide");
        const select = document.createElement("select");
        select.className = "calSelect calSelect--wide";
        seasonOptions.forEach((v) => {
          const opt = document.createElement("option");
          opt.value = v;
          opt.textContent = BLOCK_LABELS_ZH[v] || v || "—";
          if (normalizeBlockValue(w.block || "") === v) opt.selected = true;
          select.appendChild(opt);
        });
        select.addEventListener("click", (e) => e.stopPropagation());
        select.addEventListener("change", () => {
          pushHistory();
          w.block = select.value;
          applyCoachPhaseRules();
          applyAnnualVolumeRules();
          const sessionsChanged = applyCoachDayPlanRules({ start: i, end: i });
          persistState();
          updateHeader();
          renderCalendar();
          if (sessionsChanged) {
            renderCharts();
            renderWeekDetails();
          }
        });
        wrap.appendChild(select);
        cell.appendChild(wrap);
      } else if (row.type === "prioritySelect") {
        setCellText(cell, w.priority || "—");
      } else if (row.type === "phase") {
        cell.classList.add("calCell--clickable", "phaseCell");
        const btn = el("button", "phaseBtn", "");
        btn.type = "button";
        const phases = normalizePhases(w.phases);
        const selected = phases.includes(row.phase);
        if (selected) {
          cell.classList.add("is-on");
          cell.style.setProperty("--phaseBg", phaseColors[row.phase] || "var(--accent)");
        }
        btn.setAttribute("aria-label", phaseLabelZh(row.phase));
        btn.addEventListener("click", () => {
          pushHistory();
          w.phases = selected ? phases.filter((p) => p !== row.phase) : [...phases, row.phase];
          const sessionsChanged = applyCoachDayPlanRules({ start: i, end: i });
          persistState();
          renderCalendar();
          if (sessionsChanged) {
            renderCharts();
            renderWeekDetails();
          }
        });
        cell.appendChild(btn);
      } else if (row.type === "races") {
        const cols = el("div", "raceCols");
        const races = Array.isArray(w.races) ? w.races : [];
        races.slice(0, 2).forEach((r) => {
          const name = (r?.name || "").trim();
          if (!name) return;
          const item = el("div", "raceV", name);
          cols.appendChild(item);
        });
        if (races.length > 2) {
          cols.appendChild(el("div", "raceV raceV--more", `+${races.length - 2}`));
        }
        cell.appendChild(cols);
      } else if (row.type === "metric") {
        const m = computeWeekMetrics(w.index);
        if (row.key === "monotony") {
          setCellText(cell, m.monotony === null ? "—" : m.monotony.toFixed(2));
          if (m.monotony !== null && m.monotony > 2) cell.classList.add("calCell--alert");
        } else if (row.key === "acwr") {
          setCellText(cell, m.acwr === null ? "—" : m.acwr.toFixed(2));
          if (m.acwr !== null && m.acwr > 1.5) cell.classList.add("calCell--alert");
        } else {
          setCellText(cell, "—");
        }
      } else if (row.type === "text" && row.key === "volumeHrs") {
        cell.classList.add("calCell--clickable");
        const btn = el("button", "calWeekBtn", "");
        btn.type = "button";
        btn.setAttribute("aria-label", `編輯 ${weekLabelZh(w.weekNo)} 的計劃訓練量（小時）`);
        const text = w.volumeHrs ? `${w.volumeHrs}` : "—";
        btn.appendChild(el("span", "fitText", text));
        btn.addEventListener("click", () => openPlannedVolumeModal(i));
        cell.appendChild(btn);
      } else {
        setCellText(cell, w[row.key] || "");
      }

      rowEl.appendChild(cell);
    }

    calendar.appendChild(rowEl);
  });

  scheduleFitCalendarText();
}

function renderWeekPicker() {
  const select = document.getElementById("weekSelect");
  if (!select) return;
  select.replaceChildren();

  for (const w of state.weeks) {
    const opt = document.createElement("option");
    opt.value = String(w.index);
    opt.textContent = weekLabelZh(w.weekNo);
    if (w.index === state.selectedWeekIndex) opt.selected = true;
    select.appendChild(opt);
  }

  select.onchange = () => {
    const idx = Number(select.value);
    selectWeek(clamp(idx, 0, 51));
  };
}

function relabelWeekSessions(week) {
  const sessions = getWeekSessions(week);
  for (let i = 0; i < sessions.length; i++) {
    const s = sessions[i];
    if (s && typeof s === "object") s.dayLabel = dayLabelZh(i);
  }
}

const dayDragAutoScroll = {
  active: false,
  rafId: null,
  lastY: null,
  onDragOver: null,
};

function startDayDragAutoScroll() {
  if (dayDragAutoScroll.active) return;
  dayDragAutoScroll.active = true;
  dayDragAutoScroll.lastY = null;
  dayDragAutoScroll.onDragOver = (e) => {
    if (!dayDragAutoScroll.active) return;
    dayDragAutoScroll.lastY = e.clientY;
  };
  document.addEventListener("dragover", dayDragAutoScroll.onDragOver);

  const step = () => {
    if (!dayDragAutoScroll.active) return;
    const y = dayDragAutoScroll.lastY;
    if (Number.isFinite(y)) {
      const vh = window.innerHeight || 0;
      const edge = 90;
      let delta = 0;
      if (y < edge) {
        delta = -Math.round(((edge - y) / edge) * 22);
      } else if (y > vh - edge) {
        delta = Math.round(((y - (vh - edge)) / edge) * 22);
      }
      if (delta) window.scrollBy(0, delta);
    }
    dayDragAutoScroll.rafId = requestAnimationFrame(step);
  };
  dayDragAutoScroll.rafId = requestAnimationFrame(step);
}

function stopDayDragAutoScroll() {
  if (!dayDragAutoScroll.active) return;
  dayDragAutoScroll.active = false;
  if (dayDragAutoScroll.onDragOver) {
    document.removeEventListener("dragover", dayDragAutoScroll.onDragOver);
    dayDragAutoScroll.onDragOver = null;
  }
  if (dayDragAutoScroll.rafId) {
    cancelAnimationFrame(dayDragAutoScroll.rafId);
    dayDragAutoScroll.rafId = null;
  }
  dayDragAutoScroll.lastY = null;
}

function reorderDaySession(weekIndex, fromIndex, toIndex) {
  const w = state.weeks[clamp(weekIndex, 0, 51)];
  if (!w) return;
  const sessions = getWeekSessions(w);
  if (!Array.isArray(w.sessions) || !w.sessions.length) w.sessions = sessions;

  const from = clamp(fromIndex, 0, 6);
  const to = clamp(toIndex, 0, 7);
  if (to === from || to === from + 1) return;

  pushHistory();
  const item = w.sessions.splice(from, 1)[0];
  const insertAt = from < to ? Math.max(0, to - 1) : to;
  w.sessions.splice(insertAt, 0, item);
  relabelWeekSessions(w);
  persistState();
  renderCalendar();
  renderCharts();
  renderWeekDetails();
}

let weekDetailsMount = null;

function renderWeekDetails() {
  const w = state.weeks[state.selectedWeekIndex];
  let meta = document.getElementById("weekMeta");
  let weekDays = document.getElementById("weekDays");
  if ((!meta || !weekDays) && weekDetailsMount) {
    meta = weekDetailsMount.meta;
    weekDays = weekDetailsMount.weekDays;
  }
  if (!meta || !weekDays) return;

  const sessions = getWeekSessions(w);
  if (!Array.isArray(w.sessions) || !w.sessions.length) w.sessions = sessions;
  const m = computeWeekMetrics(w.index);

  const volumeLabel = `訓練量：${(m.totalMinutes / 60).toFixed(1)} 小時`;
  const srpeLabel = `s-RPE：${Math.round(m.totalLoad)} A.U.`;
  meta.textContent = `${volumeLabel} · ${srpeLabel}`;
  weekDays.replaceChildren();

  if (weekDays && weekDays.dataset.dndWired !== "1") {
    weekDays.dataset.dndWired = "1";
    weekDays.addEventListener("dragover", (e) => {
      e.preventDefault();
      if (e.dataTransfer) e.dataTransfer.dropEffect = "move";
    });
    weekDays.addEventListener("drop", (e) => {
      e.preventDefault();
      const raw = e.dataTransfer ? e.dataTransfer.getData("text/plain") : "";
      const from = Number(raw);
      if (!Number.isFinite(from)) return;

      const target = e.target && e.target.closest ? e.target.closest(".dayCard") : null;
      if (target) return;
      const w2 = state.weeks[state.selectedWeekIndex];
      const len = w2 ? getWeekSessions(w2).length : 7;
      reorderDaySession(state.selectedWeekIndex, from, len);
      stopDayDragAutoScroll();
    });
    weekDays.addEventListener("dragend", () => {
      weekDays.classList.remove("is-dragging");
      weekDays.querySelectorAll(".dayCard").forEach((n) => n.classList.remove("is-dragging", "is-dropTarget"));
      stopDayDragAutoScroll();
    });
  }

  sessions.forEach((s, i) => {
    const card = el("div", "dayCard");
    card.dataset.dayIndex = String(i);
    ensureSessionWorkouts(s);

    const titleRow = el("div", "dayTitleRow");
    const titleLeft = el("div", "dayTitleLeft");
    const dragHandle = el("div", "dayDragHandle", "⋮⋮");
    dragHandle.setAttribute("role", "button");
    dragHandle.setAttribute("aria-label", "拖曳以重新排序");
    dragHandle.draggable = true;
    dragHandle.addEventListener("dragstart", (e) => {
      if (!e.dataTransfer) return;
      e.dataTransfer.effectAllowed = "move";
      e.dataTransfer.setData("text/plain", String(i));
      if (weekDays) weekDays.classList.add("is-dragging");
      card.classList.add("is-dragging");
      startDayDragAutoScroll();
    });
    dragHandle.addEventListener("dragend", () => {
      if (weekDays) weekDays.classList.remove("is-dragging");
      if (weekDays) weekDays.querySelectorAll(".dayCard").forEach((n) => n.classList.remove("is-dragging", "is-dropTarget"));
      stopDayDragAutoScroll();
    });
    titleLeft.appendChild(dragHandle);
    titleLeft.appendChild(el("div", "dayTitle", s.dayLabel));
    titleRow.appendChild(titleLeft);

    card.addEventListener("dragover", (e) => {
      e.preventDefault();
      if (e.dataTransfer) e.dataTransfer.dropEffect = "move";
      card.classList.add("is-dropTarget");
    });
    card.addEventListener("dragleave", (e) => {
      const next = e.relatedTarget;
      if (!next || !(next instanceof Node) || !card.contains(next)) {
        card.classList.remove("is-dropTarget");
      }
    });
    card.addEventListener("drop", (e) => {
      e.preventDefault();
      e.stopPropagation();
      card.classList.remove("is-dropTarget");
      const raw = e.dataTransfer ? e.dataTransfer.getData("text/plain") : "";
      const from = Number(raw);
      if (!Number.isFinite(from)) return;

      const rect = card.getBoundingClientRect();
      const after = e.clientY > rect.top + rect.height / 2;
      const to = i + (after ? 1 : 0);
      reorderDaySession(state.selectedWeekIndex, from, to);
      stopDayDragAutoScroll();
    });

    const dayDate = addDays(w.monday, i);
    const titleRight = el("div", "dayTitleRight");
    titleRight.appendChild(el("div", "dayDate", `${formatWeekdayEnShort(dayDate)} ${formatYMD(dayDate)}`));

    const workoutControls = el("div", "dayWorkoutControls");
    workoutControls.appendChild(el("span", "muted", "當日訓練數量"));

    const workoutSelect = document.createElement("select");
    workoutSelect.className = "dayMiniSelect dayWorkoutSelect";
    for (let v = 1; v <= 10; v++) {
      const opt = document.createElement("option");
      opt.value = String(v);
      opt.textContent = String(v);
      if (clamp(Number(s.workoutsCount) || 1, 1, 10) === v) opt.selected = true;
      workoutSelect.appendChild(opt);
    }
    workoutSelect.addEventListener("change", () => {
      const nextCount = clamp(Number(workoutSelect.value) || 1, 1, 10);
      const prevCount = clamp(Number(s.workoutsCount) || 1, 1, 10);
      if (nextCount === prevCount) return;

      lockScrollPosition(() => {
        pushHistory();
        ensureSessionWorkouts(s);
        s.workoutsCount = nextCount;
        s.workouts = Array.isArray(s.workouts) ? s.workouts.slice(0, nextCount) : [];
        while (s.workouts.length < nextCount) s.workouts.push({ duration: 0, rpe: 1 });
        ensureSessionWorkouts(s);
        persistState();
        renderCalendar();
        renderCharts();
        renderWeekDetails();
      });
    });
    workoutControls.appendChild(workoutSelect);
    titleRight.appendChild(workoutControls);
    titleRow.appendChild(titleRight);
    card.appendChild(titleRow);

    const workoutsWrap = el("div", "dayWorkouts");
    card.appendChild(workoutsWrap);
    s.workouts.forEach((workout, workoutIndex) => {
      const metaRow = el("div", "dayMeta");

      const durationWrap = el("div", "dayField");
      durationWrap.appendChild(el("span", "", "時長："));
      const durationInput = document.createElement("input");
      durationInput.className = "dayMiniInput";
      durationInput.type = "number";
      durationInput.inputMode = "numeric";
      durationInput.min = "0";
      durationInput.step = "1";
      durationInput.value = Number(workout?.duration) > 0 ? String(Number(workout?.duration)) : "";
      durationInput.addEventListener("change", () => {
        const target = Array.isArray(s.workouts) ? s.workouts[workoutIndex] : null;
        if (!target) return;
        const raw = durationInput.value.trim();
        const next = raw ? Math.max(0, Math.round(Number(raw) || 0)) : 0;
        const prev = Number(target?.duration) || 0;
        if (next === prev) return;
        lockScrollPosition(() => {
          pushHistory();
          target.duration = next;
          ensureSessionWorkouts(s);
          persistState();
          renderCalendar();
          renderCharts();
          renderWeekDetails();
        });
      });
      durationWrap.appendChild(durationInput);
      durationWrap.appendChild(el("span", "muted", "分鐘"));
      metaRow.appendChild(durationWrap);

      const rpeWrap = el("div", "dayField");
      rpeWrap.appendChild(el("span", "", "RPE："));
      const rpeSelect = document.createElement("select");
      rpeSelect.className = "dayMiniSelect";
      for (let v = 1; v <= 10; v++) {
        const opt = document.createElement("option");
        opt.value = String(v);
        opt.textContent = String(v);
        if (clamp(Number(workout?.rpe) || 1, 1, 10) === v) opt.selected = true;
        rpeSelect.appendChild(opt);
      }
      rpeSelect.addEventListener("change", () => {
        const target = Array.isArray(s.workouts) ? s.workouts[workoutIndex] : null;
        if (!target) return;
        const next = clamp(Number(rpeSelect.value) || 1, 1, 10);
        const prev = clamp(Number(target?.rpe) || 1, 1, 10);
        if (next === prev) return;
        lockScrollPosition(() => {
          pushHistory();
          target.rpe = next;
          ensureSessionWorkouts(s);
          persistState();
          renderCalendar();
          renderCharts();
          renderWeekDetails();
        });
      });
      rpeWrap.appendChild(rpeSelect);
      metaRow.appendChild(rpeWrap);

      workoutsWrap.appendChild(metaRow);
    });

    const dayDateYmd = formatYMD(addDays(w.monday, i));
    const dayRaces = (Array.isArray(w.races) ? w.races : [])
      .filter((r) => (r?.date || "") === dayDateYmd)
      .map((r) => (r?.name || "").trim())
      .filter(Boolean);
    if (dayRaces.length) {
      const racesEl = el("div", "dayRaces");
      racesEl.appendChild(el("span", "dayRaces__label", "比賽："));
      racesEl.appendChild(el("span", "dayRaces__names", dayRaces.join(", ")));
      card.appendChild(racesEl);
    }

    const noteWrap = el("div", "dayNote");
    noteWrap.appendChild(el("div", "dayNote__label", "備註"));
    const noteInput = document.createElement("textarea");
    noteInput.className = "dayNote__input";
    noteInput.rows = 2;
    noteInput.value = typeof s.note === "string" ? s.note : "";
    let noteHistoryPushed = false;
    noteInput.addEventListener("focus", () => {
      noteHistoryPushed = false;
    });
    noteInput.addEventListener("input", () => {
      const next = noteInput.value.slice(0, 1200);
      if (noteInput.value !== next) noteInput.value = next;
      if (s.note === next) return;
      if (!noteHistoryPushed) {
        pushHistory();
        noteHistoryPushed = true;
      }
      s.note = next;
      persistState();
    });
    noteWrap.appendChild(noteInput);
    card.appendChild(noteWrap);

  weekDays.appendChild(card);
  });
}

function openWeekDetailsModal(weekIndex) {
  if (weekDetailsMount && weekDetailsMount.overlay && weekDetailsMount.overlay.parentNode) {
    weekDetailsMount.overlay.remove();
    weekDetailsMount = null;
  }

  const overlay = el("div", "overlay");
  const closeModal = () => {
    if (weekDetailsMount && weekDetailsMount.overlay === overlay) weekDetailsMount = null;
    overlay.remove();
  };

  overlay.addEventListener("click", (e) => {
    if (e.target === overlay) closeModal();
  });

  const modal = el("div", "modal");
  modal.classList.add("weekModal");

  const w = state.weeks[clamp(weekIndex, 0, 51)];
  const title = el("div", "modal__title", w ? weekLabelZh(w.weekNo) : "週");
  const meta = el("div", "muted", "");
  const weekDays = el("div", "weekDays");

  const actions = el("div", "modal__actions");
  const close = el("button", "btn", "關閉");
  close.type = "button";
  close.addEventListener("click", () => closeModal());
  actions.appendChild(close);

  modal.appendChild(title);
  modal.appendChild(meta);
  modal.appendChild(weekDays);
  modal.appendChild(actions);
  overlay.appendChild(modal);
  document.body.appendChild(overlay);

  weekDetailsMount = { overlay, meta, weekDays };

  selectWeek(clamp(weekIndex, 0, 51));
}

function openPlannedVolumeModal(weekIndex) {
  const idx = clamp(Number(weekIndex) || 0, 0, 51);
  const w = state.weeks[idx];
  if (!w) return;

  const overlay = el("div", "overlay");
  const onKeyDown = (e) => {
    if (e.key === "Escape") closeModal();
  };

  const closeModal = () => {
    document.removeEventListener("keydown", onKeyDown);
    overlay.remove();
  };

  overlay.addEventListener("click", (e) => {
    if (e.target === overlay) closeModal();
  });

  const modal = el("div", "modal volumeModal");
  const title = el("div", "modal__title", `計劃訓練量 · ${weekLabelZh(w.weekNo)}`);
  const subtitle = el("div", "muted", "公式固定為：過去四週平均訓練量 × 參數");

  const form = el("form", "volumeForm");

  const row1 = el("div", "volumeRow");
  row1.appendChild(el("div", "volumeLabel", "模式"));
  const modeWrap = el("div", "volumeToggle");
  const modeDirect = document.createElement("button");
  modeDirect.type = "button";
  modeDirect.className = "btn volumeToggleBtn";
  modeDirect.textContent = "直接更改";
  const modeFormula = document.createElement("button");
  modeFormula.type = "button";
  modeFormula.className = "btn volumeToggleBtn";
  modeFormula.textContent = "公式";
  let mode = w.volumeMode === "formula" ? "formula" : "direct";
  const syncMode = () => {
    modeDirect.classList.toggle("is-active", mode === "direct");
    modeFormula.classList.toggle("is-active", mode === "formula");
    row3.hidden = mode !== "direct";
    row4.hidden = mode !== "formula";
    row4b.hidden = mode !== "formula";
  };
  modeDirect.addEventListener("click", () => {
    mode = "direct";
    syncMode();
    window.setTimeout(() => direct.focus(), 0);
  });
  modeFormula.addEventListener("click", () => {
    mode = "formula";
    syncMode();
    updatePreview();
    window.setTimeout(() => factor.focus(), 0);
  });
  modeWrap.appendChild(modeDirect);
  modeWrap.appendChild(modeFormula);
  row1.appendChild(modeWrap);
  form.appendChild(row1);

  const row2 = el("div", "volumeRow");
  row2.appendChild(el("div", "volumeLabel", "套用到其他週"));
  const applyWrap = el("div", "volumeToggle");
  const applyYes = document.createElement("button");
  applyYes.type = "button";
  applyYes.className = "btn volumeToggleBtn";
  applyYes.textContent = "是";
  const applyNo = document.createElement("button");
  applyNo.type = "button";
  applyNo.className = "btn volumeToggleBtn is-active";
  applyNo.textContent = "否";
  let applyToOthers = false;
  const syncToggle = () => {
    applyYes.classList.toggle("is-active", applyToOthers);
    applyNo.classList.toggle("is-active", !applyToOthers);
  };
  applyYes.addEventListener("click", () => {
    applyToOthers = true;
    syncToggle();
  });
  applyNo.addEventListener("click", () => {
    applyToOthers = false;
    syncToggle();
  });
  applyWrap.appendChild(applyYes);
  applyWrap.appendChild(applyNo);
  row2.appendChild(applyWrap);
  form.appendChild(row2);

  const row3 = el("div", "volumeRow");
  row3.appendChild(el("div", "volumeLabel", "直接輸入（小時）"));
  const direct = document.createElement("input");
  direct.className = "input volumeInput";
  direct.type = "number";
  direct.inputMode = "decimal";
  direct.step = "0.1";
  direct.placeholder = "小時";
  direct.value = w.volumeHrs || "";
  row3.appendChild(direct);
  form.appendChild(row3);

  const row4 = el("div", "volumeRow");
  row4.hidden = true;
  row4.appendChild(el("div", "volumeLabel", "參數"));
  const factor = document.createElement("input");
  factor.className = "input volumeInput";
  factor.type = "number";
  factor.inputMode = "decimal";
  factor.step = "0.01";
  factor.placeholder = "例如：1.10";
  factor.value = Number.isFinite(Number(w.volumeFactor)) ? String(w.volumeFactor) : "1.00";
  row4.appendChild(factor);
  form.appendChild(row4);

  const row4b = el("div", "volumeRow");
  row4b.hidden = true;
  row4b.appendChild(el("div", "volumeLabel", "預覽"));
  const preview = el("div", "muted", "—");
  row4b.appendChild(preview);
  form.appendChild(row4b);

  const actions = el("div", "modal__actions");
  const cancel = el("button", "btn", "取消");
  cancel.type = "button";
  cancel.addEventListener("click", () => closeModal());
  const submit = el("button", "btn btn--primary", "確認");
  submit.type = "submit";
  actions.appendChild(cancel);
  actions.appendChild(submit);
  form.appendChild(actions);

  const updatePreview = () => {
    if (mode !== "formula") return;
    const rawFactor = String(factor.value || "").trim();
    const f = Number(rawFactor);
    if (!Number.isFinite(f)) {
      preview.textContent = "—";
      return;
    }
    const baseline = state.weeks.map((w) => Number(w.volumeHrs) || 0);
    const base = computePastMean(baseline, idx, 4);
    const out = formatVolumeHrs(base * f);
    preview.textContent = out ? `${out} 小時` : "—";
  };

  factor.addEventListener("input", updatePreview);

  form.addEventListener("submit", (e) => {
    e.preventDefault();
    const directRaw = String(direct.value || "").trim();
    if (mode === "direct" && directRaw.length === 0) return;
    if (mode === "formula" && String(factor.value || "").trim().length === 0) return;

    pushHistory();
    if (mode === "direct") {
      const n = Number(directRaw);
      const next = formatVolumeHrs(n);
      const start = applyToOthers ? 0 : idx;
      const end = applyToOthers ? state.weeks.length - 1 : idx;
      for (let k = start; k <= end; k++) {
        const wk = state.weeks[k];
        if (!wk) continue;
        wk.volumeMode = "direct";
        wk.volumeFactor = 1;
        wk.volumeHrs = next;
      }
      recomputeFormulaVolumes();
      applyCoachDayPlanRules({ start, end });
    } else if (mode === "formula") {
      const rawFactor = String(factor.value || "").trim();
      const f = Number(rawFactor);
      if (!Number.isFinite(f)) {
        showToast("參數格式不正確", { variant: "warn", durationMs: 2000 });
        return;
      }
      if (applyToOthers) {
        for (let k = 0; k < state.weeks.length; k++) {
          const wk = state.weeks[k];
          if (!wk) continue;
          if (k < 4) {
            wk.volumeMode = "direct";
            if (!Number.isFinite(Number(wk.volumeFactor))) wk.volumeFactor = 1;
          } else {
            wk.volumeMode = "formula";
            wk.volumeFactor = f;
          }
        }
      } else {
        w.volumeMode = "formula";
        w.volumeFactor = f;
      }
      recomputeFormulaVolumes();
      if (applyToOthers) applyCoachDayPlanRules({ start: 0, end: 51 });
      else applyCoachDayPlanRules({ start: idx, end: idx });
    }

    persistState();
    updateHeader();
    renderCalendar();
    renderCharts();
    closeModal();
  });

  modal.appendChild(title);
  modal.appendChild(subtitle);
  modal.appendChild(form);
  overlay.appendChild(modal);
  document.body.appendChild(overlay);

  document.addEventListener("keydown", onKeyDown);
  syncMode();
  updatePreview();
  window.setTimeout(() => (mode === "formula" ? factor.focus() : direct.focus()), 0);
}

let toastTimer = null;
function showToast(text, options) {
  const variant = options?.variant || "info";
  const durationMs = Number.isFinite(options?.durationMs) ? options.durationMs : 1400;
  let bar = document.getElementById("toast");
  if (!bar) {
    bar = el("div", "toast");
    bar.id = "toast";
    document.body.appendChild(bar);
  }

  bar.classList.toggle("toast--warn", variant === "warn");
  bar.textContent = text;
  bar.style.display = "block";

  if (toastTimer) window.clearTimeout(toastTimer);
  toastTimer = window.setTimeout(() => {
    bar.style.display = "none";
  }, durationMs);
}

function weekIndexForRaceYmd(ymd) {
  const d = parseYMD(ymd);
  if (!d) return null;
  const monday = startOfMonday(d);
  const diffWeeks = Math.round((monday.getTime() - state.startDate.getTime()) / (MS_PER_DAY * 7));
  if (!Number.isFinite(diffWeeks)) return null;
  if (diffWeeks < 0 || diffWeeks > 51) return null;
  return diffWeeks;
}

function reassignAllRacesByDate() {
  const entries = [];
  for (let i = 0; i < state.weeks.length; i++) {
    const w = state.weeks[i];
    const races = Array.isArray(w?.races) ? w.races : [];
    races.forEach((r) => {
      const name = (r?.name || "").trim();
      const date = (r?.date || "").trim();
      if (!name || !date) return;
      const dist = Number(r?.distanceKm);
      const distanceKm = Number.isFinite(dist) && dist > 0 ? dist : null;
      const kind = typeof r?.kind === "string" ? r.kind : "";
      entries.push({ name, date, distanceKm, kind });
    });
  }

  state.weeks.forEach((w) => {
    w.races = [];
  });

  entries.forEach((r) => {
    const idx = weekIndexForRaceYmd(r.date);
    if (idx === null) return;
    const w = state.weeks[idx];
    if (!w) return;
    if (!Array.isArray(w.races)) w.races = [];
    const exists = w.races.some((x) => (x?.date || "") === r.date && (x?.name || "").trim() === r.name);
    if (!exists) w.races.push({ name: r.name, date: r.date, distanceKm: r.distanceKm ?? null, kind: r.kind || "" });
  });

  state.weeks.forEach((w) => {
    if (!Array.isArray(w.races) || w.races.length === 0) w.priority = "";
  });

  applyCoachAutoRules();
}

function serializeStateForAiPlan() {
  return {
    startDate: formatYMD(state.startDate),
    weeks: state.weeks.map((w) => ({
      index: w.index,
      weekNo: w.weekNo,
      monday: formatYMD(w.monday),
      priority: typeof w.priority === "string" ? w.priority : "",
      races: Array.isArray(w.races)
        ? w.races
            .map((r) => ({
              name: typeof r?.name === "string" ? r.name.trim() : "",
              date: typeof r?.date === "string" ? r.date : "",
              distanceKm: Number.isFinite(Number(r?.distanceKm)) && Number(r.distanceKm) > 0 ? Number(r.distanceKm) : null,
              kind: typeof r?.kind === "string" ? r.kind : "",
            }))
            .filter((r) => r.name && r.date)
        : [],
      block: typeof w.block === "string" ? w.block : "",
      phases: normalizePhases(w.phases),
      volumeHrs: typeof w.volumeHrs === "string" ? w.volumeHrs : "",
    })),
  };
}

function applyAiPlanUpdates(updates) {
  if (!updates || typeof updates !== "object") return false;
  if (!Array.isArray(updates.weeks) || updates.weeks.length === 0) return false;

  pushHistory();

  updates.weeks.forEach((u) => {
    const idx = clamp(Number(u?.index), 0, 51);
    const w = state.weeks[idx];
    if (!w) return;

    const nextBlock = typeof u?.block === "string" ? normalizeBlockValue(u.block) : "";
    if (nextBlock && Object.prototype.hasOwnProperty.call(BLOCK_LABELS_ZH, nextBlock)) {
      w.block = nextBlock;
    }

    if (Array.isArray(u?.phases)) {
      w.phases = normalizePhases(u.phases);
    }

    const volRaw = typeof u?.volumeHrs === "string" ? u.volumeHrs.trim() : "";
    if (volRaw) {
      const v = Number(volRaw);
      if (Number.isFinite(v) && v > 0) {
        w.volumeMode = "direct";
        w.volumeFactor = 1;
        w.volumeHrs = formatVolumeHrs(v);
      }
    }

    if (Array.isArray(u?.sessions) && u.sessions.length) {
      const sessions = getWeekSessions(w);
      if (!Array.isArray(w.sessions) || !w.sessions.length) w.sessions = sessions;

      u.sessions.forEach((su) => {
        const dayIndex = clamp(Number(su?.dayIndex), 0, 6);
        const s = sessions[dayIndex];
        if (!s) return;

        if (Number.isFinite(Number(su?.zone))) {
          s.zone = clamp(Number(su.zone), 1, 6);
        }

        const durationMinutes = Math.max(0, Math.round(Number(su?.durationMinutes) || 0));
        const rpe = clamp(Number(su?.rpe) || 1, 1, 10);

        ensureSessionWorkouts(s);
        const count = clamp(Number(s.workoutsCount) || 1, 1, 10);
        const base = count ? Math.floor(durationMinutes / count) : durationMinutes;
        let rem = count ? durationMinutes - base * count : 0;

        for (let k = 0; k < count; k++) {
          const wk = Array.isArray(s.workouts) ? s.workouts[k] : null;
          if (!wk) continue;
          wk.duration = base + (rem > 0 ? 1 : 0);
          if (rem > 0) rem--;
          wk.rpe = rpe;
        }
        ensureSessionWorkouts(s);
      });
    }
  });

  applyCoachAutoRules();
  recomputeFormulaVolumes();
  persistState();
  updateHeader();
  renderCalendar();
  renderCharts();
  renderWeekDetails();
  return true;
}

function openAiPlanModal() {
  const overlay = el("div", "overlay");
  overlay.addEventListener("click", (e) => {
    if (e.target === overlay) overlay.remove();
  });

  const modal = el("div", "modal");
  modal.classList.add("modal--scroll");
  const title = el("div", "modal__title", "AI 助手 · 自動編排");
  const subtitle = el("div", "modal__subtitle", "按比賽日期／優先級分週期、填訓練區與訓練量，並預填 Day 時長與 RPE");

  const racesBox = el("div", "raceList");
  const rows = [];
  state.weeks.forEach((w) => {
    const races = Array.isArray(w.races) ? w.races : [];
    races.forEach((r) => {
      const name = (r?.name || "").trim();
      const date = (r?.date || "").trim();
      if (!name || !date) return;
      const dist = Number(r?.distanceKm);
      const distText = Number.isFinite(dist) && dist > 0 ? `${dist}km` : "";
      const kindRaw = String(r?.kind || "").trim();
      const kindText = kindRaw === "trail" ? "越野跑" : kindRaw === "road" ? "路跑" : kindRaw;
      rows.push({ weekNo: w.weekNo, date, name, priority: w.priority || "", distText, kindText });
    });
  });
  rows.sort((a, b) => a.date.localeCompare(b.date) || a.weekNo - b.weekNo || a.name.localeCompare(b.name));
  if (!rows.length) {
    racesBox.appendChild(el("div", "muted", "未有比賽（請先按「輸入比賽」）"));
  } else {
    rows.forEach((r) => {
      const row = el("div", "raceRow");
      const meta = [r.distText, r.kindText].filter(Boolean).join(" · ");
      const parts = [weekLabelZh(r.weekNo), r.date, r.name];
      if (meta) parts.push(meta);
      parts.push(r.priority || "—");
      row.appendChild(el("div", "raceRow__name", parts.join(" · ")));
      racesBox.appendChild(row);
    });
  }

  const form = el("form", "paceForm");
  const notes = document.createElement("textarea");
  notes.className = "input";
  notes.rows = 4;
  notes.placeholder = "補充要求（可選）：例如每週跑 5 日、長課星期日、避免高強度連續兩日…";
  form.appendChild(notes);

  const actions = el("div", "modal__actions");
  const cancel = el("button", "btn", "取消");
  cancel.type = "button";
  cancel.addEventListener("click", () => overlay.remove());
  const submit = el("button", "btn btn--primary", "生成並套用");
  submit.type = "submit";
  actions.appendChild(cancel);
  actions.appendChild(submit);

  form.addEventListener("submit", async (e) => {
    e.preventDefault();
    if (submit.disabled) return;
    if (!rows.length) {
      showToast("請先輸入最少一個比賽", { variant: "warn", durationMs: 1800 });
      return;
    }

    submit.disabled = true;
    const prevText = submit.textContent;
    submit.textContent = "生成中…";

    try {
      const autoChanged = applyCoachAutoRules();
      if (autoChanged) {
        persistState();
        renderCalendar();
        renderCharts();
      }

      const resp = await fetch("/api/ai-plan", {
        method: "POST",
        headers: { "Content-Type": "application/json", Accept: "application/json" },
        body: JSON.stringify({ state: serializeStateForAiPlan(), notes: notes.value || "" }),
      });
      let data = null;
      let rawText = "";
      const cloned = resp.clone();
      try {
        data = await resp.json();
      } catch {
        data = null;
        rawText = await cloned.text().catch(() => "");
      }
      if (!resp.ok) {
        const versionFromBody = typeof data?.serverVersion === "string" ? data.serverVersion.trim() : "";
        const versionFromHeader = (resp.headers.get("X-Server-Version") || "").trim();
        const version = versionFromBody || versionFromHeader;
        const msg =
          typeof data?.error === "string" && data.error.trim()
            ? data.error.trim()
            : rawText && String(rawText).trim()
              ? `AI 服務錯誤（HTTP ${resp.status}）`
              : `AI 服務錯誤（HTTP ${resp.status}）`;
        showToast(version ? `${msg}（${version}）` : msg, { variant: "warn", durationMs: 2600 });
        return;
      }
      const updates = data?.updates;
      if (!updates || typeof updates !== "object") {
        showToast("AI 回覆格式不正確", { variant: "warn", durationMs: 2200 });
        return;
      }

      const ok = applyAiPlanUpdates(updates);
      if (!ok) {
        showToast("未能套用 AI 編排結果", { variant: "warn", durationMs: 2200 });
        return;
      }

      showToast("已套用 AI 編排");
      overlay.remove();
    } catch {
      showToast("連線失敗（請稍後再試）", { variant: "warn", durationMs: 2200 });
    } finally {
      submit.disabled = false;
      submit.textContent = prevText;
    }
  });

  form.appendChild(actions);

  modal.appendChild(title);
  modal.appendChild(subtitle);
  modal.appendChild(racesBox);
  modal.appendChild(form);
  overlay.appendChild(modal);
  document.body.appendChild(overlay);

  window.setTimeout(() => notes.focus(), 0);
}

function openRaceInputModal() {
  const overlay = el("div", "overlay");
  overlay.addEventListener("click", (e) => {
    if (e.target === overlay) overlay.remove();
  });

  const modal = el("div", "modal");
  modal.classList.add("modal--scroll");
  const title = el("div", "modal__title", "輸入比賽");
  const subtitle = el("div", "modal__subtitle", "輸入比賽名稱、日期、距離、項目及優先級；系統會按日期更新相應週的比賽／優先級");

  const list = el("div", "raceList");
  const renderList = () => {
    list.replaceChildren();
    const rows = [];
    state.weeks.forEach((w, weekIdx) => {
      const races = Array.isArray(w.races) ? w.races : [];
      races.forEach((r, raceIdx) => {
        const name = (r?.name || "").trim();
        const date = (r?.date || "").trim();
        if (!name || !date) return;
        const dist = Number(r?.distanceKm);
        const distText = Number.isFinite(dist) && dist > 0 ? `${dist}km` : "";
        const kindRaw = String(r?.kind || "").trim();
        const kindText = kindRaw === "trail" ? "越野跑" : kindRaw === "road" ? "路跑" : kindRaw;
        rows.push({ weekIdx, weekNo: w.weekNo, raceIdx, name, date, distText, kindText });
      });
    });

    rows.sort((a, b) => a.date.localeCompare(b.date) || a.weekIdx - b.weekIdx || a.name.localeCompare(b.name));

    if (!rows.length) {
      list.appendChild(el("div", "muted", "未有比賽"));
      return;
    }

    rows.forEach((r) => {
      const row = el("div", "raceRow");
      const meta = [r.distText, r.kindText].filter(Boolean).join(" · ");
      const parts = [weekLabelZh(r.weekNo), r.date, r.name];
      if (meta) parts.push(meta);
      row.appendChild(el("div", "raceRow__name", parts.join(" · ")));
      const del = el("button", "iconBtn", "×");
      del.type = "button";
      del.addEventListener("click", () => {
        pushHistory();
        const w = state.weeks[r.weekIdx];
        if (w && Array.isArray(w.races)) {
          w.races.splice(r.raceIdx, 1);
        }
        if (w && (!Array.isArray(w.races) || w.races.length === 0)) {
          w.priority = "";
        }
        applyCoachAutoRules();
        persistState();
        renderList();
        renderCalendar();
        renderWeekDetails();
      });
      row.appendChild(del);
      list.appendChild(row);
    });
  };

  const form = el("form", "formRow");
  const name = document.createElement("input");
  name.className = "input";
  name.type = "text";
  name.placeholder = "比賽名稱";
  name.required = true;

  const date = document.createElement("input");
  date.className = "input";
  date.type = "date";
  date.required = true;

  const distanceKm = document.createElement("input");
  distanceKm.className = "input";
  distanceKm.type = "number";
  distanceKm.placeholder = "距離（公里）";
  distanceKm.min = "0";
  distanceKm.step = "0.1";
  distanceKm.required = true;

  const kind = document.createElement("select");
  kind.className = "input";
  kind.required = true;
  [
    { v: "", t: "項目" },
    { v: "road", t: "路跑" },
    { v: "trail", t: "越野跑" },
  ].forEach(({ v, t }) => {
    const opt = document.createElement("option");
    opt.value = v;
    opt.textContent = t;
    kind.appendChild(opt);
  });

  const priority = document.createElement("select");
  priority.className = "input";
  priority.required = true;
  [
    { v: "", t: "優先級" },
    { v: "A", t: "A" },
    { v: "B", t: "B" },
    { v: "C", t: "C" },
  ].forEach(({ v, t }) => {
    const opt = document.createElement("option");
    opt.value = v;
    opt.textContent = t;
    priority.appendChild(opt);
  });

  const add = el("button", "btn btn--primary", "加入");
  add.type = "submit";

  form.appendChild(name);
  form.appendChild(date);
  form.appendChild(distanceKm);
  form.appendChild(kind);
  form.appendChild(priority);
  form.appendChild(add);

  form.addEventListener("submit", (e) => {
    e.preventDefault();
    const raceName = name.value.trim();
    const raceDate = date.value;
    const raceDistanceKm = Number(String(distanceKm.value || "").trim());
    const raceKind = kind.value;
    const racePriority = priority.value;
    if (!raceName || !raceDate || !racePriority) return;
    if (!Number.isFinite(raceDistanceKm) || raceDistanceKm <= 0) {
      showToast("請輸入有效距離（公里）", { variant: "warn", durationMs: 1800 });
      return;
    }
    if (!raceKind) {
      showToast("請選擇項目", { variant: "warn", durationMs: 1800 });
      return;
    }
    const idx = weekIndexForRaceYmd(raceDate);
    if (idx === null) {
      showToast("日期不在 52 週計畫範圍內", { variant: "warn", durationMs: 1800 });
      return;
    }
    pushHistory();
    const w = state.weeks[idx];
    if (!Array.isArray(w.races)) w.races = [];
    const existingIndex = w.races.findIndex((r) => (r?.date || "") === raceDate && (r?.name || "").trim() === raceName);
    if (existingIndex >= 0) {
      const ex = w.races[existingIndex];
      if (ex && typeof ex === "object") {
        ex.distanceKm = raceDistanceKm;
        ex.kind = raceKind;
      }
    } else {
      w.races.push({ name: raceName, date: raceDate, distanceKm: raceDistanceKm, kind: raceKind });
    }
    name.value = "";
    date.value = "";
    distanceKm.value = "";
    kind.value = "";
    priority.value = "";
    w.priority = racePriority;
    applyCoachAutoRules();
    persistState();
    renderList();
    renderCalendar();
    renderWeekDetails();
  });

  const actions = el("div", "modal__actions");
  const close = el("button", "btn", "關閉");
  close.type = "button";
  close.addEventListener("click", () => overlay.remove());
  actions.appendChild(close);

  modal.appendChild(title);
  modal.appendChild(subtitle);
  modal.appendChild(list);
  modal.appendChild(form);
  modal.appendChild(actions);
  overlay.appendChild(modal);
  document.body.appendChild(overlay);

  renderList();
  window.setTimeout(() => name.focus(), 0);
}

function selectWeek(index) {
  state.selectedWeekIndex = clamp(index, 0, 51);
  persistState();
  renderCalendar();
  const weekSelect = document.getElementById("weekSelect");
  if (weekSelect) weekSelect.value = String(state.selectedWeekIndex);
  renderWeekDetails();
}

function wireTabs() {
  const tabs = Array.from(document.querySelectorAll(".tab"));
  tabs.forEach((t) => {
    t.addEventListener("click", () => {
      const key = String(t.getAttribute("data-tab") || "");
      activateTab(key, { persist: true, updateHash: true });
    });
  });

  window.addEventListener("hashchange", () => {
    const key = getTabKeyFromHash();
    if (!key) return;
    activateTab(key, { persist: true, updateHash: false });
  });
}

function wireButtons() {
  const connectBtn = document.getElementById("connectBtn");
  const connectMeta = document.getElementById("connectMeta");

  if (connectBtn && connectMeta) {
    const setConnectMeta = () => {
      connectMeta.textContent = state.connected ? "已連接" : "";
    };

    connectBtn.textContent = state.connected ? "已連接" : "連接 Strava";
    setConnectMeta();

    connectBtn.addEventListener("click", () => {
      state.connected = !state.connected;
      connectBtn.textContent = state.connected ? "已連接" : "連接 Strava";
      setConnectMeta();
      persistState();
    });
  }

  const annualVolumeInput = document.getElementById("annualVolumeInput");
  const annualVolumeApplyBtn = document.getElementById("annualVolumeApplyBtn");
  const annualVolumeClearBtn = document.getElementById("annualVolumeClearBtn");
  if (annualVolumeInput) {
    annualVolumeInput.value = Number.isFinite(state.ytdVolumeHrs) && state.ytdVolumeHrs > 0 ? String(state.ytdVolumeHrs) : "";
    annualVolumeInput.addEventListener("keydown", (e) => {
      if (e.key !== "Enter") return;
      e.preventDefault();
      annualVolumeApplyBtn?.click();
    });
  }
  if (annualVolumeApplyBtn && annualVolumeInput) {
    annualVolumeApplyBtn.addEventListener("click", () => {
      const raw = String(annualVolumeInput.value || "").trim();
      const v = Number(raw);
      if (!Number.isFinite(v) || v <= 0) {
        showToast("請輸入有效的年總訓練量（小時）", { variant: "warn", durationMs: 1800 });
        return;
      }
      pushHistory();
      applyAnnualVolumeToWeeks(v);
      applyCoachDayPlanRules({ start: 0, end: 51 });
      persistState();
      updateHeader();
      renderCalendar();
      renderCharts();
      renderWeekDetails();
      showToast("已套用年總訓練量分配");
    });
  }
  if (annualVolumeClearBtn && annualVolumeInput) {
    annualVolumeClearBtn.addEventListener("click", () => {
      pushHistory();
      state.ytdVolumeHrs = null;
      annualVolumeInput.value = "";
      persistState();
      updateHeader();
      renderCalendar();
      renderCharts();
      renderWeekDetails();
      showToast("已清除年總訓練量設定");
    });
  }

  const raceInputBtn = document.getElementById("raceInputBtn");
  if (raceInputBtn) {
    raceInputBtn.addEventListener("click", () => openRaceInputModal());
  }

  let aiPlanBtn = document.getElementById("aiPlanBtn");
  if (!aiPlanBtn && raceInputBtn && raceInputBtn.parentElement) {
    const btn = document.createElement("button");
    btn.id = "aiPlanBtn";
    btn.className = "btn";
    btn.type = "button";
    btn.textContent = "AI 編排";
    raceInputBtn.insertAdjacentElement("afterend", btn);
    aiPlanBtn = btn;
  }
  if (aiPlanBtn) aiPlanBtn.addEventListener("click", () => openAiPlanModal());

  const exportPdfBtn = document.getElementById("exportPdfBtn");
  if (exportPdfBtn) exportPdfBtn.addEventListener("click", () => exportCalendarPdf());

  const resetBtn = document.getElementById("resetBtn");
  if (resetBtn) resetBtn.addEventListener("click", () => {
    const ok = window.confirm("確定要將表格內容重設為空白？此操作會清除已儲存資料。");
    if (!ok) return;
    pushHistory();
    const connected = state.connected;
    buildInitialWeeks();
    state.ytdVolumeHrs = null;
    state.connected = connected;
    state.selectedWeekIndex = 0;
    persistState();
    updateHeader();
    renderCalendar();
    renderWeekPicker();
    renderWeekDetails();
    showToast("已重設為空白");
  });

  const generateBtn = document.getElementById("generateBtn");
  if (generateBtn) {
    generateBtn.addEventListener("click", () => {
      pushHistory();
      generatePlan();
      applyCoachDayPlanRules({ start: 0, end: 51 });
      updateHeader();
      renderCalendar();
      renderWeekDetails();
      renderCharts();
      showToast("已生成訓練計畫（示範）");
    });
  }

  const paceDistSelect = document.getElementById("paceTestDistance");
  if (paceDistSelect && paceDistSelect.dataset.filled !== "1") {
    if (!paceDistSelect.options || paceDistSelect.options.length === 0) {
      paceDistSelect.replaceChildren();
      PACE_TEST_OPTIONS.forEach((o, i) => {
        const opt = document.createElement("option");
        opt.value = String(o.meters);
        opt.textContent = o.label;
        if (i === 2) opt.selected = true;
        paceDistSelect.appendChild(opt);
      });
    }
    paceDistSelect.dataset.filled = "1";
  }

  const paceCalcBtn = document.getElementById("paceCalcBtn");
  if (paceCalcBtn) paceCalcBtn.addEventListener("click", () => computeAndRenderPaceCalculator());

  const paceResetBtn = document.getElementById("paceResetBtn");
  if (paceResetBtn) paceResetBtn.addEventListener("click", () => resetPaceCalculator());

  const paceTimeInputs = [
    document.getElementById("paceTestHours"),
    document.getElementById("paceTestMinutes"),
    document.getElementById("paceTestSeconds"),
    document.getElementById("paceTestTime"),
  ].filter(Boolean);
  paceTimeInputs.forEach((n) => {
    n.addEventListener("keydown", (e) => {
      if (e.key !== "Enter") return;
      e.preventDefault();
      computeAndRenderPaceCalculator();
    });
  });
}

function init() {
  const persisted = loadPersistedState();
  const persistedStartDate = parseYMD(persisted?.startDate);
  if (persistedStartDate) {
    state.startDate = startOfMonday(persistedStartDate);
  }
  state.ytdVolumeHrs = Number.isFinite(persisted?.ytdVolumeHrs) ? persisted.ytdVolumeHrs : null;
  buildInitialWeeks();
  if (persisted?.weeks?.length === 52) {
    const seasonOptions = ["", "Base", "Build", "Peak", "Deload", "Transition"];
    persisted.weeks.forEach((p, idx) => {
      const w = state.weeks[idx];
      if (!w) return;
      w.priority = typeof p.priority === "string" ? p.priority : "";
      w.block = typeof p.block === "string" ? normalizeBlockValue(p.block) : w.block;
      w.season = typeof p.season === "string" ? p.season : "";
      w.phases = normalizePhases(p?.phases ?? p?.phase);
      w.volumeHrs = typeof p.volumeHrs === "string" ? p.volumeHrs : "";
      w.volumeMode = typeof p.volumeMode === "string" ? p.volumeMode : "direct";
      w.volumeFactor = Number.isFinite(Number(p.volumeFactor)) ? Number(p.volumeFactor) : 1;
      w.sessions = Array.isArray(p.sessions) && p.sessions.length
        ? p.sessions.map((s, i) => {
            const next = {
              dayLabel: normalizeDayLabelZh(s?.dayLabel, i),
              workoutsCount: Number(s?.workoutsCount) || 0,
              workouts: Array.isArray(s?.workouts) ? s.workouts.map(normalizeWorkoutEntry) : [],
              duration: Number(s?.duration) || 0,
              zone: Number(s?.zone) || 0,
              rpe: clamp(Number(s?.rpe) || 1, 1, 10),
              kind: typeof s?.kind === "string" ? s.kind : "Run",
              note: typeof s?.note === "string" ? s.note : "",
            };
            ensureSessionWorkouts(next);
            return next;
          })
        : w.sessions;
      w.races = Array.isArray(p.races)
        ? p.races
            .map((r) => ({
              name: typeof r?.name === "string" ? r.name : "",
              date: typeof r?.date === "string" ? r.date : "",
              distanceKm: Number.isFinite(Number(r?.distanceKm)) && Number(r.distanceKm) > 0 ? Number(r.distanceKm) : null,
              kind: typeof r?.kind === "string" ? r.kind : "",
            }))
            .filter((r) => r.name.trim() && r.date)
        : [];

      if (!w.block && seasonOptions.includes(w.season || "")) {
        w.block = normalizeBlockValue(w.season || "");
        w.season = "";
      }
      w.block = normalizeBlockValue(w.block);
      if (!w.block) w.block = "Base";
      getWeekSessions(w);
      if (Array.isArray(w.sessions) && w.sessions.length) {
        w.sessions.forEach((s) => {
          const raw = typeof s?.note === "string" ? s.note : "";
          if (!raw) return;
          const parts = splitAutoNote(raw);
          if (!parts.has || !parts.legacy) return;
          const next = mergeAutoNote(raw, parts.auto);
          if (next !== raw) s.note = next;
        });
      }
    });
  }
  if (typeof persisted?.selectedWeekIndex === "number") {
    state.selectedWeekIndex = clamp(persisted.selectedWeekIndex, 0, 51);
  }
  if (typeof persisted?.connected === "boolean") {
    state.connected = persisted.connected;
  }
  recomputeFormulaVolumes();
  reassignAllRacesByDate();
  updateHeader();
  wireCalendarSizer();
  renderCalendar();
  renderWeekPicker();
  renderWeekDetails();
  renderCharts();
  wireTabs();
  wireButtons();
  let initialTab = getTabKeyFromHash();
  if (!initialTab) {
    try {
      initialTab = String(localStorage.getItem(ACTIVE_TAB_KEY) || "");
    } catch {
      initialTab = "";
    }
  }
  activateTab(initialTab || "plan", { persist: true, updateHash: false });
  persistState();
}

init();
