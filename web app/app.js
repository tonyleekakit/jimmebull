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
    
    // Ensure formula mode properties are set
    w.volumeMode = "formula";

    // Apply auto factor logic if enabled
    if (w.volumeFactorAuto === true) {
       const autoFactor = defaultVolumeFactorForWeekIndex(i);
       if (Number.isFinite(autoFactor)) w.volumeFactor = autoFactor;
    }
    
    // Safety check for factor
    if (!Number.isFinite(Number(w.volumeFactor))) w.volumeFactor = 1;

    const f = Number(w.volumeFactor);
    const factor = Number.isFinite(f) ? f : 1;
    
    // Always recompute volumeHrs based on past 4-week average * factor
    // The "base" is the average of the last 4 weeks' *effective* volume
    // effective[i] stores the actual number used for subsequent calculations
    let base = computePastMean(effective, i, 4);
    
    // Fallback: If base is 0 (e.g., first week or after long break), 
    // treat base as 1 so that Volume = Factor. 
    // This allows users to set an initial volume by setting the factor directly (or via the volume input).
    if (base <= 0) {
       base = Number.isFinite(state.planBaseVolume) && state.planBaseVolume > 0 ? state.planBaseVolume : 1;
    }

    const out = formatVolumeHrs(base * factor);
    
    w.volumeHrs = out;
    effective[i] = Number(out) || 0;
  }
}

function optimizePlanBaseVolume(targetAnnualHrs) {
  if (!targetAnnualHrs || targetAnnualHrs <= 0) return;

  // We want to find state.planBaseVolume such that sum(volumeHrs) ~= targetAnnualHrs.
  // Since volumeHrs is monotonically increasing with planBaseVolume, we can use binary search.
  
  let low = 0;
  // A safe upper bound: if factors are ~1.0, base ~ target/52. 
  // If factors are small, base needs to be larger. 
  // Let's use a generous upper bound.
  let high = Math.max(100, targetAnnualHrs); 

  // 20 iterations is enough for precision ~ target/2^20
  for (let i = 0; i < 20; i++) {
    const mid = (low + high) / 2;
    state.planBaseVolume = mid;
    recomputeFormulaVolumes();
    
    const currentTotal = state.weeks.reduce((acc, w) => acc + (Number(w.volumeHrs) || 0), 0);
    
    if (Math.abs(currentTotal - targetAnnualHrs) < 1) break;
    
    if (currentTotal < targetAnnualHrs) {
      low = mid;
    } else {
      high = mid;
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

const WORKOUT_HELPER_CONFIG = {
  long: {
    label: "有氧耐力 (RPE 3-4)",
    rpe: 3,
    rpeText: "3-4",
    type: "continuous",
    durationMin: 30,
    durationStep: 5,
    durationMax: 180
  },
  tempo: {
    label: "節奏跑 (RPE 5-6)",
    rpe: 5,
    rpeText: "5-6",
    type: "continuous",
    durationMin: 25,
    durationStep: 5,
    durationMax: 60
  },
  threshold: {
    label: "乳酸閾值 (RPE 7-8)",
    rpe: 7,
    rpeText: "7-8",
    type: "interval",
    durationMin: 6,
    durationMax: 12,
    durationStep: 0.5,
    setsOptions: [3, 4, 5],
    restRatio: 0.25,
    restText: "跑休 4:1"
  },
  vo2max: {
    label: "最大攝氧量 (RPE 8-9)",
    rpe: 9,
    rpeText: "8-9",
    type: "interval",
    durationMin: 1,
    durationMax: 5,
    durationStep: 0.25,
    minTotal: 5,
    maxTotal: 15,
    restRatios: [1, 0.5],
    restText: "跑休 1:1 或 1:0.5"
  },
  anaerobic: {
    label: "無氧耐力 (RPE 9-10)",
    rpe: 10,
    rpeText: "9-10",
    type: "interval",
    durationMin: 0.5,
    durationMax: 1,
    durationStep: 5 / 60,
    minTotal: 5,
    maxTotal: 15,
    restRatio: 1,
    restText: "跑休 1:1"
  }
};

function openWorkoutHelperModal(session, onImport) {
  const formatDuration = (m) => {
    if (m < 1) return `${Math.round(m * 60)} 秒`;
    const mins = Math.floor(m);
    const secs = Math.round((m - mins) * 60);
    return secs === 0 ? `${mins} 分鐘` : `${mins} 分 ${secs} 秒`;
  };

  // Use existing modal structure styles (modal--auth, authForm, etc.)
  const overlay = el("div", "overlay");
  const modal = el("div", "modal modal--auth");
  
  // Header
  const header = el("div", "authModal__header");
  const headerText = el("div", "authModal__headerText");
  const title = el("div", "modal__title", "課表小助手");
  const subtitle = el("div", "modal__subtitle", "自動計算並匯入訓練內容");
  headerText.appendChild(title);
  headerText.appendChild(subtitle);
  
  // Close Button
  const closeIcon = svgEl("svg", { viewBox: "0 0 24 24", fill: "none", stroke: "currentColor", "stroke-width": "2" });
  closeIcon.appendChild(svgEl("path", { d: "M6 6l12 12" }));
  closeIcon.appendChild(svgEl("path", { d: "M18 6L6 18" }));
  const closeIconBtn = document.createElement("button");
  closeIconBtn.type = "button";
  closeIconBtn.className = "iconBtn authModal__close";
  closeIconBtn.appendChild(closeIcon);
  closeIconBtn.addEventListener("click", () => overlay.remove());

  header.appendChild(headerText);
  header.appendChild(closeIconBtn);

  // Form
  const form = el("div", "authForm"); 
  
  // Helper to create row (styled like auth fields)
  const createRow = (label, input) => {
    const row = el("label", "authField");
    const l = el("span", "authField__label", label);
    row.appendChild(l);
    input.classList.add("input", "authInput"); 
    input.style.height = "40px";
    row.appendChild(input);
    form.appendChild(row);
  };

  // 1. Zone Select
  const zoneSelect = document.createElement("select");
  Object.keys(WORKOUT_HELPER_CONFIG).forEach(key => {
    const opt = document.createElement("option");
    opt.value = key;
    opt.textContent = WORKOUT_HELPER_CONFIG[key].label;
    zoneSelect.appendChild(opt);
  });

  // 2. Duration Select
  const durationSelect = document.createElement("select");

  // 3. Sets Select
  const setsSelect = document.createElement("select");

  // 4. Rest Select
  const restSelect = document.createElement("select");

  // Render rows
  createRow("訓練區間", zoneSelect);
  createRow("時長 (每組/總時長)", durationSelect);
  createRow("組數", setsSelect);
  createRow("組間休息", restSelect);

  // Actions (Footer)
  const actions = el("div", "authActions");
  actions.style.marginTop = "20px";
  
  const cancelBtn = el("button", "btn authBackBtn", "取消");
  cancelBtn.type = "button";
  cancelBtn.addEventListener("click", () => overlay.remove());

  const importBtn = el("button", "btn btn--primary authSubmitBtn", "匯入課表");
  importBtn.type = "button";
  importBtn.addEventListener("click", () => {
    const zoneKey = zoneSelect.value;
    const config = WORKOUT_HELPER_CONFIG[zoneKey];
    
    const repDur = Number(durationSelect.value);
    const sets = Number(setsSelect.value);
    const rest = Number(restSelect.value);
    
    let totalDuration = 0;
    let noteText = "";

    if (config.type === "continuous") {
      totalDuration = repDur;
      noteText = `${config.label.split(" ")[0]} ${repDur} 分鐘`;
    } else {
      totalDuration = Math.round((repDur + rest) * sets);
      
      const repStr = formatDuration(repDur);
      const restStr = formatDuration(rest);
      noteText = `熱身 10-15 分鐘\n${sets} x ${repStr} 間歇, 休息 ${restStr}\n緩和 10-15 分鐘`;
    }

    // Add VDOT distance info if available
    if (typeof state !== "undefined" && Number.isFinite(state.vdot) && ["tempo", "threshold", "vo2max", "anaerobic"].includes(zoneKey)) {
      const vdot = state.vdot;
      const zoneRanges = {
        tempo: { lo: 0.78, hi: 0.86, mid: 0.82 },
        threshold: { lo: 0.86, hi: 0.92, mid: 0.89 },
        vo2max: { lo: 0.92, hi: 0.99, mid: 0.955 },
        anaerobic: { lo: 0.99, hi: 1.06, mid: 1.025 }
      };
      const z = zoneRanges[zoneKey];
      if (z) {
        const paceSlow = paceSecondsPerKmFromVdotFraction(vdot, z.lo);
        const paceFast = paceSecondsPerKmFromVdotFraction(vdot, z.hi);
        if (Number.isFinite(paceSlow) && Number.isFinite(paceFast)) {
          const p1 = formatPaceFromSecondsPerKm(Math.max(paceSlow, paceFast));
          const p2 = formatPaceFromSecondsPerKm(Math.min(paceSlow, paceFast));
          const rangeStr = `${p1}-${p2}`;
          
          const midPace = paceSecondsPerKmFromVdotFraction(vdot, z.mid);
          if (Number.isFinite(midPace) && midPace > 0) {
            const speedMetersPerMin = 1000 / (midPace / 60);
            const estMeters = repDur * speedMetersPerMin;
            const trackDists = [100, 200, 300, 400, 600, 800, 1000, 1200, 1500, 1600, 2000, 3000, 4000, 5000, 6000, 8000, 10000];
            const closest = trackDists.reduce((prev, curr) => Math.abs(curr - estMeters) < Math.abs(prev - estMeters) ? curr : prev);
            
            if (config.type === "continuous") {
              noteText = `${config.label.split(" ")[0]} ${repDur} 分鐘 (約 ${closest}m @ ${rangeStr}/km)`;
            } else {
              const repStr = formatDuration(repDur);
              const restStr = formatDuration(rest);
              noteText = `熱身 10-15 分鐘\n${sets} x ${repStr} 間歇 (約 ${closest}m @ ${rangeStr}/km), 休息 ${restStr}\n緩和 10-15 分鐘`;
            }
          }
        }
      }
    }

    lockScrollPosition(() => {
      pushHistory();
      
      if (session.workoutsCount < 1) {
        session.workoutsCount = 1;
        ensureSessionWorkouts(session);
      }
      
      if (session.workouts[0]) {
        session.workouts[0].duration = Math.round(totalDuration);
        session.workouts[0].rpe = config.rpe;
      }
      
      session.note = noteText;
      
      if (onImport) onImport();
      overlay.remove();
    });
  });
  
  actions.appendChild(cancelBtn);
  actions.appendChild(importBtn);

  form.appendChild(actions);

  modal.appendChild(header);
  modal.appendChild(form);
  overlay.appendChild(modal);

  // Logic
  const updateSetsAndRest = (config) => {
    setsSelect.innerHTML = "";
    restSelect.innerHTML = "";
    
    const repDur = Number(durationSelect.value);
    if (!Number.isFinite(repDur) || repDur <= 0) return;
    
    // Calculate Sets options
    if (config.setsOptions) {
      config.setsOptions.forEach(s => {
        const opt = document.createElement("option");
        opt.value = s;
        opt.textContent = `${s} 組`;
        setsSelect.appendChild(opt);
      });
    } else {
      const minSets = Math.ceil(config.minTotal / repDur);
      const maxSets = Math.floor(config.maxTotal / repDur);
      const start = Math.max(1, minSets);
      const end = Math.max(start, maxSets); 
      
      for (let s = start; s <= end; s++) {
        const opt = document.createElement("option");
        opt.value = s;
        opt.textContent = `${s} 組 (總時長 ${formatDuration(s * repDur)})`;
        setsSelect.appendChild(opt);
      }
    }

    // Calculate Rest options
    const ratios = Array.isArray(config.restRatios) 
      ? config.restRatios 
      : [config.restRatio || 1];
      
    // Handle Min/Max if range provided (legacy)
    if (!config.restRatios && (config.restRatioMin || config.restRatioMax)) {
      const restMin = repDur * (config.restRatioMin || config.restRatio);
      const restMax = repDur * (config.restRatioMax || config.restRatio);
      
      const rOpt = document.createElement("option");
      rOpt.value = restMin;
      rOpt.textContent = `${formatDuration(restMin)} (${config.restText})`;
      restSelect.appendChild(rOpt);

      if (restMax > restMin) {
        const rOpt2 = document.createElement("option");
        rOpt2.value = restMax;
        rOpt2.textContent = formatDuration(restMax);
        restSelect.appendChild(rOpt2);
      }
      return;
    }

    // Standard list of ratios
    ratios.forEach(ratio => {
      const val = repDur * ratio;
      const opt = document.createElement("option");
      opt.value = val;
      
      let ratioLabel = "";
      if (Math.abs(ratio - 1) < 0.01) ratioLabel = "1:1";
      else if (Math.abs(ratio - 0.5) < 0.01) ratioLabel = "1:0.5";
      else if (Math.abs(ratio - 0.25) < 0.01) ratioLabel = "4:1";
      else ratioLabel = `1:${ratio}`;
      
      const desc = config.restRatios ? `(跑休 ${ratioLabel})` : `(${config.restText})`;
      
      opt.textContent = `${formatDuration(val)} ${desc}`;
      restSelect.appendChild(opt);
    });
  };

  const updateFields = () => {
    const zoneKey = zoneSelect.value;
    const config = WORKOUT_HELPER_CONFIG[zoneKey];
    
    durationSelect.innerHTML = "";
    setsSelect.innerHTML = "";
    restSelect.innerHTML = "";

    const step = config.durationStep || 1;
    for (let m = config.durationMin; m <= config.durationMax + 0.0001; m += step) {
      const opt = document.createElement("option");
      opt.value = m;
      opt.textContent = formatDuration(m);
      durationSelect.appendChild(opt);
    }

    if (config.type === "continuous") {
      const sOpt = document.createElement("option");
      sOpt.value = "1";
      sOpt.textContent = "1 組";
      setsSelect.appendChild(sOpt);
      setsSelect.disabled = true;

      const rOpt = document.createElement("option");
      rOpt.value = "0";
      rOpt.textContent = "—";
      restSelect.appendChild(rOpt);
      restSelect.disabled = true;

    } else {
      setsSelect.disabled = false;
      restSelect.disabled = false;
      updateSetsAndRest(config);
    }
  };

  zoneSelect.addEventListener("change", updateFields);
  durationSelect.addEventListener("change", () => {
    const config = WORKOUT_HELPER_CONFIG[zoneSelect.value];
    if (config.type === "interval") updateSetsAndRest(config);
  });

  // Init
  updateFields();

  document.body.appendChild(overlay);
}

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
    if (k) {
      t.classList.toggle("is-active", k === key);
    }
  });

  // Only manage panels if we are in the main SPA mode (detected by calendarShell)
  if (document.getElementById("calendarShell")) {
    document.querySelectorAll(".tabPanel").forEach((p) => p.classList.remove("is-active"));
    const panel = document.getElementById(`tab-${key}`);
    if (panel) panel.classList.add("is-active");
  }

  if (options?.persist) {
    try {
      localStorage.setItem(ACTIVE_TAB_KEY, key);
    } catch {}
  }

  if (options?.updateHash) {
    const nextHash = `#${key}`;
    if (window.location.hash !== nextHash) window.location.hash = nextHash;
  }

  // Update Page Title for basic SEO (SPA)
  const baseTitle = "訓練調控";
  const titles = {
    plan: "訓練調控 | 年度訓練計畫",
    howto: "如何使用 | 訓練調控",
    pace: "跑步配速計算機 | VDOT 訓練區間計算 | 訓練調控",
    blog: "部落格 | 跑步科學文章 | 訓練調控"
  };
  document.title = titles[key] || baseTitle;

  if (key === "charts") renderCharts();
  if (key === "blog") renderBlog();
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

  // 不再將 VDOT 寫入訓練調控頁面或用戶狀態
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

function wireHowTo() {
  const readShot = (key) => {
    try {
      return String(localStorage.getItem(`${HOWTO_SHOT_PREFIX}${key}`) || "");
    } catch {
      return "";
    }
  };

  const setShotUI = (wrap, dataUrl) => {
    const img = wrap.querySelector(".howtoShot__img");
    if (img) img.src = dataUrl || "";
    wrap.classList.toggle("has-image", Boolean(dataUrl));
  };

  document.querySelectorAll("[data-howto-shot-key]").forEach((wrap) => {
    const key = String(wrap.getAttribute("data-howto-shot-key") || "").trim();
    if (!key) return;
    const stored = readShot(key);
    if (stored) {
      setShotUI(wrap, stored);
      return;
    }

    const img = wrap.querySelector(".howtoShot__img");
    const staticSrc = img ? String(img.getAttribute("data-static-src") || "").trim() : "";
    if (!img || !staticSrc) {
      setShotUI(wrap, "");
      return;
    }

    if (img.dataset.howtoStaticWired !== "1") {
      img.dataset.howtoStaticWired = "1";
      img.addEventListener("error", () => setShotUI(wrap, ""));
    }
    setShotUI(wrap, staticSrc);
  });
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

function normalizeRacePriorityValue(value) {
  const raw = String(value || "").trim();
  if (!raw) return "";
  if (raw === "重要") return "A";
  if (raw === "不重要") return "C";
  const v = raw.toUpperCase();
  if (v === "A") return "A";
  if (v === "B") return "A";
  if (v === "C") return "C";
  return "";
}

function racePriorityLabelZh(value) {
  const v = normalizeRacePriorityValue(value);
  if (v === "A") return "重要";
  if (v === "C") return "不重要";
  return "";
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

function isRaceWeekByIndex(weekIndex) {
  const idx = clamp(Number(weekIndex) || 0, 0, 51);
  const w = state?.weeks?.[idx];
  if (!w) return false;
  return Array.isArray(w.races) && w.races.length > 0;
}

function defaultVolumeFactorForWeekIndex(weekIndex) {
  const idx = clamp(Number(weekIndex) || 0, 0, 51);
  const w = state?.weeks?.[idx];
  const block = normalizeBlockValue(w?.block || "") || "Base";

  if (block === "Peak") {
    if (isRaceWeekByIndex(idx)) return 0.6;
    if (idx < 51 && isRaceWeekByIndex(idx + 1)) return 0.8;
    return 1;
  }

  if (block === "Base") return 1.2;
  if (block === "Deload") return 0.6;
  if (block === "Build") return 1.1;
  if (block === "Transition") return 0.5;
  return 1;
}

function refreshAutoVolumeFactors() {
  if (!state || !Array.isArray(state.weeks) || state.weeks.length !== 52) return false;
  let changed = false;
  for (let i = 0; i < 52; i++) {
    const w = state.weeks[i];
    if (!w) continue;
    // Always refresh auto factor if enabled, regardless of mode
    if (w.volumeFactorAuto !== true) continue;
    const next = defaultVolumeFactorForWeekIndex(i);
    if (Number(w.volumeFactor) !== next) {
      w.volumeFactor = next;
      changed = true;
    }
  }
  if (changed) recomputeFormulaVolumes();
  return changed;
}

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
    const p = normalizeRacePriorityValue(w.priority);
    if (p !== "A") return;
    if (!Array.isArray(w.races) || w.races.length === 0) return;
    const idx = clamp(Number(w.index), 0, 51);

    setBlock(idx, "Peak");
    setBlock(idx - 1, "Peak");
    setBlock(idx + 1, "Transition");
    // 賽前第 5 週為減量期
    setBlock(idx - 5, "Deload");
    // 其餘賽前週為建立期（排除第 5 週）
    for (let d = 2; d <= 8; d++) {
      if (d === 5) continue;
      setBlock(idx - d, "Build");
    }
    return;
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

  // Rule: The week before any Build-start is Deload
  for (let j = 0; j < 52; j++) {
    if (out[j] !== "Build") continue;
    const prev = j - 1;
    if (prev >= 0 && out[prev] !== "Build" && out[prev] !== "Peak" && out[prev] !== "Transition") {
      setBlock(prev, "Deload");
    }
  }

  // Rule: The three weeks before any Deload are Base (unless overridden by Peak/Transition/Build)
  for (let i = 0; i < 52; i++) {
    if (out[i] !== "Deload") continue;
    for (let t = 1; t <= 3; t++) {
      const k = i - t;
      if (k >= 0 && out[k] !== "Build" && out[k] !== "Peak" && out[k] !== "Transition") {
        setBlock(k, "Base");
      }
    }
  }

  // Fixed alignment 3:1 anchored at each Build-start (A/B races):
  // For each Build-start j, align the preceding segment [prevStart+1 .. j-1]
  // so that j-1 is Deload, and counting backwards every 4th week is Deload, others Base.
  const buildStarts = [];
  for (let j = 0; j < 52; j++) {
    if (out[j] === "Build") {
      const prev = j - 1;
      const isStart = prev < 0 || out[prev] !== "Build";
      if (isStart) buildStarts.push(j);
    }
  }
  if (buildStarts.length) {
    for (let idx = 0; idx < buildStarts.length; idx++) {
      const j = buildStarts[idx];
      const prevStart = idx > 0 ? buildStarts[idx - 1] : -1;
      const segStart = Math.max(0, prevStart + 1);
      const segEnd = Math.max(0, j - 1);
      const anchorDeload = j - 1;
      for (let k = segStart; k <= segEnd; k++) {
        if (out[k] === "Build" || out[k] === "Peak" || out[k] === "Transition") continue;
        const diff = anchorDeload - k;
        if (diff % 4 === 0) {
          setBlock(k, "Deload");
        } else {
          setBlock(k, "Base");
        }
      }
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
    if (w.blockAuto === false) continue;
    const next = blocks[i] || "Base";
    if (normalizeBlockValue(w.block || "") !== next) {
      w.block = next;
      w.blockAuto = true;
      // If block changes by rule, reset volume factor to auto
      w.volumeFactorAuto = true;
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

function baseFocusGeneric(totalCycles, cycleIndex) {
  const t = Number(totalCycles);
  const c = Number(cycleIndex);
  if (!Number.isFinite(t) || !Number.isFinite(c) || t < 2 || c < 1 || c > t) return "";

  if (t >= 13) {
    if (c >= 1 && c <= 3) return "Anaerobic";
    if (c >= 4 && c <= 6) return "VO2Max";
    if (c >= 7 && c <= 9) return "Threshold";
    return "Tempo";
  }
  if (t === 12) {
    if (c >= 1 && c <= 3) return "Anaerobic";
    if (c >= 4 && c <= 6) return "VO2Max";
    if (c >= 7 && c <= 9) return "Threshold";
    return "Tempo";
  }
  if (t === 11) {
    if (c >= 1 && c <= 2) return "Anaerobic";
    if (c >= 3 && c <= 5) return "VO2Max";
    if (c >= 6 && c <= 8) return "Threshold";
    return "Tempo";
  }
  if (t === 10) {
    if (c >= 1 && c <= 2) return "Anaerobic";
    if (c >= 3 && c <= 4) return "VO2Max";
    if (c >= 5 && c <= 7) return "Threshold";
    return "Tempo";
  }
  if (t === 9) {
    if (c >= 1 && c <= 2) return "Anaerobic";
    if (c >= 3 && c <= 4) return "VO2Max";
    if (c >= 5 && c <= 6) return "Threshold";
    return "Tempo";
  }
  if (t === 8) {
    if (c >= 1 && c <= 2) return "Anaerobic";
    if (c >= 3 && c <= 4) return "VO2Max";
    if (c >= 5 && c <= 6) return "Threshold";
    return "Tempo";
  }
  if (t === 7) {
    if (c === 1) return "Anaerobic";
    if (c >= 2 && c <= 3) return "VO2Max";
    if (c >= 4 && c <= 5) return "Threshold";
    return "Tempo";
  }
  if (t === 6) {
    if (c === 1) return "Anaerobic";
    if (c === 2) return "VO2Max";
    if (c >= 3 && c <= 4) return "Threshold";
    return "Tempo";
  }
  if (t === 5) {
    if (c === 1) return "Anaerobic";
    if (c === 2) return "VO2Max";
    if (c === 3) return "Threshold";
    return "Tempo";
  }
  if (t === 4) {
    if (c === 1) return "Anaerobic";
    if (c === 2) return "VO2Max";
    if (c === 3) return "Threshold";
    return "Tempo";
  }
  if (t === 3) {
    if (c === 1) return "VO2Max";
    if (c === 2) return "Threshold";
    return "Tempo";
  }
  if (t === 2) {
    if (c === 1) return "VO2Max";
    return "Threshold";
  }
  return "Tempo";
}

function cycleIndexInfoForWeek(weekIndex, raceWeekIndex) {
  const j = Number(raceWeekIndex);
  if (!Number.isFinite(j) || j <= 0) return null;

  // 找出此賽事「前一場重要賽事」的結束週，做為本次循環計算的起點
  // 若無前一場賽事，預設從 0 開始
  let startIndex = 0;
  for (let r = j - 1; r >= 0; r--) {
    const wk = state.weeks[r];
    const pr = normalizeRacePriorityValue(wk?.priority);
    if (pr === "A") {
      startIndex = r + 1;
      break;
    }
  }

  const cycles = [];
  // 從 startIndex 開始掃描到 raceWeekIndex (j)
  for (let k = startIndex; k < j; k++) {
    const wk = state.weeks[k];
    const isDeload = normalizeBlockValue(wk?.block || "") === "Deload";
    if (!isDeload) continue;
    let bcount = 0;
    // 往前數 Base，但不超過 startIndex
    for (let p = k - 1; p >= startIndex; p--) {
      const bp = state.weeks[p];
      if (normalizeBlockValue(bp?.block || "") === "Base") bcount++;
      else break;
    }
    const isFirst = cycles.length === 0;
    if (bcount >= 3 || (isFirst && bcount >= 1)) {
      const start = k - bcount;
      const end = k;
      cycles.push({ start, end });
    }
  }
  if (!cycles.length) return null;
  for (let idx = 0; idx < cycles.length; idx++) {
    const seg = cycles[idx];
    if (Number(weekIndex) >= seg.start && Number(weekIndex) <= seg.end) {
      return { totalCycles: cycles.length, cycleIndex: idx + 1 };
    }
  }
  // 若該週不在任何循環內（例如最後幾個 Base 週後面接 Peak 而非 Deload），
  // 視同最後一個循環或依照剩餘週數處理？
  // 這裡回傳 null 會導致預設行為（僅 Aerobic Endurance），
  // 但依照截圖狀況，可能會有 orphan base weeks。
  // 若希望 orphan weeks 也被分配重點，可視為下一個循環（cycleIndex + 1）？
  // 但目前邏輯是：若不在循環內，回傳 cycleIndex: null -> baseFocusFor5k 預設 return Tempo (若 c 不在範圍)。
  // 為了讓 orphan weeks (如 48-49) 也能分配，我們可以檢查它是否屬於「最後一個未完成的循環」。
  
  // 檢查是否為最後一段 Base (無 Deload)
  if (Number(weekIndex) > cycles[cycles.length - 1].end && Number(weekIndex) < j) {
    // 檢查該週是否為 Base
    if (normalizeBlockValue(state.weeks[weekIndex]?.block || "") === "Base") {
       // 視為第 N+1 個循環
       return { totalCycles: cycles.length + 1, cycleIndex: cycles.length + 1 };
    }
  }

  return { totalCycles: cycles.length, cycleIndex: null };
}

function baseFocusFor10k(totalCycles, cycleIndex) {
  const t = Number(totalCycles);
  const c = Number(cycleIndex);
  if (!Number.isFinite(t) || !Number.isFinite(c) || t < 2 || c < 1 || c > t) return "";
  if (t === 13) {
    if (c >= 1 && c <= 3) return "Anaerobic";
    if (c >= 4 && c <= 6) return "VO2Max";
    if (c >= 7 && c <= 9) return "Threshold";
    return "Tempo";
  }
  if (t === 12) {
    if (c >= 1 && c <= 3) return "Anaerobic";
    if (c >= 7 && c <= 9) return "VO2Max";
    if (c >= 10 && c <= 12) return "Threshold";
    return (c >= 4 && c <= 6) ? "Tempo" : "Tempo";
  }
  if (t === 11) {
    if (c >= 1 && c <= 2) return "Anaerobic";
    if (c >= 6 && c <= 8) return "VO2Max";
    if (c >= 9 && c <= 11) return "Threshold";
    return (c >= 3 && c <= 5) ? "Tempo" : "Tempo";
  }
  if (t === 10) {
    if (c >= 1 && c <= 2) return "Anaerobic";
    if (c >= 5 && c <= 7) return "VO2Max";
    if (c >= 8 && c <= 10) return "Threshold";
    return (c >= 3 && c <= 4) ? "Tempo" : "Tempo";
  }
  if (t === 9) {
    if (c >= 1 && c <= 2) return "Anaerobic";
    if (c >= 5 && c <= 6) return "VO2Max";
    if (c >= 7 && c <= 9) return "Threshold";
    return (c >= 3 && c <= 4) ? "Tempo" : "Tempo";
  }
  if (t === 8) {
    if (c >= 1 && c <= 2) return "Anaerobic";
    if (c >= 5 && c <= 6) return "VO2Max";
    if (c >= 7 && c <= 8) return "Threshold";
    return (c >= 3 && c <= 4) ? "Tempo" : "Tempo";
  }
  if (t === 7) {
    if (c === 1) return "Anaerobic";
    if (c >= 4 && c <= 5) return "VO2Max";
    if (c >= 6 && c <= 7) return "Threshold";
    return (c >= 2 && c <= 3) ? "Tempo" : "Tempo";
  }
  if (t === 6) {
    if (c === 1) return "Anaerobic";
    if (c >= 3 && c <= 4) return "VO2Max";
    if (c >= 5 && c <= 6) return "Threshold";
    return c === 2 ? "Tempo" : "Tempo";
  }
  if (t === 5) {
    if (c === 1) return "Anaerobic";
    if (c === 3) return "VO2Max";
    if (c >= 4 && c <= 5) return "Threshold";
    return c === 2 ? "Tempo" : "Tempo";
  }
  if (t === 4) {
    if (c === 1) return "Anaerobic";
    if (c === 3) return "VO2Max";
    if (c === 4) return "Threshold";
    return c === 2 ? "Tempo" : "Tempo";
  }
  if (t === 3) {
    if (c === 2) return "VO2Max";
    if (c === 3) return "Threshold";
    return "Tempo";
  }
  if (t === 2) {
    if (c === 1) return "Anaerobic";
    return "Tempo";
  }
  return "Tempo";
}

function baseFocusFor15k(totalCycles, cycleIndex) {
  const t = Number(totalCycles);
  const c = Number(cycleIndex);
  if (!Number.isFinite(t) || !Number.isFinite(c) || t < 2 || c < 1 || c > t) return "";
  
  if (t === 13) {
    if (c >= 1 && c <= 3) return "Anaerobic";
    if (c >= 4 && c <= 6) return "VO2Max";
    if (c >= 7 && c <= 9) return "Tempo";
    if (c >= 10 && c <= 13) return "Threshold";
    return "Tempo";
  }
  if (t === 12) {
    if (c >= 1 && c <= 3) return "Anaerobic";
    if (c >= 4 && c <= 6) return "VO2Max";
    if (c >= 7 && c <= 9) return "Tempo";
    if (c >= 10 && c <= 12) return "Threshold";
    return "Tempo";
  }
  if (t === 11) {
    if (c >= 1 && c <= 2) return "Anaerobic";
    if (c >= 3 && c <= 5) return "VO2Max";
    if (c >= 6 && c <= 8) return "Tempo";
    if (c >= 9 && c <= 11) return "Threshold";
    return "Tempo";
  }
  if (t === 10) {
    if (c >= 1 && c <= 2) return "Anaerobic";
    if (c >= 3 && c <= 4) return "VO2Max";
    if (c >= 5 && c <= 7) return "Tempo";
    if (c >= 8 && c <= 10) return "Threshold";
    return "Tempo";
  }
  if (t === 9) {
    if (c >= 1 && c <= 2) return "Anaerobic";
    if (c >= 3 && c <= 4) return "VO2Max";
    if (c >= 5 && c <= 6) return "Tempo";
    if (c >= 7 && c <= 9) return "Threshold";
    return "Tempo";
  }
  if (t === 8) {
    if (c >= 1 && c <= 2) return "Anaerobic";
    if (c >= 3 && c <= 4) return "VO2Max";
    if (c >= 5 && c <= 6) return "Tempo";
    if (c >= 7 && c <= 8) return "Threshold";
    return "Tempo";
  }
  if (t === 7) {
    if (c === 1) return "Anaerobic";
    if (c >= 2 && c <= 3) return "VO2Max";
    if (c >= 4 && c <= 5) return "Tempo";
    if (c >= 6 && c <= 7) return "Threshold";
    return "Tempo";
  }
  if (t === 6) {
    if (c === 1) return "Anaerobic";
    if (c === 2) return "VO2Max";
    if (c >= 3 && c <= 4) return "Tempo";
    if (c >= 5 && c <= 6) return "Threshold";
    return "Tempo";
  }
  if (t === 5) {
    if (c === 1) return "Anaerobic";
    if (c === 2) return "VO2Max";
    if (c === 3) return "Tempo";
    if (c >= 4 && c <= 5) return "Threshold";
    return "Tempo";
  }
  if (t === 4) {
    if (c === 1) return "Anaerobic";
    if (c === 2) return "VO2Max";
    if (c === 3) return "Tempo";
    if (c === 4) return "Threshold";
    return "Tempo";
  }
  if (t === 3) {
    if (c === 1) return "VO2Max";
    if (c === 2) return "Tempo";
    if (c === 3) return "Threshold";
    return "Tempo";
  }
  if (t === 2) {
    if (c === 1) return "Anaerobic";
    if (c === 2) return "VO2Max";
    return "Tempo";
  }
  return "Tempo";
}

function baseFocusForHalfMarathon(totalCycles, cycleIndex) {
  const t = Number(totalCycles);
  const c = Number(cycleIndex);
  if (!Number.isFinite(t) || !Number.isFinite(c) || t < 2 || c < 1 || c > t) return "";
  
  if (t === 13) {
    if (c >= 1 && c <= 3) return "Anaerobic";
    if (c >= 4 && c <= 6) return "VO2Max";
    if (c >= 7 && c <= 9) return "Threshold";
    if (c >= 10 && c <= 13) return "Tempo";
    return "Tempo";
  }
  if (t === 12) {
    if (c >= 1 && c <= 3) return "Anaerobic";
    if (c >= 4 && c <= 6) return "VO2Max";
    if (c >= 7 && c <= 9) return "Threshold";
    if (c >= 10 && c <= 12) return "Tempo";
    return "Tempo";
  }
  if (t === 11) {
    if (c >= 1 && c <= 2) return "Anaerobic";
    if (c >= 3 && c <= 5) return "VO2Max";
    if (c >= 6 && c <= 8) return "Threshold";
    if (c >= 9 && c <= 11) return "Tempo";
    return "Tempo";
  }
  if (t === 10) {
    if (c >= 1 && c <= 2) return "Anaerobic";
    if (c >= 3 && c <= 4) return "VO2Max";
    if (c >= 5 && c <= 7) return "Threshold";
    if (c >= 8 && c <= 10) return "Tempo";
    return "Tempo";
  }
  if (t === 9) {
    if (c >= 1 && c <= 2) return "Anaerobic";
    if (c >= 3 && c <= 4) return "VO2Max";
    if (c >= 5 && c <= 6) return "Threshold";
    if (c >= 7 && c <= 9) return "Tempo";
    return "Tempo";
  }
  if (t === 8) {
    if (c >= 1 && c <= 2) return "Anaerobic";
    if (c >= 3 && c <= 4) return "VO2Max";
    if (c >= 5 && c <= 6) return "Threshold";
    if (c >= 7 && c <= 8) return "Tempo";
    return "Tempo";
  }
  if (t === 7) {
    if (c === 1) return "Anaerobic";
    if (c >= 2 && c <= 3) return "VO2Max";
    if (c >= 4 && c <= 5) return "Threshold";
    if (c >= 6 && c <= 7) return "Tempo";
    return "Tempo";
  }
  if (t === 6) {
    if (c === 1) return "Anaerobic";
    if (c === 2) return "VO2Max";
    if (c >= 3 && c <= 4) return "Threshold";
    if (c >= 5 && c <= 6) return "Tempo";
    return "Tempo";
  }
  if (t === 5) {
    if (c === 1) return "Anaerobic";
    if (c === 2) return "VO2Max";
    if (c === 3) return "Threshold";
    if (c >= 4 && c <= 5) return "Tempo";
    return "Tempo";
  }
  if (t === 4) {
    if (c === 1) return "Anaerobic";
    if (c === 2) return "VO2Max";
    if (c === 3) return "Threshold";
    if (c === 4) return "Tempo";
    return "Tempo";
  }
  if (t === 3) {
    if (c === 1) return "VO2Max";
    if (c === 2) return "Threshold";
    if (c === 3) return "Tempo";
    return "Tempo";
  }
  if (t === 2) {
    if (c === 1) return "Anaerobic";
    if (c === 2) return "VO2Max";
    return "Tempo";
  }
  return "Tempo";
}

function baseFocusForMarathon(totalCycles, cycleIndex) {
  const t = Number(totalCycles);
  const c = Number(cycleIndex);
  if (!Number.isFinite(t) || !Number.isFinite(c) || t < 2 || c < 1 || c > t) return "";
  
  if (t === 13) {
    if (c >= 1 && c <= 3) return "Anaerobic";
    if (c >= 4 && c <= 6) return "VO2Max";
    if (c >= 7 && c <= 9) return "Threshold";
    if (c >= 10 && c <= 13) return "Tempo";
    return "Tempo";
  }
  if (t === 12) {
    if (c >= 1 && c <= 3) return "Anaerobic";
    if (c >= 4 && c <= 6) return "VO2Max";
    if (c >= 7 && c <= 9) return "Threshold";
    if (c >= 10 && c <= 12) return "Tempo";
    return "Tempo";
  }
  if (t === 11) {
    if (c >= 1 && c <= 2) return "Anaerobic";
    if (c >= 3 && c <= 5) return "VO2Max";
    if (c >= 6 && c <= 8) return "Threshold";
    if (c >= 9 && c <= 11) return "Tempo";
    return "Tempo";
  }
  if (t === 10) {
    if (c >= 1 && c <= 2) return "Anaerobic";
    if (c >= 3 && c <= 4) return "VO2Max";
    if (c >= 5 && c <= 7) return "Threshold";
    if (c >= 8 && c <= 10) return "Tempo";
    return "Tempo";
  }
  if (t === 9) {
    if (c >= 1 && c <= 2) return "Anaerobic";
    if (c >= 3 && c <= 4) return "VO2Max";
    if (c >= 5 && c <= 6) return "Threshold";
    if (c >= 7 && c <= 9) return "Tempo";
    return "Tempo";
  }
  if (t === 8) {
    if (c >= 1 && c <= 2) return "Anaerobic";
    if (c >= 3 && c <= 4) return "VO2Max";
    if (c >= 5 && c <= 6) return "Threshold";
    if (c >= 7 && c <= 8) return "Tempo";
    return "Tempo";
  }
  if (t === 7) {
    if (c === 1) return "Anaerobic";
    if (c >= 2 && c <= 3) return "VO2Max";
    if (c >= 4 && c <= 5) return "Threshold";
    if (c >= 6 && c <= 7) return "Tempo";
    return "Tempo";
  }
  if (t === 6) {
    if (c === 1) return "Anaerobic";
    if (c === 2) return "VO2Max";
    if (c >= 3 && c <= 4) return "Threshold";
    if (c >= 5 && c <= 6) return "Tempo";
    return "Tempo";
  }
  if (t === 5) {
    if (c === 1) return "Anaerobic";
    if (c === 2) return "VO2Max";
    if (c === 3) return "Threshold";
    if (c >= 4 && c <= 5) return "Tempo";
    return "Tempo";
  }
  if (t === 4) {
    if (c === 1) return "Anaerobic";
    if (c === 2) return "VO2Max";
    if (c === 3) return "Threshold";
    if (c === 4) return "Tempo";
    return "Tempo";
  }
  if (t === 3) {
    if (c === 1) return "VO2Max";
    if (c === 2) return "Threshold";
    if (c === 3) return "Tempo";
    return "Tempo";
  }
  if (t === 2) {
    if (c === 1) return "Anaerobic";
    if (c === 2) return "VO2Max";
    return "Tempo";
  }
  return "Tempo";
}

function getGenericCycleInfo(weekIndex) {
  const idx = Number(weekIndex);
  if (!Number.isFinite(idx) || idx < 0 || idx > 51) return null;
  
  const isBaseOrDeload = (k) => {
    const w = state.weeks[k];
    if (!w) return false;
    const b = normalizeBlockValue(w.block || "");
    return b === "Base" || b === "Deload";
  };

  if (!isBaseOrDeload(idx)) return null;

  // Find start of this contiguous Base/Deload sequence
  let start = idx;
  while (start > 0 && isBaseOrDeload(start - 1)) {
    start--;
  }

  // Find end
  let end = idx;
  while (end < 51 && isBaseOrDeload(end + 1)) {
    end++;
  }

  // Count cycles in this sequence
  let totalCycles = 0;
  let currentCycleIndex = 0;
  let c = 0;
  let cycleStart = start;
  
  for (let k = start; k <= end; k++) {
    const w = state.weeks[k];
    const b = normalizeBlockValue(w.block || "");
    
    // Cycle ends if we hit Deload OR we hit the end
    if (b === "Deload" || k === end) {
      c++;
      if (idx >= cycleStart && idx <= k) {
        currentCycleIndex = c;
      }
      cycleStart = k + 1;
    }
  }
  
  totalCycles = c;
  return { totalCycles, cycleIndex: currentCycleIndex };
}

function computeCoachPhasesByRules() {
  const out = new Array(52).fill(null).map(() => []);

  const nextTargetRaceByWeek = new Array(52).fill(null);
  const nextTargetRaceWeekIndex = new Array(52).fill(null);
  for (let i = 0; i < 52; i++) {
    for (let j = i; j < 52; j++) {
      const w = state.weeks[j];
      if (!w) continue;
      const pr = normalizeRacePriorityValue(w.priority);
      if (pr !== "A") continue;
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
      if (Number.isFinite(raceWeekIdx)) {
        const info = cycleIndexInfoForWeek(i, raceWeekIdx);
        if (info && Number.isFinite(info.totalCycles)) {
          const extra = baseFocusGeneric(info.totalCycles, info.cycleIndex || 1);
          if (extra) {
            out[i] = ["Aerobic Endurance", extra];
            continue;
          }
        }
      } else {
        const info = getGenericCycleInfo(i);
        if (info && Number.isFinite(info.totalCycles)) {
          const extra = baseFocusGeneric(info.totalCycles, info.cycleIndex || 1);
          if (extra) {
            out[i] = ["Aerobic Endurance", extra];
            continue;
          }
        }
      }
      const buildPhases = normalizePhases(phasesForRaceDistance(race?.distanceKm, race?.kind));
      const relevance = intensityRelevanceForRace(race?.distanceKm, race?.kind);
      const candidates = relevance.filter((p) => !buildPhases.includes(p));
      const infoAny = Number.isFinite(raceWeekIdx) ? cycleIndexInfoForWeek(i, raceWeekIdx) : null;
      let extra = "";
      if (infoAny && Number.isFinite(infoAny.totalCycles) && candidates.length > 0) {
        const total = Math.max(1, Number(infoAny.totalCycles));
        const idxIn = Math.max(1, Number(infoAny.cycleIndex || 1));
        const len = candidates.length;
        const pos = clamp(len - Math.ceil((idxIn * len) / total), 0, len - 1);
        extra = candidates[pos] || candidates[0] || "";
      } else {
        const weeksToRace = Number.isFinite(raceWeekIdx) ? Math.max(0, raceWeekIdx - i) : null;
        const pickIdx = Number.isFinite(weeksToRace) ? clamp(Math.floor(weeksToRace / 4), 0, candidates.length - 1) : 0;
        extra = candidates[pickIdx] || candidates[0] || "";
      }
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
      out[i] = ["Aerobic Endurance", ...phasesForRaceDistance(race?.distanceKm, race?.kind)];
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
    if (w.phasesAuto === false) continue;
    const cur = normalizePhases(w.phases);
    const next = normalizePhases(phases[i]);
    if (cur.length !== next.length || cur.some((x, idx) => x !== next[idx])) {
      w.phases = next;
      w.phasesAuto = true;
      changed = true;
    }
  }
  return changed;
}

function applyCoachAutoRules() {
  const a = applyCoachBlockRules();
  const b = applyCoachPhaseRules();
  return a || b;
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
  if (t === "Race") {
    const m = Math.max(0, Math.round(Number(plan?.minutes) || 0));
    return { min: m, max: m };
  }
  return { min: 0, max: 0 };
}

function normalizePlansNoShortAerobic(plans, options) {
  const minAerobic = Number.isFinite(WORKOUT_HELPER_CONFIG?.long?.durationMin) ? Number(WORKOUT_HELPER_CONFIG.long.durationMin) : 30;
  if (!Array.isArray(plans) || !Number.isFinite(minAerobic) || minAerobic <= 0) return plans;

  const mins = plans.map((p) => Math.max(0, Math.round(Number(p?.minutes) || 0)));
  const caps = plans.map((p) => minuteCapsForPlan(p, options));

  const typeAt = (i) => String(plans[i]?.type || "");
  const easyIdx = mins.map((_, i) => i).filter((i) => typeAt(i) === "Easy");
  const longIdx = mins.map((_, i) => i).filter((i) => typeAt(i) === "Long");
  const qualityIdx = mins.map((_, i) => i).filter((i) => typeAt(i) === "Quality");
  const raceIdx = mins.map((_, i) => i).filter((i) => typeAt(i) === "Race");

  const addTo = (i, amount) => {
    if (amount <= 0) return 0;
    const hi = caps[i].max;
    const room = Math.max(0, hi - mins[i]);
    const take = Math.min(room, amount);
    mins[i] += take;
    return take;
  };

  for (const i of easyIdx) {
    if (!(mins[i] > 0 && mins[i] < minAerobic)) continue;
    const need = minAerobic - mins[i];

    let donor = -1;
    let best = -1;
    for (const j of easyIdx) {
      if (j === i) continue;
      if (mins[j] >= minAerobic + need && mins[j] > best) {
        best = mins[j];
        donor = j;
      }
    }

    if (donor >= 0) {
      mins[donor] -= need;
      mins[i] += need;
      continue;
    }

    const removed = mins[i];
    mins[i] = 0;

    if (plans[i] && typeof plans[i] === "object") {
      plans[i].type = "Rest";
      plans[i].minutes = 0;
      plans[i].phase = "";
      plans[i].race = null;
      if ("rpeOverride" in plans[i]) plans[i].rpeOverride = 1;
    }

    let remaining = removed;
    const easyReceivers = easyIdx.filter((j) => j !== i && mins[j] >= minAerobic).sort((a, b) => mins[b] - mins[a]);
    for (const j of easyReceivers) {
      if (remaining <= 0) break;
      remaining -= addTo(j, remaining);
    }
    if (remaining > 0) {
      for (const group of [longIdx, qualityIdx, raceIdx]) {
        for (const j of group) {
          if (remaining <= 0) break;
          remaining -= addTo(j, remaining);
        }
      }
    }
  }

  for (let i = 0; i < plans.length; i++) {
    if (plans[i] && typeof plans[i] === "object") plans[i].minutes = mins[i];
  }
  return plans;
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
      // Race week should be a taper, so target ~70% of chronic load
      const desiredTotalLoad = chronicDaily * 7 * 0.7;
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

    normalizePlansNoShortAerobic(base, easyCapOptions);
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
  normalizePlansNoShortAerobic(base, options);
  return { plans: base, volumeHrsOverride };
}

function raceEntriesForWeek(week) {
  if (!week) return [];
  // Ensure monday is a Date object, handling potential string from JSON
  const monday = week.monday instanceof Date ? week.monday : (week.monday ? new Date(week.monday) : null);
  if (!monday || isNaN(monday.getTime())) return [];

  const races = Array.isArray(week.races) ? week.races : [];
  const out = [];
  for (const r of races) {
    const name = String(r?.name || "").trim();
    const date = String(r?.date || "").trim();
    // Relaxed name check - if name is empty, it's still a valid race if date exists
    if (!date) continue;
    
    // Assign default name if missing
    if (!name) {
       // Optional: could derive name from distance/kind if needed, but empty is fine if logic handles it
       // or just let it be empty string
    }

    const d = parseYMD(date);
    if (!d) continue;

    // Use a more robust diff calculation to handle potential timezone/DST offsets
    const diffMs = d.getTime() - monday.getTime();
    const dayIndex = Math.round(diffMs / MS_PER_DAY);
    
    // Allow slight tolerance if needed, but strict 0-6 should work if dates are correct
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
    const pr = normalizeRacePriorityValue(w.priority);
    if (pr !== "A") continue;
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
  const end = clamp(Number(weekIndex) || 0, 0, 51);
  let count = 0;
  for (let i = 0; i <= end; i++) {
    const w = state.weeks[i];
    if (!w) continue;
    const phases = normalizePhases(w.phases);
    if (phases.includes(p)) count++;
  }
  return Math.max(0, count - 1);
}

function estimateRaceMinutes(distanceKm, kind) {
  const d = Number(distanceKm);
  const k = String(kind || "").trim();
  if (!Number.isFinite(d) || d <= 0) return 75;

  // 若非越野跑且有 VDOT，嘗試用 VDOT 預估完賽時間
  if (k !== "trail" && state.vdot && state.vdot > 0) {
    const sec = solveRaceTimeSeconds(d * 1000, state.vdot);
    if (Number.isFinite(sec) && sec > 0) {
      // 允許範圍放寬到 10 ~ 600 分鐘 (10小時)，以涵蓋超馬或慢速全馬
      return clamp(Math.round(sec / 60), 10, 600);
    }
  }

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

function getDistRangeStr(minutes, vdot, loFrac, hiFrac) {
  if (!Number.isFinite(vdot) || vdot <= 0 || minutes <= 0) return "";
  const slowSec = paceSecondsPerKmFromVdotFraction(vdot, loFrac);
  const fastSec = paceSecondsPerKmFromVdotFraction(vdot, hiFrac);
  if (!Number.isFinite(slowSec) || !Number.isFinite(fastSec)) return "";
  
  // 計算平均配速與預估距離 (公尺)
  const avgPace = (slowSec + fastSec) / 2; 
  const avgSpeed = 60 / avgPace; // km/min
  const distM = minutes * avgSpeed * 1000;

  // 使用者指定的標準運動場距離
  const stds = [100, 200, 300, 400, 800, 1000, 1200, 1600, 2000, 3000, 4000];
  
  let closest = stds[0];
  let minDiff = Math.abs(distM - closest);
  
  for (const s of stds) {
    const diff = Math.abs(distM - s);
    if (diff < minDiff) {
      minDiff = diff;
      closest = s;
    }
  }

  const paceRange = `${formatPaceFromSecondsPerKm(slowSec)}–${formatPaceFromSecondsPerKm(fastSec)}`;
  
  // 若距離遠大於列表最大值 (例如 > 4500m 的長距離跑)，則顯示 km
  if (distM > 4500) {
    return `（約 ${(distM / 1000).toFixed(2)}km @ ${paceRange}/km）`;
  }
  
  return `（約 ${closest}m @ ${paceRange}/km）`;
}

function sessionPlanForDay(weekIndex, day, ctx) {
  const minutes = Math.max(0, Math.round(Number(day?.minutes) || 0));
  const type = String(day?.type || "").trim();
  const phase = String(day?.phase || "").trim();
  const race = day?.race || null;
  const block = String(ctx?.block || "").trim();
  const rpeOverride = Number.isFinite(Number(day?.rpeOverride)) ? clamp(Number(day.rpeOverride), 1, 10) : null;

  // Use state.vdot if available
  const vdot = Number.isFinite(state.vdot) ? state.vdot : null;

  if (type === "Rest" || minutes <= 0) {
    return { zone: 1, rpe: 1, workoutMinutes: 0, noteBody: buildAutoNoteBodyForPlan({ title: "休息／伸展", minutes: 0, rpeText: "—" }) };
  }

  if (type === "Race" && race) {
    const kindText = race.kind === "trail" ? "越野跑" : race.kind === "road" ? "路跑" : "";
    const distText = race.distanceKm ? `${race.distanceKm}km` : "";
    const meta = [distText, kindText].filter(Boolean).join(" · ");
    const title = meta ? `比賽：${race.name}（${meta}）` : `比賽：${race.name}`;
    const details = ["熱身 15分鐘 + 比賽 + 放鬆 10分鐘"];
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
    const idx = phaseStreakIndex(weekIndex, phase);
    const main = minutes > 0 ? minutes : clamp(30 + idx * 5, 30, 45);
    const distInfo = vdot ? getDistRangeStr(main, vdot, 0.78, 0.86) : "";
    const details = [`主課：節奏跑 ${main}分鐘${distInfo}（30–45分鐘，可每週 +5分鐘）`, "另加：熱身／放鬆"];
    return { zone: 3, rpe: rpeOverride ?? 6, workoutMinutes: main, noteBody: buildAutoNoteBodyForPlan({ title: "節奏", details, minutes: main, rpeText: "5–6" }) };
  }
  if (phase === "Threshold") {
    const idx = phaseStreakIndex(weekIndex, phase);
    const repMin = clamp(6 + Math.floor(idx / 2) * 2, 6, 12);
    const restMin = Math.max(1, Math.round(repMin / 4));
    const sets = minutes > 0 ? Math.max(1, Math.round(minutes / repMin)) : clamp(3 + Math.floor(idx / 2), 3, 5);
    const workoutMinutes = minutes > 0 ? minutes : Math.max(0, Math.round(sets * repMin));
    const distInfo = vdot ? getDistRangeStr(repMin, vdot, 0.86, 0.92) : "";
    const details = [`主課：${sets} × ${repMin}分鐘${distInfo}（總量 ${workoutMinutes}分鐘，跑/休 4:1，休 ${restMin}分鐘）`];
    details.push("另加：熱身／放鬆");
    return {
      zone: 4,
      rpe: rpeOverride ?? 8,
      workoutMinutes,
      noteBody: buildAutoNoteBodyForPlan({ title: "乳酸閾值", details, minutes: workoutMinutes, rpeText: "7–8" }),
    };
  }
  if (phase === "VO2Max") {
    const idx = phaseStreakIndex(weekIndex, phase);
    const workTotal = minutes > 0 ? minutes : clamp(5 + idx * 2, 5, 15);
    const repMin = workTotal <= 6 ? 1 : workTotal <= 9 ? 2 : workTotal <= 12 ? 3 : 4;
    const reps = Math.max(1, Math.round(workTotal / repMin));
    const distInfo = vdot ? getDistRangeStr(repMin, vdot, 0.92, 0.99) : "";
    const details = [`主課：${reps} × ${repMin}分鐘${distInfo}（跑/休 1:1）`, "總量：由 5分鐘 逐週加到最多 15分鐘", "另加：熱身／放鬆"];
    const workoutMinutes = workTotal;
    return {
      zone: 5,
      rpe: rpeOverride ?? 9,
      workoutMinutes,
      noteBody: buildAutoNoteBodyForPlan({ title: "最大攝氧量", details, minutes: workoutMinutes, rpeText: "8–9" }),
    };
  }
  if (phase === "Anaerobic") {
    const idx = phaseStreakIndex(weekIndex, phase);
    const workTotal = minutes > 0 ? minutes : clamp(5 + idx * 2, 5, 15);
    const repSec = workTotal <= 8 ? 30 : 60;
    const repMin = repSec === 30 ? 0.5 : 1;
    const reps = Math.max(1, Math.round(workTotal / repMin));
    const distInfo = vdot ? getDistRangeStr(repMin, vdot, 0.99, 1.06) : "";
    const details = [`主課：${reps} × ${repSec}秒${distInfo}（跑/休 1:1）`, "總量：由 5分鐘 逐週加到最多 15分鐘", "另加：熱身／放鬆"];
    const workoutMinutes = workTotal;
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

function qualityWorkoutMinutesForPhase(weekIndex, phase) {
  const p = String(phase || "").trim();
  const idx = phaseStreakIndex(weekIndex, p);

  if (p === "Tempo") {
    return clamp(30 + idx * 5, 30, 45);
  }
  if (p === "Threshold") {
    const repMin = clamp(6 + Math.floor(idx / 2) * 2, 6, 12);
    const sets = clamp(3 + Math.floor(idx / 2), 3, 5);
    return Math.max(0, Math.round(sets * repMin));
  }
  if (p === "VO2Max") {
    return clamp(5 + idx * 2, 5, 15);
  }
  if (p === "Anaerobic") {
    return clamp(5 + idx * 2, 5, 15);
  }
  return clamp(50, 30, 110);
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

  const minAerobic = Number.isFinite(WORKOUT_HELPER_CONFIG?.long?.durationMin) ? Number(WORKOUT_HELPER_CONFIG.long.durationMin) : 30;
  if (Number.isFinite(minAerobic) && minAerobic > 0 && easyIdx.length) {
    for (const i of easyIdx) {
      if (!(mins[i] > 0 && mins[i] < minAerobic)) continue;
      const need = minAerobic - mins[i];

      let donor = -1;
      let best = -1;
      for (const j of easyIdx) {
        if (j === i) continue;
        if (mins[j] >= minAerobic + need && mins[j] > best) {
          best = mins[j];
          donor = j;
        }
      }

      if (donor >= 0) {
        mins[donor] -= need;
        mins[i] += need;
        continue;
      }

      const removed = mins[i];
      mins[i] = 0;

      let remaining = removed;
      const easyReceivers = easyIdx.filter((j) => j !== i && mins[j] >= minAerobic).sort((a, b) => mins[b] - mins[a]);
      for (const j of easyReceivers) {
        if (remaining <= 0) break;
        remaining -= addTo(j, remaining);
      }
      if (remaining > 0) {
        for (const group of [longIdx, qualityIdx, raceIdx]) {
          for (const j of group) {
            if (remaining <= 0) break;
            remaining -= addTo(j, remaining);
          }
        }
      }
    }
  }

  return plans.map((p, i) => {
    const minutes = mins[i];
    if (String(p?.type || "") === "Easy" && !(minutes > 0)) {
      return { ...p, type: "Rest", minutes: 0, phase: "", race: null };
    }
    return { ...p, minutes };
  });
}

function computeLongRunMinutesForCycle(weekIndex, targetRaceDistKm) {
  // Find current cycle index
  // A cycle is typically 3 Base + 1 Deload (4 weeks)
  // If the start is short (e.g. 1-2 base), it counts as cycle 1
  
  let cycleIndex = 1;
  let currentCycleWeeks = 0;
  
  // Iterate from start to current week to determine cycle number
  for (let i = 0; i <= weekIndex; i++) {
    const w = state.weeks[i];
    if (!w) continue;
    
    // Check if this week starts a new cycle (simple heuristic: if previous was deload and current is base)
    // Or just count weeks if we assume standard structure.
    // However, user said "3 base + 1 deload" is a cycle.
    // But also "if start has only 1-2 base, it counts as a cycle".
    
    // Let's count "Deload" weeks encountered. Each deload completes a cycle?
    // Or better: Cycle 1 starts at week 0.
    // If we encounter a transition from Deload -> Base, increment cycle index.
    
    if (i > 0) {
      const prev = state.weeks[i-1];
      const prevBlock = normalizeBlockValue(prev?.block || "");
      const currBlock = normalizeBlockValue(w.block || "");
      
      if ((prevBlock === "Deload" || prevBlock === "Transition") && currBlock === "Base") {
        cycleIndex++;
      }
    }
  }
  
  // Logic based on race distance
  const dist = Number(targetRaceDistKm);
  
  if (dist <= 5) { // 5km logic
    let m = 75;
    if (cycleIndex === 1) m = 60;
    else if (cycleIndex === 2) m = 65;
    else if (cycleIndex === 3) m = 70;
    return Math.min(m, 75);
  } else if (dist <= 10) { // 10km logic
    let m = 90;
    if (cycleIndex === 1) m = 60;
    else if (cycleIndex === 2) m = 65;
    else if (cycleIndex === 3) m = 70;
    else if (cycleIndex === 4) m = 75;
    else if (cycleIndex === 5) m = 80;
    else if (cycleIndex === 6) m = 85;
    return Math.min(m, 90);
  } else { // Half Marathon logic
    let m = 120;
    if (cycleIndex === 1) m = 60;
    else if (cycleIndex === 2) m = 65;
    else if (cycleIndex === 3) m = 70;
    else if (cycleIndex === 4) m = 75;
    else if (cycleIndex === 5) m = 80;
    else if (cycleIndex === 6) m = 85;
    else if (cycleIndex === 7) m = 90;
    else if (cycleIndex === 8) m = 95;
    else if (cycleIndex === 9) m = 100;
    else if (cycleIndex === 10) m = 105;
    else if (cycleIndex === 11) m = 110;
    else if (cycleIndex === 12) m = 115;
    return Math.min(m, 120);
  }
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
  const pr = normalizeRacePriorityValue(w.priority);
  const isRaceWeek = races.length > 0 && pr === "A";
  const next = nextTargetRaceFromWeekIndex(idx);
  
  // Find the main race (longest distance) for context
  let mainRace = null;
  if (isRaceWeek) {
    mainRace = races.reduce((prev, curr) => {
      return (Number(curr.distanceKm) || 0) > (Number(prev.distanceKm) || 0) ? curr : prev;
    }, races[0]);
  }
  
  const raceContext = isRaceWeek ? mainRace : next?.race || null;

  const intensityOrder = intensityRelevanceForRace(raceContext?.distanceKm, raceContext?.kind);
  const intensityPhases = phases.filter((p) => p === "Tempo" || p === "Threshold" || p === "VO2Max" || p === "Anaerobic");
  intensityPhases.sort((a, b) => (intensityOrder.indexOf(a) < 0 ? 99 : intensityOrder.indexOf(a)) - (intensityOrder.indexOf(b) < 0 ? 99 : intensityOrder.indexOf(b)));

  const day = new Array(7).fill(null).map(() => ({ type: "Rest", minutes: 0, phase: "", race: null }));

  if (isRaceWeek) {
    // Place all races on their respective days
    for (const r of races) {
      const dIdx = clamp(Number(r.dayIndex) || 0, 0, 6);
      day[dIdx] = { 
        type: "Race", 
        minutes: estimateRaceMinutes(r.distanceKm, r.kind), 
        phase: "", 
        race: r 
      };
    }
  }

  const strengthDay = 3;
  if (day[strengthDay].type === "Rest") {
    day[strengthDay] = { type: "Rest", minutes: 0, phase: "", race: null };
  }

  const hasLong = (block === "Base" || block === "Build" || (block === "Peak" && !isRaceWeek)) && !isRaceWeek;
  const longDay = hasLong ? 6 : null;

  const setIfRest = (d, next) => {
    if (day[d]?.type === "Rest") day[d] = next;
  };

  const isRecoveryBlock = block === "Deload" || block === "Transition";
  if (isRecoveryBlock) {
    [0, 2, 4, 6].forEach((d) => {
      setIfRest(d, { type: "Easy", minutes: 0, phase: "", race: null });
    });
  } else if (block === "Peak") {
    const q1 = intensityPhases[0] || "VO2Max";
    const q2 = intensityPhases[1] || (q1 === "VO2Max" ? "Threshold" : "VO2Max");
    setIfRest(1, { type: "Quality", minutes: 0, phase: q1, race: null });
    if (!isRaceWeek) setIfRest(4, { type: "Quality", minutes: 0, phase: q2, race: null });
    [0, 2, 5].forEach((d) => {
      setIfRest(d, { type: "Easy", minutes: 0, phase: "", race: null });
    });
    if (Number.isFinite(longDay)) setIfRest(longDay, { type: "Long", minutes: 0, phase: "", race: null });
  } else if (block === "Build") {
    const q1 = intensityPhases[0] || "Tempo";
    const q2 = intensityPhases[1] || "Threshold";
    setIfRest(1, { type: "Quality", minutes: 0, phase: q1, race: null });
    setIfRest(4, { type: "Quality", minutes: 0, phase: q2, race: null });
    [0, 2, 5].forEach((d) => {
      setIfRest(d, { type: "Easy", minutes: 0, phase: "", race: null });
    });
    if (Number.isFinite(longDay)) setIfRest(longDay, { type: "Long", minutes: 0, phase: "", race: null });
  } else {
    const q = intensityPhases[0] || phases.find((p) => p === "Tempo" || p === "Threshold" || p === "VO2Max" || p === "Anaerobic") || "Tempo";
    setIfRest(1, { type: "Quality", minutes: 0, phase: q, race: null });
    [0, 2, 4, 5].forEach((d) => {
      setIfRest(d, { type: "Easy", minutes: 0, phase: "", race: null });
    });
    if (Number.isFinite(longDay)) setIfRest(longDay, { type: "Long", minutes: 0, phase: "", race: null });
  }

  const plans = day.map((d) => ({ ...d }));
  let options =
    block === "Peak"
      ? { minEasy: 0, maxEasy: 90, minLong: 30, maxLong: 120, minQuality: 5, maxQuality: 110 }
      : block === "Deload" || block === "Transition"
        ? { minEasy: 0, maxEasy: 90, minLong: 0, maxLong: 0, minQuality: 0, maxQuality: 0 }
        : { minEasy: 0, maxEasy: 120, minLong: 30, maxLong: 180, minQuality: 5, maxQuality: 110 };

  const basePlans = plans.map((p) => ({ ...p, minutes: Math.max(0, Math.round(Number(p.minutes) || 0)) }));
  const baseIdx = blockStreakIndex(idx, "Base");
  const buildIdx = blockStreakIndex(idx, "Build");

  if (block === "Base" && !isRaceWeek) {
    const race = nextTargetRaceFromWeekIndex(idx)?.race;
    const dist = race ? Number(race.distanceKm) : 21.1; // Default to HM logic if no race
    const l = computeLongRunMinutesForCycle(idx, dist);
    
    // Enforce constraints: Easy runs <= Long run, Long run <= 120
    options.maxEasy = l;
    options.maxLong = l;

    for (let i = 0; i < basePlans.length; i++) {
       if (String(basePlans[i]?.type || "") === "Quality") basePlans[i].minutes = qualityWorkoutMinutesForPhase(idx, basePlans[i]?.phase);
       if (String(basePlans[i]?.type || "") === "Long") basePlans[i].minutes = l;
     }

     const maxEasy = l;
     for (let i = 0; i < basePlans.length; i++) {
       if (String(basePlans[i]?.type || "") === "Easy") {
         basePlans[i].minutes = Math.min(basePlans[i].minutes, maxEasy);
       }
     }
  } else if (block === "Build" && !isRaceWeek) {
    const race = nextTargetRaceFromWeekIndex(idx)?.race;
    const dist = race ? Number(race.distanceKm) : 21.1;
    let l = 120; // Default max for HM
    if (dist <= 5) l = 75;
    else if (dist <= 10) l = 90;
    
    // Enforce constraints: Easy runs <= Long run, Long run <= 120
    options.maxEasy = l;
    options.maxLong = l;

    for (let i = 0; i < basePlans.length; i++) {
      if (String(basePlans[i]?.type || "") === "Quality") basePlans[i].minutes = qualityWorkoutMinutesForPhase(idx, basePlans[i]?.phase);
      if (String(basePlans[i]?.type || "") === "Long") basePlans[i].minutes = l;
    }
    
    const maxEasy = l;
    for (let i = 0; i < basePlans.length; i++) {
      if (String(basePlans[i]?.type || "") === "Easy") {
        basePlans[i].minutes = Math.min(basePlans[i].minutes, maxEasy);
      }
    }
  } else if (block === "Peak") {
    const l = clamp(70, 45, 120);

    // Enforce constraints
    options.maxEasy = l;
    options.maxLong = l;

    for (let i = 0; i < basePlans.length; i++) {
      if (String(basePlans[i]?.type || "") === "Quality") basePlans[i].minutes = qualityWorkoutMinutesForPhase(idx, basePlans[i]?.phase);
      if (!isRaceWeek && String(basePlans[i]?.type || "") === "Long") basePlans[i].minutes = l;
    }

    const maxEasy = l;
    for (let i = 0; i < basePlans.length; i++) {
      if (String(basePlans[i]?.type || "") === "Easy") {
        basePlans[i].minutes = Math.min(basePlans[i].minutes, maxEasy);
      }
    }
  }

  if (isRaceWeek) {
    options.maxEasy = 45;
    
    // 1. Force day before race to Rest (unless it's a race itself)
    for (const r of races) {
      const dIdx = clamp(Number(r.dayIndex) || 0, 0, 6);
      const prev = dIdx - 1;
      if (prev >= 0 && String(basePlans[prev]?.type || "") !== "Race") {
        basePlans[prev] = { type: "Rest", minutes: 0, phase: "", race: null };
      }
    }

    // 2. Clamp Easy runs to 45 mins
    for (let i = 0; i < basePlans.length; i++) {
      if (String(basePlans[i]?.type || "") === "Easy") {
        basePlans[i].minutes = Math.min(basePlans[i].minutes, 45);
      }
    }
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
      if (chunks < easyIndices.length) {
        const i = easyIndices[chunks];
        basePlans[i].minutes = Math.max(0, Math.round(rem));
      } else {
        const i = easyIndices[chunks % easyIndices.length];
        basePlans[i].minutes = (Number(basePlans[i].minutes) || 0) + rem;
      }
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

  const raceDay = isRaceWeek && raceContext ? clamp(Number(raceContext.dayIndex) || 0, 0, 6) : null;
  const constrained = applyLoadConstraintsToPlans(idx, rebalanced, { block, isRaceWeek, raceDay, minuteOptions: options });
  return { block, targetMinutes, plans: constrained.plans, volumeHrsOverride: constrained.volumeHrsOverride };
}

function applyAutoSessionsForWeek(weekIndex) {
  const idx = clamp(Number(weekIndex) || 0, 0, 51);
  const w = state.weeks[idx];
  if (!w) return { ok: false, reason: "Invalid week index" };

  const dayPlan = computeWeekDayPlans(idx);
  if (!dayPlan || !Array.isArray(dayPlan.plans) || dayPlan.plans.length !== 7) {
    return { ok: false, reason: "Missing weekly volume" };
  }

  const sessions = getWeekSessions(w);
  if (!Array.isArray(w.sessions) || !w.sessions.length) w.sessions = sessions;

  if (typeof dayPlan.volumeHrsOverride === "string") {
    w.volumeHrs = dayPlan.volumeHrsOverride;
    w.volumeMode = "direct";
    w.volumeFactor = 1;
  }

  for (let dayIndex = 0; dayIndex < 7; dayIndex++) {
    const p = dayPlan.plans[dayIndex] || { type: "Rest", minutes: 0, phase: "", race: null, rpeOverride: 1 };
    const s = sessions[dayIndex];
    if (!s) continue;

    if (dayIndex === 3 && String(p?.type || "") !== "Race") {
      s.kind = "Strength";
      s.zone = 1;
      s.rpe = 1;
      s.workoutsCount = 1;
      s.workouts = [{ duration: 0, rpe: 1 }];
      s.note = "肌力訓練";
      ensureSessionWorkouts(s);
      continue;
    }

    const plan = sessionPlanForDay(idx, p, { block: dayPlan.block });
    const plannedMinutes = Math.max(0, Math.round(Number(p?.minutes) || 0));

    s.kind = String(p?.type || "") === "Race" ? "Race" : "Run";
    s.zone = clamp(Number(plan.zone) || 1, 1, 6);
    s.rpe = clamp(Number(plan.rpe) || 1, 1, 10);
    s.workoutsCount = 1;
    s.workouts = [{ duration: plannedMinutes, rpe: s.rpe }];
    s.note = typeof plan.noteBody === "string" ? plan.noteBody : "";
    ensureSessionWorkouts(s);
  }

  return { ok: true, volumeHrsOverride: typeof dayPlan.volumeHrsOverride === "string" ? dayPlan.volumeHrsOverride : null };
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
  {
    const card = makeCard(
      "壓力水平",
      renderLineChartSvg(strainSeries, {
        unit: "A.U.",
        tickEvery: 4,
        yMax: strainMax,
        tooltip: (i, v) => `${weekLabelZh(i + 1)}：${Math.round(v)} A.U.`,
      }),
    );
    const tip = el(
      "div",
      "chartTip chartTip--warn",
      "如壓力水平只花一週便由低點飆升至高點，而且異常高於過往水平，即使其他指標正常，亦需注意過度訓練風險！",
    );
    card.appendChild(tip);
    root.appendChild(card);
  }
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
const CLOUD_META_KEY = "trainingPlanDemo_cloudMeta_v1";
const SUPABASE_URL = String(globalThis?.__SUPABASE__?.url || "").trim();
const SUPABASE_ANON_KEY = String(globalThis?.__SUPABASE__?.anonKey || "").trim();
const SUPABASE_STATE_TABLE = "training_state";
const supabaseClient =
  SUPABASE_URL && SUPABASE_ANON_KEY && typeof supabase?.createClient === "function"
    ? supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY, {
        auth: { persistSession: true, autoRefreshToken: true, detectSessionInUrl: true, flowType: "pkce" },
      })
    : null;

let authUser = null;
let cloudSaveTimer = null;
let cloudSavePendingJson = "";
let cloudSaveInFlight = false;

let calendarBaseVars = null;

function canPersistTrainingState() {
  return true;
}

function clearPersistedTrainingState() {
  try {
    localStorage.removeItem(STORAGE_KEY);
    localStorage.removeItem(CLOUD_META_KEY);
    state.ytdVolumeHrs = null;
    state.vdot = null;
    state.weeks = [];
    buildInitialWeeks();
    state.annualVolumeSettings = { startWeeklyHrs: null, maxUpPct: 12, maxDownPct: 25 };
  } catch {}
}

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
  if (!canPersistTrainingState()) return null;
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
  if (!canPersistTrainingState()) return;
  try {
    const payload = {
      startDate: formatYMD(state.startDate),
      isPaid: state.isPaid === true,
      ytdVolumeHrs: Number.isFinite(state.ytdVolumeHrs) ? state.ytdVolumeHrs : null,
      vdot: Number.isFinite(state.vdot) ? state.vdot : null,
      planStarted: state.planStarted === true,
      annualVolumeSettings: state.annualVolumeSettings,
      weeks: state.weeks.map((w) => ({
        races: Array.isArray(w.races) ? w.races : [],
        priority: normalizeRacePriorityValue(w.priority) || "",
        block: w.block || "",
        blockAuto: w.blockAuto === false ? false : true,
        season: w.season || "",
        phases: normalizePhases(w.phases),
        phasesAuto: w.phasesAuto === false ? false : true,
        volumeHrs: w.volumeHrs || "",
        volumeMode: w.volumeMode || "direct",
        volumeFactor: Number.isFinite(Number(w.volumeFactor)) ? Number(w.volumeFactor) : 1,
        volumeFactorAuto: w.volumeFactorAuto === true,
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
      manualAdvanceWeeks: state.manualAdvanceWeeks || 0,
    };
    localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
    scheduleCloudSave(payload);
  } catch {}
}

function applyPersistedTrainingState(persisted) {
  const persistedStartDate = parseYMD(persisted?.startDate);
  if (persistedStartDate) {
    state.startDate = startOfMonday(persistedStartDate);
  }
  state.isPaid = persisted?.isPaid === true;
  state.ytdVolumeHrs = Number.isFinite(persisted?.ytdVolumeHrs) ? persisted.ytdVolumeHrs : null;
  state.vdot = Number.isFinite(persisted?.vdot) ? persisted.vdot : null;
  state.planBaseVolume = Number.isFinite(persisted?.planBaseVolume) ? persisted.planBaseVolume : null;
  state.planStarted = persisted?.planStarted === true;
  state.manualAdvanceWeeks = Number.isFinite(persisted?.manualAdvanceWeeks) ? persisted.manualAdvanceWeeks : 0;
  state.annualVolumeSettings =
    persisted?.annualVolumeSettings && typeof persisted.annualVolumeSettings === "object"
      ? {
          startWeeklyHrs: Number.isFinite(Number(persisted.annualVolumeSettings.startWeeklyHrs)) ? Number(persisted.annualVolumeSettings.startWeeklyHrs) : null,
          maxUpPct: Number.isFinite(Number(persisted.annualVolumeSettings.maxUpPct)) ? Number(persisted.annualVolumeSettings.maxUpPct) : 12,
          maxDownPct: Number.isFinite(Number(persisted.annualVolumeSettings.maxDownPct)) ? Number(persisted.annualVolumeSettings.maxDownPct) : 25,
        }
      : { startWeeklyHrs: null, maxUpPct: 12, maxDownPct: 25 };
  buildInitialWeeks();
  if (persisted?.weeks?.length === 52) {
    const seasonOptions = ["", "Base", "Build", "Peak", "Deload", "Transition"];
    persisted.weeks.forEach((p, idx) => {
      const w = state.weeks[idx];
      if (!w) return;
      w.priority = normalizeRacePriorityValue(typeof p.priority === "string" ? p.priority : "");
      w.block = typeof p.block === "string" ? normalizeBlockValue(p.block) : w.block;
       w.blockAuto = p?.blockAuto === false ? false : true;
      w.season = typeof p.season === "string" ? p.season : "";
      w.phases = normalizePhases(p?.phases ?? p?.phase);
      w.phasesAuto = p?.phasesAuto === false ? false : true;
      w.volumeHrs = typeof p.volumeHrs === "string" ? p.volumeHrs : "";
      w.volumeMode = typeof p.volumeMode === "string" ? p.volumeMode : "direct";
      w.volumeFactor = Number.isFinite(Number(p.volumeFactor)) ? Number(p.volumeFactor) : 1;
      w.volumeFactorAuto = p.volumeFactorAuto === true;
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
    });
  }
  if (typeof persisted?.selectedWeekIndex === "number") {
    state.selectedWeekIndex = clamp(persisted.selectedWeekIndex, 0, 51);
  }
  if (typeof persisted?.connected === "boolean") {
    state.connected = persisted.connected;
  }
}

function readCloudMeta() {
  try {
    const raw = localStorage.getItem(CLOUD_META_KEY);
    if (!raw) return { updatedAt: null };
    const parsed = JSON.parse(raw);
    const updatedAt = Number.isFinite(Number(parsed?.updatedAt)) ? Number(parsed.updatedAt) : null;
    return { updatedAt };
  } catch {
    return { updatedAt: null };
  }
}

function writeCloudMeta(meta) {
  try {
    const updatedAt = Number.isFinite(Number(meta?.updatedAt)) ? Number(meta.updatedAt) : null;
    localStorage.setItem(CLOUD_META_KEY, JSON.stringify({ updatedAt }));
  } catch {}
}

function coerceEpochMs(value) {
  const n = Number(value);
  if (Number.isFinite(n) && n > 0) return n;
  if (typeof value === "string") {
    const t = Date.parse(value);
    if (Number.isFinite(t) && t > 0) return t;
  }
  return null;
}

async function apiJson(path, init) {
  const res = await fetch(path, { ...(init || {}), credentials: "include" });
  let data = null;
  try {
    data = await res.json();
  } catch {
    data = null;
  }
  if (!res.ok) {
    const msg = typeof data?.error === "string" ? data.error : `Request failed (${res.status})`;
    const err = new Error(msg);
    err.status = res.status;
    err.data = data;
    throw err;
  }
  return data;
}

async function fetchCloudState() {
  if (!supabaseClient) throw new Error("Supabase not configured");
  if (!authUser?.id) throw new Error("Not logged in");
  const { data, error } = await supabaseClient
    .from(SUPABASE_STATE_TABLE)
    .select("state_json, updated_at")
    .eq("user_id", authUser.id)
    .maybeSingle();
  if (error) throw new Error(String(error.message || "Cloud fetch failed"));

  const updatedAt = coerceEpochMs(data?.updated_at);
  const state = data?.state_json && typeof data.state_json === "object" ? data.state_json : null;
  return { ok: true, state, updatedAt };
}

async function putCloudState(statePayload) {
  if (!supabaseClient) throw new Error("Supabase not configured");
  if (!authUser?.id) throw new Error("Not logged in");
  const updatedAt = Date.now();
  const { error } = await supabaseClient
    .from(SUPABASE_STATE_TABLE)
    .upsert({ user_id: authUser.id, state_json: statePayload, updated_at: updatedAt }, { onConflict: "user_id" });
  if (error) throw new Error(String(error.message || "Cloud save failed"));
  return { ok: true, updatedAt };
}

function scheduleCloudSave(payload) {
  if (!authUser) return;
  try {
    cloudSavePendingJson = JSON.stringify(payload);
  } catch {
    cloudSavePendingJson = "";
    return;
  }
  if (cloudSaveTimer) return;
  cloudSaveTimer = window.setTimeout(async () => {
    cloudSaveTimer = null;
    if (cloudSaveInFlight) return;
    if (!authUser) return;
    const nextJson = cloudSavePendingJson;
    cloudSavePendingJson = "";
    if (!nextJson) return;
    cloudSaveInFlight = true;
    try {
      const next = JSON.parse(nextJson);
      const out = await putCloudState(next);
      const updatedAt = Number.isFinite(Number(out?.updatedAt)) ? Number(out.updatedAt) : null;
      if (updatedAt) writeCloudMeta({ updatedAt });
    } catch (e) {
      cloudSavePendingJson = nextJson;
    } finally {
      cloudSaveInFlight = false;
      if (cloudSavePendingJson) scheduleCloudSave(JSON.parse(cloudSavePendingJson));
    }
  }, 1500);
}

function buildAuthModal(initialMode) {
  const overlay = el("div", "overlay overlay--auth");
  const modal = el("div", "modal modal--auth");
  const title = el("div", "modal__title", initialMode === "register" ? "註冊" : "登入");
  const subtitle = el("div", "modal__subtitle", "登入後會把你嘅表格改動保存到雲端，跨裝置都可以繼續用。");

  const closeIcon = svgEl("svg", { viewBox: "0 0 24 24", fill: "none", stroke: "currentColor", "stroke-width": "2" });
  closeIcon.appendChild(svgEl("path", { d: "M6 6l12 12" }));
  closeIcon.appendChild(svgEl("path", { d: "M18 6L6 18" }));
  const closeIconBtn = document.createElement("button");
  closeIconBtn.type = "button";
  closeIconBtn.className = "iconBtn authModal__close";
  closeIconBtn.appendChild(closeIcon);

  const header = el("div", "authModal__header");
  const headerText = el("div", "authModal__headerText");
  headerText.appendChild(title);
  headerText.appendChild(subtitle);
  header.appendChild(headerText);
  header.appendChild(closeIconBtn);

  const form = el("form", "authForm");

  const providers = el("div", "authProviders");
  const googleBtn = document.createElement("button");
  googleBtn.type = "button";
  googleBtn.className = "btn authProviderBtn authProviderBtn--google";
  const googleIcon = el("span", "authProviderBtn__icon authProviderBtn__icon--google");
  const googleIconSvg = svgEl("svg", { viewBox: "0 0 24 24", role: "img", "aria-hidden": "true" });
  const googleIconText = svgEl("text", {
    x: "12",
    y: "16",
    "text-anchor": "middle",
    "font-size": "12",
    "font-weight": "700",
    "font-family": 'ui-sans-serif, system-ui, -apple-system, "Segoe UI", Roboto, Arial, sans-serif',
    fill: "currentColor",
  });
  googleIconText.textContent = "G";
  googleIconSvg.appendChild(googleIconText);
  googleIcon.appendChild(googleIconSvg);
  const googleLabel = el("span", "authProviderBtn__label", "用 Google 登入");
  googleBtn.appendChild(googleIcon);
  googleBtn.appendChild(googleLabel);

  const emailChoiceBtn = document.createElement("button");
  emailChoiceBtn.type = "button";
  emailChoiceBtn.className = "btn authProviderBtn authProviderBtn--email";
  const emailIcon = el("span", "authProviderBtn__icon authProviderBtn__icon--email");
  const emailIconSvg = svgEl("svg", { viewBox: "0 0 24 24", fill: "none", stroke: "currentColor", "stroke-width": "2" });
  emailIconSvg.appendChild(svgEl("path", { d: "M4 7h16v10H4z" }));
  emailIconSvg.appendChild(svgEl("path", { d: "M4 7l8 6 8-6" }));
  emailIcon.appendChild(emailIconSvg);
  const emailChoiceLabel = el("span", "authProviderBtn__label", "用電郵登入");
  emailChoiceBtn.appendChild(emailIcon);
  emailChoiceBtn.appendChild(emailChoiceLabel);

  providers.appendChild(googleBtn);
  providers.appendChild(emailChoiceBtn);

  const divider = el("div", "authDivider", "或");

  const emailSection = el("div", "authEmail");
  emailSection.hidden = true;
  const backBtn = el("button", "btn authBackBtn", "返回");
  backBtn.type = "button";

  const emailRow = el("label", "authField");
  emailRow.appendChild(el("span", "authField__label", "Email"));
  const emailInput = document.createElement("input");
  emailInput.className = "input authInput";
  emailInput.type = "email";
  emailInput.autocomplete = "email";
  emailInput.placeholder = "name@example.com";
  emailRow.appendChild(emailInput);

  const passRow = el("label", "authField");
  passRow.appendChild(el("span", "authField__label", "密碼"));
  const passInput = document.createElement("input");
  passInput.className = "input authInput";
  passInput.type = "password";
  passInput.autocomplete = initialMode === "register" ? "new-password" : "current-password";
  passInput.placeholder = "至少 8 個字元";
  passRow.appendChild(passInput);

  const switchBtn = el("button", "btn authLinkBtn", initialMode === "register" ? "已有帳號？登入" : "未有帳號？註冊");
  switchBtn.type = "button";

  const submitBtn = el("button", "btn btn--primary authSubmitBtn", initialMode === "register" ? "建立帳號" : "登入");
  submitBtn.type = "submit";

  const actions = el("div", "authActions");
  actions.appendChild(submitBtn);

  const notice = el("div", "authNotice");
  notice.hidden = true;

  const footer = el("div", "authFooter");
  footer.appendChild(switchBtn);

  emailSection.appendChild(backBtn);
  emailSection.appendChild(emailRow);
  emailSection.appendChild(passRow);
  emailSection.appendChild(actions);
  emailSection.appendChild(notice);

  form.appendChild(providers);
  form.appendChild(divider);
  form.appendChild(emailSection);
  form.appendChild(footer);

  modal.appendChild(header);
  modal.appendChild(form);
  overlay.appendChild(modal);

  let awaitingEmailConfirm = false;

  const syncModeUi = () => {
    const mode = form.dataset.mode === "register" ? "register" : "login";
    title.textContent = mode === "register" ? "註冊" : "登入";
    submitBtn.textContent = awaitingEmailConfirm ? "已發送確認信" : mode === "register" ? "建立帳號" : "登入";
    emailChoiceLabel.textContent = mode === "register" ? "用電郵註冊" : "用電郵登入";
    passInput.autocomplete = mode === "register" ? "new-password" : "current-password";
    switchBtn.textContent = awaitingEmailConfirm ? "我已確認電郵 → 登入" : mode === "register" ? "已有帳號？登入" : "未有帳號？註冊";
  };

  const setMode = (mode) => {
    form.dataset.mode = mode === "register" ? "register" : "login";
    syncModeUi();
  };

  const setAwaitingConfirm = (on, targetEmail) => {
    awaitingEmailConfirm = Boolean(on);
    notice.hidden = !awaitingEmailConfirm;
    if (awaitingEmailConfirm) {
      const e = String(targetEmail || "").trim();
      notice.textContent = e ? `已發送確認電郵到 ${e}。請到收件匣／垃圾郵件夾確認後，再回到此頁登入。` : "已發送確認電郵。請到收件匣／垃圾郵件夾確認後，再回到此頁登入。";
      passInput.value = "";
    } else {
      notice.textContent = "";
    }
    emailInput.disabled = awaitingEmailConfirm;
    passInput.disabled = awaitingEmailConfirm;
    submitBtn.disabled = awaitingEmailConfirm;
    syncModeUi();
  };

  setMode(initialMode === "register" ? "register" : "login");

  closeIconBtn.addEventListener("click", () => overlay.remove());
  overlay.addEventListener("click", (e) => {
    if (e.target === overlay) overlay.remove();
  });
  switchBtn.addEventListener("click", () => {
    if (awaitingEmailConfirm) {
      setAwaitingConfirm(false);
      setMode("login");
      window.setTimeout(() => passInput.focus(), 0);
      return;
    }
    setMode(form.dataset.mode === "register" ? "login" : "register");
  });

  const redirectTo = () => {
    try {
      const u = new URL(String(document.baseURI || window.location.href || ""));
      u.hash = "";
      u.search = "";
      return u.toString();
    } catch {
      return String(window.location.href || "").split("#")[0];
    }
  };

  googleBtn.addEventListener("click", async () => {
    if (!supabaseClient) {
      showToast("未設定 Supabase（請填入 SUPABASE_URL / SUPABASE_ANON_KEY）", { variant: "warn", durationMs: 2200 });
      return;
    }
    googleBtn.disabled = true;
    try {
      const { error } = await supabaseClient.auth.signInWithOAuth({ provider: "google", options: { redirectTo: redirectTo() } });
      if (error) throw new Error(error.message);
    } catch (err) {
      googleBtn.disabled = false;
      showToast(String(err?.message || "登入失敗"), { variant: "warn", durationMs: 2200 });
    }
  });

  const openEmail = () => {
    setAwaitingConfirm(false);
    providers.hidden = true;
    divider.hidden = true;
    emailSection.hidden = false;
    window.setTimeout(() => emailInput.focus(), 0);
  };

  const closeEmail = () => {
    setAwaitingConfirm(false);
    providers.hidden = false;
    divider.hidden = false;
    emailSection.hidden = true;
    window.setTimeout(() => emailChoiceBtn.focus(), 0);
  };

  emailChoiceBtn.addEventListener("click", () => openEmail());
  backBtn.addEventListener("click", () => closeEmail());

  form.addEventListener("submit", async (e) => {
    e.preventDefault();
    const mode = form.dataset.mode === "register" ? "register" : "login";
    const email = String(emailInput.value || "").trim();
    const password = String(passInput.value || "");
    if (!email || !password) {
      showToast("請輸入 Email 同密碼", { variant: "warn", durationMs: 1600 });
      return;
    }
    setAwaitingConfirm(false);
    submitBtn.disabled = true;
    switchBtn.disabled = true;
    backBtn.disabled = true;
    googleBtn.disabled = true;
    emailChoiceBtn.disabled = true;
    closeIconBtn.disabled = true;
    try {
      if (!supabaseClient) throw new Error("未設定 Supabase");

      if (mode === "register") {
        const { error } = await supabaseClient.auth.signUp({ email, password, options: { emailRedirectTo: redirectTo() } });
        if (error) throw new Error(error.message);
        setAwaitingConfirm(true, email);
        showToast("已發送確認電郵，請到郵箱確認註冊後再登入", { durationMs: 2600 });
        return;
      }

      const { error } = await supabaseClient.auth.signInWithPassword({ email, password });
      if (error) throw new Error(error.message);
      const { data } = await supabaseClient.auth.getSession();
      const user = data?.session?.user || null;
      authUser = user ? { id: user.id, email: user.email || email } : null;
      overlay.remove();
      syncAuthUi();
      if (authUser) await afterLoginSync();
      showToast("已登入");
    } catch (err) {
      showToast(String(err?.message || "登入失敗"), { variant: "warn", durationMs: 2200 });
    } finally {
      const lock = awaitingEmailConfirm;
      submitBtn.disabled = lock;
      emailInput.disabled = lock;
      passInput.disabled = lock;
      switchBtn.disabled = false;
      backBtn.disabled = false;
      googleBtn.disabled = false;
      emailChoiceBtn.disabled = false;
      closeIconBtn.disabled = false;
    }
  });

  window.setTimeout(() => emailChoiceBtn.focus(), 0);
  return overlay;
}

function syncAuthUi() {
  const btn = document.getElementById("authBtn");
  if (!btn) return;
  btn.textContent = authUser ? "登出" : "登入";
  btn.classList.toggle("btn--danger", Boolean(authUser));
  btn.classList.toggle("btn--primary", !authUser);
}

async function afterLoginSync() {
  let remote = null;
  try {
    remote = await fetchCloudState();
  } catch {
    remote = null;
  }

  const remoteState = remote?.state;
  const remoteUpdatedAt = Number.isFinite(Number(remote?.updatedAt)) ? Number(remote.updatedAt) : null;
  const meta = readCloudMeta();
  const local = loadPersistedState();

  if (remoteState && remoteUpdatedAt && (!meta.updatedAt || remoteUpdatedAt > meta.updatedAt)) {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(remoteState));
      writeCloudMeta({ updatedAt: remoteUpdatedAt });
      window.location.reload();
      return;
    } catch {}
  }

  if (!remoteState && local) {
    try {
      const out = await putCloudState(local);
      const updatedAt = Number.isFinite(Number(out?.updatedAt)) ? Number(out.updatedAt) : null;
      if (updatedAt) writeCloudMeta({ updatedAt });
    } catch {}
  }
}

async function wireAuth() {
  const btn = document.getElementById("authBtn");
  if (!btn) return;

  let lastSyncedUserId = "";
  const runAfterLoginSyncForCurrentUser = async () => {
    const uid = String(authUser?.id || "");
    if (!uid) return;
    if (uid === lastSyncedUserId) return;
    lastSyncedUserId = uid;
    await afterLoginSync();
  };

  btn.addEventListener("click", async () => {
    if (authUser) {
      if (supabaseClient) {
        try {
          const { error } = await supabaseClient.auth.signOut();
          if (error) throw new Error(error.message);
        } catch {}
      }
      authUser = null;
      if (cloudSaveTimer) window.clearTimeout(cloudSaveTimer);
      cloudSaveTimer = null;
      cloudSavePendingJson = "";
      cloudSaveInFlight = false;
      
      lastSyncedUserId = "";
      syncAuthUi();
      showToast("已登出");
      return;
    }
    if (!supabaseClient) {
      showToast("未設定 Supabase（請填入 SUPABASE_URL / SUPABASE_ANON_KEY）", { variant: "warn", durationMs: 2400 });
      return;
    }
    const overlay = buildAuthModal("login");
    document.body.appendChild(overlay);
  });

  if (supabaseClient) {
    try {
      const { data } = await supabaseClient.auth.getSession();
      const user = data?.session?.user || null;
      authUser = user ? { id: user.id, email: user.email || "" } : null;
    } catch {
      authUser = null;
    }

    supabaseClient.auth.onAuthStateChange(async (_event, session) => {
      const user = session?.user || null;
      const next = user ? { id: user.id, email: user.email || "" } : null;
      const wasIn = Boolean(authUser);
      authUser = next;
      syncAuthUi();
      updateHeader();
      if (!authUser) {
        if (cloudSaveTimer) window.clearTimeout(cloudSaveTimer);
        cloudSaveTimer = null;
        cloudSavePendingJson = "";
        cloudSaveInFlight = false;
        
        lastSyncedUserId = "";
        return;
      }
      if (!wasIn) await runAfterLoginSyncForCurrentUser();
    });
  } else {
    authUser = null;
  }

  syncAuthUi();
  if (authUser) await runAfterLoginSyncForCurrentUser();
}

const state = {
  connected: false,
  isPaid: false,
  startDate: startOfMonday(new Date("2025-03-03T00:00:00")),
  ytdVolumeHrs: null,
  vdot: null,
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
    planBaseVolume: Number.isFinite(state.planBaseVolume) ? state.planBaseVolume : null,
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
      priority: normalizeRacePriorityValue(w.priority) || "",
      block: w.block || "",
      blockAuto: w.blockAuto === false ? false : true,
      season: w.season || "",
      phases: normalizePhases(w.phases),
      phasesAuto: w.phasesAuto === false ? false : true,
      volumeHrs: w.volumeHrs || "",
      volumeMode: w.volumeMode || "direct",
      volumeFactor: Number.isFinite(Number(w.volumeFactor)) ? Number(w.volumeFactor) : 1,
      volumeFactorAuto: w.volumeFactorAuto === true,
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
  state.planBaseVolume = Number.isFinite(snapshot.planBaseVolume) ? snapshot.planBaseVolume : null;
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
    w.priority = normalizeRacePriorityValue(typeof p.priority === "string" ? p.priority : "");
    w.block = typeof p.block === "string" ? p.block : "";
    w.blockAuto = p?.blockAuto === false ? false : true;
    w.season = typeof p.season === "string" ? p.season : "";
    w.phases = normalizePhases(p?.phases ?? p?.phase);
    w.phasesAuto = p?.phasesAuto === false ? false : true;
    w.volumeHrs = typeof p.volumeHrs === "string" ? p.volumeHrs : "";
    w.volumeMode = typeof p.volumeMode === "string" ? p.volumeMode : "direct";
    w.volumeFactor = Number.isFinite(Number(p.volumeFactor)) ? Number(p.volumeFactor) : 1;
    w.volumeFactorAuto = p.volumeFactorAuto === true;
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
  syncAnnualVolumeInputs();
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
      blockAuto: true,
      season: "",
      phases: [],
      phasesAuto: true,
      volumeHrs: "",
      volumeMode: "direct",
      volumeFactor: 1.2,
      volumeFactorAuto: true,
      sessions: buildDefaultSessions(),
    });
  }

  state.weeks = weeks;
}

function generatePlan() {
  const rng = seededRandom(20250303);

  for (const w of state.weeks) {
    const base = 4.5 + rng() * 2.5;
    const volume = normalizeRacePriorityValue(w.priority) === "A" ? base * 0.75 : base;
    w.volumeHrs = `${volume.toFixed(1)}`;
    w.volumeMode = "direct";
    w.volumeFactor = 1;
    w.volumeFactorAuto = false;
  }
}

function syncAnnualVolumeInputs() {
  const annualVolumeInput = document.getElementById("annualVolumeInput");

  if (annualVolumeInput) {
    annualVolumeInput.value = Number.isFinite(state.ytdVolumeHrs) && state.ytdVolumeHrs > 0 ? String(state.ytdVolumeHrs) : "";
  }
}

function updateHeader() {
  const dateRangeEl = document.getElementById("dateRange");
  const volumeTotalEl = document.getElementById("volumeTotal");
  const plannedVolume52El = document.getElementById("plannedVolume52");

  if (!dateRangeEl && !volumeTotalEl && !plannedVolume52El) return;

  const endDate = addDays(state.startDate, 52 * 7 - 1);
  if (dateRangeEl) {
    dateRangeEl.textContent = `${formatMD(state.startDate)} / ${state.startDate.getFullYear()} - ${formatMD(endDate)} / ${endDate.getFullYear()}`;
  }

  const totalHrs = state.weeks.reduce((sum, w) => sum + (Number(w.volumeHrs) || 0), 0);
  if (volumeTotalEl) {
    volumeTotalEl.textContent = `訓練總量：${totalHrs ? totalHrs.toFixed(1) : "—"} 小時`;
  }
  if (plannedVolume52El) {
    let text = "";
    if (Number.isFinite(state.ytdVolumeHrs) && state.ytdVolumeHrs > 0) {
      const target = state.ytdVolumeHrs;
      const diff = Math.round((totalHrs - target) * 10) / 10;
      text = `年總訓練量：${target.toFixed(1)} 小時 · 已分配：${totalHrs ? totalHrs.toFixed(1) : "—"} 小時 · 差：${diff.toFixed(1)} 小時`;
    } else {
      text = `計劃訓練量（52週）：${totalHrs ? totalHrs.toFixed(1) : "—"} 小時`;
    }
    
    plannedVolume52El.textContent = "";
    plannedVolume52El.appendChild(document.createTextNode(text));
  }
}

function renderCalendar() {
  const calendar = document.getElementById("calendar");
  if (!calendar) return;
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
    { label: "ACWR(訓練量)", key: "volumeFactor", type: "acwrInput" },
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
      if (i === state.selectedWeekIndex) {
        cell.classList.add("is-current");
      }

      if (row.type === "weekBtn") {
        setCellText(cell, String(w.weekNo));
      } else if (row.type === "date") {
        setCellText(cell, formatMD(w.monday));
      } else if (row.type === "blockSelect") {
        const val = normalizeBlockValue(w.block || "");
        setCellText(cell, BLOCK_LABELS_ZH[val] || val || "—");
      } else if (row.type === "prioritySelect") {
        setCellText(cell, racePriorityLabelZh(w.priority) || "—");
      } else if (row.type === "phase") {
        cell.classList.add("phaseCell");
        const phases = normalizePhases(w.phases);
        const selected = phases.includes(row.phase);
        if (selected) {
          cell.classList.add("is-on");
          cell.style.setProperty("--phaseBg", phaseColors[row.phase] || "var(--accent)");
        }
        // Read-only: no button, no interaction
      } else if (row.type === "races") {
        const cols = el("div", "raceCols");
        const races = Array.isArray(w.races) ? w.races : [];
        races.slice(0, 2).forEach((r) => {
          const name = (r?.name || "").trim();
          const dist = Number(r?.distanceKm);
          const distText = Number.isFinite(dist) && dist > 0 ? `${dist}km` : "";
          const text = name ? (distText ? `${name}（${distText}）` : name) : distText;
          if (!text) return;
          const item = el("div", "raceV", text);
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
          if (m.acwr !== null && m.acwr > 1.5) cell.classList.add("calCell--acwrAlert");
        } else {
          setCellText(cell, "—");
        }
      } else if (row.type === "acwrInput") {
        const f = Number(w.volumeFactor);
        const val = Number.isFinite(f) ? f.toFixed(1) : "1.0";
        setCellText(cell, val);
        if (Number.isFinite(f) && f > 1.5) cell.classList.add("calCell--acwrAlert");
      } else if (row.type === "text" && row.key === "volumeHrs") {
        setCellText(cell, w.volumeHrs ? String(w.volumeHrs) : "");
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

  const maxAllowed = getUnlockedWeekIndex();

  for (const w of state.weeks) {
    // If plan started, only show up to maxAllowed
    if (state.planStarted && w.index > maxAllowed) break;

    const opt = document.createElement("option");
    opt.value = String(w.index);
    opt.textContent = weekLabelZh(w.weekNo);
    if (w.index === state.selectedWeekIndex) opt.selected = true;
    select.appendChild(opt);
  }
  
  // If current selection is out of bounds, reset it
  if (state.planStarted && state.selectedWeekIndex > maxAllowed) {
    selectWeek(maxAllowed);
    return; // selectWeek will trigger re-render
  }

  select.onchange = () => {
    const idx = Number(select.value);
    selectWeek(clamp(idx, 0, 51));
  };
}

function getUnlockedWeekIndex() {
  if (!state.planStarted || !state.startDate) return 51;
  const now = new Date();
  const diff = now - state.startDate;
  // Weeks passed since start (0-based)
  // Week 1 starts at t=0. Week 2 starts at t=7days.
  // So if t < 7 days, index is 0.
  let idx = Math.floor(diff / (7 * 24 * 3600 * 1000));
  if (idx < 0) idx = 0;
  
  // Add manual advance (for testing)
  idx += (state.manualAdvanceWeeks || 0);
  
  return clamp(idx, 0, 51);
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

  // Locked View for Unpaid Users (Week 2 onwards)
  if (!state.isPaid && state.selectedWeekIndex > 0) {
    const lockedWrap = el("div", "dayCard");
    lockedWrap.style.textAlign = "center";
    lockedWrap.style.padding = "40px 20px";
    lockedWrap.style.display = "flex";
    lockedWrap.style.flexDirection = "column";
    lockedWrap.style.alignItems = "center";
    lockedWrap.style.gap = "16px";

    const lockIcon = el("div", "", "🔒");
    lockIcon.style.fontSize = "48px";
    
    const lockTitle = el("div", "", "此內容需要升級");
    lockTitle.style.fontSize = "18px";
    lockTitle.style.fontWeight = "600";
    
    const lockDesc = el("div", "muted", "您目前只能查看第 1 週的訓練內容。請升級以解鎖完整 52 週訓練計畫。");
    
    const upgradeBtn = el("button", "btn btn--primary", "立即升級");
    upgradeBtn.onclick = () => openPaymentModal();

    lockedWrap.appendChild(lockIcon);
    lockedWrap.appendChild(lockTitle);
    lockedWrap.appendChild(lockDesc);
    lockedWrap.appendChild(upgradeBtn);
    
    weekDays.appendChild(lockedWrap);
    return; // Stop rendering details
  }

  if (weekDays) {
    // Read-only: No drag and drop listeners
  }

  sessions.forEach((s, i) => {
    const card = el("div", "dayCard");
    card.dataset.dayIndex = String(i);
    ensureSessionWorkouts(s);

    const titleRow = el("div", "dayTitleRow");
    const titleLeft = el("div", "dayTitleLeft");
    // Read-only: No drag handle
    titleLeft.appendChild(el("div", "dayTitle", s.dayLabel));
    titleRow.appendChild(titleLeft);

    // Read-only: No drag listeners on card

    const dayDate = addDays(w.monday, i);
    const titleRight = el("div", "dayTitleRight");
    titleRight.appendChild(el("div", "dayDate", `${formatWeekdayEnShort(dayDate)} ${formatYMD(dayDate)}`));

    const workoutControls = el("div", "dayWorkoutControls");
    workoutControls.appendChild(el("span", "muted", "當日訓練數量："));
    workoutControls.appendChild(el("span", "", String(s.workoutsCount || 1)));
    // Read-only: No helper/reset buttons or select

    titleRight.appendChild(workoutControls);
    titleRow.appendChild(titleRight);
    card.appendChild(titleRow);

    const workoutsWrap = el("div", "dayWorkouts");
    card.appendChild(workoutsWrap);
    s.workouts.forEach((workout, workoutIndex) => {
      const metaRow = el("div", "dayMeta");

      const durationWrap = el("div", "dayField");
      durationWrap.appendChild(el("span", "", "時長："));
      const durationVal = Number(workout?.duration) > 0 ? String(Number(workout?.duration)) : "0";
      durationWrap.appendChild(el("span", "", durationVal));
      durationWrap.appendChild(el("span", "muted", " 分鐘"));
      metaRow.appendChild(durationWrap);

      const rpeWrap = el("div", "dayField");
      rpeWrap.appendChild(el("span", "", "RPE："));
      const rpeVal = String(clamp(Number(workout?.rpe) || 1, 1, 10));
      rpeWrap.appendChild(el("span", "", rpeVal));
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
    const noteText = el("div", "dayNote__text", typeof s.note === "string" ? s.note : "");
    noteWrap.appendChild(noteText);
    card.appendChild(noteWrap);

  weekDays.appendChild(card);
  });

  // Add Complete Button (Strict Mode) if showing the latest unlocked week or the one before it
  const unlocked = getUnlockedWeekIndex();
  // Show button if this is the current active week (unlocked) OR if it's the previous week but we haven't moved yet (unlocked - 1)
  // Actually, user wants "Complete This Week" on Week 1. If date passed, move to Week 2.
  // So we show it on the currently selected week as long as it's <= unlocked.
  // And strictly, we only need it on the latest week to "advance".
  if (state.planStarted && state.selectedWeekIndex <= unlocked && state.selectedWeekIndex < 51) {
    const btnRow = el("div", "dayCard");
    btnRow.style.textAlign = "center";
    btnRow.style.padding = "16px";
    btnRow.style.cursor = "pointer";
    btnRow.style.backgroundColor = "#e0f2fe"; // Light blue
    
    const btn = el("button", "btn btn--primary", "完成此週");
    btn.onclick = () => {
      // Re-check time
      const currentUnlocked = getUnlockedWeekIndex();
      if (currentUnlocked > state.selectedWeekIndex) {
        // Time has passed, allowed to advance
        
        // Payment Wall: If completing Week 1 (index 0) to go to Week 2 (index 1)
        if (state.selectedWeekIndex === 0 && !state.isPaid) {
          openPaymentModal();
          return;
        }

        pushHistory();
        selectWeek(state.selectedWeekIndex + 1);
        showToast("已完成此週，進入下一週");
      } else {
        // Time has not passed
        alert("這週還沒結束");
      }
    };
    
    btnRow.appendChild(btn);
    weekDays.appendChild(btnRow);
  }
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

function openPaymentModal() {
  const overlay = el("div", "overlay");
  
  const modal = el("div", "modal");
  modal.style.maxWidth = "400px";
  modal.style.textAlign = "center";

  const title = el("div", "modal__title", "解鎖完整計畫");
  title.style.marginBottom = "16px";
  
  const desc = el("div", "", "您已完成第 1 週的體驗。如需繼續瀏覽第 2 週及之後的訓練內容，請先付款解鎖完整計畫。");
  desc.style.marginBottom = "24px";
  desc.style.lineHeight = "1.5";
  desc.style.color = "var(--text)";

  const actions = el("div", "modal__actions");
  actions.style.justifyContent = "center";
  actions.style.gap = "16px";

  const cancelBtn = el("button", "btn", "稍後再說");
  cancelBtn.onclick = () => overlay.remove();

  const payBtn = el("button", "btn btn--primary", "立即付款 (模擬)");
  payBtn.onclick = () => {
    // 模擬後端付款驗證
    // 實際開發時，此處應跳轉至 Stripe 付款頁面或呼叫後端 API
    simulatePaymentSuccess(); 
  };
  
  function simulatePaymentSuccess() {
    state.isPaid = true;
    persistState();
    overlay.remove();
    
    // Auto advance after payment if we were on Week 1
    if (state.selectedWeekIndex === 0) {
        pushHistory();
        selectWeek(state.selectedWeekIndex + 1);
    } else {
        // Refresh current view to unlock
        renderWeekDetails();
    }
    showToast("付款成功！已解鎖完整計畫");
  }

  actions.appendChild(cancelBtn);
  actions.appendChild(payBtn);

  modal.appendChild(title);
  modal.appendChild(desc);
  modal.appendChild(actions);
  
  overlay.appendChild(modal);
  
  // Close on overlay click
  overlay.addEventListener("click", (e) => {
    if (e.target === overlay) overlay.remove();
  });

  document.body.appendChild(overlay);
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
      if (!date) return;
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
    const exists = w.races.some((x) => (x?.date || "") === r.date && (x?.name || "").trim() === (r.name || ""));
    if (!exists) w.races.push({ name: r.name || "", date: r.date, distanceKm: r.distanceKm ?? null, kind: r.kind || "" });
  });

  state.weeks.forEach((w) => {
    const hasRaces = Array.isArray(w.races) && w.races.length > 0;
    if (!hasRaces) {
      w.priority = "";
      return;
    }
    w.priority = normalizeRacePriorityValue(w.priority) || "C";
  });

  applyCoachAutoRules();
  refreshAutoVolumeFactors();
}

function openRaceInputModal() {
  const overlay = el("div", "overlay");
  overlay.addEventListener("click", (e) => {
    if (e.target === overlay) overlay.remove();
  });

  const modal = el("div", "modal");
  modal.classList.add("modal--scroll");
  const title = el("div", "modal__title", "輸入比賽");
  const subtitle = el("div", "modal__subtitle", "輸入比賽日期、距離及優先級；系統會按日期更新相應週的比賽／優先級");

  const list = el("div", "raceList");
  const renderList = () => {
    list.replaceChildren();
    const rows = [];
    state.weeks.forEach((w, weekIdx) => {
      const races = Array.isArray(w.races) ? w.races : [];
      races.forEach((r, raceIdx) => {
        const date = (r?.date || "").trim();
        if (!date) return;
        const dist = Number(r?.distanceKm);
        const distText = Number.isFinite(dist) && dist > 0 ? `${dist}km` : "";
        const pr = normalizeRacePriorityValue(w.priority);
        rows.push({ weekIdx, weekNo: w.weekNo, raceIdx, date, distText, pr });
      });
    });

    rows.sort((a, b) => a.date.localeCompare(b.date) || a.weekIdx - b.weekIdx);

    if (!rows.length) {
      list.appendChild(el("div", "muted", "未有比賽"));
      return;
    }

    rows.forEach((r) => {
      const row = el("div", "raceRow");
      const prLabel = racePriorityLabelZh(r.pr) || r.pr;
      const meta = [r.distText, prLabel].filter(Boolean).join(" · ");
      const parts = [weekLabelZh(r.weekNo), r.date];
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
        refreshAutoVolumeFactors();
        persistState();
        renderList();
        updateHeader();
        renderCalendar();
        renderCharts();
        renderWeekDetails();
      });
      row.appendChild(del);
      list.appendChild(row);
    });
  };

  const form = el("form", "formRow");
  const date = document.createElement("input");
  date.className = "input";
  date.type = "date";
  date.required = true;
  {
    const s = state.startDate instanceof Date ? state.startDate : null;
    const minD = s ? addDays(s, 70) : new Date();
    date.min = formatYMD(minD);
  }

  const distanceKm = document.createElement("select");
  distanceKm.className = "input";
  [
    { v: "5", t: "5 公里" },
    { v: "10", t: "10 公里" },
    { v: "15", t: "15 公里" },
    { v: "21.1", t: "半程馬拉松" },
    { v: "42.195", t: "馬拉松" },
  ].forEach(({ v, t }) => {
    const opt = document.createElement("option");
    opt.value = v;
    opt.textContent = t;
    distanceKm.appendChild(opt);
  });

  const priority = document.createElement("select");
  priority.className = "input";
  priority.required = true;
  [
    { v: "", t: "優先級" },
    { v: "A", t: "重要" },
    { v: "C", t: "不重要" },
  ].forEach(({ v, t }) => {
    const opt = document.createElement("option");
    opt.value = v;
    opt.textContent = t;
    priority.appendChild(opt);
  });

  const add = el("button", "btn btn--primary", "加入");
  add.type = "submit";

  form.appendChild(date);
  form.appendChild(distanceKm);
  form.appendChild(priority);
  form.appendChild(add);

  form.addEventListener("submit", (e) => {
    e.preventDefault();
    const raceDate = date.value;
    const raceDistanceKm = Number(String(distanceKm.value || "").trim());
    const racePriority = normalizeRacePriorityValue(priority.value);
    if (!raceDate || !racePriority) return;
    if (!Number.isFinite(raceDistanceKm) || raceDistanceKm <= 0) {
      showToast("請輸入有效距離（公里）", { variant: "warn", durationMs: 1800 });
      return;
    }
    {
      const idx0 = weekIndexForRaceYmd(raceDate);
      if (idx0 !== null && idx0 < 10) {
        showToast("比賽日期需距離開始日期至少 10 週", { variant: "warn", durationMs: 1800 });
        return;
      }
    }
    const aCount = state.weeks.filter((w) => normalizeRacePriorityValue(w?.priority) === "A").length;
    const nextPr = racePriority;
    if (nextPr === "A" && aCount >= 2) {
      showToast("重要最多只可選取兩場", { variant: "warn" });
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
    const existingIndex = w.races.findIndex((r) => (r?.date || "") === raceDate);
    if (existingIndex >= 0) {
      const ex = w.races[existingIndex];
      if (ex && typeof ex === "object") {
        ex.distanceKm = raceDistanceKm;
      }
    } else {
      w.races.push({ name: "", date: raceDate, distanceKm: raceDistanceKm, kind: "" });
    }
    date.value = "";
    distanceKm.selectedIndex = 0;
    priority.value = "";
    w.priority = racePriority;
    // Force auto volume factor when race/priority is updated
    w.volumeFactorAuto = true;
    applyCoachAutoRules();
    refreshAutoVolumeFactors();
    persistState();
    renderList();
    updateHeader();
    renderCalendar();
    renderCharts();
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
  window.setTimeout(() => date.focus(), 0);
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
      if (!key) return; // Allow normal navigation for links without data-tab
      activateTab(key, { persist: true, updateHash: true });
    });
  });

  window.addEventListener("hashchange", () => {
    const key = getTabKeyFromHash();
    if (!key) return;
    activateTab(key, { persist: true, updateHash: false });
  });
}

function openDesignWizard() {
  const overlay = el("div", "overlay");
  const modal = el("div", "modal");
  
  const wizardState = {
    step: 1,
    startDate: "",
    annualVolume: "",
    races: [], // Array of race objects
    currentRace: { date: "", distance: "", priority: "A" },
    vdot: null
  };

  const renderStep = () => {
    modal.innerHTML = "";

    const title = el("div", "modal__title", "開始設計計劃");
    const content = el("div", "modal__content");
    content.style.padding = "20px 0";
    
    const actions = el("div", "modal__actions");
    const nextBtn = el("button", "btn btn--primary", "下一步");
    const prevBtn = el("button", "btn", "上一步");
    const cancelBtn = el("button", "btn btn--reset", "取消");

    cancelBtn.onclick = () => {
      const lbl = document.querySelector('.calRow[data-row-key="monday"] .calLabel .fitText');
      if (lbl) lbl.textContent = "星期一";
      overlay.remove();
    };

    if (wizardState.step === 1) {
       const label = el("div", "fieldLabel", "1. 計劃開始日期");
       label.style.fontWeight = "bold";
       label.style.marginBottom = "8px";
       
       const input = document.createElement("input");
       input.type = "date";
       input.className = "input";
       const todayYmd = formatYMD(new Date());
       input.min = todayYmd;
       if (wizardState.startDate) {
         input.value = wizardState.startDate;
       } else if (state.startDate) {
         const ymd = formatYMD(state.startDate);
         input.value = ymd < todayYmd ? todayYmd : ymd;
       }
       const setWeekdayLabel = (ymd) => {
         const lbl = document.querySelector('.calRow[data-row-key="monday"] .calLabel .fitText');
         const d = parseYMD(ymd);
         if (!lbl || !d) return;
         const names = ["星期日","星期一","星期二","星期三","星期四","星期五","星期六"];
         lbl.textContent = names[d.getDay()] || "星期一";
       };
       const desc = el("div", "", "建議開始日期設於星期一或星期日");
       desc.style.fontSize = "12px";
       desc.style.color = "var(--muted)";
       desc.style.marginBottom = "8px";
       
       content.appendChild(label);
       content.appendChild(desc);
       content.appendChild(input);
       setWeekdayLabel(input.value || todayYmd);
       input.addEventListener("input", () => setWeekdayLabel(input.value || todayYmd));
       input.addEventListener("change", () => setWeekdayLabel(input.value || todayYmd));

      nextBtn.onclick = () => {
        if (!input.value) return showToast("請選擇日期", { variant: "warn" });
        if (String(input.value) < todayYmd) {
          return showToast("請選擇今天或以後的日期", { variant: "warn" });
        }
        
        // 限制只可選擇星期一
        const d = parseYMD(input.value);
        if (d) {
          const day = d.getDay(); // 0=Sun, 1=Mon
          if (day !== 1) {
            return showToast("計劃開始日期必須是星期一", { variant: "warn" });
          }
        }

        wizardState.startDate = input.value;
        wizardState.step++;
        renderStep();
      };
       
       actions.appendChild(cancelBtn);
       actions.appendChild(nextBtn);

    } else if (wizardState.step === 2) {
       const label = el("div", "fieldLabel", "2. 年總訓練量（小時）");
       label.style.fontWeight = "bold";
       label.style.marginBottom = "8px";

       const desc = el("div", "", "例如比上一個訓練年度增加10%");
       desc.style.fontSize = "12px";
       desc.style.color = "var(--text-muted)";
       desc.style.marginBottom = "8px";

       const input = document.createElement("select");
       input.className = "input";
       [
         { label: "150 小時（初階）", value: 150 },
         { label: "200 小時（初中階）", value: 200 },
         { label: "250 小時（中階）", value: 250 },
         { label: "300 小時（中高階）", value: 300 },
         { label: "350 小時（高階）", value: 350 },
       ].forEach((o) => {
         const opt = document.createElement("option");
         opt.value = String(o.value);
         opt.textContent = o.label;
         input.appendChild(opt);
       });
       if (Number(state.ytdVolumeHrs)) {
         input.value = String(state.ytdVolumeHrs);
       }
       if (Number(wizardState.annualVolume)) {
         input.value = String(wizardState.annualVolume);
       }

       content.appendChild(label);
       content.appendChild(desc);
       content.appendChild(input);

       prevBtn.onclick = () => {
         wizardState.step--;
         renderStep();
       };
       nextBtn.onclick = () => {
         if (!input.value) {
           return showToast("請選擇年總訓練量", { variant: "warn" });
         }
         wizardState.annualVolume = Number(input.value);
         wizardState.step++;
         renderStep();
       };

       actions.appendChild(prevBtn);
       actions.appendChild(nextBtn);

   } else if (wizardState.step === 3) {
      const label = el("div", "fieldLabel", "3. 推測VDOT");
      label.style.fontWeight = "bold";
      label.style.marginBottom = "8px";
      content.appendChild(label);
      
      const form = el("div");
      form.style.display = "grid";
      form.style.gap = "8px";
      form.style.gridTemplateColumns = "1fr";
      
      const distWrap = el("label", "paceField");
      distWrap.appendChild(el("span", "muted", "測試距離"));
      const distSelect = document.createElement("select");
      distSelect.className = "input";
      [
        { label: "1 英里", meters: 1609.344 },
        { label: "3 公里", meters: 3000 },
        { label: "5 公里", meters: 5000 },
        { label: "10 公里", meters: 10000 },
        { label: "半馬", meters: 21097.5 },
        { label: "全馬", meters: 42195 },
      ].forEach((o, i) => {
        const opt = document.createElement("option");
        opt.value = String(o.meters);
        opt.textContent = o.label;
        if (i === 0) opt.selected = true;
        distSelect.appendChild(opt);
      });
      distWrap.appendChild(distSelect);
      
      const timeWrap = el("div", "paceField");
      timeWrap.appendChild(el("span", "muted", "時間（時:分:秒）"));
      const timeGrid = el("div", "paceTimeGrid");
      const hEl = document.createElement("input");
      hEl.className = "input paceTimeInput";
      hEl.type = "number";
      hEl.inputMode = "numeric";
      hEl.min = "0";
      hEl.placeholder = "時";
      const mEl = document.createElement("input");
      mEl.className = "input paceTimeInput";
      mEl.type = "number";
      mEl.inputMode = "numeric";
      mEl.min = "0";
      mEl.max = "59";
      mEl.placeholder = "分";
      const sEl = document.createElement("input");
      sEl.className = "input paceTimeInput";
      sEl.type = "number";
      sEl.inputMode = "numeric";
      sEl.min = "0";
      sEl.max = "59";
      sEl.placeholder = "秒";
      timeGrid.appendChild(hEl);
      timeGrid.appendChild(mEl);
      timeGrid.appendChild(sEl);
      timeWrap.appendChild(timeGrid);
      
      form.appendChild(distWrap);
      form.appendChild(timeWrap);
      const meta = el("div", "muted", "");
      form.appendChild(meta);
      content.appendChild(form);
      
      const parseTimeSec = () => {
        const h = Number(hEl.value || 0);
        const m = Number(mEl.value || 0);
        const s = Number(sEl.value || 0);
        if (!Number.isFinite(h) || h < 0) return null;
        if (!Number.isFinite(m) || m < 0 || m > 59) return null;
        if (!Number.isFinite(s) || s < 0 || s > 59) return null;
        return Math.floor(h) * 3600 + Math.floor(m) * 60 + Math.floor(s);
      };
      const recalc = () => {
        const distMeters = Number(distSelect.value);
        const tSec = parseTimeSec();
        if (!Number.isFinite(distMeters) || distMeters <= 0 || !Number.isFinite(tSec) || tSec <= 0) {
          wizardState.vdot = null;
          meta.textContent = "";
          return;
        }
        const v = vdotFromRace(distMeters, tSec);
        if (Number.isFinite(v) && v > 0) {
          wizardState.vdot = v;
          const pace = tSec / (distMeters / 1000);
          meta.textContent = `測試配速：${formatPaceFromSecondsPerKm(pace)} / 公里 · VDOT：${v.toFixed(1)}`;
        } else {
          wizardState.vdot = null;
          meta.textContent = "";
        }
      };
      distSelect.addEventListener("change", recalc);
      hEl.addEventListener("input", recalc);
      mEl.addEventListener("input", recalc);
      sEl.addEventListener("input", recalc);
      
      prevBtn.onclick = () => {
        wizardState.step--;
        renderStep();
      };
      nextBtn.onclick = () => {
        const distMeters = Number(distSelect.value);
        const tSec = parseTimeSec();
        if (!Number.isFinite(distMeters) || distMeters <= 0 || !Number.isFinite(tSec) || tSec <= 0) {
          return showToast("請輸入有效的測試距離及時間", { variant: "warn" });
        }
        const v = vdotFromRace(distMeters, tSec);
        if (!Number.isFinite(v) || v <= 0) {
          return showToast("無法計算（請檢查時間格式）", { variant: "warn" });
        }
        wizardState.vdot = v;
        wizardState.step++;
        renderStep();
      };
      actions.appendChild(prevBtn);
      actions.appendChild(nextBtn);
   } else if (wizardState.step === 4) {
      const label = el("div", "fieldLabel", "4. 輸入比賽");
      label.style.fontWeight = "bold";
      label.style.marginBottom = "8px";
 
      // List of added races
      if (wizardState.races.length > 0) {
        const list = el("div");
        list.style.marginBottom = "16px";
        list.style.border = "1px solid var(--border)";
        list.style.borderRadius = "4px";
        list.style.padding = "8px";
        
        wizardState.races.forEach((r, idx) => {
          const row = el("div");
          row.style.display = "flex";
          row.style.justifyContent = "space-between";
          row.style.alignItems = "center";
          row.style.marginBottom = "4px";
          row.style.fontSize = "14px";
          
          row.textContent = `${r.date} (${Number(r.distance || 0)}km) [${racePriorityLabelZh(r.priority) || r.priority}]`;
          
          const del = el("button", "btn btn--small btn--reset", "✕");
          del.style.color = "var(--warn)";
          del.onclick = () => {
            wizardState.races.splice(idx, 1);
            renderStep();
          };
          row.appendChild(del);
          list.appendChild(row);
        });
        content.appendChild(list);
      }
 
      const formTitle = el("div", "fieldLabel", "４. 輸入比賽");
      formTitle.style.fontSize = "14px";
      formTitle.style.fontWeight = "bold";
      formTitle.style.marginBottom = "4px";
      content.appendChild(formTitle);
 
      const baseStart =
        wizardState.startDate
          ? parseYMD(wizardState.startDate)
          : (state.startDate instanceof Date ? state.startDate : null);
      const minRaceDate = baseStart ? addDays(baseStart, 70) : new Date();
      const minRaceYmd = formatYMD(minRaceDate);
      const form = el("div");
      form.style.display = "flex";
      form.style.flexDirection = "column";
      form.style.gap = "8px";
 
      const dateInput = document.createElement("input");
      dateInput.className = "input";
      dateInput.type = "date";
      dateInput.min = minRaceYmd;
      dateInput.value = wizardState.currentRace.date && String(wizardState.currentRace.date) < minRaceYmd
        ? minRaceYmd
        : wizardState.currentRace.date;
 
      const distInput = document.createElement("select");
      distInput.className = "input";
      [
        { value: 5, label: "5 公里" },
        { value: 10, label: "10 公里" },
        { value: 21.1, label: "半程馬拉松" },
      ].forEach(({ value, label }) => {
        const opt = document.createElement("option");
        opt.value = String(value);
        opt.textContent = label;
        distInput.appendChild(opt);
      });
      if (Number(wizardState.currentRace.distance)) {
        distInput.value = String(Number(wizardState.currentRace.distance));
      }
 
      const prioritySelect = document.createElement("select");
      prioritySelect.className = "input";
      [
        { v: "", t: "優先級" },
        { v: "A", t: "重要" },
        { v: "C", t: "不重要" },
      ].forEach(({ v, t }) => {
        const opt = document.createElement("option");
        opt.value = v;
        opt.textContent = t;
        if (v === wizardState.currentRace.priority) opt.selected = true;
        prioritySelect.appendChild(opt);
      });
 
      const updateState = () => {
        wizardState.currentRace.date = dateInput.value;
        wizardState.currentRace.distance = Number(distInput.value);
        wizardState.currentRace.priority = prioritySelect.value;
      };
      dateInput.oninput = updateState;
      distInput.onchange = updateState;
      prioritySelect.onchange = updateState;
 
      form.appendChild(dateInput);
      form.appendChild(distInput);
      form.appendChild(prioritySelect);
      content.appendChild(form);
 
      const addBtn = el("button", "btn", "加入此比賽");
      addBtn.style.marginTop = "8px";
      addBtn.style.width = "100%";
      addBtn.onclick = () => {
        if (!dateInput.value) return showToast("請選擇比賽日期", { variant: "warn" });
        if (String(dateInput.value) < minRaceYmd) return showToast("比賽日期需距離開始日期至少 10 週", { variant: "warn" });
        if (!prioritySelect.value) return showToast("請選擇優先級", { variant: "warn" });
        const nextPr = normalizeRacePriorityValue(prioritySelect.value);
        const aCount = wizardState.races.filter((r) => normalizeRacePriorityValue(r.priority) === "A").length;
        if (nextPr === "A" && aCount >= 2) return showToast("重要最多只可選取兩場", { variant: "warn" });
 
        wizardState.races.push({ 
           ...wizardState.currentRace, 
           name: "比賽", // Add default name
           priority: nextPr 
        });
        wizardState.currentRace = { date: "", distance: "", priority: "A" };
        renderStep();
      };
      content.appendChild(addBtn);
 
      prevBtn.onclick = () => {
        wizardState.step--;
        renderStep();
      };
      const finishBtn = el("button", "btn btn--primary", "完成及建立");
      finishBtn.onclick = () => {
        try {
          const hasPending = !!dateInput.value;
          if (hasPending) {
            if (!dateInput.value || !prioritySelect.value) {
              return showToast("請先加入比賽或清空輸入欄", { variant: "warn" });
            }
            if (String(dateInput.value) < minRaceYmd) {
              return showToast("比賽日期需距離開始日期至少 10 週", { variant: "warn" });
            }
            const aCount = wizardState.races.filter((r) => normalizeRacePriorityValue(r.priority) === "A").length;
            const nextPr = normalizeRacePriorityValue(prioritySelect.value);
            if (nextPr === "A" && aCount >= 2) return showToast("重要最多只可選取兩場", { variant: "warn" });
            wizardState.races.push({ ...wizardState.currentRace, priority: nextPr });
          }
          const aCountAll = wizardState.races.filter((r) => normalizeRacePriorityValue(r.priority) === "A").length;
          if (aCountAll > 2) return showToast("重要最多只可選取兩場", { variant: "warn" });
          
          applyWizard(wizardState);
          overlay.remove();
        } catch (e) {
          console.error(e);
          showToast(`建立計劃時發生錯誤: ${e.message}`, { variant: "warn", durationMs: 3000 });
        }
      };
 
      actions.appendChild(prevBtn);
      actions.appendChild(finishBtn);
    }

    modal.appendChild(title);
    modal.appendChild(content);
    modal.appendChild(actions);
  };

  overlay.appendChild(modal);
  document.body.appendChild(overlay);
  renderStep();
}

function applyWizard(data) {
  pushHistory();

  if (data.startDate) {
    const d = parseYMD(data.startDate);
    if (d) {
      const monday = startOfMonday(d);
      state.startDate = monday;
      for (let i = 0; i < state.weeks.length; i++) {
        const w = state.weeks[i];
        if (!w) continue;
        w.monday = addDays(monday, i * 7);
      }
    }
  }

  const vol = Number(data.annualVolume);
  if (Number.isFinite(vol) && vol > 0) {
    state.ytdVolumeHrs = vol;
  } else {
    state.ytdVolumeHrs = null;
  }
  syncAnnualVolumeInputs();

  if (Array.isArray(data.races)) {
    data.races.forEach(r => {
      if (!r.date) return;

      const raceDate = r.date;
      const raceName = typeof r.name === "string" ? r.name : "";
      const raceDist = Number(r.distance);
      const raceKind = r.kind || "road";
      const racePriority = normalizeRacePriorityValue(r.priority || "A");
      
      const d = parseYMD(raceDate);
      if (d && state.startDate) {
         // Align to Monday start
         const startOfStart = startOfMonday(state.startDate);
         const diff = d.getTime() - startOfStart.getTime();
         const days = Math.round(diff / MS_PER_DAY);
         const weekIdx = Math.floor(days / 7);
         
         if (weekIdx >= 0 && weekIdx < 52) {
            const w = state.weeks[weekIdx];
            if (w) {
                if (!Array.isArray(w.races)) w.races = [];
                w.races.push({
                   name: raceName,
                   date: raceDate,
                   distanceKm: Number.isFinite(raceDist) ? raceDist : null,
                   kind: raceKind
                });
                w.priority = racePriority;
                w.volumeFactorAuto = true;
            }
         }
      }
    });
  }

  if (data.vdot && Number.isFinite(data.vdot)) {
    state.vdot = data.vdot;
  }
  // 不記錄設計精靈推測的 VDOT 至訓練調控狀態

  reassignAllRacesByDate();
   applyCoachAutoRules();
   refreshAutoVolumeFactors();
   
   // Apply optimization if volume is set
  if (state.ytdVolumeHrs) {
      optimizePlanBaseVolume(state.ytdVolumeHrs);
   } else {
      recomputeFormulaVolumes();
   }

  state.selectedWeekIndex = 0;
  // Auto-generate sessions for all 52 weeks
  for (let i = 0; i < 52; i++) {
    const w = state.weeks[i];
    if (w && Number(w.volumeHrs) > 0) {
      applyAutoSessionsForWeek(i);
    }
  }
   
   state.planStarted = true;
   persistState();
   updateHeader();
   renderCalendar();
  renderCharts();
  renderWeekPicker();
  renderWeekDetails();
  syncPlanStartedUi();
  showToast("已建立計劃！");
}

function syncPlanStartedUi() {
  const startBtn = document.getElementById("startWizardBtn");
  const annualVolumeClearBtn = document.getElementById("annualVolumeClearBtn");
  const autoFillAllBtn = document.getElementById("autoFillAllBtn");
  const exportPdfBtn = document.getElementById("exportPdfBtn");
  const resetBtn = document.getElementById("resetBtn");
  const controls = [
    annualVolumeClearBtn,
    autoFillAllBtn,
    exportPdfBtn,
  ];
  
  if (startBtn) {
    startBtn.style.display = "";
    startBtn.textContent = state.planStarted ? "重設計劃" : "開始跑步計劃";
    if (state.planStarted) {
      startBtn.classList.remove("btn--primary");
      startBtn.classList.add("btn--reset");
    } else {
      startBtn.classList.add("btn--primary");
      startBtn.classList.remove("btn--reset");
    }
  }
  
  if (resetBtn) resetBtn.style.display = "none";

  controls.forEach((el) => {
    if (!el) return;
    el.style.display = state.planStarted ? "" : "none";
  });
}

function wireButtons() {
  const startWizardBtn = document.getElementById("startWizardBtn");
  if (startWizardBtn) {
    startWizardBtn.addEventListener("click", () => {
      if (state.planStarted) {
        const ok = window.confirm("確定要重啟計劃？所有資料將重置。");
        if (!ok) return;
        pushHistory();
        
        state.planStarted = false;
        state.ytdVolumeHrs = null;
        state.vdot = null;
        state.annualVolumeSettings = { startWeeklyHrs: null, maxUpPct: 12, maxDownPct: 25 };
        state.weeks = [];
        state.manualAdvanceWeeks = 0;
        buildInitialWeeks();
        
        state.selectedWeekIndex = 0;
        
        persistState();
        updateHeader();
        renderCalendar();
        renderWeekPicker();
        renderWeekDetails();
        renderCharts();
        syncPlanStartedUi();
        showToast("已重置計劃");
      } else {
        openDesignWizard();
      }
    });
  }
  syncPlanStartedUi();

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
  const annualVolumeClearBtn = document.getElementById("annualVolumeClearBtn");
  const autoFillAllBtn = document.getElementById("autoFillAllBtn");
  if (annualVolumeInput) {
    annualVolumeInput.value = Number.isFinite(state.ytdVolumeHrs) && state.ytdVolumeHrs > 0 ? String(state.ytdVolumeHrs) : "";
    annualVolumeInput.addEventListener("keydown", (e) => {
      if (e.key !== "Enter") return;
      e.preventDefault();
      
      const raw = String(annualVolumeInput.value || "").trim();
      const v = Number(raw);
      if (!Number.isFinite(v) || v <= 0) {
        showToast("請輸入有效的年總訓練量（小時）", { variant: "warn", durationMs: 1800 });
        return;
      }
      pushHistory();
      state.ytdVolumeHrs = v;
      
      // Recalibrate plan based on new target
      optimizePlanBaseVolume(state.ytdVolumeHrs);

      persistState();
      updateHeader();
      renderCalendar();
      renderCharts();
      renderWeekDetails();
      showToast("已更新年總訓練量設定並重新計算課表");
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

  if (autoFillAllBtn) {
    autoFillAllBtn.addEventListener("click", () => {
      const ok = window.confirm("確定要自動填充範例課表？此操作會覆蓋現有每日課表內容。");
      if (!ok) return;
      pushHistory();
      let changed = 0;
      for (let i = 0; i < 52; i++) {
        const w = state.weeks[i];
        if (!w) continue;
        if (!(Number(w.volumeHrs) > 0)) continue;
        const out = applyAutoSessionsForWeek(i);
        if (out.ok) changed++;
      }
      persistState();
      updateHeader();
      renderCalendar();
      renderCharts();
      renderWeekDetails();
      showToast(changed ? `已自動填充 ${changed} 週課表` : "未有可填充的週（請先設定訓練量）", { durationMs: 2200 });
    });
  }



  const exportPdfBtn = document.getElementById("exportPdfBtn");
  if (exportPdfBtn) exportPdfBtn.addEventListener("click", () => exportCalendarPdf());


  const generateBtn = document.getElementById("generateBtn");
  if (generateBtn) {
    generateBtn.addEventListener("click", () => {
      pushHistory();
      generatePlan();
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

async function init() {
  if (typeof state.planStarted !== "boolean") state.planStarted = false;
  state.ytdVolumeHrs = null;
  state.annualVolumeSettings = { startWeeklyHrs: null, maxUpPct: 12, maxDownPct: 25 };
  buildInitialWeeks();
  updateHeader();
  wireCalendarSizer();
  renderCalendar();
  renderWeekPicker();
  renderWeekDetails();
  renderCharts();
  wireTabs();
  wireButtons();
  await wireAuth();
  
  const persisted = loadPersistedState();
  if (persisted) {
    applyPersistedTrainingState(persisted);
    
    // Auto-advance to current unlocked week if plan started
    if (state.planStarted) {
      state.selectedWeekIndex = getUnlockedWeekIndex();
    }

    updateHeader();
    renderCalendar();
    renderWeekPicker();
    renderWeekDetails();
    renderCharts();
    syncPlanStartedUi();
  }

  if (document.getElementById("calendarShell")) {
    let initialTab = getTabKeyFromHash();
    if (!initialTab) {
      try {
        initialTab = String(localStorage.getItem(ACTIVE_TAB_KEY) || "");
      } catch {
        initialTab = "";
      }
    }
    activateTab(initialTab || "plan", { persist: true, updateHash: false });
  }
  persistState();
}

init();

const BLOG_POSTS = [
  {
    id: "p3",
    title: "【2026最完整科學教學】如何從0打造個人跑步計劃?沒有教練也能進步神速!",
    date: "2026-01-08",
    url: "./blog/blog(8-1-2026).html",
    excerpt: "打造屬於自己的全年跑步訓練藍圖，從基礎期到賽季階段循序漸進提升能力。",
    author: "傑哥, 香港中文大學-運動醫學碩士",
    social: "Instagram/Threads: @kitgordont"
  },
  {
    id: "p1",
    title: "每週跑量應該加多少？10% 加量法為何未必適合你（附更安全做法）",
    date: "2026-01-01",
    url: "./blog/how-to-increase-mileage.html",
    excerpt: "解析每週跑量增幅的風險與更穩陣的做法。固定 10% 並非萬靈丹，根據單調度與 ACWR 的組合調整更安全。",
    author: "傑哥, 香港中文大學-運動醫學碩士",
    social: "Instagram/Threads: @kitgordont"
  },
  {
    id: "p2",
    title: "沒有教練也能變強：Borg RPE 體感強度量表實戰教學（附網站自動計算負荷）",
    date: "2026-01-04",
    url: "./blog/borg-rpe-scale-guide.html",
    excerpt: "很多跑者都有一個迷思：要變強是否一定要買昂貴手錶或請教練？本文介紹最有效的強度調控工具：Borg RPE 體感強度量表，教你如何利用體感量化訓練負荷。",
    author: "傑哥, 香港中文大學-運動醫學碩士",
    social: "Instagram/Threads: @kitgordont"
  }
];

function renderBlog(filterText = "") {
  const listRoot = document.getElementById("blogList");
  if (!listRoot) return;

  listRoot.replaceChildren();
  const term = String(filterText || "").trim().toLowerCase();
  const input = document.getElementById("blogSearchInput");
  const userSearch = input && String(input.dataset.user || "") === "1";

  let filtered = BLOG_POSTS.filter((p) => {
    if (!term) return true;
    const t = (p.title || "").toLowerCase();
    const e = (p.excerpt || "").toLowerCase();
    return t.includes(term) || e.includes(term);
  });

  if (filtered.length === 0) {
    if (!userSearch) {
      filtered = BLOG_POSTS.slice();
    } else {
      listRoot.appendChild(el("div", "muted", "沒有找到符合的文章"));
      return;
    }
  }

  filtered.forEach((p) => {
    // Create an anchor tag for SEO-friendly linking
    const row = el("a", "blogItem");
    row.href = p.url;
    // row.target = "_blank"; // Optional: open in new tab? Better for retention to keep app open.
    
    const left = el("div", "blogItem__text", p.title);
    
    // Construct meta string
    const metaParts = [p.date];
    if (p.author) metaParts.push(p.author);
    if (p.social) metaParts.push(p.social);
    const right = el("div", "blogItem__meta", metaParts.join(" · "));
    
    // const desc = el("div", "blogItem__excerpt", p.excerpt);
    
    row.appendChild(left);
    row.appendChild(right);
    // row.appendChild(desc);
    
    listRoot.appendChild(row);
  });
  
  // Clean up title if switching back from article view (legacy cleanup)
  document.title = "部落格";
}

// Legacy functions removed for pure static linking
function updateBlogUrl(postId) {}
function checkBlogUrl() {}

// Ensure blog list renders only after BLOG_POSTS is defined
try {
  if (document.getElementById("blogList")) {
    renderBlog("");
  }
} catch {}
